"""
Google Sheets synchronisation.

Design rule
-----------
The Sheet owns Order Status. The importer owns everything else.

Status is the one column humans edit, and several people may edit it
concurrently. So on every sync we read the Sheet first, keep whatever
status is already recorded against each PO line, and only then write
back. The importer never overwrites a status that a person has set.

Every row carries its own PO Number and Line # (no merged/blank
continuation cells as in the Excel file) because the "PO Number|Line #"
pair is the key used to carry status forward.
"""

from __future__ import annotations

import logging
from typing import Any

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

import tcs_engine as engine

logger = logging.getLogger(__name__)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
]

SHEET_COLUMNS = engine.EXCEL_COLUMNS + [engine.HIDDEN_KEY_COLUMN]

STATUS_COLUMN = "Order Status"


def _client(credentials_path: str | None, credentials_info: dict | None):
    if credentials_path:
        creds = Credentials.from_service_account_file(
            credentials_path, scopes=SCOPES
        )
        return gspread.authorize(creds)

    if credentials_info:
        creds = Credentials.from_service_account_info(
            credentials_info, scopes=SCOPES
        )
        return gspread.authorize(creds)

    # No explicit credentials: use Application Default Credentials,
    # which on Cloud Run resolves to the service's own identity.
    import google.auth

    creds, _ = google.auth.default(scopes=SCOPES)
    return gspread.authorize(creds)


def _open_worksheet(
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
):
    spreadsheet = _client(credentials_path, credentials_info).open_by_key(
        spreadsheet_id
    )

    try:
        return spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        return spreadsheet.add_worksheet(
            title=worksheet_name, rows=1000, cols=len(SHEET_COLUMNS)
        )


META_HEADER = ["Setting", "Value"]


def _open_meta_worksheet(
    spreadsheet_id: str,
    meta_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
    create: bool = False,
):
    spreadsheet = _client(credentials_path, credentials_info).open_by_key(
        spreadsheet_id
    )

    try:
        return spreadsheet.worksheet(meta_name)
    except gspread.WorksheetNotFound:
        if not create:
            return None
        return spreadsheet.add_worksheet(title=meta_name, rows=50, cols=2)


def read_meta(
    spreadsheet_id: str,
    meta_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> dict[str, str]:
    """
    Small key/value store kept in its own worksheet (e.g. when the last
    refresh ran). The Sheet is the source of truth, so this survives
    Cloud Run restarts. Returns {} if the worksheet doesn't exist yet.
    """

    worksheet = _open_meta_worksheet(
        spreadsheet_id, meta_name, credentials_path, credentials_info
    )

    if worksheet is None:
        return {}

    meta: dict[str, str] = {}

    for row in worksheet.get_all_values()[1:]:
        if row and row[0].strip():
            meta[row[0].strip()] = row[1] if len(row) > 1 else ""

    return meta


def write_meta(
    values: dict[str, Any],
    spreadsheet_id: str,
    meta_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> None:
    """Merge `values` into the Meta worksheet (creating it if needed)."""

    worksheet = _open_meta_worksheet(
        spreadsheet_id,
        meta_name,
        credentials_path,
        credentials_info,
        create=True,
    )

    current: dict[str, str] = {}

    for row in worksheet.get_all_values()[1:]:
        if row and row[0].strip():
            current[row[0].strip()] = row[1] if len(row) > 1 else ""

    for key, value in values.items():
        current[key] = "" if value is None else str(value)

    rows = [META_HEADER] + [[k, v] for k, v in current.items()]

    worksheet.clear()
    worksheet.update(rows, value_input_option="RAW")


def row_key(po_number: Any, line_number: Any) -> str:
    po = engine.normalize_whitespace(po_number)
    line = engine.normalize_whitespace(line_number)

    if not po or not line:
        return ""

    return f"{po}|{line}"


def read_sheet(
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> pd.DataFrame:
    """
    Read the Sheet into a DataFrame using the engine's schema.
    """

    worksheet = _open_worksheet(
        spreadsheet_id, worksheet_name, credentials_path, credentials_info
    )

    records = worksheet.get_all_records()

    if not records:
        return engine.create_empty_dataframe()

    df = pd.DataFrame(records)

    return engine.normalize_dataframe_schema(df)


def read_status_map(
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> dict[str, str]:
    """
    Return {"PO Number|Line #": "Order Status"} for everything already
    in the Sheet. This is what protects human edits.
    """

    worksheet = _open_worksheet(
        spreadsheet_id, worksheet_name, credentials_path, credentials_info
    )

    try:
        records = worksheet.get_all_records()
    except Exception as exc:
        logger.warning("Could not read existing sheet rows: %s", exc)
        return {}

    status_map: dict[str, str] = {}

    for record in records:
        key = row_key(record.get("PO Number"), record.get("Line #"))
        status = engine.normalize_whitespace(record.get(STATUS_COLUMN, ""))

        if key and status:
            status_map[key] = status

    return status_map


def apply_status_map(
    df: pd.DataFrame, status_map: dict[str, str]
) -> pd.DataFrame:
    """
    Overlay statuses recorded in the Sheet onto freshly parsed rows.
    The Sheet wins wherever it has a value; newly discovered lines
    keep the importer's default.
    """

    if df.empty or not status_map:
        return df

    df = df.copy()

    df[STATUS_COLUMN] = [
        status_map.get(row_key(po, line))
        or engine.normalize_whitespace(current)
        or engine.DEFAULT_ORDER_STATUS
        for po, line, current in zip(
            df["PO Number"], df["Line #"], df[STATUS_COLUMN]
        )
    ]

    return df


def update_statuses_from_map(
    status_map: dict[str, str],
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> tuple[int, int]:
    """
    Write Order Status for exactly the rows whose "PO Number|Line #"
    key appears in status_map. Every other cell in the sheet is left
    completely untouched -- including rows added by a Gmail refresh
    that ran after the file being uploaded was downloaded, and
    including any other column a person may have edited directly in
    the Sheet in the meantime.

    This is intentionally NOT a full write_sheet() rewrite: reading
    the live sheet and patching only the matched Status cells means a
    stale or partial upload can never delete or corrupt anything --
    at worst it's a no-op for rows it doesn't recognise.

    Returns (matched, changed):
      matched  -- how many status_map keys were found in the sheet
      changed  -- of those, how many actually had a different value
                  written (so uploading the same file twice reports
                  changed=0 the second time)
    """

    if not status_map:
        return 0, 0

    worksheet = _open_worksheet(
        spreadsheet_id, worksheet_name, credentials_path, credentials_info
    )

    all_values = worksheet.get_all_values()

    if not all_values:
        return 0, 0

    header = all_values[0]

    try:
        po_col = header.index("PO Number")
        line_col = header.index("Line #")
        status_col = header.index(STATUS_COLUMN)
    except ValueError as exc:
        raise RuntimeError(
            "The Sheet is missing expected columns. Run a Gmail "
            "refresh at least once before uploading a status file."
        ) from exc

    def cell(row: list[str], index: int) -> str:
        return engine.normalize_whitespace(row[index]) if index < len(row) else ""

    updates: list[gspread.Cell] = []
    matched = 0

    for sheet_row_number, row in enumerate(all_values[1:], start=2):

        key = row_key(cell(row, po_col), cell(row, line_col))

        if not key or key not in status_map:
            continue

        matched += 1

        new_status = engine.normalize_whitespace(status_map[key])
        current_status = cell(row, status_col)

        if new_status and new_status != current_status:
            updates.append(
                gspread.Cell(sheet_row_number, status_col + 1, new_status)
            )

    if updates:
        worksheet.update_cells(updates, value_input_option="USER_ENTERED")

    return matched, len(updates)


def write_sheet(
    df: pd.DataFrame,
    spreadsheet_id: str,
    worksheet_name: str,
    credentials_path: str | None = None,
    credentials_info: dict | None = None,
) -> int:
    """
    Replace the Sheet contents with df. Every row gets its own PO
    Number / Line # so the status join key always resolves.
    """

    worksheet = _open_worksheet(
        spreadsheet_id, worksheet_name, credentials_path, credentials_info
    )

    out = df.copy()

    for column in SHEET_COLUMNS:
        if column not in out.columns:
            out[column] = ""

    out = out[SHEET_COLUMNS]

    out[engine.HIDDEN_KEY_COLUMN] = [
        row_key(po, line) for po, line in zip(out["PO Number"], out["Line #"])
    ]

    values = [SHEET_COLUMNS]

    for record in out.to_dict("records"):
        row = []
        for column in SHEET_COLUMNS:
            value = record[column]
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                row.append(value)
            elif column in engine.MULTILINE_COLUMNS:
                row.append(engine.normalize_multiline(value))
            else:
                row.append(engine.normalize_whitespace(value))
        values.append(row)

    worksheet.clear()
    worksheet.update(values, value_input_option="USER_ENTERED")

    _format_sheet(worksheet)

    return len(values) - 1


def _format_sheet(worksheet) -> None:
    """
    Light formatting only. Sheets cannot reproduce the Excel merges,
    PO outer borders or chart, so it is treated as the shared live
    view for editing status -- the Excel download stays the formatted
    deliverable.
    """

    try:
        worksheet.freeze(rows=1)

        last_col = gspread.utils.rowcol_to_a1(1, len(SHEET_COLUMNS))
        last_col = "".join(ch for ch in last_col if ch.isalpha())

        worksheet.format(
            f"A1:{last_col}1",
            {
                "backgroundColor": {"red": 0.12, "green": 0.22, "blue": 0.39},
                "textFormat": {
                    "bold": True,
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                },
                "verticalAlignment": "MIDDLE",
            },
        )

        # Attachment names sit one per line inside a cell, so the column
        # must wrap for the line breaks to show.
        for column_name in engine.MULTILINE_COLUMNS:
            if column_name in SHEET_COLUMNS:
                col_letter = "".join(
                    ch
                    for ch in gspread.utils.rowcol_to_a1(
                        1, SHEET_COLUMNS.index(column_name) + 1
                    )
                    if ch.isalpha()
                )
                worksheet.format(
                    f"{col_letter}2:{col_letter}",
                    {"wrapStrategy": "WRAP", "verticalAlignment": "TOP"},
                )
    except Exception as exc:
        logger.warning("Sheet formatting skipped: %s", exc)