from __future__ import annotations

import email
import email.utils
import imaplib
import logging
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime
from email import policy
from email.header import decode_header
from email.message import Message
from pathlib import Path
from typing import Any, Callable

# ============================================================
# TCS SAP ARIBA PURCHASE ORDER IMPORTER
# ============================================================
#
# Purpose
# -------
# Imports TCS SAP Ariba Purchase Order emails from Gmail into:
#
#   Documents\TCS Order\tcs_ariba_purchase_orders.xlsx
#
# Design
# ------
# 1. One Excel row = one Ariba PO line item.
# 2. PO-level fields are repeated internally and vertically merged in Excel.
# 3. Line-level fields belong only to their own product.
# 4. Unique key = PO Number + Line #.
# 5. Existing lines are UPDATED, not blindly skipped.
# 6. Existing Order Status is preserved during an update.
# 7. PO numbers are sorted descending; line numbers ascending.
# 8. Product Name and Description are split using the email's own
#    structure (a blank line / double <br><br> in Ariba's description
#    cell separates the product category line from the model and
#    technical description that follows), not guessed keywords.
# 9. Old unwanted columns are removed from the workbook schema.
# 10. A backup is created before replacing an existing workbook.
# 11. The saved workbook uses enterprise-style formatting: a title
#     banner, bold header, borders, zebra striping per PO, and
#     conditional coloring of Order Status.
# 12. process_emails() accepts an optional progress_callback so a
#     caller (the web console) can report percentage-complete and
#     estimate time remaining while a long Gmail scan runs.
# ============================================================

# ============================================================
# AUTO-INSTALL REQUIRED PACKAGES
# ============================================================

REQUIRED_PACKAGES = {
    "pandas": "pandas",
    "bs4": "beautifulsoup4",
    "openpyxl": "openpyxl",
}


def ensure_dependencies_installed() -> None:
    import importlib

    missing: list[str] = []

    for import_name, pip_name in REQUIRED_PACKAGES.items():
        try:
            importlib.import_module(import_name)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print("Installing missing package(s): " + ", ".join(missing))
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", *missing]
        )
        importlib.invalidate_caches()


ensure_dependencies_installed()

import pandas as pd
from bs4 import BeautifulSoup, Tag
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

# ============================================================
# CONFIGURATION
# ============================================================

IMAP_SERVER = "imap.gmail.com"
IMAP_PORT = 993
EMAIL_ACCOUNT = "microlineindia@gmail.com"
PASSWORD = os.getenv("TCS_GMAIL_APP_PASSWORD", "").replace(" ", "").strip()

SEARCH_SENDER = "ordersender-prod@ansmtp.ariba.com"
SEARCH_SUBJECT_FRAGMENT = (
    "Tata Consultancy Services Limited sent a new Purchase Order"
)
MAILBOX = "INBOX"

ORDER_DIR = Path("C:/TCS Orders")
OUTPUT_FILE = ORDER_DIR / "tcs_ariba_purchase_orders.xlsx"
LOG_FILE = ORDER_DIR / "tcs_order_import.log"
BACKUP_DIR = ORDER_DIR / "Backups"
SHEET_NAME = "TCS Orders"

EXCEL_SAVE_RETRIES = 5
EXCEL_SAVE_RETRY_DELAY = 2
DEFAULT_ORDER_STATUS = "Pending"

# ============================================================
# EXCEL SCHEMA
# ============================================================

COMMON_PO_COLUMNS = [
    "PO Number",
    "Order Date",
    "Total PO Amount",
    "PR Number",
    "Requester",
    "Delivery Location",
    "Buyer Contact",
    "Payment Terms",
    "Attachments",
    "Buyer Tax ID",
    "Supplier Tax ID",
    "Order Status",
]

LINE_ITEM_COLUMNS = [
    "Line #",
    "Product Name",
    "Description",
    "Qty",
    "Unit",
    "Unit Price",
    "Subtotal",
    "HSN/SAC Code",
    "Other Charges",
    "Commodity Code",
]

EXCEL_COLUMNS = (
    COMMON_PO_COLUMNS[:2]
    + LINE_ITEM_COLUMNS
    + COMMON_PO_COLUMNS[2:]
)

HIDDEN_KEY_COLUMN = "PO Line Key"
ALL_COLUMNS = EXCEL_COLUMNS + [HIDDEN_KEY_COLUMN]

MONEY_COLUMNS = {"Unit Price", "Subtotal", "Total PO Amount"}
CENTER_COLUMNS = {"Line #", "Qty", "Unit", "HSN/SAC Code"}
WRAP_COLUMNS = {
    "Product Name",
    "Description",
    "Other Charges",
    "Commodity Code",
    "Requester",
    "Delivery Location",
    "Buyer Contact",
    "Payment Terms",
    "Attachments",
}

# Columns whose cells hold several values, one per line (Alt+Enter style).
MULTILINE_COLUMNS = {"Attachments"}
MULTILINE_SEPARATOR = "\n"

# ============================================================
# LOGGING
# ============================================================


def configure_logging() -> None:
    ORDER_DIR.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.INFO)

    if root.handlers:
        return

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )

    file_handler = logging.FileHandler(LOG_FILE, encoding="utf-8")
    file_handler.setFormatter(formatter)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)

    root.addHandler(file_handler)
    root.addHandler(console_handler)


logger = logging.getLogger(__name__)

# ============================================================
# GENERAL HELPERS
# ============================================================


def normalize_whitespace(value: Any) -> str:
    if value is None:
        return ""

    text = str(value).replace("\xa0", " ")
    text = text.replace("\u200b", "")
    text = text.replace("\u200c", "")
    text = text.replace("\u200d", "")

    return re.sub(r"\s+", " ", text).strip()


def normalize_multiline(value: Any) -> str:
    """
    Like normalize_whitespace(), but keeps line breaks: each line is
    cleaned on its own, empty lines are dropped, and the lines are
    re-joined with a single newline. Used for cells that hold a list
    (e.g. attachment names) so the Alt+Enter style breaks survive.
    """

    if value is None:
        return ""

    lines = [
        normalize_whitespace(line)
        for line in re.split(r"[\r\n]+", str(value))
    ]

    return MULTILINE_SEPARATOR.join(line for line in lines if line)


def clean_header_text(value: str | None) -> str:
    if not value:
        return ""

    result: list[str] = []

    try:
        for text, encoding in decode_header(value):
            if isinstance(text, bytes):
                decoded = None

                if encoding:
                    try:
                        decoded = text.decode(encoding, errors="replace")
                    except (LookupError, UnicodeDecodeError):
                        pass

                if decoded is None:
                    for fallback in ("utf-8", "latin-1"):
                        try:
                            decoded = text.decode(fallback, errors="replace")
                            break
                        except UnicodeDecodeError:
                            pass

                result.append(decoded or "")
            else:
                result.append(str(text))

    except Exception:
        return normalize_whitespace(value)

    return normalize_whitespace("".join(result))


def first_non_empty(*values: Any) -> str:
    for value in values:
        cleaned = normalize_whitespace(value)
        if cleaned:
            return cleaned
    return ""


def dedupe_preserve_order(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []

    for value in values:
        cleaned = normalize_whitespace(value)
        if cleaned and cleaned not in seen:
            seen.add(cleaned)
            result.append(cleaned)

    return result


def normalize_label(value: str) -> str:
    value = normalize_whitespace(value)
    value = value.rstrip(":")
    return re.sub(r"[^a-z0-9]+", " ", value.lower()).strip()

# ============================================================
# FORMATTING HELPERS
# ============================================================


def format_indian_currency(amount_str: str, currency: str) -> str:
    cleaned = (amount_str or "").replace(",", "").strip()

    if not cleaned:
        return ""

    try:
        value = float(cleaned)
    except ValueError:
        return normalize_whitespace(f"{amount_str} {currency}")

    sign = "-" if value < 0 else ""
    value = abs(value)

    integer_part = int(value)
    decimal_part = f"{value:.2f}".split(".")[1]
    digits = str(integer_part)

    if len(digits) <= 3:
        grouped = digits
    else:
        last_three = digits[-3:]
        rest = digits[:-3]
        groups: list[str] = []

        while len(rest) > 2:
            groups.insert(0, rest[-2:])
            rest = rest[:-2]

        if rest:
            groups.insert(0, rest)

        grouped = ",".join(groups) + "," + last_three

    return f"{sign}{grouped}.{decimal_part} {currency}".strip()


def format_order_date_from_parts(day: str, month: str, year: str) -> str:
    try:
        dt = datetime.strptime(f"{day} {month} {year}", "%d %b %Y")
        return f"{dt:%b} {dt.day}, {dt.year}"
    except ValueError:
        return normalize_whitespace(f"{day} {month} {year}")


def to_float(value: Any) -> float | str:
    """
    Convert any representation of a money amount -- a raw email
    string like "220,000.00 INR", an older Indian-grouped string
    like "7,42,800.00 INR" from a previous version of this script,
    or an already-numeric value -- into a plain Python float so it
    can be stored as a real Excel number (summable, sortable,
    filterable) rather than text.

    Returns "" if no usable number is found, so empty cells stay
    empty rather than becoming 0.
    """

    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_whitespace(value)

    if not text:
        return ""

    cleaned = re.sub(r"[^0-9.]", "", text)

    if not cleaned or cleaned == ".":
        return ""

    try:
        return float(cleaned)
    except ValueError:
        return ""

def to_int(value: Any) -> int | str:
    """
    Convert a whole-number string (line number, quantity) into a real
    int so Excel treats it as a number -- this removes the green
    "number stored as text" warning triangles and makes numeric
    sorting behave correctly. Non-numeric values pass through
    unchanged.
    """

    if isinstance(value, int):
        return value

    text = normalize_whitespace(value)

    if not text:
        return ""

    try:
        as_float = float(text)
    except ValueError:
        return text

    if as_float.is_integer():
        return int(as_float)

    return text

# ============================================================
# EMAIL BODY
# ============================================================


def decode_part_payload(part: Message) -> str:
    try:
        payload = part.get_payload(decode=True)

        if payload is None:
            payload = part.get_payload()
            return payload if isinstance(payload, str) else ""

        charset = part.get_content_charset()

        if charset:
            try:
                return payload.decode(charset, errors="replace")
            except (LookupError, UnicodeDecodeError):
                pass

        for encoding in ("utf-8", "latin-1"):
            try:
                return payload.decode(encoding, errors="replace")
            except UnicodeDecodeError:
                pass

    except Exception as exc:
        logger.warning("Could not decode MIME part: %s", exc)

    return ""


def extract_html_and_text_body(msg: Message) -> tuple[str, str]:
    html_body = ""
    text_body = ""

    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            disposition = str(part.get("Content-Disposition") or "").lower()

            if "attachment" in disposition:
                continue

            if content_type == "text/html" and not html_body:
                html_body = decode_part_payload(part)

            elif content_type == "text/plain" and not text_body:
                text_body = decode_part_payload(part)

    else:
        payload = decode_part_payload(msg)

        if msg.get_content_type() == "text/html":
            html_body = payload
        else:
            text_body = payload

    return html_body, text_body

# ============================================================
# REGEX HELPERS
# ============================================================


def extract_first(
    text: str,
    patterns: list[str],
    flags: int = re.IGNORECASE,
) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, flags)
        if match:
            return normalize_whitespace(match.group(1))

    return ""

# ============================================================
# ARIBA METADATA
# ============================================================

METADATA_KEY_MAP = {
    "hsn sac code": "hsn_sac_code",
    "req line no": "req_line_no",
    "requester": "requester",
    "pr no": "pr_number",
    "goods or services": "goods_or_services",
    "po line type": "po_line_type",
    "warranty in days": "warranty_in_days",
    "hormonized system of nomenclature": "harmonized_system_of_nomenclature",
    "harmonized system of nomenclature": "harmonized_system_of_nomenclature",
    "other charges": "other_charges",
    "client billing": "client_billing",
    "questionnaires": "questionnaires",
    "commodity code": "commodity_code",
    "manufacturer part id": "manufacturer_part_id",
    "manufacturer name": "manufacturer_name",
}

METADATA_FIELD_RE = re.compile(r"^(.+?)\s*:\s*(.*)$", re.IGNORECASE)


def empty_metadata() -> dict[str, str]:
    return {value: "" for value in METADATA_KEY_MAP.values()}


def parse_metadata_text(text: str) -> dict[str, str]:
    result = empty_metadata()

    raw_lines = [
        normalize_whitespace(x)
        for x in re.split(r"[\r\n]+", text or "")
        if normalize_whitespace(x)
    ]

    for raw_line in raw_lines:
        if ":" not in raw_line:
            continue

        match = METADATA_FIELD_RE.match(raw_line)
        if not match:
            continue

        label = normalize_label(match.group(1))
        value = normalize_whitespace(match.group(2))
        field = METADATA_KEY_MAP.get(label)

        if field and value:
            result[field] = value

    return result

# ============================================================
# HTML HELPERS
# ============================================================


def direct_row_cells(row: Tag) -> list[Tag]:
    """
    Return ONLY the direct cells of this row.

    This is critical for SAP Ariba because "Other Information" contains
    nested tables. Recursive find_all() mixes nested metadata cells into
    the line-item columns.
    """
    return row.find_all(["td", "th"], recursive=False)


def cell_text(cell: Tag) -> str:
    return normalize_whitespace(cell.get_text(" ", strip=True))


def cell_lines(cell: Tag) -> list[str]:
    """
    Extract non-empty visual lines from a cell (used for metadata cells,
    where blank-line structure doesn't matter).
    """
    clone = BeautifulSoup(str(cell), "html.parser").find()

    if clone is None:
        return []

    for br in clone.find_all("br"):
        br.replace_with("\n")

    text = clone.get_text("\n", strip=True)

    result: list[str] = []

    for raw in text.splitlines():
        cleaned = normalize_whitespace(raw)
        if cleaned:
            result.append(cleaned)

    return result


def split_cell_by_blank_line(cell: Tag) -> list[list[str]]:
    """
    Split a cell's visual content into blocks, using a blank line
    (a double <br><br>, i.e. two or more consecutive line breaks) as
    the block separator.

    Ariba's own PO template structurally separates a product's
    category/title line from its model number and technical
    description with exactly this kind of blank line, e.g.:

        Specialized Hardware Lab Equipment Components and
        Accessories - EMG ARMBAND
        <br><br>
        MR-ARB-8-V04 MindRove Armband
        <br>
        Wearable EMG Sensor with 8 Channel

    This lets us split Product Name from Description using the
    email's own structure instead of guessing from keywords.
    """

    clone = BeautifulSoup(str(cell), "html.parser").find()

    if clone is None:
        return []

    for br in clone.find_all("br"):
        br.replace_with("\n")

    raw_text = clone.get_text()
    raw_text = raw_text.replace("\r\n", "\n").replace("\r", "\n")

    blocks_raw = re.split(r"\n\s*\n+", raw_text)

    blocks: list[list[str]] = []

    for block in blocks_raw:
        lines = [normalize_whitespace(x) for x in block.split("\n")]
        lines = [x for x in lines if x]
        if lines:
            blocks.append(lines)

    return blocks


def split_product_name_and_description_from_cell(
    cell: Tag,
) -> tuple[str, str]:
    """
    Split a line item's description cell into (Product Name,
    Description) using the blank-line structure described above.

    If the source has no blank-line separator, there is no distinct
    product-category line, so the whole text is kept as Product Name
    (Description is left blank rather than guessed).
    """

    blocks = split_cell_by_blank_line(cell)

    if not blocks:
        return "", ""

    if len(blocks) == 1:
        text = normalize_whitespace(" ".join(blocks[0]))
        return text, ""

    product_name = normalize_whitespace(" ".join(blocks[0]))
    description = normalize_whitespace(
        " ".join(word for block in blocks[1:] for word in block)
    )

    return product_name, description


def is_amount(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"[\d,]+(?:\.\d+)?\s+[A-Z]{3}",
            normalize_whitespace(value),
            re.IGNORECASE,
        )
    )


def is_quantity(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"\d+(?:\.\d+)?\s*\([A-Za-z]+\)",
            normalize_whitespace(value),
        )
    )


def is_date(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}",
            normalize_whitespace(value),
        )
    )


def is_type(value: str) -> bool:
    return bool(
        re.fullmatch(
            r"(?:Material|Service|Services|Goods)",
            normalize_whitespace(value),
            re.IGNORECASE,
        )
    )

# ============================================================
# LINE ITEM PARSER
# ============================================================

LINE_NUMBER_RE = re.compile(r"^\d+$")
QTY_RE = re.compile(r"^(\d+(?:\.\d+)?)\s*\(([^)]+)\)$")


def is_line_item_row(row: Tag) -> bool:
    cells = direct_row_cells(row)

    if not cells:
        return False

    first = cell_text(cells[0])

    if not LINE_NUMBER_RE.fullmatch(first):
        return False

    texts = [cell_text(c) for c in cells[1:] if cell_text(c)]

    return (
        any(is_quantity(x) for x in texts)
        or any(is_amount(x) for x in texts)
        or len(cells) >= 3
    )


def find_description_cell(
    rows: list[Tag], start_index: int, end_index: int
) -> Tag | None:
    """
    Within the region following a line-item row (up to the next line
    item or the end of the table), find the cell holding the item's
    full product description.

    Ariba puts this on its own row, separate from the numbered line
    row itself -- it is the first row in the region whose direct
    cells contain meaningful text that is not a known "Label: value"
    metadata field and not the "Other Information" section title.
    """

    for row_index in range(start_index, end_index):

        row = rows[row_index]
        cells = direct_row_cells(row)
        texts = [cell_text(cell) for cell in cells]
        nonempty = [text for text in texts if text]

        if not nonempty:
            continue

        first_label = normalize_label(nonempty[0])

        if first_label == "other information":
            continue

        if any(first_label == normalize_label(key) for key in METADATA_KEY_MAP):
            continue

        if len(nonempty) == 1 and LINE_NUMBER_RE.fullmatch(nonempty[0]):
            continue

        if len(nonempty) >= 2:
            continue

        for cell in cells:
            if cell_text(cell):
                return cell

    return None


def parse_line_item_row(row: Tag) -> dict[str, str]:
    """
    Parse the main Ariba line row without relying on fixed HTML indexes.
    """

    cells = direct_row_cells(row)

    line_number = cell_text(cells[0]) if cells else ""

    text_candidates: list[tuple[Tag, str]] = []
    amount_cells: list[str] = []
    quantity_raw = ""
    need_by_raw = ""
    item_type = ""

    for cell in cells[1:]:
        text = cell_text(cell)

        if not text:
            continue

        if is_quantity(text):
            if not quantity_raw:
                quantity_raw = text
            continue

        if is_date(text):
            if not need_by_raw:
                need_by_raw = text
            continue

        if is_amount(text):
            amount_cells.append(text)
            continue

        if is_type(text):
            item_type = text
            continue

        text_candidates.append((cell, text))

    fallback_text = ""

    if text_candidates:
        _, fallback_text = max(
            text_candidates, key=lambda candidate: len(candidate[1])
        )

    unit_price = ""
    subtotal = ""

    if len(amount_cells) >= 2:
        unit_price = amount_cells[-2]
        subtotal = amount_cells[-1]
    elif len(amount_cells) == 1:
        subtotal = amount_cells[0]

    qty = ""
    unit = ""

    qty_match = QTY_RE.match(quantity_raw)

    if qty_match:
        qty = qty_match.group(1)
        unit = qty_match.group(2)

    return {
        "line_number": line_number,
        "product_name": "",
        "description": fallback_text,
        "qty": qty,
        "unit": unit,
        "unit_price": unit_price,
        "subtotal": subtotal,
        "need_by_raw": need_by_raw,
        "type": item_type,
    }

# ============================================================
# METADATA EXTRACTION FROM AN ARIBA LINE BLOCK
# ============================================================


def metadata_from_line_region(
    rows: list[Tag],
    start_index: int,
    end_index: int,
) -> dict[str, str]:
    """
    Extract metadata only from the current line's region.
    """

    result = empty_metadata()

    for row_index in range(start_index, end_index):
        row = rows[row_index]
        cells = direct_row_cells(row)

        if not cells:
            continue

        texts = [cell_text(cell) for cell in cells]
        nonempty = [text for text in texts if text]

        if len(nonempty) >= 2:
            label = normalize_label(nonempty[0])
            field = METADATA_KEY_MAP.get(label)

            if field:
                value = normalize_whitespace(" ".join(nonempty[1:]))
                if value:
                    result[field] = value
                    continue

        for cell in cells:
            for line in cell_lines(cell):
                if ":" not in line:
                    continue

                match = METADATA_FIELD_RE.match(line)
                if not match:
                    continue

                label = normalize_label(match.group(1))
                value = normalize_whitespace(match.group(2))
                field = METADATA_KEY_MAP.get(label)

                if field and value:
                    result[field] = value

    return result

# ============================================================
# HTML LINE ITEM BLOCK PARSER
# ============================================================


def parse_html_line_items(soup: BeautifulSoup) -> list[dict[str, str]]:
    """
    Parse line items and their immediately following Other Information.
    """

    rows = soup.find_all("tr")

    line_row_indices = [
        index for index, row in enumerate(rows) if is_line_item_row(row)
    ]

    logger.info(
        "Detected %d possible Ariba line-item rows.", len(line_row_indices)
    )

    items: list[dict[str, str]] = []

    for position, row_index in enumerate(line_row_indices):
        row = rows[row_index]

        next_line_index = (
            line_row_indices[position + 1]
            if position + 1 < len(line_row_indices)
            else len(rows)
        )

        item = parse_line_item_row(row)

        description_cell = find_description_cell(
            rows, row_index + 1, next_line_index
        )

        if description_cell is not None:
            product_name, description = (
                split_product_name_and_description_from_cell(
                    description_cell
                )
            )
            item["product_name"] = product_name
            item["description"] = description

        metadata = metadata_from_line_region(
            rows, row_index + 1, next_line_index
        )

        item.update(
            {
                "hsn_sac_code": metadata["hsn_sac_code"],
                "other_charges": metadata["other_charges"],
                "commodity_code": metadata["commodity_code"],
                "pr_number": metadata["pr_number"],
                "requester": metadata["requester"],
            }
        )

        useful = any(
            normalize_whitespace(item.get(key, ""))
            for key in (
                "product_name",
                "description",
                "qty",
                "unit_price",
                "subtotal",
            )
        )

        if not useful:
            logger.warning(
                "Ignoring incomplete candidate row %s.",
                item.get("line_number", ""),
            )
            continue

        items.append(item)

    return items

# ============================================================
# PLAIN TEXT FALLBACK
# ============================================================


def parse_text_line_items(text: str) -> list[dict[str, str]]:
    """
    Conservative fallback for plain-text Ariba emails (no HTML body).
    """

    lines = [
        normalize_whitespace(line)
        for line in text.splitlines()
        if normalize_whitespace(line)
    ]

    items: list[dict[str, str]] = []
    current: dict[str, str] | None = None
    metadata_lines: list[str] = []

    def finalize() -> None:
        nonlocal current, metadata_lines

        if current is None:
            return

        metadata = parse_metadata_text("\n".join(metadata_lines))

        current["hsn_sac_code"] = metadata["hsn_sac_code"]
        current["other_charges"] = metadata["other_charges"]
        current["commodity_code"] = metadata["commodity_code"]
        current["pr_number"] = metadata["pr_number"]
        current["requester"] = metadata["requester"]

        items.append(current)
        current = None
        metadata_lines = []

    for line in lines:
        line_match = re.match(r"^(\d+)\s+(.+)$", line)

        if line_match:
            possible_number = line_match.group(1)

            if len(possible_number) <= 4:
                finalize()

                raw_product = line_match.group(2)

                current = {
                    "line_number": possible_number,
                    "product_name": "",
                    "description": normalize_whitespace(raw_product),
                    "qty": "",
                    "unit": "",
                    "unit_price": "",
                    "subtotal": "",
                    "need_by_raw": "",
                    "hsn_sac_code": "",
                    "other_charges": "",
                    "commodity_code": "",
                    "pr_number": "",
                    "requester": "",
                }

                continue

        if current is None:
            continue

        qty_match = re.search(
            r"Qty\s*:\s*(\d+(?:\.\d+)?)\s*(?:\(([^)]+)\)|([A-Za-z]+))",
            line,
            re.IGNORECASE,
        )

        if qty_match:
            current["qty"] = qty_match.group(1)
            current["unit"] = first_non_empty(
                qty_match.group(2), qty_match.group(3)
            )

        unit_match = re.search(
            r"Unit\s*:\s*([\d,]+(?:\.\d+)?)\s+([A-Z]{3})",
            line,
            re.IGNORECASE,
        )

        if unit_match:
            current["unit_price"] = format_indian_currency(
                unit_match.group(1), unit_match.group(2).upper()
            )

        total_match = re.search(
            r"Total\s*:\s*([\d,]+(?:\.\d+)?)\s+([A-Z]{3})",
            line,
            re.IGNORECASE,
        )

        if total_match:
            current["subtotal"] = format_indian_currency(
                total_match.group(1), total_match.group(2).upper()
            )

        if METADATA_FIELD_RE.match(line):
            metadata_lines.append(line)

    finalize()

    return items

# ============================================================
# PO PARSER
# ============================================================


ATTACHMENT_RE = re.compile(
    r"Has\s+Attachment\s+(.+?)\s*\(\s*[A-Za-z0-9.+-]+/[A-Za-z0-9.+-]+\s*\)",
    re.IGNORECASE,
)

MIME_SUFFIX_RE = re.compile(
    r"\s*\(\s*[A-Za-z0-9.+-]+/[A-Za-z0-9.+-]+\s*\)\s*$"
)


def extract_attachment_names(soup: BeautifulSoup | None, full_text: str) -> str:
    """
    Return the attachment file names listed in the email's "Attachments"
    section, one per line, in email order and without duplicates.

    In the Ariba HTML, "Has Attachment" is only the alt text of a small
    paperclip <img>; the file name is the plain text that follows the
    icon, e.g.
        <img alt="Has Attachment" ...> Quote_A.pdf (application/pdf)
    so it never appears in soup.get_text(). We therefore read the text
    after each icon. For plain-text emails (or if the markup changes)
    we fall back to matching "Has Attachment <name> (<mime>)" in the text.
    """

    names: list[str] = []

    if soup is not None:
        for img in soup.find_all(
            "img", alt=re.compile(r"has\s+attachment", re.IGNORECASE)
        ):
            parts: list[str] = []

            for sibling in img.next_siblings:
                if isinstance(sibling, Tag):
                    break
                parts.append(str(sibling))

            name = MIME_SUFFIX_RE.sub("", normalize_whitespace("".join(parts)))

            if name:
                names.append(name)

    if not names:
        names = ATTACHMENT_RE.findall(full_text or "")

    return MULTILINE_SEPARATOR.join(dedupe_preserve_order(names))


def parse_tcs_purchase_order(
    html_content: str,
    text_content: str,
    email_subject: str,
    email_date: str,
    message_id: str,
) -> dict[str, Any] | None:

    soup = BeautifulSoup(html_content or "", "html.parser")

    full_text = (
        normalize_whitespace(soup.get_text(" ", strip=True))
        if html_content
        else normalize_whitespace(text_content)
    )

    po_number = extract_first(
        full_text,
        [
            r"Purchase\s*Order\s*\([^)]*\)\s*(\d+)",
            r"Purchase\s+Order\s*(?:No\.?|Number|#)?\s*[:\-]?\s*(\d+)",
        ],
    )

    if not po_number:
        po_number = extract_first(
            email_subject,
            [r"Purchase\s+Order\s+(\d+)", r"\bPO\s*[-#: ]\s*(\d+)"],
        )

    if not po_number:
        logger.warning(
            "Could not determine PO number from email subject=%r.",
            email_subject,
        )
        return None

    amount_match = re.search(
        r"Amount\s*:\s*([\d,]+(?:\.\d+)?)\s*([A-Z]{3})",
        full_text,
        re.IGNORECASE,
    )

    total_amount = amount_match.group(1) if amount_match else ""

    order_date_match = re.search(
        r"Order\s+submitted\s+on\s*:\s*\w+\s+(\d{1,2})\s+(\w{3})\s+(\d{4})",
        full_text,
        re.IGNORECASE,
    )

    if order_date_match:
        order_date = format_order_date_from_parts(*order_date_match.groups())
    else:
        try:
            parsed = email.utils.parsedate_to_datetime(email_date)
            order_date = f"{parsed:%b} {parsed.day}, {parsed.year}"
        except Exception:
            order_date = normalize_whitespace(email_date)

    delivery_match = re.search(
        r"Ship\s+All\s+Items\s+To\s*(.*?)\s*Ship\s+To\s+Code\s*:",
        full_text,
        re.IGNORECASE,
    )

    delivery_location = (
        normalize_whitespace(delivery_match.group(1)) if delivery_match else ""
    )

    contact_match = re.search(
        r"Name\s*:\s*([^()]+?)\s*\(([^)]+)\)\s*"
        r"Contact\s+Number\s*:\s*(\S+)\s*Email\s*:\s*(\S+)",
        full_text,
        re.IGNORECASE,
    )

    if contact_match:
        buyer_contact = (
            f"{normalize_whitespace(contact_match.group(1))} "
            f"({contact_match.group(3)} / {contact_match.group(4)})"
        )
    else:
        buyer_contact_match = re.search(
            r"(Mr\.|Ms\.|Mrs\.|Dr\.)\s+([A-Za-z .'-]+)\s*\(([^)]+)\)",
            full_text,
            re.IGNORECASE,
        )

        buyer_contact = (
            normalize_whitespace(buyer_contact_match.group(0))
            if buyer_contact_match
            else ""
        )

    payment_terms = extract_first(
        full_text, [r"Payment\s+Terms\s*:\s*(.*?)(?=Buyer\s+Tax\s+ID)"]
    )

    buyer_tax_id = extract_first(full_text, [r"Buyer\s+Tax\s+ID\s*:\s*(\S+)"])

    supplier_tax_id = extract_first(
        full_text, [r"Supplier\s+Tax\s+ID\s*:\s*(\S+)"]
    )

    attachments = extract_attachment_names(soup, full_text)

    if html_content:
        items = parse_html_line_items(soup)
    else:
        items = parse_text_line_items(text_content)

    unique_items: dict[str, dict[str, str]] = {}

    for item in items:
        line_no = normalize_whitespace(item.get("line_number", ""))

        if not line_no:
            continue

        if line_no not in unique_items:
            unique_items[line_no] = item
            continue

        old_score = sum(
            bool(normalize_whitespace(value))
            for value in unique_items[line_no].values()
        )
        new_score = sum(
            bool(normalize_whitespace(value)) for value in item.values()
        )

        if new_score > old_score:
            unique_items[line_no] = item

    items = list(unique_items.values())

    pr_numbers = dedupe_preserve_order(
        [item.get("pr_number", "") for item in items]
    )
    requesters = dedupe_preserve_order(
        [item.get("requester", "") for item in items]
    )

    pr_number = "; ".join(pr_numbers)
    requester = "; ".join(requesters)

    if not pr_number:
        pr_number = "; ".join(
            dedupe_preserve_order(
                re.findall(
                    r"PR\s*No\.?\s*:\s*([A-Za-z0-9\-_/]+)",
                    full_text,
                    re.IGNORECASE,
                )
            )
        )

    if not requester:
        requester = "; ".join(
            dedupe_preserve_order(
                re.findall(
                    r"Requester\s*:\s*(.*?)\s*PR\s*No\.",
                    full_text,
                    re.IGNORECASE,
                )
            )
        )

    return {
        "po_number": po_number,
        "order_date": order_date,
        "items": items,
        "total_amount": total_amount,
        "pr_number": pr_number,
        "requester": requester,
        "delivery_location": delivery_location,
        "buyer_contact": buyer_contact,
        "payment_terms": payment_terms,
        "attachments": attachments,
        "buyer_tax_id": buyer_tax_id,
        "supplier_tax_id": supplier_tax_id,
        "order_status": DEFAULT_ORDER_STATUS,
        "message_id": message_id,
    }

# ============================================================
# GMAIL
# ============================================================


def connect_to_mailbox() -> imaplib.IMAP4_SSL:
    logger.info("Connecting to %s:%s", IMAP_SERVER, IMAP_PORT)

    mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
    mail.login(EMAIL_ACCOUNT, PASSWORD)

    logger.info("Gmail authentication successful.")

    status, _ = mail.select(MAILBOX, readonly=True)

    if status != "OK":
        raise RuntimeError(f"Unable to select mailbox: {MAILBOX}")

    return mail


def find_ariba_emails(mail: imaplib.IMAP4_SSL) -> list[bytes]:
    criterion = f'(FROM "{SEARCH_SENDER}" SUBJECT "{SEARCH_SUBJECT_FRAGMENT}")'

    status, messages = mail.search(None, criterion)

    if status != "OK":
        raise RuntimeError("IMAP search failed.")

    if not messages or not messages[0]:
        return []

    return messages[0].split()


def fetch_message(mail: imaplib.IMAP4_SSL, email_id: bytes) -> Message:
    status, msg_data = mail.fetch(email_id, "(RFC822)")

    if status != "OK":
        raise RuntimeError(f"Unable to fetch email ID {email_id!r}")

    for response_part in msg_data:
        if isinstance(response_part, tuple):
            return email.message_from_bytes(
                response_part[1], policy=policy.default
            )

    raise RuntimeError(f"No message content for email ID {email_id!r}")

# ============================================================
# EXCEL DATA
# ============================================================


def create_empty_dataframe() -> pd.DataFrame:
    return pd.DataFrame(columns=ALL_COLUMNS)


def normalize_dataframe_schema(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize both the new workbook and older malformed workbooks.
    Unwanted old columns are intentionally discarded.
    """

    df = df.copy()

    df = df.loc[
        :,
        [
            column
            for column in df.columns
            if not str(column).startswith("Unnamed:")
        ],
    ]

    for column in ALL_COLUMNS:
        if column not in df.columns:
            df[column] = ""

    df = df[ALL_COLUMNS].fillna("")

    line_content_columns = [
        "Line #",
        "Product Name",
        "Description",
        "Qty",
        "Unit Price",
        "Subtotal",
    ]

    has_content = (
        df[line_content_columns]
        .apply(lambda col: col.map(lambda v: bool(normalize_whitespace(v))))
        .any(axis=1)
    )

    df = df[has_content].reset_index(drop=True)

    for column in COMMON_PO_COLUMNS:
        cleaner = (
            normalize_multiline
            if column in MULTILINE_COLUMNS
            else normalize_whitespace
        )
        values = df[column].astype(str).map(cleaner)
        values = values.replace("", pd.NA)

        if column in MULTILINE_COLUMNS:
            # Optional per-PO field: a PO with no attachments must stay
            # blank rather than inherit the previous PO's names, so only
            # fill within the same PO.
            values = values.groupby(df["PO Number"]).ffill()
        else:
            values = values.ffill()

        df[column] = values.fillna("").astype(str)

    df[HIDDEN_KEY_COLUMN] = df.apply(
        lambda row: (
            f"{normalize_whitespace(row['PO Number'])}|"
            f"{normalize_whitespace(row['Line #'])}"
            if normalize_whitespace(row["PO Number"])
            and normalize_whitespace(row["Line #"])
            else ""
        ),
        axis=1,
    )

    for column in MONEY_COLUMNS:
        df[column] = df[column].map(to_float)

    for column in ("Line #", "Qty"):
        df[column] = df[column].map(to_int)

    return df


def load_existing_excel() -> pd.DataFrame:
    if not OUTPUT_FILE.exists():
        logger.info(
            "Workbook does not exist yet. A new workbook will be created."
        )
        return create_empty_dataframe()

    logger.info("Loading existing workbook: %s", OUTPUT_FILE)

    try:
        df = pd.read_excel(
            OUTPUT_FILE,
            sheet_name=SHEET_NAME,
            dtype=str,
            header=HEADER_ROW - 1,
        )
    except ValueError:
        logger.warning(
            "Sheet '%s' not found. Starting with an empty workbook.",
            SHEET_NAME,
        )
        return create_empty_dataframe()
    except Exception as exc:
        raise RuntimeError(f"Could not read Excel file: {exc}") from exc

    return normalize_dataframe_schema(df)


def read_uploaded_excel_statuses(file_obj: Any) -> dict[str, str]:
    """
    Extract {"PO Number|Line #": "Order Status"} from an uploaded
    workbook -- normally a copy of the one downloaded from this same
    console, with Order Status hand-edited.

    Editing status in Excel is much faster than in the Sheet, because
    Order Status is one of the columns merged across every line of a
    PO: change it once at the top of the PO's block and every line
    inherits it visually. That merge is purely a display trick,
    though -- the cells below the first are genuinely empty in the
    file. This function relies on normalize_dataframe_schema()'s
    existing forward-fill step (originally written to cope with that
    same merge when reloading the local Windows workbook) to recover
    the intended status for every line before building the map.

    file_obj can be a file path or a file-like object (e.g. the
    bytes from an HTTP upload wrapped in io.BytesIO).
    """

    df = pd.read_excel(
        file_obj,
        sheet_name=SHEET_NAME,
        dtype=str,
        header=HEADER_ROW - 1,
    )

    df = normalize_dataframe_schema(df)

    status_map: dict[str, str] = {}

    for _, row in df.iterrows():
        po = normalize_whitespace(row["PO Number"])
        line = normalize_whitespace(row["Line #"])
        status = normalize_whitespace(row["Order Status"])

        if po and line and status:
            status_map[f"{po}|{line}"] = status

    return status_map

# ============================================================
# ROW / UPSERT
# ============================================================


def item_to_excel_row(
    po_data: dict[str, Any],
    item: dict[str, str],
    existing_status: str = "",
) -> dict[str, str]:
    status = first_non_empty(
        existing_status, po_data.get("order_status"), DEFAULT_ORDER_STATUS
    )

    po_number = normalize_whitespace(po_data.get("po_number", ""))
    line_number = normalize_whitespace(item.get("line_number", ""))

    return {
        "PO Number": po_number,
        "Order Date": po_data.get("order_date", ""),
        "Total PO Amount": to_float(po_data.get("total_amount", "")),
        "PR Number": po_data.get("pr_number", ""),
        "Requester": po_data.get("requester", ""),
        "Delivery Location": po_data.get("delivery_location", ""),
        "Buyer Contact": po_data.get("buyer_contact", ""),
        "Payment Terms": po_data.get("payment_terms", ""),
        "Attachments": po_data.get("attachments", ""),
        "Buyer Tax ID": po_data.get("buyer_tax_id", ""),
        "Supplier Tax ID": po_data.get("supplier_tax_id", ""),
        "Order Status": status,
        "Line #": to_int(line_number),
        "Product Name": item.get("product_name", ""),
        "Description": item.get("description", ""),
        "Qty": to_int(item.get("qty", "")),
        "Unit": item.get("unit", ""),
        "Unit Price": to_float(item.get("unit_price", "")),
        "Subtotal": to_float(item.get("subtotal", "")),
        "HSN/SAC Code": item.get("hsn_sac_code", ""),
        "Other Charges": item.get("other_charges", ""),
        "Commodity Code": item.get("commodity_code", ""),
        HIDDEN_KEY_COLUMN: f"{po_number}|{line_number}",
    }


def dataframe_line_key(row: pd.Series) -> str:
    key = normalize_whitespace(row.get(HIDDEN_KEY_COLUMN, ""))

    if key:
        return key

    po = normalize_whitespace(row.get("PO Number", ""))
    line = normalize_whitespace(row.get("Line #", ""))

    if po and line:
        return f"{po}|{line}"

    return ""


def build_existing_index(df: pd.DataFrame) -> dict[str, int]:
    index: dict[str, int] = {}

    for row_index, row in df.iterrows():
        key = dataframe_line_key(row)
        if key:
            index[key] = int(row_index)

    return index


def upsert_parsed_po(
    df: pd.DataFrame, po_data: dict[str, Any]
) -> tuple[pd.DataFrame, int, int]:
    df = normalize_dataframe_schema(df)
    existing_index = build_existing_index(df)

    added = 0
    updated = 0

    for item in po_data.get("items", []):
        line_number = normalize_whitespace(item.get("line_number", ""))

        if not line_number:
            continue

        key = f'{po_data["po_number"]}|{line_number}'

        if key in existing_index:
            row_index = existing_index[key]

            old_status = normalize_whitespace(
                df.at[row_index, "Order Status"]
            )

            row = item_to_excel_row(po_data, item, existing_status=old_status)

            for column in ALL_COLUMNS:
                df.at[row_index, column] = row.get(column, "")

            updated += 1

        else:
            row = item_to_excel_row(po_data, item)

            df = pd.concat(
                [df, pd.DataFrame([row], columns=ALL_COLUMNS)],
                ignore_index=True,
            )

            existing_index[key] = len(df) - 1
            added += 1

    return df, added, updated

# ============================================================
# SORTING
# ============================================================


def sort_orders_descending(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sort: PO Number descending, Line # ascending.
    """

    df = normalize_dataframe_schema(df)

    if df.empty:
        return df

    working = df.copy()

    working["__po_sort"] = working["PO Number"].map(
        lambda value: (
            int(normalize_whitespace(value))
            if normalize_whitespace(value).isdigit()
            else -1
        )
    )

    working["__line_sort"] = working["Line #"].map(
        lambda value: (
            int(normalize_whitespace(value))
            if normalize_whitespace(value).isdigit()
            else 999999
        )
    )

    working = working.sort_values(
        by=["__po_sort", "__line_sort"],
        ascending=[False, True],
        kind="stable",
    )

    return (
        working.drop(columns=["__po_sort", "__line_sort"])
        .reset_index(drop=True)
    )

# ============================================================
# BACKUP
# ============================================================


def create_backup() -> None:
    if not OUTPUT_FILE.exists():
        return

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = BACKUP_DIR / f"tcs_ariba_purchase_orders_{timestamp}.xlsx"

    shutil.copy2(OUTPUT_FILE, backup_file)

    logger.info("Backup created: %s", backup_file)

# ============================================================
# EXCEL FORMATTING (ENTERPRISE STYLE)
# ============================================================

TITLE_ROW = 1
HEADER_ROW = 2
DATA_START_ROW = 3

NAVY_FILL = PatternFill(fill_type="solid", fgColor="1F3864")
LIGHT_BAND_FILL = PatternFill(fill_type="solid", fgColor="F2F5FA")
WHITE_FILL = PatternFill(fill_type="solid", fgColor="FFFFFF")

STATUS_FILLS = {
    "pending": PatternFill(fill_type="solid", fgColor="FDE9C8"),
    "completed": PatternFill(fill_type="solid", fgColor="D9EAD3"),
    "delivered": PatternFill(fill_type="solid", fgColor="D9EAD3"),
    "closed": PatternFill(fill_type="solid", fgColor="D9EAD3"),
    "cancelled": PatternFill(fill_type="solid", fgColor="F4CCCC"),
    "canceled": PatternFill(fill_type="solid", fgColor="F4CCCC"),
    "rejected": PatternFill(fill_type="solid", fgColor="F4CCCC"),
    "on hold": PatternFill(fill_type="solid", fgColor="FFF2CC"),
}

THIN_SIDE = Side(style="thin", color="BFBFBF")
THIN_BORDER = Border(
    left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE
)

PO_BORDER_SIDE = Side(style="medium", color="1F3864")


def status_fill(status: str) -> PatternFill | None:
    return STATUS_FILLS.get(normalize_whitespace(status).lower())


def write_title_banner(worksheet: Worksheet, total_pos: int) -> None:
    last_column_letter = get_column_letter(len(EXCEL_COLUMNS))

    worksheet.merge_cells(
        f"A{TITLE_ROW}:{last_column_letter}{TITLE_ROW}"
    )

    timestamp = datetime.now().strftime("%d-%b-%Y %I:%M %p")

    title_cell = worksheet.cell(TITLE_ROW, 1)
    title_cell.value = (
        f"TCS SAP Ariba Purchase Orders   |   "
        f"Total POs: {total_pos}   |   Last Updated: {timestamp}"
    )
    title_cell.font = Font(bold=True, size=13, color="FFFFFF")
    title_cell.fill = NAVY_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    worksheet.row_dimensions[TITLE_ROW].height = 26


def apply_excel_formatting(worksheet: Worksheet, total_pos: int) -> None:

    write_title_banner(worksheet, total_pos)

    header_map = {
        normalize_whitespace(cell.value): cell.column
        for cell in worksheet[HEADER_ROW]
    }

    for cell in worksheet[HEADER_ROW]:
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = NAVY_FILL
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = THIN_BORDER

    worksheet.row_dimensions[HEADER_ROW].height = 32

    widths = {
        "PO Number": 16,
        "Order Date": 16,
        "Line #": 8,
        "Product Name": 46,
        "Description": 46,
        "Qty": 9,
        "Unit": 9,
        "Unit Price": 18,
        "Subtotal": 18,
        "HSN/SAC Code": 16,
        "Other Charges": 20,
        "Commodity Code": 26,
        "Total PO Amount": 20,
        "PR Number": 14,
        "Requester": 30,
        "Delivery Location": 46,
        "Buyer Contact": 36,
        "Payment Terms": 40,
        "Attachments": 48,
        "Buyer Tax ID": 18,
        "Supplier Tax ID": 18,
        "Order Status": 16,
        HIDDEN_KEY_COLUMN: 26,
    }

    for name, width in widths.items():
        if name in header_map:
            worksheet.column_dimensions[
                get_column_letter(header_map[name])
            ].width = width

    if HIDDEN_KEY_COLUMN in header_map:
        worksheet.column_dimensions[
            get_column_letter(header_map[HIDDEN_KEY_COLUMN])
        ].hidden = True

    po_col = header_map.get("PO Number")
    status_col = header_map.get("Order Status")

    band_on = False
    previous_po = None

    for row_number in range(DATA_START_ROW, worksheet.max_row + 1):

        current_po = normalize_whitespace(
            worksheet.cell(row_number, po_col).value if po_col else ""
        )

        if current_po and current_po != previous_po:
            band_on = not band_on

        if current_po:
            previous_po = current_po

        row_fill = LIGHT_BAND_FILL if band_on else WHITE_FILL

        row_height = 48

        attachments_col = header_map.get("Attachments")
        if attachments_col:
            attachment_value = worksheet.cell(row_number, attachments_col).value
            if attachment_value:
                est_lines = sum(
                    max(1, -(-len(line) // 44))
                    for line in str(attachment_value).split("\n")
                )
                row_height = max(row_height, 15 * est_lines + 4)

        worksheet.row_dimensions[row_number].height = row_height

        for column_name, column_number in header_map.items():

            cell = worksheet.cell(row_number, column_number)
            cell.border = THIN_BORDER

            if column_name in WRAP_COLUMNS:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
            elif column_name in CENTER_COLUMNS:
                cell.alignment = Alignment(
                    horizontal="center", vertical="top", wrap_text=True
                )
            elif column_name in MONEY_COLUMNS:
                cell.alignment = Alignment(horizontal="right", vertical="top")
                if isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##,##0.00" INR"'
            else:
                cell.alignment = Alignment(vertical="top")

            if (
                status_col
                and column_number == status_col
                and status_fill(cell.value)
            ):
                cell.fill = status_fill(cell.value)
            else:
                cell.fill = row_fill

    worksheet.freeze_panes = f"A{DATA_START_ROW}"

    if worksheet.max_row >= HEADER_ROW:
        last_column_letter = get_column_letter(len(EXCEL_COLUMNS))
        worksheet.auto_filter.ref = (
            f"A{HEADER_ROW}:{last_column_letter}{worksheet.max_row}"
        )

    worksheet.sheet_properties.tabColor = "1F3864"

# ============================================================
# MERGE COMMON PO FIELDS
# ============================================================


def merge_common_po_fields(
    worksheet: Worksheet, groups: list[tuple[int, int]]
) -> None:
    """
    Vertically merge common PO-level cells. Line-level fields are
    never merged.
    """

    header_map = {
        normalize_whitespace(cell.value): cell.column
        for cell in worksheet[HEADER_ROW]
    }

    if "PO Number" not in header_map:
        return

    for row, end_row in groups:

        if end_row > row:
            for column_name in COMMON_PO_COLUMNS:
                column_number = header_map.get(column_name)

                if not column_number:
                    continue

                worksheet.merge_cells(
                    start_row=row,
                    start_column=column_number,
                    end_row=end_row,
                    end_column=column_number,
                )

                top_cell = worksheet.cell(row, column_number)
                top_cell.alignment = Alignment(
                    vertical="center", horizontal="left", wrap_text=True
                )

                if column_name == "Order Status":
                    top_cell.alignment = Alignment(
                        vertical="center", horizontal="center", wrap_text=True
                    )
                elif column_name in MONEY_COLUMNS:
                    top_cell.alignment = Alignment(
                        vertical="center", horizontal="right"
                    )


def compute_po_groups(worksheet: Worksheet) -> list[tuple[int, int]]:
    """
    Return [(start_row, end_row), ...] for each PO group.

    Must be called BEFORE merge_common_po_fields(), while the PO
    Number column still has a value on every row.
    """

    header_map = {
        normalize_whitespace(cell.value): cell.column
        for cell in worksheet[HEADER_ROW]
    }

    po_col = header_map.get("PO Number")

    if not po_col:
        return []

    groups: list[tuple[int, int]] = []
    row = DATA_START_ROW

    while row <= worksheet.max_row:
        po_number = normalize_whitespace(worksheet.cell(row, po_col).value)

        if not po_number:
            row += 1
            continue

        end_row = row

        while end_row + 1 <= worksheet.max_row:
            next_po = normalize_whitespace(
                worksheet.cell(end_row + 1, po_col).value
            )
            if next_po != po_number:
                break
            end_row += 1

        groups.append((row, end_row))
        row = end_row + 1

    return groups

# ============================================================
# PO OUTER BORDERS
# ============================================================


def apply_po_outer_borders(
    worksheet: Worksheet, groups: list[tuple[int, int]]
) -> None:
    """
    Draw a medium-weight navy border around each PO group, spanning
    all visible columns. Must run AFTER merge_common_po_fields().
    """

    header_map = {
        normalize_whitespace(cell.value): cell.column
        for cell in worksheet[HEADER_ROW]
    }

    if "PO Number" not in header_map:
        return

    first_col = 1
    last_col = len(EXCEL_COLUMNS)

    if last_col < first_col:
        return

    merged_ranges = list(worksheet.merged_cells.ranges)

    def find_merge(cell) -> Any:
        for mr in merged_ranges:
            if (
                mr.min_row <= cell.row <= mr.max_row
                and mr.min_col <= cell.column <= mr.max_col
            ):
                return mr
        return None

    def style_owner(row_number: int, column_number: int):
        cell = worksheet.cell(row_number, column_number)
        mr = find_merge(cell)
        if mr is not None:
            return worksheet.cell(mr.min_row, mr.min_col)
        return cell

    def add_sides(cell, **sides) -> None:
        border = cell.border
        cell.border = Border(
            left=sides.get("left") or border.left,
            right=sides.get("right") or border.right,
            top=sides.get("top") or border.top,
            bottom=sides.get("bottom") or border.bottom,
        )

    for start_row, end_row in groups:

        for column_number in range(first_col, last_col + 1):
            add_sides(
                style_owner(start_row, column_number), top=PO_BORDER_SIDE
            )

        for column_number in range(first_col, last_col + 1):
            add_sides(
                style_owner(end_row, column_number), bottom=PO_BORDER_SIDE
            )

        for row_number in range(start_row, end_row + 1):
            add_sides(
                style_owner(row_number, first_col), left=PO_BORDER_SIDE
            )
            add_sides(
                style_owner(row_number, last_col), right=PO_BORDER_SIDE
            )

# ============================================================
# SUMMARY SHEET
# ============================================================


def compute_summary_stats(df: pd.DataFrame) -> dict[str, Any]:
    """
    Compute PO-level statistics from the finalized, sorted DataFrame.
    """

    if df.empty:
        return {
            "total_pos": 0,
            "total_lines": 0,
            "total_value": 0.0,
            "status_counts": {},
        }

    po_level = df[df["PO Number"].map(normalize_whitespace) != ""]
    po_level = po_level.drop_duplicates(subset=["PO Number"], keep="first")

    total_pos = po_level["PO Number"].nunique()
    total_lines = len(df)

    total_value = sum(
        value
        for value in po_level["Total PO Amount"].tolist()
        if isinstance(value, (int, float))
    )

    status_counts: dict[str, int] = {}

    for status in po_level["Order Status"].map(normalize_whitespace):
        key = status or "Unspecified"
        status_counts[key] = status_counts.get(key, 0) + 1

    return {
        "total_pos": int(total_pos),
        "total_lines": int(total_lines),
        "total_value": float(total_value),
        "status_counts": status_counts,
    }


def write_summary_sheet(workbook, stats: dict[str, Any]) -> None:
    """
    Write (or replace) a "Summary" sheet with headline totals, a
    status breakdown table, and a bar chart.
    """

    if "Summary" in workbook.sheetnames:
        del workbook["Summary"]

    ws = workbook.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "1F3864"

    ws.merge_cells("A1:D1")
    title_cell = ws["A1"]
    title_cell.value = "TCS SAP Ariba Purchase Orders \u2014 Summary"
    title_cell.font = Font(bold=True, size=14, color="FFFFFF")
    title_cell.fill = NAVY_FILL
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    labels_and_values = [
        ("Total Purchase Orders", stats["total_pos"], None),
        ("Total Line Items", stats["total_lines"], None),
        ("Total Order Value", stats["total_value"], '#,##,##0.00" INR"'),
        (
            "Last Updated",
            datetime.now().strftime("%d-%b-%Y %I:%M %p"),
            None,
        ),
    ]

    row = 3

    for label, value, number_format in labels_and_values:
        label_cell = ws.cell(row, 1, label)
        label_cell.font = Font(bold=True)

        value_cell = ws.cell(row, 2, value)

        if number_format:
            value_cell.number_format = number_format

        row += 1

    row += 1

    breakdown_title_row = row
    ws.cell(breakdown_title_row, 1, "Order Status Breakdown").font = Font(
        bold=True, size=12
    )

    header_row = breakdown_title_row + 1

    ws.cell(header_row, 1, "Status")
    ws.cell(header_row, 2, "Count")

    for col in (1, 2):
        header_cell = ws.cell(header_row, col)
        header_cell.font = Font(bold=True, color="FFFFFF")
        header_cell.fill = NAVY_FILL
        header_cell.alignment = Alignment(horizontal="center")

    data_row = header_row + 1

    for status, count in sorted(stats["status_counts"].items()):
        ws.cell(data_row, 1, status)
        ws.cell(data_row, 2, count)
        data_row += 1

    last_data_row = data_row - 1

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 16

    if last_data_row >= header_row + 1:
        chart = BarChart()
        chart.type = "col"
        chart.title = "Purchase Orders by Status"
        chart.y_axis.title = "Count"
        chart.x_axis.title = "Status"
        chart.style = 10

        data_ref = Reference(
            ws, min_col=2, min_row=header_row, max_row=last_data_row
        )
        category_ref = Reference(
            ws, min_col=1, min_row=header_row + 1, max_row=last_data_row
        )

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(category_ref)
        chart.width = 16
        chart.height = 9

        ws.add_chart(chart, f"D{breakdown_title_row}")


def apply_print_setup(worksheet: Worksheet) -> None:
    """
    Sensible defaults for printing/exporting to PDF.
    """

    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.print_title_rows = f"{TITLE_ROW}:{HEADER_ROW}"


def atomic_save_excel(df: pd.DataFrame) -> None:
    ORDER_DIR.mkdir(parents=True, exist_ok=True)

    df = sort_orders_descending(df)

    total_pos = len(
        {
            normalize_whitespace(x)
            for x in df["PO Number"].tolist()
            if normalize_whitespace(x)
        }
    )

    temp_file: Path | None = None

    try:
        with tempfile.NamedTemporaryFile(
            suffix=".xlsx", prefix="tcs_orders_", dir=ORDER_DIR, delete=False
        ) as temp:
            temp_file = Path(temp.name)

        with pd.ExcelWriter(temp_file, engine="openpyxl") as writer:
            df.to_excel(
                writer,
                sheet_name=SHEET_NAME,
                index=False,
                startrow=HEADER_ROW - 1,
            )

            worksheet = writer.book[SHEET_NAME]

            apply_excel_formatting(worksheet, total_pos)

            po_groups = compute_po_groups(worksheet)

            merge_common_po_fields(worksheet, po_groups)
            apply_po_outer_borders(worksheet, po_groups)
            apply_print_setup(worksheet)

            stats = compute_summary_stats(df)
            write_summary_sheet(writer.book, stats)

            writer.book.active = 0

        for attempt in range(1, EXCEL_SAVE_RETRIES + 1):
            try:
                os.replace(temp_file, OUTPUT_FILE)
                temp_file = None

                logger.info("Excel saved successfully: %s", OUTPUT_FILE)
                return

            except PermissionError as exc:
                if attempt == EXCEL_SAVE_RETRIES:
                    raise RuntimeError(
                        "Excel file appears to be open or locked. "
                        "Close the workbook and run the importer again."
                    ) from exc

                logger.warning(
                    "Excel file is locked. Retrying %d/%d...",
                    attempt,
                    EXCEL_SAVE_RETRIES,
                )
                time.sleep(EXCEL_SAVE_RETRY_DELAY)

    finally:
        if temp_file and temp_file.exists():
            try:
                temp_file.unlink()
            except OSError:
                pass

# ============================================================
# EMAIL PROCESSING
# ============================================================


def process_emails(
    mail: imaplib.IMAP4_SSL,
    email_ids: list[bytes],
    existing_df: pd.DataFrame,
    progress_callback: Callable[[int, int], None] | None = None,
) -> tuple[pd.DataFrame, int, int, int]:
    """
    progress_callback(processed_count, total_count) is called once
    after each email is handled (whether it succeeded, was skipped,
    or failed), so a caller can report percentage-complete and
    estimate time remaining during a long scan.
    """

    df = normalize_dataframe_schema(existing_df)

    processed_messages = 0
    added_lines = 0
    updated_lines = 0
    total = len(email_ids)

    for index, email_id in enumerate(email_ids, start=1):
        try:
            logger.info("Processing email %d/%d...", index, total)

            msg = fetch_message(mail, email_id)

            message_id = normalize_whitespace(msg.get("Message-ID", ""))

            if not message_id:
                message_id = "IMAP-ID-" + email_id.decode(errors="ignore")

            subject = clean_header_text(msg.get("Subject"))
            email_date = normalize_whitespace(msg.get("Date", ""))

            html_body, text_body = extract_html_and_text_body(msg)

            po_data = parse_tcs_purchase_order(
                html_content=html_body,
                text_content=text_body,
                email_subject=subject,
                email_date=email_date,
                message_id=message_id,
            )

            if po_data is None:
                logger.warning("No PO number found. Email skipped.")
                processed_messages += 1
                continue

            if not po_data["items"]:
                logger.warning(
                    "PO %s found, but no line items were parsed.",
                    po_data["po_number"],
                )
                processed_messages += 1
                continue

            df, added, updated = upsert_parsed_po(df, po_data)

            added_lines += added
            updated_lines += updated

            logger.info(
                "PO %s: %d lines parsed; %d added; %d updated.",
                po_data["po_number"],
                len(po_data["items"]),
                added,
                updated,
            )

            processed_messages += 1

        except Exception as exc:
            logger.exception(
                "Failed to process email ID %r: %s", email_id, exc
            )

        finally:
            if progress_callback is not None:
                try:
                    progress_callback(index, total)
                except Exception:
                    logger.exception("progress_callback raised")

    return df, processed_messages, added_lines, updated_lines

# ============================================================
# MAIN
# ============================================================


def main() -> int:
    configure_logging()

    logger.info("=" * 78)
    logger.info("TCS SAP ARIBA PURCHASE ORDER IMPORT STARTED")
    logger.info("=" * 78)

    if not PASSWORD:
        logger.error(
            "TCS_GMAIL_APP_PASSWORD environment variable is not set."
        )
        return 1

    ORDER_DIR.mkdir(parents=True, exist_ok=True)
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)

    try:
        existing_df = load_existing_excel()
    except Exception as exc:
        logger.exception("Unable to load existing Excel: %s", exc)
        return 1

    original_df = sort_orders_descending(existing_df)
    original_rows = len(original_df)

    mail = None

    try:
        mail = connect_to_mailbox()
        email_ids = find_ariba_emails(mail)

        if not email_ids:
            logger.info("No matching TCS Purchase Order emails found.")
            return 0

        final_df, processed, added, updated = process_emails(
            mail, email_ids, original_df
        )

    except imaplib.IMAP4.error as exc:
        logger.error("IMAP/Gmail error: %s", exc)
        return 1

    except Exception as exc:
        logger.exception("Unexpected import error: %s", exc)
        return 1

    finally:
        if mail is not None:
            try:
                mail.close()
            except Exception:
                pass
            try:
                mail.logout()
            except Exception:
                pass

    final_df = sort_orders_descending(final_df)

    original_columns = set(existing_df.columns)
    schema_requires_rewrite = original_columns != set(ALL_COLUMNS)

    if added == 0 and updated == 0 and not schema_requires_rewrite:
        logger.info("No new or changed PO line items found.")
        logger.info("Existing rows: %d", original_rows)
        logger.info("TCS SAP ARIBA PURCHASE ORDER IMPORT COMPLETED")
        return 0

    try:
        create_backup()
    except Exception as exc:
        logger.exception("Unable to create Excel backup: %s", exc)
        return 1

    try:
        atomic_save_excel(final_df)
    except Exception as exc:
        logger.exception("Excel save failed: %s", exc)
        return 1

    final_po_count = len(
        {
            normalize_whitespace(x)
            for x in final_df["PO Number"].tolist()
            if normalize_whitespace(x)
        }
    )

    logger.info("Emails processed: %d", processed)
    logger.info("Line items added: %d", added)
    logger.info("Line items updated: %d", updated)
    logger.info("Purchase orders on file: %d", final_po_count)
    logger.info("Output: %s", OUTPUT_FILE)
    logger.info("TCS SAP ARIBA PURCHASE ORDER IMPORT COMPLETED")

    return 0


if __name__ == "__main__":
    try:
        exit_code = main()
    except KeyboardInterrupt:
        logger.warning("Import interrupted by user.")
        exit_code = 130
    except Exception as exc:
        logger.exception("Fatal application error: %s", exc)
        exit_code = 1

    sys.exit(exit_code)