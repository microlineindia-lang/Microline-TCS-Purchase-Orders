"""
TCS Purchase Order console -- web service.

Endpoints
---------
GET  /                    the console UI
POST /api/refresh         start a full Gmail import (returns a job id)
GET  /api/jobs/{job_id}   poll job progress (percentage, ETA, stage, log)
GET  /api/summary         current totals, read live from the Sheet
                          (also used by the "Sync from Sheet" button --
                          it's a plain read, no Gmail involved, so it
                          returns in well under a second)
GET  /api/download        formatted .xlsx, generated on demand
POST /api/upload-status   apply Order Status edits from a downloaded,
                          hand-edited copy of that workbook back onto
                          the Sheet -- see _upload_status() below for
                          why this is safe even with a stale file
GET  /healthz             liveness probe
POST /api/refresh/scheduled
                          daily 8 AM refresh, called by a free external
                          cron (GitHub Actions); needs the X-Cron-Token
                          header; see AUTO_REFRESH_* settings below

Access
------
Every route except /healthz, /login, /auth/* and the token-protected
scheduled endpoint requires a Google sign-in from an allow-listed
account (default microlineindia@gmail.com). See auth.py.

Two refresh actions, two speeds
--------------------------------
"Refresh from Gmail" (POST /api/refresh) scans every matching email --
this is the slow one (minutes for a large mailbox) and reports
progress via /api/jobs/{id}.

"Sync from Sheet" (GET /api/summary, called directly by the button)
just re-reads the Google Sheet -- useful when a colleague has edited
Order Status and you want the header figures to reflect that without
waiting for a Gmail scan.

Why the Sheet is the source of truth
------------------------------------
Hosted containers are ephemeral: any file written to local disk
disappears when the instance scales down. Rather than bolt on a
bucket, the Google Sheet holds the data. Refresh writes to it and
Download reads from it, so the downloaded workbook always reflects
the statuses people have edited, and the service itself stays
stateless.
"""

from __future__ import annotations

# Load a local .env file (if present) BEFORE importing the modules below:
# tcs_engine reads TCS_GMAIL_APP_PASSWORD at import time. On a hosting
# platform you set real environment variables instead and there is no
# .env file, in which case this does nothing.
try:
    from dotenv import load_dotenv

    load_dotenv()
except ImportError:
    pass

import base64
import json
import logging
import os
import threading
import time
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import hmac
import io
from datetime import timedelta
from zoneinfo import ZoneInfo

from fastapi import FastAPI, File, Header, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

import auth
import sheets
import tcs_engine as engine

# ============================================================
# CONFIGURATION
# ============================================================

SPREADSHEET_ID = os.getenv("TCS_SPREADSHEET_ID", "").strip()
WORKSHEET_NAME = os.getenv("TCS_WORKSHEET_NAME", "Orders").strip() or "Orders"

# Free hosts have ephemeral disks; /tmp is always writable.
WORK_DIR = Path(os.getenv("TCS_WORK_DIR", "/tmp/tcs"))

engine.ORDER_DIR = WORK_DIR
engine.OUTPUT_FILE = WORK_DIR / "tcs_ariba_purchase_orders.xlsx"
engine.LOG_FILE = WORK_DIR / "tcs_order_import.log"
engine.BACKUP_DIR = WORK_DIR / "Backups"

WORK_DIR.mkdir(parents=True, exist_ok=True)

# ---- Last-refresh record and automatic daily refresh --------------
#
# When a refresh last ran is stored in a small "Meta" worksheet in the
# same Google Sheet (the Sheet is the source of truth; the host's local
# disk is not).
#
# The daily refresh is normally triggered by a free external cron (GitHub
# Actions) calling POST /api/refresh/scheduled with the header
# X-Cron-Token, which also wakes the app if the host put it to sleep.
# Set TCS_CRON_TOKEN to enable it.
#
# For a machine that stays on (e.g. local Windows), set
# TCS_AUTO_REFRESH_LOCAL=1 instead to run a small built-in scheduler.
# Use one or the other, not both.

META_WORKSHEET_NAME = os.getenv("TCS_META_WORKSHEET", "Meta").strip() or "Meta"
TIMEZONE_NAME = os.getenv("TCS_TIMEZONE", "Asia/Kolkata").strip() or "Asia/Kolkata"
AUTO_REFRESH_TIME = os.getenv("TCS_AUTO_REFRESH_TIME", "08:00").strip() or "08:00"
CRON_TOKEN = os.getenv("TCS_CRON_TOKEN", "").strip()
AUTO_REFRESH_LOCAL = os.getenv("TCS_AUTO_REFRESH_LOCAL", "").strip() in {
    "1",
    "true",
    "yes",
}

# True (default): POST /api/refresh/scheduled starts the import in the
# background and returns 202 straight away, so a slow Gmail scan can't hit
# a proxy timeout. Set to 0 to run it inside the request instead (what
# Cloud Run needed, because it freezes the CPU between requests).
SCHEDULED_BACKGROUND = os.getenv(
    "TCS_SCHEDULED_BACKGROUND", "1"
).strip().lower() in {"1", "true", "yes"}

# A scheduled call is skipped if a refresh already succeeded this recently
# (protects against scheduler retries and double triggers).
SCHEDULED_SKIP_MINUTES = 30

LOCAL_TZ = ZoneInfo(TIMEZONE_NAME)
AUTO_HOUR, AUTO_MINUTE = (int(p) for p in AUTO_REFRESH_TIME.split(":", 1))

STATIC_DIR = Path(__file__).parent / "static"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)
logger = logging.getLogger("tcs.console")

app = FastAPI(title="TCS Purchase Orders", docs_url=None, redoc_url=None)

# Login gate: only the addresses in TCS_ALLOWED_EMAILS (default
# microlineindia@gmail.com) can use the app. See auth.py.
auth.install(app)

# ============================================================
# JOB TRACKING
# ============================================================
#
# In-process job state. The service must therefore run as a single
# instance with a single worker for polling to stay consistent -- see
# DEPLOY.md. If you later need
# several instances, move this into Firestore or Redis; the shape
# below is deliberately easy to swap since every mutation goes
# through _new_job / _log / _progress / _set_stage / _finish.
#
# Stages, in order, so the UI can show a meaningful label beyond a
# bare percentage:
#   queued -> connecting -> reading_sheet -> scanning -> writing -> done

_jobs: dict[str, dict[str, Any]] = {}
_jobs_lock = threading.Lock()
_run_lock = threading.Lock()

STAGE_LABELS = {
    "queued": "Queued",
    "connecting": "Connecting to Gmail",
    "reading_sheet": "Reading current Sheet",
    "scanning": "Scanning purchase order emails",
    "writing": "Writing results to Sheet",
    "done": "Done",
    "error": "Failed",
}


def _new_job() -> str:
    job_id = uuid.uuid4().hex[:12]

    with _jobs_lock:
        _jobs[job_id] = {
            "id": job_id,
            "state": "queued",
            "stage": "queued",
            "stage_label": STAGE_LABELS["queued"],
            "started_at": datetime.now(timezone.utc).isoformat(),
            "finished_at": None,
            "messages": [],
            "error": None,
            "result": None,
            "total": 0,
            "processed": 0,
            "progress_pct": 0.0,
            "eta_seconds": None,
        }

    return job_id


def _log(job_id: str, message: str) -> None:
    logger.info("[%s] %s", job_id, message)

    with _jobs_lock:
        job = _jobs.get(job_id)
        if job is not None:
            job["messages"].append(
                {
                    "at": datetime.now(timezone.utc).strftime("%H:%M:%S"),
                    "text": message,
                }
            )
            job["messages"] = job["messages"][-200:]


def _set_stage(job_id: str, stage: str) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if job is not None:
            job["stage"] = stage
            job["stage_label"] = STAGE_LABELS.get(stage, stage)


def _progress(
    job_id: str, processed: int, total: int, elapsed_seconds: float
) -> None:
    """
    Update percentage-complete and a simple estimate of time
    remaining, based on the average time per email seen so far in
    THIS run. Recomputed on every call, so the estimate improves as
    the scan progresses.
    """

    pct = round((processed / total) * 100, 1) if total else 0.0

    eta_seconds = None
    if processed > 0 and total > processed:
        rate = elapsed_seconds / processed
        eta_seconds = round(rate * (total - processed))

    with _jobs_lock:
        job = _jobs.get(job_id)
        if job is not None:
            job["total"] = total
            job["processed"] = processed
            job["progress_pct"] = pct
            job["eta_seconds"] = eta_seconds


def _finish(job_id: str, **fields: Any) -> None:
    with _jobs_lock:
        job = _jobs.get(job_id)
        if job is not None:
            job.update(fields)
            job["finished_at"] = datetime.now(timezone.utc).isoformat()


# ============================================================
# CREDENTIALS
# ============================================================


_SERVICE_ACCOUNT_INFO: dict[str, Any] | None = None
_SERVICE_ACCOUNT_LOADED = False


def _service_account_info() -> dict[str, Any] | None:
    """
    GOOGLE_SERVICE_ACCOUNT_JSON holds the service-account key itself
    (the whole JSON, or the same JSON base64-encoded). This is the
    easiest option on free hosts, where you can set secrets but not
    upload files.
    """

    global _SERVICE_ACCOUNT_INFO, _SERVICE_ACCOUNT_LOADED

    if _SERVICE_ACCOUNT_LOADED:
        return _SERVICE_ACCOUNT_INFO

    _SERVICE_ACCOUNT_LOADED = True
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()

    if not raw:
        return None

    try:
        if not raw.startswith("{"):
            raw = base64.b64decode(raw).decode("utf-8")
        _SERVICE_ACCOUNT_INFO = json.loads(raw)
    except Exception as exc:
        raise RuntimeError(
            "GOOGLE_SERVICE_ACCOUNT_JSON is set but is not valid JSON "
            "(or base64-encoded JSON)."
        ) from exc

    return _SERVICE_ACCOUNT_INFO


def _sheet_kwargs() -> dict[str, Any]:
    """
    Where the Google credentials come from, in priority order:

    1. GOOGLE_SERVICE_ACCOUNT_JSON   the key contents (hosting platforms)
    2. GOOGLE_APPLICATION_CREDENTIALS path to a key file (local machine)
    3. Application Default Credentials (only works on Google Cloud)
    """

    info = _service_account_info()
    if info:
        return {"credentials_info": info}

    path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS", "").strip()

    if path and Path(path).exists():
        return {"credentials_path": path}

    return {}


def _require_config() -> None:
    if not SPREADSHEET_ID:
        raise RuntimeError(
            "TCS_SPREADSHEET_ID is not set. Add the Google Sheet id as "
            "an environment variable."
        )

    if not engine.PASSWORD:
        raise RuntimeError(
            "TCS_GMAIL_APP_PASSWORD is not set. Add it as a secret / env var."
        )


# ============================================================
# IMPORT JOB
# ============================================================


def _record_refresh(
    kwargs: dict[str, Any],
    trigger: str,
    error: str | None = None,
    result: dict[str, Any] | None = None,
) -> None:
    """
    Remember in the Meta worksheet that a refresh was attempted, and
    (if it worked) when it last succeeded. Best effort: a failure to
    write this must never fail the refresh itself.
    """

    now = datetime.now(timezone.utc).isoformat()

    values: dict[str, Any] = {
        "last_attempt_at": now,
        "last_attempt_trigger": trigger,
        "last_attempt_status": "error" if error else "ok",
        "last_attempt_error": error or "",
    }

    if not error:
        result = result or {}
        values.update(
            {
                "last_success_at": now,
                "last_success_trigger": trigger,
                "last_success_added": result.get("added", 0),
                "last_success_updated": result.get("updated", 0),
                "last_success_rows": result.get("rows", 0),
            }
        )

    try:
        sheets.write_meta(
            values, SPREADSHEET_ID, META_WORKSHEET_NAME, **kwargs
        )
    except Exception:
        logger.exception("Could not record refresh time in Meta sheet")


def _run_import(job_id: str, trigger: str = "manual") -> None:
    """
    Fetch Gmail, parse purchase orders, preserve statuses already set
    in the Sheet, then write the Sheet back. Reports stage and
    percentage progress throughout.

    `trigger` is "manual" (button) or "scheduled" (8 AM job); it is
    stored with the last-refresh time so the UI can say which.
    """

    if not _run_lock.acquire(blocking=False):
        _finish(
            job_id,
            state="error",
            stage="error",
            stage_label=STAGE_LABELS["error"],
            error="Another refresh is already running. Wait for it to finish.",
        )
        return

    mail = None
    kwargs: dict[str, Any] = {}

    try:
        with _jobs_lock:
            _jobs[job_id]["state"] = "running"

        _require_config()
        kwargs = _sheet_kwargs()

        _set_stage(job_id, "reading_sheet")
        _log(job_id, "Reading statuses already recorded in the Sheet")
        status_map = sheets.read_status_map(
            SPREADSHEET_ID, WORKSHEET_NAME, **kwargs
        )
        _log(job_id, f"Found {len(status_map)} existing status entries")

        existing_df = sheets.read_sheet(
            SPREADSHEET_ID, WORKSHEET_NAME, **kwargs
        )
        _log(job_id, f"Sheet currently holds {len(existing_df)} line items")

        _set_stage(job_id, "connecting")
        _log(job_id, "Connecting to Gmail")
        mail = engine.connect_to_mailbox()

        email_ids = engine.find_ariba_emails(mail)
        total = len(email_ids)
        _log(job_id, f"Found {total} purchase order emails")

        if not email_ids:
            _finish(
                job_id,
                state="done",
                stage="done",
                stage_label=STAGE_LABELS["done"],
                progress_pct=100.0,
                result={"added": 0, "updated": 0, "rows": len(existing_df)},
            )
            _log(job_id, "Nothing to import")
            _record_refresh(
                kwargs,
                trigger,
                result={"added": 0, "updated": 0, "rows": len(existing_df)},
            )
            return

        _set_stage(job_id, "scanning")
        with _jobs_lock:
            _jobs[job_id]["total"] = total

        parse_start = time.monotonic()

        def _on_progress(processed: int, total_count: int) -> None:
            elapsed = time.monotonic() - parse_start
            _progress(job_id, processed, total_count, elapsed)

        df, processed, added, updated = engine.process_emails(
            mail, email_ids, existing_df, progress_callback=_on_progress
        )

        _log(
            job_id,
            f"Parsed {processed} emails: {added} new lines, "
            f"{updated} updated",
        )

        df = engine.sort_orders_descending(df)

        df = sheets.apply_status_map(df, status_map)
        _log(job_id, "Preserved human-edited statuses")

        _set_stage(job_id, "writing")
        written = sheets.write_sheet(
            df, SPREADSHEET_ID, WORKSHEET_NAME, **kwargs
        )
        _log(job_id, f"Wrote {written} line items to the Sheet")

        _finish(
            job_id,
            state="done",
            stage="done",
            stage_label=STAGE_LABELS["done"],
            progress_pct=100.0,
            eta_seconds=0,
            result={"added": added, "updated": updated, "rows": written},
        )
        _record_refresh(
            kwargs,
            trigger,
            result={"added": added, "updated": updated, "rows": written},
        )

    except Exception as exc:
        logger.exception("Import failed")
        _log(job_id, f"Failed: {exc}")
        _record_refresh(kwargs, trigger, error=str(exc))
        _finish(
            job_id,
            state="error",
            stage="error",
            stage_label=STAGE_LABELS["error"],
            error=str(exc),
        )

    finally:
        if mail is not None:
            for closer in ("close", "logout"):
                try:
                    getattr(mail, closer)()
                except Exception:
                    pass

        _run_lock.release()


# ============================================================
# ROUTES
# ============================================================


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    return HTMLResponse((STATIC_DIR / "index.html").read_text("utf-8"))


@app.post("/api/refresh")
def refresh() -> JSONResponse:
    job_id = _new_job()

    thread = threading.Thread(
        target=_run_import, args=(job_id,), daemon=True
    )
    thread.start()

    return JSONResponse({"job_id": job_id})


def _parse_iso(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(value)
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def _next_auto_refresh_iso() -> str:
    now = datetime.now(LOCAL_TZ)
    candidate = now.replace(
        hour=AUTO_HOUR, minute=AUTO_MINUTE, second=0, microsecond=0
    )
    if candidate <= now:
        candidate += timedelta(days=1)
    return candidate.astimezone(timezone.utc).isoformat()


@app.post("/api/refresh/scheduled")
def refresh_scheduled(
    x_cron_token: str | None = Header(default=None),
) -> JSONResponse:
    """
    Called once a day by an external scheduler (the GitHub Actions
    workflow in .github/workflows/daily-refresh.yml). Not behind the
    Google login -- protected by the X-Cron-Token header instead.

    By default the import runs in a background thread and this returns
    202 immediately (TCS_SCHEDULED_BACKGROUND=0 runs it inline).
    """

    if not CRON_TOKEN:
        raise HTTPException(
            status_code=503, detail="Scheduled refresh is not enabled."
        )

    if not x_cron_token or not hmac.compare_digest(x_cron_token, CRON_TOKEN):
        raise HTTPException(status_code=403, detail="Forbidden")

    try:
        _require_config()
        meta = sheets.read_meta(
            SPREADSHEET_ID, META_WORKSHEET_NAME, **_sheet_kwargs()
        )
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc

    last_ok = _parse_iso(meta.get("last_success_at"))
    if last_ok and datetime.now(timezone.utc) - last_ok < timedelta(
        minutes=SCHEDULED_SKIP_MINUTES
    ):
        return JSONResponse({"skipped": "refreshed recently"})

    if _run_lock.locked():
        return JSONResponse({"skipped": "a refresh is already running"})

    job_id = _new_job()

    if SCHEDULED_BACKGROUND:
        threading.Thread(
            target=_run_import, args=(job_id, "scheduled"), daemon=True
        ).start()
        # Result (or failure) is recorded in the Meta sheet and shown on
        # the dashboard once the import finishes.
        return JSONResponse(
            {"state": "started", "job_id": job_id}, status_code=202
        )

    _run_import(job_id, trigger="scheduled")

    with _jobs_lock:
        job = dict(_jobs[job_id])

    if job["state"] == "error":
        return JSONResponse(
            {"state": "error", "error": job["error"]}, status_code=500
        )

    return JSONResponse({"state": "done", "result": job["result"]})


@app.get("/api/refresh/scheduled/status")
def refresh_scheduled_status(
    job_id: str,
    x_cron_token: str | None = Header(default=None),
) -> JSONResponse:
    """
    Lets the scheduler poll a scheduled import until it really finishes
    (and learn whether it succeeded). 404 means the app restarted and
    the job was lost. Same X-Cron-Token protection as the trigger.
    """

    if not CRON_TOKEN:
        raise HTTPException(status_code=503, detail="Not enabled.")

    if not x_cron_token or not hmac.compare_digest(x_cron_token, CRON_TOKEN):
        raise HTTPException(status_code=403, detail="Forbidden")

    with _jobs_lock:
        job = _jobs.get(job_id)

        if job is None:
            raise HTTPException(status_code=404, detail="Unknown job")

        return JSONResponse(
            {
                "state": job["state"],
                "stage": job["stage_label"],
                "progress_pct": job["progress_pct"],
                "error": job["error"],
                "result": job["result"],
            }
        )


def _scheduler_loop() -> None:
    """
    Optional in-process daily scheduler (TCS_AUTO_REFRESH_LOCAL=1).
    Runs once per day after the configured time, and also catches up
    if the app was off at that time and starts later the same day.
    """

    handled_for = None

    while True:
        time.sleep(30)

        try:
            now = datetime.now(LOCAL_TZ)
            due = now.replace(
                hour=AUTO_HOUR, minute=AUTO_MINUTE, second=0, microsecond=0
            )

            if now < due or handled_for == due.date():
                continue

            meta = sheets.read_meta(
                SPREADSHEET_ID, META_WORKSHEET_NAME, **_sheet_kwargs()
            )
            last_attempt = _parse_iso(meta.get("last_attempt_at"))

            handled_for = due.date()

            if last_attempt and last_attempt >= due.astimezone(timezone.utc):
                continue

            if _run_lock.locked():
                continue

            logger.info("Starting scheduled daily refresh")
            _run_import(_new_job(), trigger="scheduled")
        except Exception:
            logger.exception("Scheduler loop error")


@app.on_event("startup")
def _start_scheduler() -> None:
    if AUTO_REFRESH_LOCAL and SPREADSHEET_ID:
        threading.Thread(target=_scheduler_loop, daemon=True).start()
        logger.info(
            "Built-in daily refresh enabled at %s %s",
            AUTO_REFRESH_TIME,
            TIMEZONE_NAME,
        )


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str) -> JSONResponse:
    with _jobs_lock:
        job = _jobs.get(job_id)

        if job is None:
            raise HTTPException(status_code=404, detail="Unknown job")

        return JSONResponse(dict(job))


@app.get("/api/summary")
def summary() -> JSONResponse:
    """
    Totals for the header, read live from the Sheet. This is a plain
    read (no Gmail call), so it's what the "Sync from Sheet" button
    calls for a fast refresh of the numbers on screen.
    """

    try:
        _require_config()
        df = sheets.read_sheet(
            SPREADSHEET_ID, WORKSHEET_NAME, **_sheet_kwargs()
        )
    except Exception as exc:
        return JSONResponse(
            {"configured": False, "message": str(exc)}, status_code=200
        )

    stats = engine.compute_summary_stats(df)

    sheet_url = (
        f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit"
    )

    try:
        meta = sheets.read_meta(
            SPREADSHEET_ID, META_WORKSHEET_NAME, **_sheet_kwargs()
        )
    except Exception:
        logger.exception("Could not read Meta sheet")
        meta = {}

    last_refresh = None
    if meta.get("last_success_at"):
        last_refresh = {
            "at": meta["last_success_at"],
            "trigger": meta.get("last_success_trigger", ""),
            "added": meta.get("last_success_added", ""),
            "updated": meta.get("last_success_updated", ""),
        }

    last_failure = None
    if meta.get("last_attempt_status") == "error":
        last_failure = {
            "at": meta.get("last_attempt_at", ""),
            "trigger": meta.get("last_attempt_trigger", ""),
            "error": meta.get("last_attempt_error", ""),
        }

    auto_enabled = bool(CRON_TOKEN) or AUTO_REFRESH_LOCAL

    return JSONResponse(
        {
            "configured": True,
            "total_pos": stats["total_pos"],
            "total_lines": stats["total_lines"],
            "total_value": stats["total_value"],
            "status_counts": stats["status_counts"],
            "sheet_url": sheet_url,
            "last_refresh": last_refresh,
            "last_failure": last_failure,
            "auto_refresh": {
                "enabled": auto_enabled,
                "time": AUTO_REFRESH_TIME,
                "timezone": TIMEZONE_NAME,
                "next_at": _next_auto_refresh_iso() if auto_enabled else None,
            },
            "synced_at": datetime.now(timezone.utc).isoformat(),
        }
    )


@app.get("/api/download")
def download() -> FileResponse:
    """
    Build the formatted workbook on demand from current Sheet data,
    so the download always reflects the latest statuses.
    """

    try:
        _require_config()
        df = sheets.read_sheet(
            SPREADSHEET_ID, WORKSHEET_NAME, **_sheet_kwargs()
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    if df.empty:
        raise HTTPException(
            status_code=404,
            detail="No purchase orders yet. Run a refresh first.",
        )

    df = engine.sort_orders_descending(df)

    engine.atomic_save_excel(df)

    stamp = datetime.now().strftime("%Y%m%d")

    return FileResponse(
        engine.OUTPUT_FILE,
        media_type=(
            "application/vnd.openxmlformats-officedocument"
            ".spreadsheetml.sheet"
        ),
        filename=f"TCS_Purchase_Orders_{stamp}.xlsx",
    )


@app.post("/api/upload-status")
async def upload_status(file: UploadFile = File(...)) -> JSONResponse:
    """
    Take a downloaded copy of the workbook with Order Status
    hand-edited, and apply just those status changes to the Sheet.

    Why editing in Excel is worth supporting: Order Status is merged
    across every line of a PO in the download, so changing a PO's
    status once at the top updates it for every line visually --
    editing the same status on every individual row in the Sheet
    itself is the tedious alternative this replaces.

    Why this is safe to use with an old download: only rows whose PO
    Number + Line # already exist in the live Sheet are touched, and
    only their Status cell is written. New POs added by a Gmail
    refresh that happened after this file was downloaded are left
    completely alone; nothing here can delete or overwrite data,
    at worst an unmatched or unchanged row is a no-op.
    """

    if not file.filename or not file.filename.lower().endswith((".xlsx", ".xlsm")):
        raise HTTPException(
            status_code=400,
            detail="Please upload the .xlsx file downloaded from this console.",
        )

    try:
        _require_config()
        contents = await file.read()

        status_map = engine.read_uploaded_excel_statuses(io.BytesIO(contents))

        if not status_map:
            return JSONResponse(
                {
                    "matched": 0,
                    "changed": 0,
                    "message": "No Order Status values were found in that file.",
                }
            )

        matched, changed = sheets.update_statuses_from_map(
            status_map,
            SPREADSHEET_ID,
            WORKSHEET_NAME,
            **_sheet_kwargs(),
        )

    except HTTPException:
        raise
    except Exception as exc:
        logger.exception("Status upload failed")
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    unmatched = len(status_map) - matched

    message = f"Updated {changed} of {matched} matching line item(s)."
    if unmatched:
        message += f" {unmatched} row(s) in the file no longer exist in the Sheet and were skipped."

    return JSONResponse(
        {
            "matched": matched,
            "changed": changed,
            "unmatched": unmatched,
            "message": message,
        }
    )