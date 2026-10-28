# Microline TCS Purchase Orders

A Python-based workflow for importing Tata Consultancy Services (TCS) SAP Ariba purchase orders from Gmail into a Google Sheet, preserving manually updated order statuses, and exposing a lightweight web dashboard for monitoring, download, and refresh operations.

This project is designed for operational teams that need to track purchase order line items, review status changes, and maintain a clean, spreadsheet-based register without manually copying data from email notifications.

## Overview

The application does the following:

- Connects to Gmail using IMAP
- Finds TCS SAP Ariba purchase order emails
- Parses each email into structured purchase order line items
- Normalizes values such as pricing, quantities, supplier details, and product text
- Updates or appends rows in a Google Sheet
- Preserves human-edited `Order Status` values while refreshing data
- Exposes a web console to trigger manual refreshes and check progress
- Allows users to download the current workbook in Excel format
- Restricts access to allow-listed Google accounts (default `microlineindia@gmail.com`)
- Supports scheduled refreshes via a free GitHub Actions cron or a local scheduler

## Why this project exists

Microline India receives purchase orders from TCS through Ariba email notifications. Those emails contain structured PO details, but they are not easy to track in a clean operational register. This project automates the extraction and maintains a central Sheet storing the latest state of purchase orders.

The system keeps the Google Sheet as the source of truth for operational status, which prevents lost data when the host restarts and allows teammates to update order statuses without overwriting their edits during a new import.

## Project structure

- `main.py` — FastAPI web app, routes, job tracking, scheduled refresh logic
- `auth.py` — Google Sign-In login gate and email allow-list
- `.env.example` — every setting, with comments
- `DEPLOY.md` — free hosting and setup guide
- `.github/workflows/daily-refresh.yml` — free daily 8 AM trigger
- `tcs_engine.py` — email parsing, Excel generation, normalization, order processing
- `sheets.py` — Google Sheets read/write logic and status synchronization
- `static/index.html` — web dashboard UI
- `Dockerfile` — container definition for deployment
- `requirements.txt` — Python dependencies

## Core features

### Gmail import pipeline
The app connects to Gmail using the IMAP protocol and filters messages from Ariba order senders with a TCS purchase order subject pattern.

### Structured PO extraction
Each purchase order email is parsed into a tabular record set with:

- PO Number
- Order Date
- Total PO Amount
- PR Number
- Requester
- Delivery Location
- Buyer Contact
- Payment Terms
- Attachments
- Buyer Tax ID
- Supplier Tax ID
- Order Status
- Line items such as Product Name, Description, Qty, Unit, Unit Price, Subtotal, and HSN/SAC code

### Status safety
The Google Sheet is treated as the operational system of record. The app reads each existing `Order Status` value before writing new rows and overlays the saved statuses onto refreshed data. This ensures that any human edits are preserved during later imports.

### Downloadable Excel output
The system generates a workbook with a clean tabular schema and presentation formatting for downstream reporting or share-outs.

### Dashboard and job tracking
The FastAPI UI shows:

- current totals from the Google Sheet
- refresh progress
- processing stage labels
- ETA estimates
- logs from the import job
- last refresh timestamps

### Scheduled refresh support
The app supports:

- Manual refresh through the web UI
- Daily scheduled refresh via GitHub Actions
- Local scheduler support using environment-based configuration

## Technology stack

- Python 3.12
- FastAPI
- Uvicorn
- Pandas
- BeautifulSoup
- OpenPyXL
- gspread
- Google Auth
- Gmail IMAP
- Google Sheets API

## Prerequisites

Before running the app, make sure you have:

- A free Google Cloud project (no billing needed) with the Sheets API enabled
- A service account key, shared as Editor on the Sheet
- An OAuth client id for Google Sign-In
- A Gmail account with an app password for IMAP access
- A Google Sheet containing or accepting the required PO data structure
- A hosting environment if deploying remotely (see `DEPLOY.md`)

## Environment variables

See `.env.example` for a ready-to-fill template and `DEPLOY.md` for how to get each value.

### Required

- `TCS_SPREADSHEET_ID` — the Google Sheet ID to read/write
- `TCS_GMAIL_APP_PASSWORD` — Gmail app password used for IMAP access
- `GOOGLE_SERVICE_ACCOUNT_JSON` — full contents of the service-account key (or `GOOGLE_APPLICATION_CREDENTIALS` = path to the key file, for local runs)
- `GOOGLE_CLIENT_ID` — OAuth web client id for Google Sign-In
- `TCS_SESSION_SECRET` — long random string used to sign login cookies

### Optional

- `TCS_ALLOWED_EMAILS` — comma-separated accounts allowed in, default `microlineindia@gmail.com`
- `TCS_SESSION_HOURS` — login lifetime, default `12`
- `TCS_WORKSHEET_NAME` — sheet tab name, default `Orders`
- `TCS_META_WORKSHEET` — metadata sheet name, default `Meta`
- `TCS_TIMEZONE` — time zone for scheduling, default `Asia/Kolkata`
- `TCS_AUTO_REFRESH_TIME` — scheduled refresh time in `HH:MM`, default `08:00`
- `TCS_CRON_TOKEN` — secret that the daily GitHub Actions job sends as `X-Cron-Token`
- `TCS_SCHEDULED_BACKGROUND` — run scheduled imports in the background (default `1`)
- `TCS_AUTO_REFRESH_LOCAL` — built-in scheduler for an always-on local machine
- `TCS_AUTH_DISABLED` — turns login off; **local testing only**
- `PORT` — set automatically by hosting platforms

## Local setup

1. Clone the repository
2. Create a virtual environment
3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Export the required environment variables:

```bash
export TCS_SPREADSHEET_ID="your_google_sheet_id"
export TCS_GMAIL_APP_PASSWORD="your_gmail_app_password"
export GOOGLE_APPLICATION_CREDENTIALS="/path/to/service-account.json"
```

5. Start the app:

```bash
uvicorn main:app --host 0.0.0.0 --port 8000 --workers 1
```

6. Open the dashboard in a browser:

```text
http://localhost:8000/
```

## Deployment

The app runs anywhere that can run a Docker container or Python 3.12. The recommended
free setup is **Render (free web service) + a GitHub Actions daily cron**; the full
step-by-step guide, including the Google sign-in setup, is in [`DEPLOY.md`](DEPLOY.md).

Run the container locally:

```bash
docker build -t microline-tcs-purchase-orders .
docker run -p 8080:8080 --env-file .env microline-tcs-purchase-orders
```

Run only **one instance with one worker**: job progress is kept in memory.

Scheduled refresh endpoint (called by the GitHub Actions workflow):

```text
POST /api/refresh/scheduled
Header: X-Cron-Token: your-secret-token
```

## API endpoints

The service exposes the following HTTP endpoints:

- `GET /` — web dashboard (login required)
- `GET /login`, `POST /auth/google`, `POST /auth/logout`, `GET /api/me` — sign-in flow
- `GET /healthz` — health check (public)
- `POST /api/refresh` — start a new Gmail refresh job
- `GET /api/jobs/{job_id}` — poll a running job status
- `GET /api/summary` — read current sheet summary
- `GET /api/download` — download the generated Excel workbook
- `POST /api/upload-status` — apply edited status values back to the sheet
- `POST /api/refresh/scheduled` — scheduled refresh endpoint (token-protected, not behind login)

## Data flow

1. Read existing sheet data and status map
2. Connect to Gmail
3. Find Ariba order emails
4. Parse each email and normalize the content
5. Merge new and existing records by PO Number + Line
6. Preserve manually updated `Order Status`
7. Write the updated dataset back to Google Sheets
8. Store refresh metadata in the `Meta` sheet
9. Expose summary and workbook downloads to the UI

## Notes on reliability

- The Google Sheet is the system of record for order status, not the local filesystem.
- Hosted instances are ephemeral, so local generated files are not relied upon for persistence.
- Job tracking is intentionally in-memory and assumes single-instance operation for consistent polling behavior.
- The app logs refresh attempts and success events to a metadata sheet for operational visibility.

## Security considerations

- Store secrets in environment variables or cloud secret management, not in source code.
- Every route except `/healthz`, the sign-in flow and the token-protected scheduled endpoint requires a Google sign-in from an allow-listed account.
- Protect scheduled endpoints with a token header when exposed publicly.
- Restrict Google service account permissions to the minimum needed.
- Keep Gmail app-password credentials out of version control.

## License

This project is intended for internal operational use within Microline India and may be customized to match internal compliance and deployment standards.

## GitHub repository description

A Python + FastAPI service for importing TCS SAP Ariba purchase order emails from Gmail into a Google Sheet, preserving status updates, tracking refresh jobs, and generating Excel exports for operational reporting.

## Suggested short description for GitHub

Automated TCS SAP Ariba purchase-order import workflow for Microline India — Gmail parsing, Google Sheets synchronization, Excel export, and status tracking dashboard.
