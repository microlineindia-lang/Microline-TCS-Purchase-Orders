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
- Supports scheduled refreshes via Cloud Scheduler or a local scheduler

## Why this project exists

Microline India receives purchase orders from TCS through Ariba email notifications. Those emails contain structured PO details, but they are not easy to track in a clean operational register. This project automates the extraction and maintains a central Sheet storing the latest state of purchase orders.

The system keeps the Google Sheet as the source of truth for operational status, which prevents lost data during Cloud Run restarts and allows teammates to update order statuses without overwriting their edits during a new import.

## Project structure

- `main.py` — FastAPI web app, routes, job tracking, scheduled refresh logic
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
- Daily scheduled refresh via Cloud Scheduler
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

- A Google Cloud project with access to Google Sheets
- A service account JSON file or Application Default Credentials configured
- A Gmail account with an app password for IMAP access
- A Google Sheet containing or accepting the required PO data structure
- Access to Cloud Run or another hosting environment if deploying remotely

## Environment variables

Set the following environment variables before running the app:

### Required

- `TCS_SPREADSHEET_ID` — the Google Sheet ID to read/write
- `TCS_GMAIL_APP_PASSWORD` — Gmail app password used for IMAP access

### Optional

- `TCS_WORKSHEET_NAME` — sheet tab name, default: `Orders`
- `TCS_META_WORKSHEET` — metadata sheet name, default: `Meta`
- `TCS_TIMEZONE` — time zone used for scheduling, default: `Asia/Kolkata`
- `TCS_AUTO_REFRESH_TIME` — scheduled refresh time in `HH:MM`, default: `08:00`
- `TCS_CRON_TOKEN` — token to validate scheduled refresh requests
- `TCS_AUTO_REFRESH_LOCAL` — enable a local background scheduler when running on a local machine
- `GOOGLE_APPLICATION_CREDENTIALS` — local path to a service account JSON file
- `PORT` — used by container platforms such as Cloud Run

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

## Docker deployment

The repository includes a Dockerfile for deployment in containerized environments such as Google Cloud Run.

Build the image:

```bash
docker build -t microline-tcs-purchase-orders .
```

Run locally:

```bash
docker run -p 8080:8080 \
  -e TCS_SPREADSHEET_ID="your_sheet_id" \
  -e TCS_GMAIL_APP_PASSWORD="your_gmail_app_password" \
  -e GOOGLE_APPLICATION_CREDENTIALS="/tmp/credentials.json" \
  microline-tcs-purchase-orders
```

## Cloud Run deployment

This app is designed to run on Google Cloud Run with a single worker, which is important because the job state is kept in memory for polling and progress tracking.

Recommended deployment approach:

- Use Cloud Run with a single instance
- Set `TCS_SPREADSHEET_ID` and `TCS_GMAIL_APP_PASSWORD` as environment variables or secrets
- Configure Google Cloud service account credentials via Application Default Credentials
- Use Cloud Scheduler to trigger the scheduled endpoint
- Protect the scheduled endpoint with the `TCS_CRON_TOKEN` header and matching `X-Cron-Token`

Example scheduled refresh endpoint:

```text
POST /api/refresh/scheduled
Header: X-Cron-Token: your-secret-token
```

## API endpoints

The service exposes the following HTTP endpoints:

- `GET /` — web dashboard
- `GET /healthz` — health check
- `POST /api/refresh` — start a new Gmail refresh job
- `GET /api/jobs/{job_id}` — poll a running job status
- `GET /api/summary` — read current sheet summary
- `GET /api/download` — download the generated Excel workbook
- `POST /api/upload-status` — apply edited status values back to the sheet
- `POST /api/refresh/scheduled` — scheduled refresh endpoint

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
- Cloud Run instances are ephemeral, so local generated files are not relied upon for persistence.
- Job tracking is intentionally in-memory and assumes single-instance operation for consistent polling behavior.
- The app logs refresh attempts and success events to a metadata sheet for operational visibility.

## Security considerations

- Store secrets in environment variables or cloud secret management, not in source code.
- Protect scheduled endpoints with a token header when exposed publicly.
- Restrict Google service account permissions to the minimum needed.
- Keep Gmail app-password credentials out of version control.

## License

This project is intended for internal operational use within Microline India and may be customized to match internal compliance and deployment standards.

## GitHub repository description

A Python + FastAPI service for importing TCS SAP Ariba purchase order emails from Gmail into a Google Sheet, preserving status updates, tracking refresh jobs, and generating Excel exports for operational reporting.

## Suggested short description for GitHub

Automated TCS SAP Ariba purchase-order import workflow for Microline India — Gmail parsing, Google Sheets synchronization, Excel export, and status tracking dashboard.
