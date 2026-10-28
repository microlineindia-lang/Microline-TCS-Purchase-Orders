# Free deployment guide

The app is a single Python web service. The Google Sheet is the database, so the
host's disk can be thrown away at any time and nothing is lost.

**Access control:** only the Google account(s) in `TCS_ALLOWED_EMAILS` (default
`microlineindia@gmail.com`) can open the app. Everyone else is bounced to a sign-in
page and refused.

---

## Where to run it for free

| Option | Cost | Always on? | Notes |
|---|---|---|---|
| **Render – Free web service** (recommended) | Free, no card required | No – sleeps after 15 idle min, ~1 min to wake | 512 MB RAM. The GitHub Actions job below wakes it every morning. Easiest setup. |
| **Your own PC / office machine** | Free | While the PC is on | Zero limits. Run it with `TCS_AUTO_REFRESH_LOCAL=1`. Add a free Cloudflare Tunnel if you need to reach it from outside the office. |
| Oracle Cloud "Always Free" VM | Free | Yes | Truly always-on, but needs a card for sign-up and you manage the server yourself. |
| Google Cloud Run | Has a free allowance | – | Needs a billing account. The old guide used min-instances 1, which is what usually costs money. |

Not recommended: Hugging Face Docker Spaces (their docs now say Docker Spaces need a
paid plan), Fly.io (no free tier for new users), Railway (only ~$1 credit/month).
Free tiers change – check each provider's pricing page before committing.

**Free-tier trade-off:** the first visit after the app has been idle takes about a
minute to load. That is normal on free hosting.

---

## One-time Google setup (free, no billing needed)

### 1. Service account (lets the app read/write the Sheet)
1. [console.cloud.google.com](https://console.cloud.google.com) → pick or create a project.
2. **APIs & Services → Library** → enable **Google Sheets API** and **Google Drive API**.
3. **IAM & Admin → Service Accounts → Create service account** (any name, no roles needed).
4. Open it → **Keys → Add key → Create new key → JSON**. A `.json` file downloads.
5. Copy the service account's email (looks like `name@project.iam.gserviceaccount.com`).
6. Open your Google Sheet → **Share** → paste that email → **Editor**.

You will paste the *entire contents* of the JSON file into `GOOGLE_SERVICE_ACCOUNT_JSON`.
Treat that file like a password.

### 2. Sign-in (restricts access to microlineindia@gmail.com)
1. Same Cloud project → **APIs & Services → OAuth consent screen**.
   - User type **External**, fill in app name + your email, save.
   - Under **Test users** add `microlineindia@gmail.com`. (Leave the app in *Testing* –
     that is an extra lock: no one else can sign in even if they tried.)
2. **APIs & Services → Credentials → Create credentials → OAuth client ID**.
   - Type **Web application**.
   - **Authorized JavaScript origins:** add your app URL
     (e.g. `https://microline-tcs.onrender.com`) and `http://localhost:8000` for local use.
   - Leave redirect URIs empty. Create.
3. Copy the **Client ID** → this is `GOOGLE_CLIENT_ID`. (No client secret is needed.)

You'll only know the Render URL after step 1 of the Render section; come back and add
it to the authorized origins, then redeploy/restart.

### 3. Gmail app password
Already used by the existing app (`TCS_GMAIL_APP_PASSWORD`). If you need a new one:
Google Account → Security → 2-Step Verification → App passwords.

---

## Deploy on Render (free)

1. Put this project in a **private GitHub repo** (the `.gitignore` already keeps secrets out).
2. [render.com](https://render.com) → sign up → **New → Web Service** → connect the repo.
3. Settings: **Runtime: Docker**, **Instance type: Free**, **Health check path: `/healthz`**.
4. Add **Environment Variables**:

| Name | Value |
|---|---|
| `TCS_SPREADSHEET_ID` | the id in the Sheet URL (`/d/<this part>/edit`) |
| `TCS_GMAIL_APP_PASSWORD` | Gmail app password |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | full contents of the service-account JSON |
| `GOOGLE_CLIENT_ID` | OAuth client id from above |
| `TCS_ALLOWED_EMAILS` | `microlineindia@gmail.com` |
| `TCS_SESSION_SECRET` | long random string (`python -c "import secrets; print(secrets.token_urlsafe(48))"`) |
| `TCS_CRON_TOKEN` | another long random string |

   Optional: `TCS_TIMEZONE`, `TCS_AUTO_REFRESH_TIME`, `TCS_WORKSHEET_NAME`, `TCS_META_WORKSHEET`, `TCS_SESSION_HOURS`.
   See `.env.example` for the full list.
5. **Create Web Service.** When it's live, copy the URL, add it to the OAuth client's
   *Authorized JavaScript origins* (step 2 above), and open it. Sign in with
   `microlineindia@gmail.com`.

---

## Daily 8 AM refresh (replaces Cloud Scheduler)

The workflow `.github/workflows/daily-refresh.yml` runs at 08:00 IST, wakes the app,
and triggers the import. In the GitHub repo: **Settings → Secrets and variables →
Actions → New repository secret**:

- `APP_URL` – your Render URL, no trailing slash
- `TCS_CRON_TOKEN` – the **same** value you set on Render

Test it once from the **Actions** tab → *Daily TCS refresh* → **Run workflow**.

Notes: GitHub can start scheduled runs a few minutes late. In a public repo GitHub
disables scheduled workflows after 60 days with no repo activity – a private repo
is not affected the same way, but if it ever stops, re-enable it in the Actions tab.
The result of each run shows on the dashboard ("last refresh" / last failure).

---

## Run on your own PC instead

```bash
python -m venv .venv
.venv\Scripts\activate            # Windows   (Linux/Mac: source .venv/bin/activate)
pip install -r requirements.txt
copy .env.example .env            # then fill in .env
uvicorn main:app --host 127.0.0.1 --port 8000 --workers 1
```

Open <http://localhost:8000>. For the built-in daily refresh add
`TCS_AUTO_REFRESH_LOCAL=1` to `.env` (leave `TCS_CRON_TOKEN` empty and skip GitHub Actions).
To try it without Google sign-in, set `TCS_AUTH_DISABLED=1` – **local testing only**.

---

## Security notes

- Login is enforced server-side on every page and API route. The only open URLs are
  `/healthz`, the sign-in page, and `/api/refresh/scheduled` (needs `X-Cron-Token`).
- Sessions are signed, HttpOnly, `Secure` over HTTPS, and expire after 12 h.
- To add or remove a user, edit `TCS_ALLOWED_EMAILS` (comma-separated) – takes effect
  on restart; existing sessions of removed users stop working at once.
- Never commit `.env` or the service-account JSON. If either leaks, rotate it
  (new Gmail app password / new service-account key).
