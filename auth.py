"""
Access control: "Sign in with Google", restricted to an allow-list.

How it works
------------
1. /login shows Google's sign-in button.
2. Google hands the browser a signed ID token; the browser POSTs it to
   /auth/google.
3. We verify the token (signature, audience = our client id, expiry) and
   that the email is verified AND on the allow-list
   (TCS_ALLOWED_EMAILS, default: microlineindia@gmail.com).
4. If so, we set a signed, HttpOnly session cookie. Every other route
   requires that cookie. The allow-list is re-checked on every request,
   so removing an address from the env var revokes access immediately.

Fails closed: if GOOGLE_CLIENT_ID is not set, nobody can sign in and
every page is blocked (unless TCS_AUTH_DISABLED=1, for local testing).

Exempt from the login gate
--------------------------
/healthz                    liveness probe, returns {"status": "ok"}
/login, /auth/*             the sign-in flow itself
/api/refresh/scheduled      protected by its own secret header
                            (X-Cron-Token) instead -- see main.py
"""

from __future__ import annotations

import html
import json
import logging
import os
import secrets

from fastapi import FastAPI, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse
from itsdangerous import BadSignature, URLSafeTimedSerializer
from pydantic import BaseModel

logger = logging.getLogger("tcs.auth")

GOOGLE_CLIENT_ID = os.getenv("GOOGLE_CLIENT_ID", "").strip()

ALLOWED_EMAILS = {
    e.strip().lower()
    for e in os.getenv("TCS_ALLOWED_EMAILS", "microlineindia@gmail.com").split(",")
    if e.strip()
}

AUTH_DISABLED = os.getenv("TCS_AUTH_DISABLED", "").strip().lower() in {
    "1",
    "true",
    "yes",
}

try:
    SESSION_HOURS = max(1, int(os.getenv("TCS_SESSION_HOURS", "12")))
except ValueError:
    SESSION_HOURS = 12

_secret = os.getenv("TCS_SESSION_SECRET", "").strip()
if not _secret:
    # Works, but everyone is signed out whenever the app restarts.
    _secret = secrets.token_urlsafe(48)
    logger.warning(
        "TCS_SESSION_SECRET is not set; using a random one. Sessions will "
        "reset on every restart. Set TCS_SESSION_SECRET to keep people "
        "signed in."
    )

_serializer = URLSafeTimedSerializer(_secret, salt="tcs-session")

COOKIE_NAME = "tcs_session"

PUBLIC_PATHS = {"/healthz", "/login", "/auth/google", "/auth/logout"}
TOKEN_PROTECTED_PATHS = {"/api/refresh/scheduled"}


# ------------------------------------------------------------------
# Session helpers
# ------------------------------------------------------------------


def current_user(request: Request) -> str | None:
    """Return the signed-in, still-allowed email, or None."""

    if AUTH_DISABLED:
        return "local (auth disabled)"

    token = request.cookies.get(COOKIE_NAME)
    if not token:
        return None

    try:
        data = _serializer.loads(token, max_age=SESSION_HOURS * 3600)
    except BadSignature:
        return None

    email = str(data.get("email", "")).lower()
    return email if email in ALLOWED_EMAILS else None


def _is_https(request: Request) -> bool:
    proto = request.headers.get("x-forwarded-proto", request.url.scheme)
    return proto.split(",")[0].strip() == "https"


# ------------------------------------------------------------------
# Login page
# ------------------------------------------------------------------

_LOGIN_PAGE = """<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sign in &middot; TCS Purchase Orders</title>
<script src="https://accounts.google.com/gsi/client" async defer></script>
<style>
  body { margin:0; min-height:100vh; display:grid; place-items:center;
         font-family: system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
         background:#f4f6fa; color:#1f2a44; }
  .card { background:#fff; padding:36px 32px; border-radius:14px; width:min(380px, 90vw);
          box-shadow:0 6px 30px rgba(20,35,70,.10); text-align:center; }
  h1 { font-size:20px; margin:0 0 4px; }
  p.sub { margin:0 0 24px; color:#5b6b8c; font-size:14px; }
  #btn { display:flex; justify-content:center; min-height:44px; }
  #msg { margin-top:18px; font-size:14px; color:#b42318; min-height:20px; }
  .note { margin-top:22px; font-size:12px; color:#8592ad; }
</style>
</head>
<body>
  <div class="card">
    <h1>TCS Purchase Orders</h1>
    <p class="sub">Microline India &middot; authorised access only</p>
    __BODY__
    <div id="msg">__ERROR__</div>
    <div class="note">Sign in with the authorised Google account.</div>
  </div>
__SCRIPT__
</body>
</html>
"""

_LOGIN_SCRIPT = """<script>
  const CLIENT_ID = __CLIENT_ID__;
  async function onCredential(resp) {
    const msg = document.getElementById("msg");
    msg.textContent = "";
    try {
      const r = await fetch("/auth/google", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ credential: resp.credential }),
      });
      if (r.ok) { location.href = "/"; return; }
      const d = await r.json().catch(() => ({}));
      msg.textContent = d.detail || "Sign-in failed.";
    } catch (e) {
      msg.textContent = "Network error. Please try again.";
    }
  }
  window.onload = function () {
    const t = setInterval(function () {
      if (!window.google || !google.accounts) return;
      clearInterval(t);
      google.accounts.id.initialize({ client_id: CLIENT_ID, callback: onCredential });
      google.accounts.id.renderButton(document.getElementById("btn"),
        { theme: "outline", size: "large", text: "signin_with", shape: "pill" });
    }, 100);
  };
</script>"""


def _render_login(error: str = "", configured: bool = True) -> str:
    if not configured:
        body = ""
        script = ""
        error = (
            "Sign-in is not configured. Set GOOGLE_CLIENT_ID on the server."
        )
    else:
        body = '<div id="btn"></div>'
        script = _LOGIN_SCRIPT.replace(
            "__CLIENT_ID__", json.dumps(GOOGLE_CLIENT_ID)
        )

    return (
        _LOGIN_PAGE.replace("__BODY__", body)
        .replace("__ERROR__", html.escape(error))
        .replace("__SCRIPT__", script)
    )


# ------------------------------------------------------------------
# Wiring
# ------------------------------------------------------------------


class _Credential(BaseModel):
    credential: str


def install(app: FastAPI) -> None:
    """Attach the login gate and the /login, /auth/*, /api/me routes."""

    if AUTH_DISABLED:
        logger.warning(
            "TCS_AUTH_DISABLED is on: the app is OPEN to anyone who can "
            "reach it. Use this for local testing only."
        )
    else:
        logger.info(
            "Access restricted to: %s", ", ".join(sorted(ALLOWED_EMAILS))
        )
        if not GOOGLE_CLIENT_ID:
            logger.error(
                "GOOGLE_CLIENT_ID is not set -- all access is blocked."
            )

    @app.middleware("http")
    async def _login_gate(request: Request, call_next):
        path = request.url.path

        if (
            AUTH_DISABLED
            or path in PUBLIC_PATHS
            or path in TOKEN_PROTECTED_PATHS
            or current_user(request)
        ):
            response = await call_next(request)
            response.headers.setdefault("Cache-Control", "no-store")
            response.headers.setdefault("X-Content-Type-Options", "nosniff")
            return response

        wants_page = request.headers.get("sec-fetch-dest") == "document"

        if path.startswith("/api/") and not wants_page:
            return JSONResponse({"detail": "Not signed in"}, status_code=401)

        return RedirectResponse("/login", status_code=303)

    @app.get("/login", response_class=HTMLResponse)
    def login_page(request: Request) -> HTMLResponse:
        if AUTH_DISABLED or current_user(request):
            return RedirectResponse("/", status_code=303)  # type: ignore[return-value]

        return HTMLResponse(
            _render_login(configured=bool(GOOGLE_CLIENT_ID)),
            headers={"Cache-Control": "no-store"},
        )

    @app.post("/auth/google")
    def auth_google(body: _Credential, request: Request) -> JSONResponse:
        if not GOOGLE_CLIENT_ID:
            return JSONResponse(
                {"detail": "Sign-in is not configured."}, status_code=503
            )

        try:
            from google.auth.transport import requests as g_requests
            from google.oauth2 import id_token

            info = id_token.verify_oauth2_token(
                body.credential, g_requests.Request(), GOOGLE_CLIENT_ID
            )
        except Exception:
            logger.warning("Rejected an invalid Google token")
            return JSONResponse(
                {"detail": "Could not verify your Google sign-in."},
                status_code=401,
            )

        email = str(info.get("email", "")).lower()

        if not info.get("email_verified") or email not in ALLOWED_EMAILS:
            logger.warning("Sign-in refused for %s", email or "(no email)")
            return JSONResponse(
                {"detail": "This account is not authorised to use this app."},
                status_code=403,
            )

        logger.info("Signed in: %s", email)

        response = JSONResponse({"ok": True, "email": email})
        response.set_cookie(
            COOKIE_NAME,
            _serializer.dumps({"email": email}),
            max_age=SESSION_HOURS * 3600,
            httponly=True,
            secure=_is_https(request),
            samesite="lax",
            path="/",
        )
        return response

    @app.post("/auth/logout")
    def auth_logout() -> JSONResponse:
        response = JSONResponse({"ok": True})
        response.delete_cookie(COOKIE_NAME, path="/")
        return response

    @app.get("/api/me")
    def me(request: Request) -> JSONResponse:
        return JSONResponse({"email": current_user(request)})
