#!/usr/bin/env python3
"""
m365_sender.py — Send personalized emails from YOUR OWN Microsoft 365 mailbox,
reading the recipient list from a Google Sheet you "Publish to web" as CSV.

This is the FAST PATH — designed to go from zero to first send in an hour or two,
entirely self-serve, no IT ticket, no admin approval, no Google Cloud project:

  * Microsoft side: you sign in ONCE with your own Microsoft 365 account using the
    "device code" flow (the script shows a short code + URL; you approve it in a
    browser). It sends email AS YOU, using delegated Mail.Send permission you can
    usually grant yourself. It saves a refresh token so future runs are silent.
  * Google side: you click File → Share → Publish to web → CSV in your own Sheet
    and paste that link into config.env. The script just downloads that CSV. No
    service account, no API keys, no Google Cloud setup.
  * Targeting happens IN the sheet: only rows whose `Send` column is truthy get
    emailed (plus an optional config filter). Because a published sheet is
    read-only, results are recorded in a LOCAL send_log.csv — which also gives
    you resume/idempotency (nobody gets emailed twice).

Zero third-party dependencies — pure Python standard library. Just Python 3.9+.

Usage:
    1. Copy config.example.env to config.env and fill in the FAST-PATH values.
    2. Publish your Google Sheet to the web as CSV (SETUP-google-sheet.md) and
       paste its URL into config.env as CSV_URL.
    3. Edit SUBJECT_TEMPLATE / BODY_TEMPLATE below to say what you want to say.
    4. Run:
         python3 m365_sender.py --dry-run   # download + target + print. Sends nothing.
         python3 m365_sender.py --limit 1   # really send to just the first eligible row (a test)
         python3 m365_sender.py             # send to everyone eligible
       The FIRST real send prints a code + URL for a one-time browser sign-in.

Read SETUP-microsoft-365.md and SETUP-google-sheet.md first.

(Want write-back into the sheet and app-only auth instead? That's the optional
heavier path — see m365_sender_advanced.py and the appendices in the setup guides.)
"""
from __future__ import annotations

import argparse
import csv
import io
import json
import os
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlencode
from urllib.request import Request, urlopen
from urllib.error import HTTPError

# --------------------------------------------------------------------------
# 1. YOUR EMAIL CONTENT — edit these two templates.
#    Anything in {curly_braces} is replaced with the matching column from your
#    Google Sheet. So {name} pulls the "name" column, {city} pulls "city", etc.
#    Matching is case-insensitive, so {name} works whether your header is
#    "name", "Name", or "NAME". Unknown placeholders become blank (no crash).
# --------------------------------------------------------------------------
SUBJECT_TEMPLATE = "A quick note for you, {name}"

BODY_TEMPLATE = """\
Hi {name},

This is a personalized message sent just to you.

(Replace this with whatever you actually want to say. You can reference any
column from your sheet here, e.g. {name}, {city}, or {role}.)

Best,
Your Name
"""

# --------------------------------------------------------------------------
# 2. Pacing / safety caps.
#    Microsoft 365 throttles sending. Safe defaults stay well under the limits
#    (~30 messages/minute; ~10,000 recipients/day on a normal licensed mailbox).
#    Slow down if you see HTTP 429 errors.
# --------------------------------------------------------------------------
SLEEP_SECONDS_BETWEEN_SENDS = 2.0     # 2s => ~30/min
DAILY_SEND_CAP = 10000                # hard stop per run

# --------------------------------------------------------------------------
# 3. Which sheet columns hold what. Header matching is case-insensitive.
#    You normally don't change these — just use these header names in your sheet.
# --------------------------------------------------------------------------
EMAIL_COLUMN = "email"        # required: the recipient's email address
SEND_COLUMN = "Send"          # primary targeting: truthy value => eligible to send
TRUTHY = {"yes", "y", "true", "1", "x", "send", "✓", "checked"}

# Microsoft Graph delegated scopes we request. Mail.Send lets us send as the
# signed-in user; offline_access gives us a refresh token so we don't re-login
# every run; openid/profile identify the account.
GRAPH_SCOPES = "openid profile offline_access https://graph.microsoft.com/Mail.Send"
GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"

SCRIPT_DIR = Path(__file__).resolve().parent
SEND_LOG_PATH = SCRIPT_DIR / "send_log.csv"
DEFAULT_TOKEN_CACHE = SCRIPT_DIR / "token_cache.json"


# ==========================================================================
# Config loading — reads config.env (simple KEY=value lines) into os.environ.
# ==========================================================================
def load_config(path: Path) -> None:
    if not path.exists():
        sys.exit(
            f"Config file not found: {path}\n"
            "Copy config.example.env to config.env and fill in your values."
        )
    for line in path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def require(name: str) -> str:
    val = os.environ.get(name)
    if not val:
        sys.exit(f"Missing required config value: {name} (set it in config.env)")
    return val


def _norm(s: str) -> str:
    """Normalize a header/value: lower-cased, trimmed. Forgiving matching."""
    return (s or "").strip().lower()


# ==========================================================================
# Recipient list — download the "Published to web" CSV and parse it.
# A published Google Sheet CSV is a plain, public-but-unguessable URL that always
# reflects the current sheet contents. We just fetch it with the standard library.
# ==========================================================================
def load_recipients(local_file: str | None = None) -> tuple[list[dict], list[str]]:
    # MCP / manual hook: if a local CSV file is given (via --recipients-file), read
    # that instead of downloading. This is how the "Claude reads the sheet through
    # its own Google Sheets connector" path works — Claude writes the eligible rows
    # to a local CSV (same columns) and points the script at it. No CSV_URL needed.
    if local_file:
        path = Path(local_file).expanduser()
        if not path.exists():
            sys.exit(f"--recipients-file not found: {path}")
        raw = path.read_text(encoding="utf-8-sig")
    else:
        url = require("CSV_URL")
        req = Request(url, headers={"User-Agent": "m365-sender/1.0"})
        try:
            with urlopen(req, timeout=60) as resp:
                raw = resp.read().decode("utf-8-sig")   # utf-8-sig strips a BOM if present
        except HTTPError as e:
            sys.exit(
                f"Could not download the sheet CSV (HTTP {e.code}).\n"
                "Check CSV_URL in config.env. It must be the 'Publish to web → CSV' link "
                "(File → Share → Publish to web), not the normal sheet URL."
            )
        except Exception as e:
            sys.exit(f"Could not download the sheet CSV: {e}")

    reader = csv.DictReader(io.StringIO(raw))
    rows = [r for r in reader]
    headers = reader.fieldnames or []
    if not rows:
        sys.exit("The published sheet has a header row but no data rows. Nothing to do.")
    return rows, headers


def find_header(headers: list[str], wanted: str) -> str | None:
    target = _norm(wanted)
    for h in headers:
        if _norm(h) == target:
            return h
    return None


# ==========================================================================
# Microsoft auth — DEVICE CODE flow (delegated, "sign in as yourself").
# One-time interactive login: the script asks Microsoft for a device code,
# prints a short user code + URL, you approve in a browser, and Microsoft hands
# back an access token + a refresh token. We cache the refresh token so every
# later run silently mints a fresh access token with no prompt.
# ==========================================================================
def _token_endpoint() -> str:
    # "organizations" works for any work/school (M365) account without needing
    # to know your tenant. If you set TENANT_ID in config, we use that instead.
    tenant = os.environ.get("TENANT_ID", "").strip() or "organizations"
    return f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"


def _token_cache_path() -> Path:
    return Path(os.environ.get("TOKEN_CACHE", str(DEFAULT_TOKEN_CACHE))).expanduser()


def _post_form(url: str, form: dict) -> tuple[int, dict]:
    req = Request(url, data=urlencode(form).encode("utf-8"),
                  headers={"Content-Type": "application/x-www-form-urlencoded"},
                  method="POST")
    try:
        with urlopen(req, timeout=30) as resp:
            return resp.status, json.loads(resp.read())
    except HTTPError as e:
        return e.code, json.loads(e.read() or b"{}")


def _device_code_login(client_id: str) -> dict:
    """Run the interactive device-code flow; return the token response dict."""
    tenant = os.environ.get("TENANT_ID", "").strip() or "organizations"
    dc_url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode"
    status, dc = _post_form(dc_url, {"client_id": client_id, "scope": GRAPH_SCOPES})
    if status != 200 or "device_code" not in dc:
        sys.exit(f"Could not start device-code sign-in: {dc}")

    print("\n" + "=" * 66)
    print("ONE-TIME MICROSOFT SIGN-IN")
    print(dc.get("message")
          or f"Go to {dc['verification_uri']} and enter code: {dc['user_code']}")
    print("=" * 66 + "\n", flush=True)

    interval = int(dc.get("interval", 5))
    device_code = dc["device_code"]
    deadline = time.time() + int(dc.get("expires_in", 900))
    while time.time() < deadline:
        time.sleep(interval)
        status, tok = _post_form(_token_endpoint(), {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": client_id,
            "device_code": device_code,
        })
        if status == 200:
            print("Signed in. (Saving a refresh token so you won't be asked again.)")
            return tok
        err = tok.get("error")
        if err == "authorization_pending":
            continue          # you haven't approved in the browser yet — keep waiting
        if err == "slow_down":
            interval += 5
            continue
        sys.exit(f"Sign-in failed: {err} — {tok.get('error_description', '')}")
    sys.exit("Sign-in timed out. Run the script again to retry.")


def _save_refresh_token(tok: dict) -> None:
    rt = tok.get("refresh_token")
    if not rt:
        return
    path = _token_cache_path()
    path.write_text(json.dumps({"refresh_token": rt}))
    try:
        os.chmod(path, 0o600)   # keep the token file private
    except OSError:
        pass


def _refresh_access_token(client_id: str, refresh_token: str) -> dict | None:
    status, tok = _post_form(_token_endpoint(), {
        "grant_type": "refresh_token",
        "client_id": client_id,
        "refresh_token": refresh_token,
        "scope": GRAPH_SCOPES,
    })
    return tok if status == 200 else None


def get_access_token() -> str:
    """Return a usable access token, using a cached refresh token if we have one,
    otherwise running the one-time interactive device-code sign-in."""
    client_id = require("CLIENT_ID")
    cache = _token_cache_path()

    # Try the saved refresh token first (silent path).
    if cache.exists():
        try:
            rt = json.loads(cache.read_text()).get("refresh_token")
        except Exception:
            rt = None
        if rt:
            tok = _refresh_access_token(client_id, rt)
            if tok and tok.get("access_token"):
                _save_refresh_token(tok)   # Microsoft may rotate the refresh token
                print("Authenticated to Microsoft Graph (silent).")
                return tok["access_token"]
            print("Saved sign-in expired — signing in again.")

    # No/expired refresh token: do the interactive device-code login once.
    tok = _device_code_login(client_id)
    _save_refresh_token(tok)
    if not tok.get("access_token"):
        sys.exit("Sign-in returned no access token; try again.")
    return tok["access_token"]


# ==========================================================================
# Send one email via Graph's sendMail action, AS the signed-in user (/me).
# Delegated Mail.Send sends from your own mailbox. Graph returns HTTP 202 on
# success (queued for delivery) with no body.
# ==========================================================================
def send_mail(access_token: str, to_email: str, subject: str, body_text: str) -> None:
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": True,   # keep a copy in your Sent Items
    }
    req = Request(
        f"{GRAPH_API_BASE}/me/sendMail",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urlopen(req, timeout=60) as resp:
        if resp.status not in (200, 202):
            raise RuntimeError(f"Unexpected status {resp.status}")


# ==========================================================================
# Personalization — fill {placeholders} from a sheet row, tolerating missing keys
# and matching placeholder names to headers case-insensitively.
# ==========================================================================
class _Blankable(dict):
    def __missing__(self, key):  # {unknown} -> "" instead of KeyError
        return ""


def render(template: str, row: dict) -> str:
    merged = {}
    for k, v in row.items():
        merged[k] = v
        merged[_norm(k)] = v
    return template.format_map(_Blankable(merged))


# ==========================================================================
# Local send-log = the record of what happened AND the idempotency/resume store.
# (The sheet is read-only when published, so this file is the source of truth for
# "who have I already emailed for this campaign?") Keyed by CAMPAIGN_ID + email.
# ==========================================================================
def load_already_sent(campaign_id: str) -> set[str]:
    done: set[str] = set()
    if SEND_LOG_PATH.exists():
        with SEND_LOG_PATH.open(newline="") as f:
            for r in csv.DictReader(f):
                if r.get("campaign_id") == campaign_id and r.get("status") == "sent":
                    done.add((r.get("email") or "").lower())
    return done


def append_log(campaign_id: str, email: str, status: str, detail: str = "") -> None:
    new_file = not SEND_LOG_PATH.exists()
    with SEND_LOG_PATH.open("a", newline="") as f:
        w = csv.writer(f)
        if new_file:
            w.writerow(["timestamp_utc", "campaign_id", "email", "status", "detail"])
        w.writerow([datetime.now(timezone.utc).isoformat(), campaign_id,
                    email, status, detail])


# ==========================================================================
# Main loop.
# ==========================================================================
def main() -> None:
    ap = argparse.ArgumentParser(
        description="Send personalized M365 emails to recipients in a published Google Sheet.")
    ap.add_argument("--config", default=str(SCRIPT_DIR / "config.env"))
    ap.add_argument("--recipients-file", default=None,
                    help="Read recipients from this local CSV instead of CSV_URL. "
                         "Used for the 'Claude reads the sheet via its own Sheets "
                         "connector' path — it writes eligible rows here, then runs this.")
    ap.add_argument("--dry-run", action="store_true",
                    help="Download + target + render + print each email. Sends nothing.")
    ap.add_argument("--limit", type=int, default=0,
                    help="Only process the first N eligible recipients (0 = all). Great for a test.")
    args = ap.parse_args()

    load_config(Path(args.config))
    campaign_id = os.environ.get("CAMPAIGN_ID", "default")
    filter_column = os.environ.get("FILTER_COLUMN", "").strip()
    filter_value = os.environ.get("FILTER_VALUE", "").strip()

    rows, headers = load_recipients(args.recipients_file)
    email_header = find_header(headers, EMAIL_COLUMN)
    send_header = find_header(headers, SEND_COLUMN)
    if not email_header:
        sys.exit(f"No '{EMAIL_COLUMN}' column found in the sheet header. Headers: {headers}")
    if not send_header:
        print(f"WARNING: no '{SEND_COLUMN}' column found — every row is treated as "
              f"eligible. Add a '{SEND_COLUMN}' column to target specific rows.")

    print(f"Loaded {len(rows)} row(s) from the published sheet. Campaign: '{campaign_id}'."
          + (" [DRY RUN]" if args.dry_run else ""))
    if filter_column and filter_value:
        print(f"Config filter active: {filter_column} == {filter_value}")

    already_sent = load_already_sent(campaign_id)
    access_token = None if args.dry_run else get_access_token()

    sent = skipped = failed = 0
    processed = 0
    for i, row in enumerate(rows):
        sheet_row = i + 2   # +1 header, +1 for 1-based, for friendlier messages

        email = str(row.get(email_header, "") or "").strip()
        if not email:
            continue

        # --- Targeting (all in-the-sheet, plus optional config filter) ---
        if send_header is not None:
            if _norm(str(row.get(send_header, ""))) not in TRUTHY:
                continue
        if filter_column and filter_value:
            fh = find_header(headers, filter_column)
            if fh is None or _norm(str(row.get(fh, ""))) != _norm(filter_value):
                continue

        # --- Idempotency / resume from the local log ---
        if email.lower() in already_sent:
            skipped += 1
            print(f"[row {sheet_row}] SKIP: {email} already sent for '{campaign_id}'.")
            continue

        if args.limit and processed >= args.limit:
            break
        processed += 1

        if sent >= DAILY_SEND_CAP:
            print(f"Reached DAILY_SEND_CAP ({DAILY_SEND_CAP}). Stopping.")
            break

        subject = render(SUBJECT_TEMPLATE, row)
        body = render(BODY_TEMPLATE, row)

        if args.dry_run:
            print(f"\n[row {sheet_row}] --- DRY RUN -> {email} ---\n"
                  f"Subject: {subject}\n{body}")
            continue

        try:
            send_mail(access_token, email, subject, body)
        except HTTPError as e:
            failed += 1
            detail = e.read().decode("utf-8", "replace")[:300]
            hint = " (throttled — increase SLEEP_SECONDS_BETWEEN_SENDS)" if e.code == 429 else ""
            print(f"[row {sheet_row}] FAIL {email}: HTTP {e.code}{hint}")
            append_log(campaign_id, email, "failed", f"{e.code} {detail}")
            continue
        except Exception as e:  # network hiccup, etc. — isolate to this recipient
            failed += 1
            print(f"[row {sheet_row}] FAIL {email}: {e}")
            append_log(campaign_id, email, "failed", str(e))
            continue

        sent += 1
        already_sent.add(email.lower())
        append_log(campaign_id, email, "sent")
        print(f"[row {sheet_row}] SENT -> {email}")
        time.sleep(SLEEP_SECONDS_BETWEEN_SENDS)

    print(f"\nDone. sent={sent} skipped={skipped} failed={failed}. "
          f"Record + resume log: {SEND_LOG_PATH.name}")


if __name__ == "__main__":
    main()
