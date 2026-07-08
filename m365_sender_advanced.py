#!/usr/bin/env python3
"""
=============================================================================
 OPTIONAL / ADVANCED VERSION — most people should use m365_sender.py instead.
=============================================================================
Use this heavier version ONLY if you want BOTH of these and have the time:
  - App-only Microsoft auth (a permanent app + secret, needs a Global Admin to
    grant consent once), instead of signing in as yourself; and
  - Live write-back INTO the Google Sheet (Sent At / Result columns), which needs
    a Google Cloud service account + JSON key.
It requires:  pip install gspread google-auth  and a Google Cloud project.
If you're an intern on a deadline with no admin rights, STOP — use the default
fast path (m365_sender.py): sign in as yourself, read a published-CSV sheet, and
track results in a local log. See README.md.
=============================================================================

m365_sender_advanced.py — Send personalized emails through Microsoft 365
(Microsoft Graph), reading the recipient list from a Google Sheet you own.

What it does, top to bottom:
  1. Opens your Google Sheet (via a Google "service account" you shared it with)
     and reads every row as a recipient. Columns like name/email/city/role become
     merge fields AND targeting fields.
  2. Decides who to email using dead-simple, IN-THE-SHEET targeting:
       - Primary: only rows whose `Send` column is truthy (yes / true / 1 / x).
       - Optional: a config filter (e.g. FILTER_COLUMN=City, FILTER_VALUE=Houston)
         to further narrow the list without touching the sheet's Send marks.
       - Idempotency: rows already stamped Result=sent are skipped, so a re-run
         resumes cleanly and never double-emails.
  3. Authenticates to Microsoft Graph using the "app-only" client-credentials
     flow (your registered Azure app's ID + secret). No interactive login.
  4. For each eligible recipient, fills your subject/body templates with their
     row's fields and sends ONE individual email from your chosen mailbox.
  5. Writes the outcome back to that row in the Sheet — `Sent At` (timestamp) and
     `Result` (sent / failed / skipped) — so there's a visible record, and also
     appends to a local send_log.csv as a backup.

Dependencies:  pip install gspread google-auth
(The Microsoft side uses only the Python standard library — no extra install.)

Usage:
    1. Copy config.example.env to config.env and fill in your values.
    2. Lay out your Google Sheet (see SETUP-google-sheet.md) and share it with the
       service account email.
    3. Edit SUBJECT_TEMPLATE / BODY_TEMPLATE below to say what you want to say.
    4. Run:
         python3 m365_sender.py --dry-run     # read + target + print. Sends nothing.
         python3 m365_sender.py --limit 1     # really send to just the first eligible row (a test)
         python3 m365_sender.py               # send to everyone eligible, write results back

Read SETUP-microsoft-365.md and SETUP-google-sheet.md first — they walk through
the two one-time setups that produce the values this script needs.
"""
from __future__ import annotations

import argparse
import csv
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
#    Column matching is case-insensitive and ignores surrounding spaces, so
#    {name} works whether your header is "name", "Name", or "NAME".
#    Unknown placeholders are left blank rather than crashing the whole run.
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
#    Microsoft 365 throttles sending. Safe defaults below stay well under the
#    limits (roughly 30 messages/minute; ~10,000 recipients/day on a normal
#    licensed mailbox). Slow down if you see HTTP 429 errors.
# --------------------------------------------------------------------------
SLEEP_SECONDS_BETWEEN_SENDS = 2.0     # 2s => ~30/min
DAILY_SEND_CAP = 10000                # hard stop per run

# --------------------------------------------------------------------------
# 3. Which sheet columns hold what. Header matching is case-insensitive and
#    space-insensitive, so "Sent At" matches a header of "sent at" too.
#    You normally don't need to change these — just use these header names in
#    your sheet (see SETUP-google-sheet.md).
# --------------------------------------------------------------------------
EMAIL_COLUMN = "email"        # required: the recipient's email address
SEND_COLUMN = "Send"          # primary targeting: truthy value => eligible to send
SENT_AT_COLUMN = "Sent At"    # written back: UTC timestamp of the send attempt
RESULT_COLUMN = "Result"      # written back: sent / failed / skipped

# Values in the Send column that mean "yes, email this row".
TRUTHY = {"yes", "y", "true", "1", "x", "send", "✓", "checked"}

GRAPH_API_BASE = "https://graph.microsoft.com/v1.0"
SCRIPT_DIR = Path(__file__).resolve().parent
SEND_LOG_PATH = SCRIPT_DIR / "send_log.csv"


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


# ==========================================================================
# Google Sheet access.
# We authenticate with a *service account* — a robot Google identity whose JSON
# key you downloaded (see SETUP-google-sheet.md). You shared your Sheet with that
# service account's email, exactly like sharing with a coworker, which is all the
# access it needs. gspread handles the Sheets API for us.
# ==========================================================================
def open_worksheet():
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError:
        sys.exit(
            "Missing dependencies. Run:  pip install gspread google-auth\n"
            "(These let the script read/write your Google Sheet.)"
        )

    key_path = Path(require("GOOGLE_SERVICE_ACCOUNT_JSON")).expanduser()
    if not key_path.exists():
        sys.exit(
            f"Google service-account key not found: {key_path}\n"
            "Check GOOGLE_SERVICE_ACCOUNT_JSON in config.env points to your JSON key file."
        )

    # These scopes let the robot read AND write the sheets you've shared with it.
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(str(key_path), scopes=scopes)
    client = gspread.authorize(creds)

    sheet_ref = require("SHEET_ID")   # accepts a bare Sheet ID or a full sheet URL
    try:
        if sheet_ref.startswith("http"):
            spreadsheet = client.open_by_url(sheet_ref)
        else:
            spreadsheet = client.open_by_key(sheet_ref)
    except Exception as e:
        sys.exit(
            f"Could not open the Google Sheet ({e}).\n"
            "Two things to check:\n"
            "  1. SHEET_ID in config.env is the sheet's ID or full URL.\n"
            "  2. You shared the sheet with the service account email "
            "(the client_email inside your JSON key)."
        )

    tab = os.environ.get("SHEET_TAB", "").strip()
    try:
        return spreadsheet.worksheet(tab) if tab else spreadsheet.sheet1
    except Exception as e:
        sys.exit(f"Could not open tab '{tab or '(first tab)'}' in the sheet ({e}).")


def _norm(s: str) -> str:
    """Normalize a header/value for forgiving, case- and space-insensitive matching."""
    return (s or "").strip().lower()


def find_header(headers: list[str], wanted: str) -> str | None:
    """Return the actual header string in the sheet that matches `wanted`, or None."""
    target = _norm(wanted)
    for h in headers:
        if _norm(h) == target:
            return h
    return None


# ==========================================================================
# Auth — app-only client-credentials flow (Microsoft Graph).
# We POST our client_id + client_secret to the tenant's token endpoint and get
# back a short-lived access_token (valid ~60 min). No user, no interactive login.
# This is why the app needs the *application* Mail.Send permission with admin
# consent (see SETUP-microsoft-365.md) — it acts as itself, not as a person.
# ==========================================================================
def get_access_token() -> str:
    tenant = require("TENANT_ID")
    body = urlencode({
        "grant_type": "client_credentials",
        "client_id": require("CLIENT_ID"),
        "client_secret": require("CLIENT_SECRET"),
        # For client-credentials you request the ".default" scope, which grants
        # every *application* permission already consented for the app.
        "scope": "https://graph.microsoft.com/.default",
    }).encode("utf-8")
    req = Request(
        f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token",
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"},
        method="POST",
    )
    try:
        with urlopen(req, timeout=30) as resp:
            token = json.loads(resp.read())["access_token"]
            print("Authenticated to Microsoft Graph.")
            return token
    except HTTPError as e:
        detail = e.read().decode("utf-8", "replace")
        sys.exit(
            f"Authentication failed ({e.code}). Check CLIENT_ID / CLIENT_SECRET / "
            f"TENANT_ID, and that admin consent was granted for Mail.Send.\n{detail}"
        )


# ==========================================================================
# Send one email via Graph's sendMail action.
# We send AS the mailbox in SENDER_UPN (e.g. you@yourcompany.com). The app must
# be allowed to send as that mailbox (any licensed mailbox in your tenant works
# for app-only unless locked down with an Application Access Policy — see the
# Microsoft setup guide). Graph returns HTTP 202 with no body on success.
# ==========================================================================
def send_mail(access_token: str, sender_upn: str, to_email: str,
              subject: str, body_text: str) -> None:
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body_text},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
        },
        "saveToSentItems": True,   # keep a copy in the sender's Sent Items
    }
    req = Request(
        f"{GRAPH_API_BASE}/users/{sender_upn}/sendMail",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urlopen(req, timeout=60) as resp:
        # 202 Accepted = queued for delivery. urlopen only reaches here on 2xx.
        if resp.status not in (200, 202):
            raise RuntimeError(f"Unexpected status {resp.status}")


# ==========================================================================
# Personalization — fill {placeholders} from a sheet row, tolerating missing keys.
# We build a lookup that matches placeholders case-insensitively to headers, so
# {name} finds a "Name" or "name" column, and {Role} finds "role".
# ==========================================================================
class _Blankable(dict):
    def __missing__(self, key):  # {unknown} -> "" instead of KeyError
        return ""


def render(template: str, row: dict) -> str:
    # Expose each column under both its exact header and a normalized lower-case
    # key, so {name} and {Name} both resolve regardless of the header's casing.
    merged = {}
    for k, v in row.items():
        merged[k] = v
        merged[_norm(k)] = v
    return template.format_map(_Blankable(merged))


# ==========================================================================
# Local send-log backup (belt-and-suspenders alongside the write-back to the
# Sheet). Keyed by CAMPAIGN_ID + email; used for a secondary "already sent"
# check even if someone edits the sheet's Result column by hand.
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
        description="Send personalized M365 emails to recipients in a Google Sheet.")
    ap.add_argument("--config", default=str(SCRIPT_DIR / "config.env"))
    ap.add_argument("--dry-run", action="store_true",
                    help="Read + target + render + print each email. Sends nothing, writes nothing.")
    ap.add_argument("--limit", type=int, default=0,
                    help="Only process the first N eligible recipients (0 = all). Great for a test.")
    args = ap.parse_args()

    load_config(Path(args.config))
    sender_upn = require("SENDER_UPN")            # mailbox you send AS, e.g. you@company.com
    campaign_id = os.environ.get("CAMPAIGN_ID", "default")

    # Optional in-config filter (add-on to the in-sheet Send column).
    filter_column = os.environ.get("FILTER_COLUMN", "").strip()
    filter_value = os.environ.get("FILTER_VALUE", "").strip()

    # ---- Read the sheet -------------------------------------------------
    ws = open_worksheet()
    records = ws.get_all_records()                # list of dicts keyed by header row
    headers = ws.row_values(1)                    # header row, in sheet order
    if not records:
        sys.exit("The sheet has a header row but no data rows. Nothing to do.")

    # Resolve the columns we care about to their real header strings.
    email_header = find_header(headers, EMAIL_COLUMN)
    send_header = find_header(headers, SEND_COLUMN)
    sent_at_header = find_header(headers, SENT_AT_COLUMN)
    result_header = find_header(headers, RESULT_COLUMN)

    if not email_header:
        sys.exit(f"No '{EMAIL_COLUMN}' column found in the sheet's header row. "
                 f"Headers seen: {headers}")
    if not send_header:
        print(f"WARNING: no '{SEND_COLUMN}' column found — every row is treated as "
              f"eligible. Add a '{SEND_COLUMN}' column to target specific rows.")

    # Column letters for write-back (1-based index -> A1 letter).
    def col_letter(header: str | None) -> str | None:
        if not header or header not in headers:
            return None
        return _index_to_col_letter(headers.index(header) + 1)

    sent_at_col = col_letter(sent_at_header)
    result_col = col_letter(result_header)
    if not args.dry_run and (not sent_at_col or not result_col):
        print(f"NOTE: '{SENT_AT_COLUMN}' and/or '{RESULT_COLUMN}' columns are missing, "
              "so results can't be written back to the sheet (local log still records them). "
              "Add those two columns to get in-sheet tracking.")

    print(f"Loaded {len(records)} row(s) from the sheet. Campaign: '{campaign_id}'. "
          f"Sender: {sender_upn}." + (" [DRY RUN]" if args.dry_run else ""))
    if filter_column and filter_value:
        print(f"Config filter active: {filter_column} == {filter_value}")

    already_sent = load_already_sent(campaign_id)
    access_token = None if args.dry_run else get_access_token()

    sent = skipped = failed = 0
    processed = 0
    for i, row in enumerate(records):
        sheet_row = i + 2   # +1 for header, +1 because sheet rows are 1-based

        email = str(row.get(email_header, "") or "").strip()
        if not email:
            continue  # blank row / no address — silently ignore

        # --- Targeting decision (all IN-THE-SHEET, plus optional config filter) ---
        # 1. Send column must be truthy (if the column exists at all).
        if send_header is not None:
            send_val = _norm(str(row.get(send_header, "")))
            if send_val not in TRUTHY:
                continue

        # 2. Optional config filter narrows further.
        if filter_column and filter_value:
            fh = find_header(headers, filter_column)
            if fh is None or _norm(str(row.get(fh, ""))) != _norm(filter_value):
                continue

        # 3. Idempotency: already marked sent (in sheet OR local log) -> skip.
        already_result = _norm(str(row.get(result_header, ""))) if result_header else ""
        if already_result == "sent" or email.lower() in already_sent:
            skipped += 1
            print(f"[row {sheet_row}] SKIP: {email} already sent.")
            continue

        # Respect --limit against ELIGIBLE rows (not raw sheet rows).
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
            send_mail(access_token, sender_upn, email, subject, body)
        except HTTPError as e:
            failed += 1
            detail = e.read().decode("utf-8", "replace")[:300]
            hint = " (throttled — increase SLEEP_SECONDS_BETWEEN_SENDS)" if e.code == 429 else ""
            print(f"[row {sheet_row}] FAIL {email}: HTTP {e.code}{hint}")
            append_log(campaign_id, email, "failed", f"{e.code} {detail}")
            _write_back(ws, sheet_row, sent_at_col, result_col, "failed")
            continue
        except Exception as e:  # network hiccup, etc. — isolate to this recipient
            failed += 1
            print(f"[row {sheet_row}] FAIL {email}: {e}")
            append_log(campaign_id, email, "failed", str(e))
            _write_back(ws, sheet_row, sent_at_col, result_col, "failed")
            continue

        sent += 1
        already_sent.add(email.lower())
        append_log(campaign_id, email, "sent")
        _write_back(ws, sheet_row, sent_at_col, result_col, "sent")
        print(f"[row {sheet_row}] SENT -> {email}")
        time.sleep(SLEEP_SECONDS_BETWEEN_SENDS)

    print(f"\nDone. sent={sent} skipped={skipped} failed={failed}. "
          f"Local backup log: {SEND_LOG_PATH.name}")


def _index_to_col_letter(idx: int) -> str:
    """1 -> A, 2 -> B, ... 27 -> AA (spreadsheet column letters)."""
    letters = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _write_back(ws, sheet_row: int, sent_at_col: str | None,
                result_col: str | None, result: str) -> None:
    """Stamp Sent At + Result back onto the recipient's row. Best-effort:
    a write-back failure never aborts the run (the local log still has it)."""
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    try:
        if sent_at_col:
            ws.update_acell(f"{sent_at_col}{sheet_row}", now)
        if result_col:
            ws.update_acell(f"{result_col}{sheet_row}", result)
    except Exception as e:
        print(f"    (write-back to sheet row {sheet_row} failed: {e})")


if __name__ == "__main__":
    main()
