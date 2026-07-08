[README.md](https://github.com/user-attachments/files/29790074/README.md)
# Personalized email sender — build kit

**What this is (for the human skimming):** a small tool that emails a list of
people one personalized message at a time, from a Microsoft 365 mailbox, where
the list lives in a Google Sheet you already know how to edit. You mark who to
email right in the spreadsheet. You don't need to be technical — you'll hand this
folder to Claude (Claude Code) and it will do the setup and sending with you.

**What this is (for Claude reading this):** the rest of this file is written to
**you, Claude**. The person who gave you this folder is **non-technical** and is
likely working on a **short (~2 day) deadline**. Your job is to stand this tool
up with them and get a real send out, with the least possible work on their end.
Assume you have no other context. Everything you need is in this folder.

---

## Files in this folder

| File | What it is |
|------|-----------|
| `README.md` | This file — your instructions. |
| `m365_sender.py` | **Default reference implementation.** Stdlib only. Device-code sign-in (send as the user) + published-CSV sheet reading + local log. Also accepts `--recipients-file` for the MCP path. |
| `m365_sender_advanced.py` | **Optional heavier version.** App-only auth + Google service account + write-back into the sheet. Needs `pip install gspread google-auth`. |
| `config.example.env` | Copy to `config.env` and fill in. Fast-path values on top; advanced values commented at the bottom. |
| `SETUP-microsoft-365.md` | Click-by-click for the Microsoft side (fast path + app-only appendix). |
| `SETUP-google-sheet.md` | Click-by-click for the Google side (all three sheet options). |
| `sample-sheet.csv` | The column layout the tool expects. |

---

## Pick your path (do this first — ask the user, then branch)

Before touching setup, walk the user through **two quick decisions**. Recommend
the **quick / self-serve** option for each given the short timeline, but present
both and let them choose. Then follow only the branch they pick.

### DECISION 1 — Which mailbox do the emails come FROM?

- **1A · Send from the user's OWN mailbox (Quick, self-serve — recommended).**
  One-time interactive "device code" sign-in as themselves. No admin approval —
  they consent for their own account. Fastest to stand up; perfect for a
  short-term task. Uses `m365_sender.py`. → `SETUP-microsoft-365.md` main section.

- **1B · Send from a SHARED / company mailbox (More robust, takes longer).**
  An app registration with app-only client-credentials + a **Global
  Administrator** granting consent (optionally scoped to one mailbox with an
  Application Access Policy). Tradeoff: more durable — a company/team address that
  outlives this person and isn't tied to their personal login — but needs an admin
  and more setup time. Uses `m365_sender_advanced.py`. → appendix in
  `SETUP-microsoft-365.md`.

### DECISION 2 — How does the tool READ the Google Sheet?

- **2A · "Publish to web → CSV" link (Quickest — recommended).** The user
  publishes their sheet tab as a CSV link; the tool just downloads it. No
  credentials, no cloud project. Read-only, so results are recorded in a local
  `send_log.csv` (which also powers resume / no-double-send). Note to the user:
  a published sheet is reachable by anyone who has the obscure link — fine for a
  contact list, but don't publish anything sensitive. Uses `m365_sender.py` +
  `CSV_URL`. → `SETUP-google-sheet.md` Option A.

- **2B · Read it directly through YOUR Google Sheets connector (Least work — IF
  available).** If you, Claude, have a Google Sheets integration/MCP connected in
  this session, you can read the user's sheet directly — no publishing, no keys,
  nothing for them to handle. **Check first:** do you have Google Sheets tools
  available? If yes, this is the least-effort option for a non-technical person:
  read the tab, keep only rows where `Send` is truthy (and any filter), write those
  rows to a local `recipients.csv` with the same headers, and run
  `python3 m365_sender.py --recipients-file recipients.csv`. If you do NOT have a
  Sheets connector, skip this and use 2A. → `SETUP-google-sheet.md` Option B.

- **2C · Google service account + share the sheet (More robust, durable).** A
  Google Cloud service account with a JSON key; the user shares the sheet with the
  robot's email. More setup, but it can **write results back into the sheet**
  (`Sent At` / `Result` columns) instead of a local log. Uses
  `m365_sender_advanced.py`. → `SETUP-google-sheet.md` Option C.

**Recommended default for a 2-day, self-serve task: 1A + 2A** (or 1A + 2B if you
have a Sheets connector). Only reach for 1B / 2C if the user explicitly wants a
durable, shared, write-back setup and has time + an admin.

---

## How to onboard the user (the happy path: 1A + 2A)

Do these with them, in order. Keep it friendly and concrete. Explain that
targeting and record-keeping live in the spreadsheet they already know — you'll
handle everything technical.

### Step 0 — Check the basics
Confirm Python 3.9+ is available (`python3 --version`). The default path needs
**no pip installs**. Then `cp config.example.env config.env` so they have a
config file to fill in.

### Step 1 — Microsoft side (Decision 1A)
Open `SETUP-microsoft-365.md` and walk them through registering the app, turning
on public client flows, and adding delegated `Mail.Send`. They copy ONE value —
`CLIENT_ID` — into `config.env`. Leave `TENANT_ID` blank. (The actual sign-in
happens later, at first send.)

### Step 2 — Lay out their Google Sheet
Have them open (or create) a Google Sheet and set row 1 to these headers — show
them `sample-sheet.csv` as the shape:

- `name` — used for personalization (`{name}` in the message). Recommended.
- `email` — **required.** The address to send to.
- `Send` — **their targeting switch.** They put `yes` (or `TRUE`, `1`, `x`) in
  this cell for each person to email; blank / `no` rows are skipped.
- Any extra columns they want — `city`, `role`, `company`, `status`, … — usable
  both as merge fields (`{city}`) and as the optional filter.

Explain the whole targeting model in one sentence: **filter/sort your sheet
however you like, then put `yes` in the `Send` column for the people you want to
email.** That's it. (On the fast path there's no `Sent At` / `Result` write-back —
the local log tracks that. Those columns only matter for the advanced 2C path.)

### Step 3 — Publish the sheet as CSV (Decision 2A)
In `SETUP-google-sheet.md` Option A: File → Share → **Publish to web** → choose
the tab → **CSV** → Publish. Copy that link into `config.env` as `CSV_URL`.
(If instead you're doing 2B via your own connector, skip publishing — see that
option; if 2C, follow Option C.)

### Step 4 — Write the message
Edit `SUBJECT_TEMPLATE` and `BODY_TEMPLATE` near the top of `m365_sender.py`
together. Anything in `{curly_braces}` is replaced per person from their sheet
columns (`{name}`, `{city}`, …). Draft copy WITH the user; don't send anything
they haven't approved.

### Step 5 — Dry run (sends nothing)
```
python3 m365_sender.py --dry-run
```
This downloads the sheet, applies targeting, and prints each rendered email.
Confirm the right people are selected and the merge fields look right. Fix the
sheet or templates and repeat until it's clean.

### Step 6 — Send to exactly one (the real test)
```
python3 m365_sender.py --limit 1
```
The **first** real send triggers the one-time Microsoft sign-in: the script
prints a short code and `https://microsoft.com/devicelogin`. Have the user open
it, enter the code, sign in, and approve. The script sends to the first eligible
row and saves a refresh token so they're never prompted again. Best practice:
make that first row the user's OWN address so they can eyeball the result.

### Step 7 — Full send
```
python3 m365_sender.py
```
Sends to everyone eligible, ~30/minute, one personalized email each, isolating
per-recipient failures. Every outcome is written to `send_log.csv`. **Re-running
is safe** — anyone already sent for this `CAMPAIGN_ID` is skipped, so an
interrupted run resumes cleanly. For a brand-new campaign later, change
`CAMPAIGN_ID` in `config.env`.

---

## Branch notes (when the user picked a heavier option)

- **1B (shared/app-only mailbox):** follow the appendix in
  `SETUP-microsoft-365.md`, set `TENANT_ID` / `CLIENT_ID` / `CLIENT_SECRET` /
  `SENDER_UPN` in `config.env`, and run `m365_sender_advanced.py`. Needs an admin
  to grant consent once.
- **2C (service-account + write-back):** follow `SETUP-google-sheet.md` Option C,
  `pip install gspread google-auth`, set `GOOGLE_SERVICE_ACCOUNT_JSON` / `SHEET_ID`
  / `SHEET_TAB`, add `Sent At` and `Result` columns to the sheet, and run
  `m365_sender_advanced.py`. Results are written back into the sheet AND a local
  log. The advanced script pairs 1B-style app-only auth with 2C write-back.
- **Mixing:** the advanced script uses app-only auth (1B) together with
  service-account sheet access (2C). The fast script (`m365_sender.py`) covers
  1A + (2A or 2B). If the user wants an unusual mix (e.g. own-mailbox 1A but
  write-back 2C), tell them the simplest supported combos are "both quick"
  (fast script) or "both robust" (advanced script), and pick the closest.

---

## Guardrails (please honor these)

- **Never send anything the user hasn't seen and approved.** Always dry-run, then
  `--limit 1` to their own address, before a full send.
- **Keep `config.env` and `token_cache.json` private.** They can send mail as the
  user. Don't commit, share, or paste them.
- **This is for personalized outreach to a known list, not mass cold marketing.**
  M365 throttles and polices bulk mail; for true mass sending point them at a
  dedicated provider (SendGrid, Amazon SES). Say so if their list looks like cold
  marketing at scale.
- **If a step fails, read the script's error message** — they're written to be
  actionable (wrong CLIENT_ID, un-published sheet, throttling, etc.).
