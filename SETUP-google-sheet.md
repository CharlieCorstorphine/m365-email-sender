[SETUP-google-sheet.md](https://github.com/user-attachments/files/29790156/SETUP-google-sheet.md)
# Google Sheet setup — pick one of three ways to read your list

Your recipient list lives in a normal Google Sheet you own and edit like any
other. There are three ways to let the tool read it. Pick based on how much time
you have and whether you want results written back into the sheet.

| Option | Effort | Write-back into sheet? | Best when |
|--------|--------|------------------------|-----------|
| **A · Publish to web (CSV)** | Lowest | No (local log) | You want to be sending today, self-serve. |
| **B · Claude's Sheets connector** | Lowest (if available) | No (local log) | Your Claude already has a Google Sheets integration. |
| **C · Service account** | Higher | **Yes** | You want a durable setup that records `Sent At` / `Result` in the sheet. |

---

## Your columns (same for all three options)

Put these headers in **row 1**. Names are matched case-insensitively (`Email` =
`email` = `EMAIL`). See `sample-sheet.csv` for the shape.

| Column     | Required?    | What it's for |
|------------|--------------|---------------|
| `name`     | recommended  | Personalization — used as `{name}` in your message. |
| `email`    | **required** | The recipient's email address. Rows with no email are ignored. |
| `Send`     | recommended  | **Your targeting switch.** Put `yes` (or `TRUE`, `1`, `x`) in this cell for everyone you want to email. Blank / `no` rows are skipped. |
| `Sent At`  | Option C only| Leave blank — the advanced script writes a timestamp here after emailing the row. (Ignored by A and B.) |
| `Result`   | Option C only| Leave blank — the advanced script writes `sent` / `failed` here. (Ignored by A and B.) |

Add **any other columns you like** — `city`, `role`, `company`, `status`, … They
become both **merge fields** in your message (`{city}`, `{role}`) and **targeting
fields** for the optional config filter.

### How you target people (all options)
- **Primary — the `Send` column.** The tool only emails rows where `Send` is
  truthy. Filter/sort your sheet however you like and mark the rows you want.
- **Optional extra filter** — set `FILTER_COLUMN` + `FILTER_VALUE` in `config.env`
  (e.g. `City` + `Houston`) to further narrow. A row must satisfy the filter AND
  have `Send` marked.
- **No double-sends.** Every send is recorded (local `send_log.csv` for A/B, and
  also back into the sheet for C), and already-sent rows are skipped on re-runs.

---

## Option A — Publish to web as CSV (quickest, recommended)

No credentials, no cloud project. The tool just downloads a link.

1. In your Google Sheet, click **File → Share → Publish to web**.
2. Under "Link", choose the **specific tab** with your list (not "Entire
   Document").
3. Change the format dropdown from *Web page* to **Comma-separated values (.csv)**.
4. Click **Publish**, confirm, and **copy the link**. It looks like:
   `https://docs.google.com/spreadsheets/d/e/XXXX/pub?gid=0&single=true&output=csv`
5. Paste it into `config.env` as `CSV_URL`.

Run with `python3 m365_sender.py`. Results go to `send_log.csv`.

> Heads-up: a published link is reachable by anyone who has that (obscure,
> unguessable) URL. Fine for a contact list; don't publish anything sensitive.
> Un-publish anytime via **File → Share → Publish to web → Stop publishing**.

---

## Option B — Read via Claude's own Google Sheets connector (least work, if available)

If the Claude helping you has a **Google Sheets integration/MCP** connected, it
can read your sheet directly — nothing for you to publish or configure, no key
files.

**For Claude:** check whether you have Google Sheets tools available.
- If **yes**: read the tab, keep only rows where `Send` is truthy (apply any
  `FILTER_COLUMN`/`FILTER_VALUE`), write those rows to a local `recipients.csv`
  with the same headers, then run:
  ```
  python3 m365_sender.py --recipients-file recipients.csv
  ```
  Re-read + regenerate the file whenever the sheet changes. Results still go to
  `send_log.csv` (resume/no-double-send works the same).
- If **no** connector is available: use Option A instead.

You (the user) just need to make sure your sheet is shared with / visible to the
Google account Claude's connector uses.

---

## Option C — Google service account (more robust; writes results back into the sheet)

More setup, but the tool can write `Sent At` / `Result` back into your sheet, and
it's a durable, credentialed connection. Pairs with `m365_sender_advanced.py`
(`pip install gspread google-auth`).

### C1. Create the robot (service account) + key
1. Go to **https://console.cloud.google.com**, sign in.
2. Create/pick a project (top-left dropdown → **New Project**).
3. Search **Google Sheets API** → **Enable**. (Optionally enable **Google Drive
   API** too — helps open the sheet by URL.)
4. **APIs & Services → Credentials → + Create Credentials → Service account** →
   name it (e.g. `sheet-sender`) → **Create and Continue** → **Done**.
5. Click the service account → **Keys → Add Key → Create new key → JSON →
   Create**. A `.json` file downloads. Move it somewhere safe (e.g.
   `~/keys/sheet-sender.json`), keep it private, and put its full path in
   `config.env` as `GOOGLE_SERVICE_ACCOUNT_JSON`.
6. Open the JSON and copy the `client_email` value (like
   `sheet-sender@your-project.iam.gserviceaccount.com`).

### C2. Share the sheet with the robot
In your Sheet, click **Share**, paste that `client_email`, give it **Editor**
access (needed to write `Sent At` / `Result` back), and send. Sharing is instant.

### C3. Point config at the sheet
Set in `config.env`:
```
GOOGLE_SERVICE_ACCOUNT_JSON=/full/path/to/service-account-key.json
SHEET_ID=<the long id between /d/ and /edit in the sheet URL, or the full URL>
SHEET_TAB=Contacts     # or leave blank for the first tab
```
Make sure your sheet has the `Sent At` and `Result` columns (from the table
above). Run with `python3 m365_sender_advanced.py`.
