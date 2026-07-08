[SETUP-microsoft-365.md](https://github.com/user-attachments/files/29790189/SETUP-microsoft-365.md)
# Microsoft 365 setup — the FAST PATH (sign in as yourself)

Goal: let the script send email from **your own** Microsoft 365 mailbox, using a
one-time interactive sign-in. No admin approval, no client secret, no waiting on
IT. You'll end with ONE value for `config.env`: `CLIENT_ID`.

> Why this way: you're signing in as yourself and granting the script permission
> to send mail *on your own behalf* (delegated **Mail.Send**). Most organizations
> let regular users consent to this themselves — so you don't need to be a Global
> Administrator. The script signs in once via a "device code" (it shows a short
> code + link, you approve in a browser) and saves a refresh token so later runs
> are silent.

---

## Step by step (about 10 minutes)

### 1. Open the app-registration portal
Go to **https://entra.microsoft.com** (or https://portal.azure.com), sign in with
your Microsoft 365 account, and open **Applications → App registrations**.

### 2. Register a new app
Click **+ New registration**.
- **Name:** anything, e.g. `My Email Sender`.
- **Supported account types:** **"Accounts in this organizational directory only"**.
- Leave "Redirect URI" blank.
- Click **Register**.

### 3. Copy your Client ID
On the app's **Overview** page, copy **Application (client) ID** → this is your
`CLIENT_ID`. (You can ignore the Tenant ID on the fast path; leaving `TENANT_ID`
blank in config uses "organizations", which works for normal M365 accounts.)

### 4. Turn ON "public client flows"
Left menu → **Authentication** → scroll to **Advanced settings** → set
**"Allow public client flows"** to **Yes** → **Save**. (This is what enables the
device-code sign-in. There's no secret to create on the fast path.)

### 5. Add the delegated "send mail" permission
Left menu → **API permissions → + Add a permission → Microsoft Graph →
Delegated permissions**. Search **Mail.Send**, tick it, **Add permissions**.
(`offline_access` — for the refresh token — is included automatically when the
script signs in; you don't need to add it manually.)

- You do **not** need the "Grant admin consent" button on the fast path. When you
  first sign in (step 7), Microsoft shows you a consent screen and — if your org
  allows user consent — you approve it yourself.

### 6. Fill in your config file
Copy `config.example.env` to `config.env` and set:
```
CLIENT_ID=...        (from step 3)
TENANT_ID=           (leave blank unless your admin says otherwise)
```

### 7. First send = one-time sign-in
Run a real test send:
```
python3 m365_sender.py --limit 1
```
The first time, the script prints something like:
> To sign in, use a web browser to open **https://microsoft.com/devicelogin**
> and enter the code **ABCD-EFGH**.

Open that link, enter the code, sign in with your M365 account, and approve. The
script then sends, and saves a refresh token to `token_cache.json` so you're
**never asked again** (until the token eventually expires).

---

## Things to know (gotchas)

- **"Need admin approval" on the consent screen?** Your org has turned off
  user self-consent. Ask an admin to click **Grant admin consent** on your app's
  **API permissions** page once — that's the only admin touch, and far lighter
  than the app-only path.
- **The mailbox that sends is YOUR mailbox** (the account you sign in with). It
  must be a real, licensed M365 mailbox.
- **Sending limits.** A normal M365 mailbox can send to roughly **10,000
  recipients per day** and about **30 messages per minute**. The script paces
  itself (2s between sends) and stops at 10,000 per run. `HTTP 429` = throttled;
  raise `SLEEP_SECONDS_BETWEEN_SENDS` near the top of `m365_sender.py` and re-run
  (already-sent people are skipped automatically).
- **Bulk/marketing email is different.** M365 is built for normal business mail.
  For thousands of cold/marketing emails, use a dedicated bulk provider (SendGrid,
  Amazon SES). This script is for personalized outreach to a known list.
- **Keep `config.env` and `token_cache.json` private** — the token can send mail
  as you. Don't share or commit them.

---
---

## APPENDIX — OPTIONAL heavier path (Decision 1B): app-only auth / shared mailbox

Do this if you want to send from a **shared or company mailbox** (a team address
that outlives any one person and isn't tied to an individual's login) and/or run
fully unattended as a service with no person signed in — **and** you can get a
Global Administrator to approve it once. More robust and durable, but more setup.
It pairs with `m365_sender_advanced.py`. Most people on a short timeline should use
the fast path above (Decision 1A).

### A1. Register the app (steps 1–3 above), but you'll also create a secret.

### A2. Add an APPLICATION permission (not delegated)
**API permissions → + Add a permission → Microsoft Graph → Application
permissions** → search **Mail.Send** → tick → **Add permissions**.

### A3. Create a client secret
**Certificates & secrets → + New client secret** → set an expiry → **Add** →
**immediately copy the secret's `Value`** (shown once). This is `CLIENT_SECRET`.

### A4. Grant admin consent (requires a Global Administrator)
On **API permissions**, click **"Grant admin consent for [your org]"**. The
Mail.Send status must flip to green **Granted**. Without this, every send is
rejected.

### A5. Config values for the advanced script
In `config.env`, also copy the Tenant ID (Overview page) and set:
```
TENANT_ID=...
CLIENT_ID=...
CLIENT_SECRET=...        (from A3)
SENDER_UPN=you@yourcompany.com   (the licensed mailbox to send FROM)
```
Then run `python3 m365_sender_advanced.py` instead of `m365_sender.py`.

Notes: app-only can send as ANY mailbox in your tenant by default; an admin can
restrict it with an **Application Access Policy** (`New-ApplicationAccessPolicy`).
The secret expires — sending stops until you make a new one and update config.
