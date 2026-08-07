# Gmail Triage Summary

Sends a digest email of **starred** vs **non-starred** headlines to your general
inbox, on a schedule.

## Why this exists

The Gmail triage (which stars the emails that matter) runs inside a Claude
session that only has the Gmail **connector** tools. That connector can create
drafts but has **no send capability** — so the "email me a summary" step could
never actually deliver anything. An unsent draft just sits in the Drafts folder
and never reaches an inbox, which is why no summary was arriving.

This job performs the delivery step through Gmail's own IMAP + SMTP endpoints
instead, so the digest genuinely lands in the inbox. It's decoupled from the
triage: it only *reads back* the starred/unstarred state and emails the summary.

## Setup (one-time)

1. On the triaged Gmail account, enable **2-Step Verification**, then create an
   **App Password** (Google Account → Security → App passwords). Also make sure
   **IMAP is enabled** (Gmail → Settings → Forwarding and POP/IMAP).

2. In this repo: **Settings → Secrets and variables → Actions**, add secrets:
   - `GMAIL_ADDRESS` — the triaged account, e.g. `you@gmail.com`
   - `GMAIL_APP_PASSWORD` — the 16-character App Password (not your login password)

3. Optional repository **variables** (same page, "Variables" tab):
   - `SUMMARY_TO` — where to send the digest (defaults to `GMAIL_ADDRESS`)
   - `LOOKBACK_HOURS` — how far back to scan (defaults to `24`)
   - `SUMMARY_MAILBOX` — mailbox/label to scan (defaults to `INBOX`)

## Running it

- **Automatically:** the workflow runs daily (see the `cron` in
  `.github/workflows/gmail-triage-summary.yml`).
- **On demand:** Actions tab → *Gmail Triage Summary* → *Run workflow*.
- **Locally:**

  ```bash
  export GMAIL_ADDRESS="you@gmail.com"
  export GMAIL_APP_PASSWORD="xxxx xxxx xxxx xxxx"
  export SUMMARY_TO="you@hotmail.co.uk"   # optional
  python scripts/gmail_triage_summary.py
  ```

The script uses only the Python standard library, so no dependencies are needed.
