#!/usr/bin/env python3
"""
Gmail Triage Summary — email a digest of starred / non-starred headlines.

The Gmail triage stars the emails that matter and leaves the rest unstarred.
This script reads that triaged state back out of the mailbox and sends a
single summary email to the general inbox, listing the starred headlines
(the ones flagged as important) separately from the non-starred ones.

Why this exists
---------------
The triage runs inside a Claude session that only has the Gmail *connector*
tools available. That connector can create drafts but has no "send" tool, so
the summary step could never actually deliver an email — an unsent draft never
leaves the Drafts folder. This job runs the delivery step through Gmail's own
IMAP/SMTP endpoints instead, so the digest genuinely arrives in the inbox.

It uses only the Python standard library (imaplib / smtplib / email), so the
workflow needs no pip install.

Environment / secrets
----------------------
  GMAIL_ADDRESS       Gmail account that is being triaged (e.g. you@gmail.com)
  GMAIL_APP_PASSWORD  Google App Password for that account (NOT the login
                      password). Requires 2-Step Verification + an App Password,
                      and IMAP enabled in Gmail settings.
  SUMMARY_TO          Where to send the digest. Optional; defaults to
                      GMAIL_ADDRESS (i.e. the account's own inbox).
  LOOKBACK_HOURS      How far back to scan. Optional; defaults to 24.
  SUMMARY_MAILBOX     Mailbox/label to scan. Optional; defaults to "INBOX".

Usage
-----
  python scripts/gmail_triage_summary.py
"""

from __future__ import annotations

import email
import email.utils
import html
import imaplib
import os
import smtplib
import sys
from datetime import datetime, timedelta, timezone
from email.header import decode_header, make_header
from email.message import EmailMessage

IMAP_HOST = "imap.gmail.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465


def _require(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        sys.exit(f"ERROR: required environment variable {name} is not set.")
    return value


def _decode(raw: str | None) -> str:
    """Decode a possibly RFC 2047-encoded header into a plain string."""
    if not raw:
        return ""
    try:
        return str(make_header(decode_header(raw))).strip()
    except Exception:
        return raw.strip()


def _sender_name(from_header: str) -> str:
    name, addr = email.utils.parseaddr(from_header)
    name = _decode(name)
    return name or addr or from_header


class Headline:
    __slots__ = ("subject", "sender", "date")

    def __init__(self, subject: str, sender: str, date: str):
        self.subject = subject
        self.sender = sender
        self.date = date


def fetch_headlines(
    imap: imaplib.IMAP4_SSL, mailbox: str, since: datetime
) -> tuple[list[Headline], list[Headline]]:
    """Return (starred, non_starred) headlines newer than `since`."""
    status, _ = imap.select(mailbox, readonly=True)
    if status != "OK":
        sys.exit(f"ERROR: could not open mailbox {mailbox!r}.")

    # IMAP SINCE has day granularity; we filter to the exact window afterwards.
    since_str = since.strftime("%d-%b-%Y")

    def search(criteria: str) -> list[bytes]:
        status, data = imap.search(None, criteria, "SINCE", since_str)
        if status != "OK" or not data or not data[0]:
            return []
        return data[0].split()

    starred_ids = set(search("FLAGGED"))
    all_ids = search("ALL")

    starred: list[Headline] = []
    non_starred: list[Headline] = []

    for msg_id in all_ids:
        status, data = imap.fetch(
            msg_id, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM DATE)])"
        )
        if status != "OK" or not data or not isinstance(data[0], tuple):
            continue

        msg = email.message_from_bytes(data[0][1])
        date_raw = msg.get("Date", "")
        parsed = email.utils.parsedate_to_datetime(date_raw) if date_raw else None
        if parsed is not None:
            if parsed.tzinfo is None:
                parsed = parsed.replace(tzinfo=timezone.utc)
            if parsed < since:
                continue  # outside the exact lookback window

        headline = Headline(
            subject=_decode(msg.get("Subject")) or "(no subject)",
            sender=_sender_name(msg.get("From", "")),
            date=parsed.strftime("%a %d %b %H:%M") if parsed else "",
        )
        (starred if msg_id in starred_ids else non_starred).append(headline)

    return starred, non_starred


def build_message(
    starred: list[Headline],
    non_starred: list[Headline],
    lookback_hours: int,
    sender: str,
    recipient: str,
) -> EmailMessage:
    now = datetime.now(timezone.utc)
    subject = (
        f"Gmail Triage Summary — {len(starred)} starred, "
        f"{len(non_starred)} other ({now:%d %b %Y})"
    )

    def text_block(title: str, items: list[Headline]) -> str:
        if not items:
            return f"{title} (0)\n  (none)\n"
        lines = [f"{title} ({len(items)})"]
        for h in items:
            meta = " · ".join(p for p in (h.sender, h.date) if p)
            lines.append(f"  • {h.subject}")
            if meta:
                lines.append(f"      {meta}")
        return "\n".join(lines) + "\n"

    text = (
        f"Triage summary for the last {lookback_hours}h "
        f"(as of {now:%Y-%m-%d %H:%M UTC}).\n\n"
        + text_block("STARRED", starred)
        + "\n"
        + text_block("NOT STARRED", non_starred)
    )

    def html_block(title: str, items: list[Headline], color: str) -> str:
        if not items:
            rows = "<li style='color:#888'>(none)</li>"
        else:
            rows = ""
            for h in items:
                meta = " · ".join(
                    html.escape(p) for p in (h.sender, h.date) if p
                )
                rows += (
                    "<li style='margin:0 0 8px'>"
                    f"<span style='font-weight:600'>{html.escape(h.subject)}</span>"
                    + (f"<br><span style='color:#888;font-size:12px'>{meta}</span>" if meta else "")
                    + "</li>"
                )
        return (
            f"<h3 style='margin:20px 0 8px;color:{color}'>"
            f"{html.escape(title)} <span style='color:#888;font-weight:400'>"
            f"({len(items)})</span></h3><ul style='padding-left:18px;margin:0'>{rows}</ul>"
        )

    html_body = (
        "<div style='font-family:-apple-system,Segoe UI,Roboto,sans-serif;"
        "max-width:640px;margin:0 auto;color:#1a1a1a'>"
        "<h2 style='margin:0 0 4px'>Gmail Triage Summary</h2>"
        f"<p style='color:#888;margin:0 0 8px;font-size:13px'>Last {lookback_hours}h · "
        f"as of {now:%Y-%m-%d %H:%M UTC}</p>"
        + html_block("Starred", starred, "#b8860b")
        + html_block("Not starred", non_starred, "#444")
        + "</div>"
    )

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = recipient
    msg.set_content(text)
    msg.add_alternative(html_body, subtype="html")
    return msg


def main() -> None:
    gmail_address = _require("GMAIL_ADDRESS")
    app_password = _require("GMAIL_APP_PASSWORD")
    recipient = os.environ.get("SUMMARY_TO", "").strip() or gmail_address
    mailbox = os.environ.get("SUMMARY_MAILBOX", "").strip() or "INBOX"
    try:
        lookback_hours = int(os.environ.get("LOOKBACK_HOURS", "24"))
    except ValueError:
        lookback_hours = 24

    since = datetime.now(timezone.utc) - timedelta(hours=lookback_hours)

    print(f"Connecting to {IMAP_HOST} as {gmail_address} …")
    imap = imaplib.IMAP4_SSL(IMAP_HOST)
    try:
        imap.login(gmail_address, app_password)
        starred, non_starred = fetch_headlines(imap, mailbox, since)
    finally:
        try:
            imap.logout()
        except Exception:
            pass

    print(f"Found {len(starred)} starred and {len(non_starred)} non-starred "
          f"message(s) in the last {lookback_hours}h.")

    msg = build_message(
        starred, non_starred, lookback_hours, gmail_address, recipient
    )

    print(f"Sending summary to {recipient} via {SMTP_HOST} …")
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.login(gmail_address, app_password)
        smtp.send_message(msg)

    print("✅ Triage summary sent.")


if __name__ == "__main__":
    main()
