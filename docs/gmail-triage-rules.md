# Gmail Triage — operating rules

Canonical rule set for the triage on **mickeyperrytrading@gmail.com**.

This file exists because the rules previously lived only in chat history and in
the hourly run's own output. When the hourly job stopped, the rules were
effectively lost and the triage could not be restarted without reconstructing
them. Treat this file as the source of truth; the scheduled Routine's prompt is
generated from it.

## The two moves

Triage is a **routing** job, not just a flagging job. Every in-scope message
ends up in exactly one of two places, and in both cases it **leaves the inbox**:

| Classification | Star | Destination label | Remove `INBOX`? |
|---|---|---|---|
| **Noise** | yes | `INBOX/Resolved` (`Label_1`) | **yes** |
| **Substantive** — needed by the RNS system | no | `INBOX/To Process` (`Label_34`) | **yes** |

`INBOX/To Process` is the queue the Windows RNS ingest (`watch-ons.ps1` →
`process-articles.py`) consumes. It marks what it has taken with
`ADVFN Processed` (`Label_31`) and `INBOX/To Process/Processed` (`Label_35`).
Do **not** file substantive announcements straight to `INBOX/Resolved` — that
starves the blog's RNS feed.

> **Critical Gmail detail.** Adding `INBOX/Resolved` does *not* archive a
> message. `INBOX/Resolved` is a nested user label, not the system inbox.
> A message only leaves the inbox when the system `INBOX` label is
> **removed**. Star + label without removing `INBOX` is the exact failure
> that left mail visibly starred but still sitting in the inbox.

## Scope

**In scope:** ADVFN news alerts only — `news_alert@advfn.com`.

**Never touch** (leave in the inbox, unstarred, unlabelled):

- `marketalerts@alertshub.ft.com` — FT price/volume alerts
- `clientmanagement@cmcmarkets.co.uk` — CMC Daily Statement and account mail
- `no-reply@plaud.ai` — Plaud AutoFlow notifications

## Noise — star, then file to `INBOX/Resolved`

Regulatory and administrative chaff:

- Takeover Panel dealing disclosures: Form 8.3, Form 8.5, Form 8 (DD),
  Form 38.5a / 38.5b (EPT/RI and EPT/NON-RI), Irish equivalents
- US regulatory reposts: Form 6-K, Form 10-Q, Form 8-K, Form 4, Form 425,
  Schedule 13G/A
- PDMR notifications and "Dealing in Securities" / share-scheme dealings
- Broker and research notes (e.g. Edison)
- Broker, adviser and Nomad appointments
- Notice of results date / earnings call scheduling
- Product marketing, CSR and charity PR
- Secondary news-wire coverage and commentary articles
- Cross-ticker reposts and exact duplicates
- Shareholder class-action solicitations
- Index notices (e.g. FTSE Russell) and bare subjects with no announcement type

## Substantive — leave unstarred, file to `INBOX/To Process`

Anything price-forming or genuinely company-specific:

- Results: interim, quarterly, half-year, full-year
- Trading updates and profit warnings
- Acquisitions and disposals
- Dividends: declarations, currency exchange rates, key information
- Bond offerings, tender offers, capital raises
- Board and leadership changes; organisational restructuring
- Takeover Code substantive announcements (e.g. Rule 2.8 responses)
- Material operational and contract milestones

## Not covered by the rules — flag, do not guess

Leave in place and list under "Flagged for review" in the run summary:

- Auditor appointments / audit tender conclusions
- Net Asset Value (NAV) notices
- Investment trust portfolio disclosures (e.g. "Ten Largest Investments")

## Run summary

Each run writes a summary listing, in order: items starred as noise, items left
as substantive, and items flagged for review. Recipient:
`mickey.perry@mallingmarkwell.co.uk`.

The Gmail connector can create drafts but **cannot send**. A summary left as a
draft never arrives. Either read the summary in the session output, or use the
SMTP delivery job in `scripts/gmail_triage_summary.py` (PR #4) to actually
deliver it.

## Verification

A run is only complete when the inbox is genuinely clear. After acting, re-query
`in:inbox` and confirm no in-scope thread remains. Two traps:

1. **Thread vs message.** Gmail keeps a whole thread in the inbox if *any* one
   of its messages still carries `INBOX`. Long ADVFN threads accumulate new
   messages all day, so operate per-message, then confirm at thread level.
2. **Arrivals mid-run.** Mail landing while the run is in progress will be
   missed. The next hourly run picks it up; that is why the schedule matters
   more than any single pass.
