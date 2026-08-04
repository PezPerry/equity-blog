---
name: transcripts
description: Pull new Plaud transcripts on request, keep them as plain transcripts (no summaries), categorise them, and merge same-category recordings into single pasteable blocks. Use when the user says "grab my new transcripts", "pull my transcripts", "download my Plaud notes", or similar. This is a manual, on-demand pull — never an automation.
---

# Plaud transcript pull & categorise

Mickey dictates voice notes on a Plaud recorder throughout the day. When asked
(and only when asked — nothing here runs automatically, and nothing is pushed
into Obsidian), pull the new transcripts, categorise them, and hand back plain
transcript blocks ready to paste into other Claude Code sessions.

## Steps

1. **List recordings** with the Plaud MCP tools (load via ToolSearch:
   `mcp__Plaud__list_files`, `mcp__Plaud__get_transcript`). Default scope is
   today's recordings; if the user names a date range or says "since last
   time", use that instead. If unsure whether older recordings were already
   processed, ask or say which ones you skipped.
2. **Fetch plain transcripts** — the default `transaction` block. Do NOT use
   the outline block and do NOT summarise; the deliverable is the transcript
   itself.
3. **Categorise each recording** by content, not by filename. Recurring
   categories:
   - **Mission Control / system build** — dashboard, task-system, and
     automation design instructions
   - **Website updates** — split by site if mixed: Equity & Markets Insight,
     Malling Markwell, Lactose Free, Botswana Mana
   - **Newsletter updates**
   - **Task capture** — dictated admin / purchase / banking / reminder /
     research / email tasks
   - **Emails & messages to draft**
   - **Fragments** — clips too short or cut off to be usable; list them so
     Mickey can re-record, but do nothing else with them
   Invent a sensible new category if a recording fits none of these.
4. **Merge same-category recordings into one block.** Multiple website-update
   recordings become a single pasteable update; ditto for each other
   category. Label each source recording inside the block with its start time.
5. **Output** everything in the final chat message as fenced plain-text blocks
   under category headings, so each block can be copied straight into another
   Claude Code session as one instruction.

## Transcript fidelity

Keep wording verbatim. The only permitted edits are obvious speech-to-text
mangles of known names — never rephrase content. Known proper nouns:
Malling Markwell (not "Morling/Mailing"), Base44 (not "Base64"), Plaud,
ShareScope, Deebot, EcoVac, Mounjaro, Emilie, Letty, Bexleyheath, Snodland,
Tunbridge Wells, Monkey Puzzle, T1 Roofing, Bluewater, Brent crude.

## Hard rules

- Plain transcripts only — no summaries, breakdowns, or bullet rewrites of
  what was said (the task-capture category may keep the user's own
  one-line-per-task dictation style).
- Never write transcripts into Obsidian or any other system automatically;
  the user pastes the blocks where they want them.
- Flag fragments and truncated clips rather than silently dropping them.
