# Paperclip AI Orchestration Plan — Equity Markets Blog

**Goal:** move from ad-hoc Claude Code sessions (+ Obsidian for notes, Codex for
image generation and research) to a persistent, self-hosted orchestration layer
— [Paperclip](https://github.com/paperclipai/paperclip) — where a small "org"
of AI agents runs the blog's production pipeline continuously, with Mickey
approving anything that publishes.

**Status of this document:** scoped and ready to execute. The companion
`bootstrap.sh` and `agents/` role files in this folder are the execution kit —
a Claude Code session running locally on your machine can drive the whole
install (see "How the hands-off install works" below).

---

## 1. What Paperclip is (and isn't)

Paperclip (launched March 2026, open source, ~50k GitHub stars) is a
self-hosted control plane for AI agents:

- **Node.js server + React dashboard**, backed by PostgreSQL. No cloud
  account required; everything runs on your hardware.
- **Org chart:** agents get roles, titles, reporting lines, and permissions.
  Work is routed as **tickets**; agents delegate to each other through the
  chart.
- **Heartbeats:** agents wake on cron schedules or events (task assignment,
  @-mention). Session state persists across heartbeats.
- **Budgets:** each agent gets a monthly token budget and auto-pauses at
  100% — this is the main cost-control mechanism.
- **Adapters:** Claude Code, Codex, Cursor, Gemini CLI, OpenCode, Bash, and
  HTTP webhooks. Our stack maps cleanly: Claude Code for research/writing/
  publishing, Codex for cover images (as today).
- **Governance:** every mutating action is traced to an actor; human-approval
  gates can be placed on specific actions.

What it is *not*: it doesn't replace the agents themselves. Claude Code still
does the work; Paperclip decides **when agents run, what they work on, what
they may spend, and what needs your sign-off**. Your Obsidian vault and this
repo remain the sources of truth.

**Caveat to hold in mind:** the project is ~5 months old and moving fast.
Expect breaking changes between releases; pin the installed version and
upgrade deliberately, not automatically.

## 2. Current vs target architecture

**Today**
```
You → Claude Code (manual sessions) → equity-blog repo → publish
You → Codex (manual) → cover images + research
You → Obsidian → notes / drafts / watchlists
```

**Target**
```
                    ┌──────────────── Paperclip server (your machine) ────────────────┐
RNS feeds / earnings│  Editor-in-Chief (Claude Code) — triage, QA, assigns tickets    │
calendar / your     │    ├─ Research Analyst (Claude Code) — RNS watch, company deep- │
tickets ───────────▶│    │   dives → writes findings to Obsidian vault                │
                    │    ├─ Writer (Claude Code) — findings → blog-page HTML in house │
                    │    │   theme (em-analysis-theme.css, apply-analysis-theme.py)   │
                    │    ├─ Image Producer (Codex) — cover art per article            │
                    │    └─ Publisher (Claude Code) — indexes, front page, git push   │
                    │         └── ⛔ human-approval gate: nothing goes live without   │
                    │             Mickey approving the ticket in the dashboard        │
                    └──────────────────────────────────────────────────────────────────┘
```

## 3. Where it runs

Paperclip needs an always-on-ish host to be useful (heartbeats fire on
schedule). Options, in order of recommendation:

1. **Your main computer** — zero extra cost, simplest start. Heartbeats only
   fire while the machine is awake, which is actually fine for a pilot: the
   backlog drains when you're at the desk. Start here.
2. **A small always-on box later** (Mac mini / NUC / £5–10-per-month VPS) —
   move to this only once the pilot proves out and you want overnight RNS
   monitoring. The install is identical; Postgres data migrates with
   `pg_dump`.

Requirements: Node.js 20+, pnpm 9.15+, PostgreSQL, macOS or Linux (Windows
via WSL2). The managed installer handles most of this; `bootstrap.sh` in this
folder verifies prerequisites first.

## 4. How the hands-off install works

This cloud session cannot reach your computer — that's a hard isolation
boundary, and it's the right one. The way you get "Claude does everything" is
to run Claude Code **locally** and let it drive:

1. Install the Claude Code desktop app (or CLI) on your machine if not
   already there, and open this repo.
2. Check out this branch and say:
   *"Read paperclip/PLAN.md and run paperclip/bootstrap.sh — do the full
   Paperclip setup."*
3. Local Claude Code runs the preflight, the checksum-verified installer, the
   onboarding, connects the Claude Code and Codex adapters, and creates the
   org from the role files in `paperclip/agents/`. You approve the permission
   prompts as they come; everything else is automated.

Total hands-on time for you: roughly 15–30 minutes of watching and approving.

## 5. Phased rollout

### Phase 0 — Prerequisites (½ day, mostly automated)
- Run `bootstrap.sh --preflight` locally: checks OS, Node 20+, pnpm, disk,
  and that `claude` and `codex` CLIs are installed and authenticated.
- Decide the vault path: agents need read/write access to the Obsidian vault
  folder (it's just Markdown on disk — no plugin needed).

### Phase 1 — Install Paperclip (1 hour)
- `bootstrap.sh --install`: downloads `install.sh` from paperclip.ing,
  **verifies the SHA-256 checksum**, runs it, then `paperclipai onboard`.
- Confirm dashboard loads on localhost. Keep it loopback-only (do not expose
  to the network) until/unless it moves to a server, at which point enable
  authenticated mode.
- Pin the installed version in `paperclip/VERSION` in this repo.

### Phase 2 — Wire the adapters (1 hour)
- Connect the **Claude Code adapter** (uses your existing Claude
  subscription/API auth — no new spend category).
- Connect the **Codex adapter** for image generation, same as your current
  usage.
- Smoke-test each with a trivial ticket ("say hello, report your tools").

### Phase 3 — Create the org (1–2 hours)
- Create the five roles from `paperclip/agents/*.md` (they are written as
  role briefs; local Claude Code pastes/adapts them into whatever format the
  installed Paperclip version expects — the format is still evolving, so the
  briefs are deliberately tool-agnostic).
- Set budgets (see §7) and heartbeats:
  - Research Analyst: hourly on weekdays 07:00–17:00 UK (RNS windows).
  - Editor-in-Chief: every 2 hours, plus event wake on ticket assignment.
  - Writer / Image Producer / Publisher: event-driven only (woken by
    tickets), no cron.
- **Set the human-approval gate on the Publisher**: any git push / front-page
  rebuild requires Mickey's approval on the ticket. This is non-negotiable in
  the pilot.

### Phase 4 — Pilot: one workflow end-to-end (1–2 weeks)
Pick the highest-value, lowest-risk loop: **RNS → draft article**.
- Analyst watches the RNS feed for companies in `companies.json`, files a
  findings ticket per material announcement.
- Editor-in-Chief triages (skip / short note / full article), assigns Writer.
- Writer produces `{slug}-blog-page.html` in the house theme; Image Producer
  attaches a cover; Publisher stages everything and requests your approval.
- You approve or reject from the dashboard. **Nothing publishes without you.**
- Success criteria before Phase 5: ≥5 articles through the loop, zero
  unapproved pushes, token spend within budget, output quality you'd have
  shipped anyway.

### Phase 5 — Expand (ongoing)
- Add the earnings-calendar loop (pre-write scheduled results coverage).
- Add long-reads and the monthly financial-calendar page rebuild.
- Consider relaxing the approval gate to post-hoc review for low-risk,
  formulaic updates (e.g. `rns-articles.json` refreshes) — your call, later.
- Optionally migrate the server to an always-on box.

### Phase 6 — Operations
- Weekly: skim the audit log and token-spend dashboard (5 min).
- Monthly: review budgets; pinned-version upgrade if release notes warrant.
- Keep telemetry disabled (bootstrap sets the env var).

## 6. Org design summary

| Role | Adapter | Wakes | Budget guide* | Key outputs |
|---|---|---|---|---|
| Editor-in-Chief | Claude Code | 2-hourly + events | 15% | Triage decisions, QA notes |
| Research Analyst | Claude Code | Hourly, weekdays | 30% | Findings tickets, vault notes |
| Writer | Claude Code | Event | 30% | `*-blog-page.html` drafts |
| Image Producer | Codex | Event | 10% | `*-cover.png` |
| Publisher | Claude Code | Event, **gated** | 15% | Index updates, git pushes |

\* share of whatever monthly total you set; start small (see §7).

Full role briefs live in `paperclip/agents/`.

## 7. Costs

- **Software:** £0 (open source, self-hosted).
- **Tokens:** the real cost. Heartbeats are the danger — an hourly agent that
  finds nothing still burns a wake-up's worth of tokens. Mitigations are
  built into the role briefs (check-cheap-first patterns: "diff the RNS feed;
  if nothing new, end the heartbeat immediately") and enforced by Paperclip's
  auto-pause budgets. **Start with a deliberately low monthly cap** for the
  whole org in month one, observe, then raise it. Your existing Claude and
  Codex subscriptions may absorb much of this depending on plan limits.
- **Hardware:** £0 for the pilot (your machine). £5–10/month VPS or a
  one-off Mac mini/NUC only if Phase 5 justifies it.

## 8. Risks and mitigations

| Risk | Mitigation |
|---|---|
| Project is 5 months old; breaking changes | Pin version; manual upgrades only; the repo + vault remain source of truth, so Paperclip is disposable/re-installable at any time |
| Runaway token spend | Per-agent budgets with auto-pause; low first-month cap; check-cheap-first heartbeat patterns |
| Agent publishes something wrong | Human-approval gate on Publisher; git history makes every publish revertible |
| Curl-pipe installer | bootstrap.sh downloads and **verifies SHA-256** before executing |
| Dashboard exposed | Loopback-only in pilot; authenticated mode + firewall if moved to a server |
| Secrets leakage into prompts | Use Paperclip's scoped-secrets store; never put credentials in role briefs or the repo |
| Orchestrator down / machine asleep | Nothing breaks — tickets queue; the blog just doesn't update until it wakes |

## 9. Rollback

Everything Paperclip touches is git-tracked (this repo) or plain Markdown
(vault). Rollback = stop the Paperclip service and go back to manual Claude
Code sessions. No lock-in, no data migration needed in either direction.

## 10. Decision points reserved for Mickey

1. Monthly token cap for the org (needed at Phase 3).
2. Obsidian vault path to grant the agents (Phase 0).
3. When/whether to relax the publish approval gate (Phase 5, earliest).
4. When/whether to move to always-on hardware (Phase 5).
