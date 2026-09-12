# Paperclip AI setup kit

Everything needed to move the Equity Markets blog onto
[Paperclip](https://github.com/paperclipai/paperclip) orchestration.

| File | Purpose |
|---|---|
| `PLAN.md` | The full scoped plan: architecture, phases, budgets, risks |
| `bootstrap.sh` | Preflight + checksum-verified install + onboarding |
| `agents/*.md` | Role briefs for the five-agent org |

## Quick start (on your local machine)

Open this repo in a **local** Claude Code session and say:

> Read `paperclip/PLAN.md`, then run `paperclip/bootstrap.sh --preflight`,
> fix anything it flags, run `--install`, and set up the org from
> `paperclip/agents/` per the plan. Set the Publisher approval gate first.

Approve the permission prompts as they appear. Budget ~15–30 minutes of
supervision.

**Note:** the role briefs are written tool-agnostically because Paperclip's
config format is still evolving; the local session should adapt them to
whatever the installed version expects (AGENTS.md files, dashboard forms,
etc.) without changing their content or boundaries.
