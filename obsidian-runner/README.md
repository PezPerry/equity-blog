# Obsidian website-notes runner

Watches the Obsidian folders **General Website Updates**, **Website Updates** and
**Equity and Markets Insight**. When a new note shows up, the runner passes it to
Claude Code inside this repo. Claude makes the site change or writes the article
and commits it. The runner then pushes to `main`, and the existing
*Publish Article to Base44* workflow publishes any new article and rotates the
front page.

It runs on the PC that holds the vault. The vault lives on local disk, so a
cloud job can't see it.

## Setup (once, on the Windows PC)

1. You need Python 3, git with push access to this repo, and Claude Code
   (`claude`) logged in on the PC.
2. Copy `config.example.json` to `config.json` and set:
   - `vault_path`: the Obsidian vault folder (the one with `.obsidian` in it).
   - `repo_path`: a clone of equity-blog **used only by the runner**. It pulls
     and resets that checkout, so don't share it with your day-to-day work.
   - `watch_folders`: folder names in the vault. Give them as vault-relative
     paths, or as bare names found anywhere in the vault.
3. Check that it finds your notes:
   `python obsidian_runner.py --dry-run`
4. Choose what happens to the notes already in those folders:
   - **Recommended: let the first run go through them.** Claude checks each
     note against the site. Notes it finds already done are marked
     `already-done` and moved to `Resolved`. The rest are actioned like new
     notes.
   - To leave them all untouched and handle only notes added from now on, run
     `python obsidian_runner.py --mark-existing` first.
5. Try one pass by hand: `python obsidian_runner.py --once`
6. Schedule it: `powershell -ExecutionPolicy Bypass -File install-task.ps1`
   (every 10 minutes; add `-Minutes 5` to change it).

## Per-note controls (frontmatter)

After handling a note, the runner writes these fields into it, so the result
shows up in Obsidian:

| Field | Meaning |
|---|---|
| `website_status` | `done` (committed and pushed), `already-done` (the site already covers it), `no-change` (Claude didn't act; see the summary), `failed` (retried up to `max_attempts`) |
| `website_commits` | Short SHAs of the pushed commits |
| `website_summary` | Claude's one-line account of what it did |

Notes that end up `done` or `already-done` move into a `Resolved` subfolder of
their watched folder, which is created if missing. The runner never scans
`Resolved`, and a name clash gets a number added (`note 2.md`). Set
`resolved_folder` to `""` in `config.json` to keep notes where they are.
Obsidian only fixes `[[links]]` for moves made inside the app, so links to a
moved note from other notes may need re-pointing.

- Set `website_status: redo` to run a note again (for example, after editing it).
- Set `website_status: skip` to keep the runner off a note.

Notes are left alone until they have gone unchanged for `settle_seconds`
(2 minutes by default), so a note still being typed or synced isn't picked up
half-finished.

## Logs and state

- `runner.log`: every scan and outcome.
- `state.json`: which notes have been handled.

Both files are git-ignored, along with `config.json`.
