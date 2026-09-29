#!/usr/bin/env python3
"""Pick up new Obsidian notes about the website and action them with Claude Code.

Runs on the machine that holds the Obsidian vault (the vault is local, not in the
cloud). Each run scans the watched vault folders ("General Website Updates",
"Website Updates", "Equity and Markets Insight" by default), and for every note it
has not handled before it:

  1. pulls the latest equity-blog main,
  2. hands the note to Claude Code headless (`claude -p`) inside the repo, which
     makes the change or writes the article and commits it,
  3. pushes the commit(s), which fires the existing Publish-to-Base44 workflow,
  4. writes the outcome back into the note's frontmatter, so it shows in Obsidian.

Notes are keyed by their path in the vault. A note is handled once; to run it
again set `website_status: redo` in its frontmatter, and to keep the runner off a
note set `website_status: skip`.

Usage:
    python obsidian_runner.py --once            # one scan, then exit (Task Scheduler)
    python obsidian_runner.py                   # keep polling every poll_seconds
    python obsidian_runner.py --dry-run         # list what would be processed
    python obsidian_runner.py --mark-existing   # treat every note there now as done,
                                                # so only notes added later are run
    python obsidian_runner.py --config PATH     # default: config.json beside this file

Standard library only, so it runs on a stock Windows Python.
"""

import argparse
import datetime
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_CONFIG = os.path.join(HERE, "config.json")

DEFAULTS = {
    "vault_path": "",
    "watch_folders": ["General Website Updates", "Website Updates", "Equity and Markets Insight"],
    "repo_path": os.path.dirname(HERE),
    "branch": "main",
    "poll_seconds": 300,
    # A note still being typed or synced keeps changing; wait until it has sat
    # untouched this long before acting on it.
    "settle_seconds": 120,
    "max_attempts": 3,
    # Notes that are done (actioned now, or found already done) move into this
    # subfolder of their watched folder. Set to "" to leave them where they are.
    "resolved_folder": "Resolved",
    "claude_command": "claude",
    "claude_args": [
        "--permission-mode", "acceptEdits",
        "--allowedTools", "Read", "Edit", "Write", "Glob", "Grep",
        "Bash(git add:*)", "Bash(git commit:*)", "Bash(git status:*)",
        "Bash(git diff:*)", "Bash(git log:*)", "Bash(python:*)", "Bash(python3:*)",
    ],
    "claude_timeout_seconds": 1800,
    "state_file": os.path.join(HERE, "state.json"),
    "log_file": os.path.join(HERE, "runner.log"),
    "lock_file": os.path.join(HERE, "runner.lock"),
}

PROMPT = """You are running unattended from the Obsidian website-notes runner.

A new note has appeared in Mickey's Obsidian vault, in the folder "{folder}".
Notes in that folder are about the Equity & Markets Insight website, whose source
is this repository (published to GitHub Pages, and to Base44 by
.github/workflows/publish-to-base44.yml when a new article HTML lands on main).

Read the note below and action it in this repository:
- If it asks for a change to the site (a fix, a tweak, a new section, a data
  update), make that change, following the conventions of the files you touch.
- If it is an article or article material, produce the article as a new HTML page
  in the repo root, matching the structure, meta tags (title, description,
  keywords, tile-* tags) and theme of the most recent comparable articles, with a
  cover image only if the note supplies one.
- Leave anything the note does not ask for alone.

Commit your work with a clear message that names the note. Do NOT push; the
runner pushes. Do not ask questions: nobody is watching. If the note cannot be
actioned safely (unclear, needs information you do not have, needs credentials,
or is not about this website), make no commit.

The note may be older than the current site. If everything it asks for is
already in place in this repository, make no commit and report it as
already-done.

Finish your reply with exactly two lines:
RUNNER-OUTCOME: changed | already-done | not-actioned
RUNNER-SUMMARY: one sentence on what you did, what already covers the note, or
why you did nothing.

--- NOTE: {name} ---
{body}
--- END NOTE ---
"""

FRONTMATTER_KEYS = ("website_status", "website_processed", "website_commits", "website_summary")


# ---------------------------------------------------------------- utilities

def now_iso():
    return datetime.datetime.now().replace(microsecond=0).isoformat()


class Logger:
    def __init__(self, path):
        self.path = path

    def __call__(self, msg):
        line = "%s  %s" % (now_iso(), msg)
        print(line, flush=True)
        try:
            with open(self.path, "a", encoding="utf-8") as fh:
                fh.write(line + "\n")
        except OSError:
            pass


def load_config(path):
    cfg = dict(DEFAULTS)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as fh:
            cfg.update(json.load(fh))
    if not cfg["vault_path"]:
        sys.exit("vault_path is not set. Copy config.example.json to config.json and fill it in.")
    for key in ("vault_path", "repo_path", "state_file", "log_file", "lock_file"):
        cfg[key] = os.path.abspath(os.path.expanduser(os.path.expandvars(cfg[key])))
    return cfg


def load_state(path):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return {"notes": {}}


def save_state(path, state):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump(state, fh, indent=2, sort_keys=True)
    os.replace(tmp, path)


def acquire_lock(path, log):
    """Stop a Task Scheduler run from overlapping a long Claude job."""
    if os.path.exists(path):
        age = time.time() - os.path.getmtime(path)
        if age < 4 * 3600:
            log("Another run holds %s (%.0f min old); exiting." % (path, age / 60))
            return False
        log("Removing stale lock (%.1f h old)." % (age / 3600))
    with open(path, "w") as fh:
        fh.write(str(os.getpid()))
    return True


# ---------------------------------------------------------------- frontmatter

FM_RE = re.compile(r"\A---\r?\n(.*?)\r?\n---\r?\n?", re.S)


def read_frontmatter(text):
    m = FM_RE.match(text)
    if not m:
        return {}
    out = {}
    for line in m.group(1).splitlines():
        if ":" in line and not line.startswith((" ", "\t", "-")):
            key, val = line.split(":", 1)
            out[key.strip()] = val.strip().strip("'\"")
    return out


def strip_runner_keys(text):
    """The note body as the author wrote it, without the runner's own fields."""
    m = FM_RE.match(text)
    if not m:
        return text
    kept = [l for l in m.group(1).splitlines() if l.split(":", 1)[0].strip() not in FRONTMATTER_KEYS]
    rest = text[m.end():]
    return ("---\n%s\n---\n%s" % ("\n".join(kept), rest)) if kept else rest


def write_frontmatter(path, updates):
    """Set top-level scalar keys in a note's YAML frontmatter, creating it if absent."""
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    rendered = {k: json.dumps(str(v), ensure_ascii=False) for k, v in updates.items()}
    m = FM_RE.match(text)
    if m:
        lines = m.group(1).splitlines()
        seen = set()
        for i, line in enumerate(lines):
            key = line.split(":", 1)[0].strip()
            if key in rendered and not line.startswith((" ", "\t")):
                lines[i] = "%s: %s" % (key, rendered[key])
                seen.add(key)
        lines += ["%s: %s" % (k, v) for k, v in rendered.items() if k not in seen]
        text = "---\n%s\n---\n%s" % ("\n".join(lines), text[m.end():])
    else:
        block = "\n".join("%s: %s" % kv for kv in rendered.items())
        text = "---\n%s\n---\n%s" % (block, text)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)


# ---------------------------------------------------------------- discovery

def resolve_folders(cfg, log):
    """Map each watched folder name to a directory in the vault.

    A name is tried as a path relative to the vault root first; failing that, the
    first directory anywhere in the vault with that exact name is used, so the
    folders can sit inside a parent folder without extra config.
    """
    vault = cfg["vault_path"]
    found = {}
    for name in cfg["watch_folders"]:
        direct = os.path.join(vault, name)
        if os.path.isdir(direct):
            found[name] = direct
            continue
        for root, dirs, _ in os.walk(vault):
            dirs[:] = [d for d in dirs if not d.startswith(".")]
            if os.path.basename(root) == os.path.basename(name) and root != vault:
                found[name] = root
                break
        else:
            log("Watched folder not found in vault: %s" % name)
    return found


def candidate_notes(cfg, folders):
    for label, directory in folders.items():
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith(".") and d != cfg["resolved_folder"]]
            for fn in sorted(files):
                if fn.lower().endswith(".md"):
                    path = os.path.join(root, fn)
                    yield label, path, os.path.relpath(path, cfg["vault_path"]).replace("\\", "/")


def move_to_resolved(cfg, watched_dir, path):
    """Move a finished note into the Resolved subfolder of its watched folder.

    Obsidian only rewrites [[links]] for moves made inside the app, so links to a
    moved note from elsewhere in the vault will need Obsidian's link repair.
    """
    target_dir = os.path.join(watched_dir, cfg["resolved_folder"])
    os.makedirs(target_dir, exist_ok=True)
    stem, ext = os.path.splitext(os.path.basename(path))
    target, n = os.path.join(target_dir, stem + ext), 1
    while os.path.exists(target):
        n += 1
        target = os.path.join(target_dir, "%s %d%s" % (stem, n, ext))
    os.replace(path, target)
    return target


def note_hash(body):
    return hashlib.sha256(body.encode("utf-8")).hexdigest()[:16]


# ---------------------------------------------------------------- git + claude

def git(cfg, *args, check=True):
    res = subprocess.run(["git"] + list(args), cwd=cfg["repo_path"], capture_output=True,
                         text=True, encoding="utf-8", errors="replace")
    if check and res.returncode != 0:
        raise RuntimeError("git %s failed: %s" % (" ".join(args), (res.stderr or res.stdout).strip()))
    return res.stdout.strip()


def prepare_repo(cfg):
    if git(cfg, "status", "--porcelain"):
        raise RuntimeError("repo has uncommitted changes; the runner will not work on a dirty checkout")
    git(cfg, "checkout", cfg["branch"])
    git(cfg, "pull", "--rebase", "origin", cfg["branch"])
    return git(cfg, "rev-parse", "HEAD")


def push_with_retry(cfg, log):
    # The RNS/ONS pipeline pushes to main often, so a plain push can lose a race.
    for attempt in range(1, 6):
        res = subprocess.run(["git", "push", "origin", cfg["branch"]], cwd=cfg["repo_path"],
                             capture_output=True, text=True)
        if res.returncode == 0:
            return
        log("  push rejected (attempt %d); rebasing and retrying" % attempt)
        git(cfg, "pull", "--rebase", "origin", cfg["branch"])
        time.sleep(2 ** attempt)
    raise RuntimeError("could not push after 5 attempts")


def run_claude(cfg, folder, name, body):
    prompt = PROMPT.format(folder=folder, name=name, body=body)
    # which() finds the claude.cmd shim npm installs on Windows, so no shell is
    # needed (and cmd.exe never sees the parentheses in the tool patterns).
    exe = shutil.which(cfg["claude_command"]) or cfg["claude_command"]
    cmd = [exe, "-p", "--output-format", "json"] + cfg["claude_args"]
    res = subprocess.run(cmd, input=prompt, cwd=cfg["repo_path"], capture_output=True,
                         text=True, encoding="utf-8", errors="replace",
                         timeout=cfg["claude_timeout_seconds"])
    if res.returncode != 0:
        raise RuntimeError("claude exited %d: %s" % (res.returncode, (res.stderr or res.stdout)[-800:]))
    try:
        text = json.loads(res.stdout).get("result", "")
    except ValueError:
        text = res.stdout
    m = re.search(r"RUNNER-SUMMARY:\s*(.+)", text)
    summary = (m.group(1).strip() if m else text.strip().splitlines()[-1] if text.strip() else "")[:400]
    m = re.search(r"RUNNER-OUTCOME:\s*([a-z-]+)", text)
    return (m.group(1) if m else ""), summary


def process_note(cfg, log, folder, path, rel, body):
    before = prepare_repo(cfg)
    outcome, summary = run_claude(cfg, folder, os.path.basename(path), body)
    after = git(cfg, "rev-parse", "HEAD")
    if git(cfg, "status", "--porcelain"):
        # Claude left edits uncommitted; don't let them leak into the next note.
        git(cfg, "reset", "--hard", after)
        git(cfg, "clean", "-fd")
        log("  discarded uncommitted edits Claude left behind")
    if after == before:
        return ("already-done" if outcome == "already-done" else "no-change"), [], summary
    commits = git(cfg, "rev-list", "--reverse", "%s..%s" % (before, after)).split()
    push_with_retry(cfg, log)
    return "done", [c[:7] for c in commits], summary


# ---------------------------------------------------------------- main loop

def scan(cfg, log, dry_run=False, mark_existing=False):
    state = load_state(cfg["state_file"])
    notes = state.setdefault("notes", {})
    folders = resolve_folders(cfg, log)
    for label, path, rel in candidate_notes(cfg, folders):
        try:
            with open(path, encoding="utf-8") as fh:
                raw = fh.read()
        except (OSError, UnicodeDecodeError) as exc:
            log("Cannot read %s: %s" % (rel, exc))
            continue
        fm_status = read_frontmatter(raw).get("website_status", "").lower()
        body = strip_runner_keys(raw)
        entry = notes.get(rel)

        if fm_status == "skip":
            continue
        if fm_status != "redo":
            if entry and entry["status"] in ("done", "already-done", "no-change", "baseline"):
                continue
            if entry and entry["status"] == "failed" and entry.get("attempts", 0) >= cfg["max_attempts"]:
                continue
        if not body.strip():
            continue
        if time.time() - os.path.getmtime(path) < cfg["settle_seconds"]:
            log("Waiting for %s to settle" % rel)
            continue

        if mark_existing:
            notes[rel] = {"status": "baseline", "hash": note_hash(body), "at": now_iso()}
            log("Baseline: %s" % rel)
            continue
        if dry_run:
            log("Would process [%s] %s" % (label, rel))
            continue

        log("Processing [%s] %s" % (label, rel))
        attempts = (entry or {}).get("attempts", 0) + 1 if fm_status != "redo" else 1
        try:
            status, commits, summary = process_note(cfg, log, label, path, rel, body)
        except Exception as exc:  # keep going with the other notes
            status, commits, summary = "failed", [], str(exc)[:400]
        log("  -> %s %s %s" % (status, " ".join(commits), summary))

        notes[rel] = {"status": status, "hash": note_hash(body), "at": now_iso(),
                      "commits": commits, "summary": summary, "attempts": attempts}
        save_state(cfg["state_file"], state)
        try:
            write_frontmatter(path, {
                "website_status": status,
                "website_processed": now_iso(),
                "website_commits": " ".join(commits),
                "website_summary": summary,
            })
        except OSError as exc:
            log("  could not update note frontmatter: %s" % exc)

        if status in ("done", "already-done") and cfg["resolved_folder"]:
            try:
                new_path = move_to_resolved(cfg, folders[label], path)
            except OSError as exc:
                log("  could not move note to %s: %s" % (cfg["resolved_folder"], exc))
            else:
                new_rel = os.path.relpath(new_path, cfg["vault_path"]).replace("\\", "/")
                notes[new_rel] = notes.pop(rel)
                save_state(cfg["state_file"], state)
                log("  moved to %s" % new_rel)

    save_state(cfg["state_file"], state)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--config", default=DEFAULT_CONFIG)
    ap.add_argument("--once", action="store_true", help="scan once and exit")
    ap.add_argument("--dry-run", action="store_true", help="list notes that would be processed")
    ap.add_argument("--mark-existing", action="store_true",
                    help="record every current note as already handled")
    args = ap.parse_args()

    cfg = load_config(args.config)
    log = Logger(cfg["log_file"])
    if not os.path.isdir(cfg["vault_path"]):
        sys.exit("vault_path does not exist: %s" % cfg["vault_path"])

    if args.dry_run or args.mark_existing:
        scan(cfg, log, dry_run=args.dry_run, mark_existing=args.mark_existing)
        return
    if not acquire_lock(cfg["lock_file"], log):
        return
    try:
        while True:
            os.utime(cfg["lock_file"])  # keep a long-running poller's lock fresh
            scan(cfg, log)
            if args.once:
                break
            time.sleep(cfg["poll_seconds"])
    finally:
        try:
            os.remove(cfg["lock_file"])
        except OSError:
            pass


if __name__ == "__main__":
    main()
