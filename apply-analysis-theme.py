#!/usr/bin/env python3
"""Roll the shared front-page design system across the company analysis pages.

Every ``*-blog-page.html`` in the repo root is built from the same markup
vocabulary (hero-banner / article-container / meta-bar / summary-box /
verdict-box / h2-wrapper / table-label ...), so a single stylesheet plus a
single enhancement script restyles all of them consistently.

This script only ever INSERTS two marker-guarded blocks:

  * in <head>, after the article's own inline <style> so it wins on equal
    specificity:   DM Mono webfont + em-analysis-theme.css
  * before </body>:  em-analysis-embed.js

It never touches article content, figures, tables, charts or wording, and it
is idempotent - a page that already carries the markers is left alone.

Usage:
    python apply-analysis-theme.py            # inject into every analysis page
    python apply-analysis-theme.py --check    # exit 1 if any page is missing it
    python apply-analysis-theme.py --revert   # strip the blocks back out again
"""

import argparse
import glob
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
PATTERN = os.path.join(HERE, "*-blog-page.html")

CSS_START = "<!-- EM-THEME:START"
CSS_END = "<!-- EM-THEME:END -->"
JS_START = "<!-- EM-THEME-JS:START"
JS_END = "<!-- EM-THEME-JS:END -->"

HEAD_BLOCK = """<!-- EM-THEME:START - shared front-page design system; injected by apply-analysis-theme.py. Do not hand-edit. -->
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<link rel="stylesheet" href="em-analysis-theme.css">
<!-- EM-THEME:END -->
"""

BODY_BLOCK = """<!-- EM-THEME-JS:START - injected by apply-analysis-theme.py. Do not hand-edit. -->
<script src="em-analysis-embed.js" defer></script>
<!-- EM-THEME-JS:END -->
"""


def find_last(source, tag):
    """Index of the last occurrence of a closing tag, case-insensitively."""
    matches = list(re.finditer(re.escape(tag), source, re.IGNORECASE))
    return matches[-1].start() if matches else -1


def strip_block(source, start_marker, end_marker):
    start = source.find(start_marker)
    if start == -1:
        return source
    end = source.find(end_marker)
    if end == -1:
        return source
    end += len(end_marker)
    while end < len(source) and source[end] in "\r\n":
        end += 1
    return source[:start] + source[end:]


def inject(path):
    """Return (changed, reason) for one page."""
    with open(path, encoding="utf-8") as fh:
        source = fh.read()
    original = source

    if CSS_START not in source:
        head = find_last(source, "</head>")
        if head == -1:
            return False, "no </head>"
        source = source[:head] + HEAD_BLOCK + source[head:]

    if JS_START not in source:
        body = find_last(source, "</body>")
        if body == -1:
            return False, "no </body>"
        source = source[:body] + BODY_BLOCK + source[body:]

    if source == original:
        return False, "already themed"

    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(source)
    return True, "themed"


def revert(path):
    with open(path, encoding="utf-8") as fh:
        source = fh.read()
    updated = strip_block(source, CSS_START, CSS_END)
    updated = strip_block(updated, JS_START, JS_END)
    if updated == source:
        return False
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(updated)
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--revert", action="store_true")
    args = parser.parse_args()

    pages = sorted(glob.glob(PATTERN))
    if not pages:
        raise SystemExit("ERROR: no *-blog-page.html files found in %s" % HERE)

    if args.check:
        missing = []
        for path in pages:
            with open(path, encoding="utf-8") as fh:
                source = fh.read()
            if CSS_START not in source or JS_START not in source:
                missing.append(os.path.basename(path))
        if missing:
            print("Analysis pages missing the shared theme (%d):" % len(missing))
            for name in missing:
                print("  " + name)
            print("Run: python apply-analysis-theme.py")
            return 1
        print("All %d analysis pages carry the shared theme." % len(pages))
        return 0

    if args.revert:
        n = sum(1 for path in pages if revert(path))
        print("Reverted %d of %d analysis pages." % (n, len(pages)))
        return 0

    changed = 0
    for path in pages:
        ok, reason = inject(path)
        name = os.path.basename(path)
        if ok:
            changed += 1
            print("  themed   %s" % name)
        elif reason == "already themed":
            print("  skipped  %s (already themed)" % name)
        else:
            print("  WARNING  %s (%s)" % (name, reason))
    print("Done: %d of %d analysis pages updated." % (changed, len(pages)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
