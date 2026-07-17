#!/usr/bin/env python3
"""Rebuild the front-page tiles from front-page-tiles.json.

The front page shows the 10 newest entries in the registry: the newest entry with
kind "analysis" takes the hero, the next 9 (any kind) fill the bento slots in date
order, and anything older drops off. Ties on date are broken by registry order.

Usage:
    python build-front-page.py                       # rebuild from the registry
    python build-front-page.py --add ARTICLE.html --date YYYY-MM-DD
                                                     # read an article's tile-* meta
                                                     # tags into the registry, then rebuild
    python build-front-page.py --check               # verify front-page.html is up to date
"""

import argparse
import html
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REGISTRY = os.path.join(HERE, "front-page-tiles.json")
FRONT_PAGE = os.path.join(HERE, "front-page.html")

HERO_START = "<!-- HERO:START"
HERO_END = "<!-- HERO:END -->"
TILES_START = "<!-- TILES:START"
TILES_END = "<!-- TILES:END -->"

VALID_PILLS = {"buy", "hold", "sell", "noact", "watch", "oval"}


def esc(text):
    return html.escape(str(text), quote=True)


def load_registry():
    with open(REGISTRY, encoding="utf-8") as fh:
        data = json.load(fh)
    slots = data["_slots"]
    tiles = data["tiles"]
    for tile in tiles:
        pill = tile.get("pill", "watch")
        if pill not in VALID_PILLS:
            raise SystemExit(
                "ERROR: tile '%s' uses pill '%s', which has no CSS rule. Valid: %s"
                % (tile["slug"], pill, ", ".join(sorted(VALID_PILLS)))
            )
        if tile.get("kind", "analysis") not in ("analysis", "companion"):
            raise SystemExit(
                "ERROR: tile '%s' has kind '%s'; expected 'analysis' or 'companion'."
                % (tile["slug"], tile["kind"])
            )
    return data, slots, tiles


def select(tiles, slots):
    """Return (hero, [slot tiles]) newest-first. Stable sort keeps registry order on ties."""
    ordered = sorted(tiles, key=lambda t: t["date"], reverse=True)

    hero = next((t for t in ordered if t.get("kind", "analysis") == "analysis"), None)
    if hero is None:
        raise SystemExit("ERROR: no tile with kind 'analysis' — the hero cannot be filled.")

    rest = [t for t in ordered if t is not hero]
    chosen = rest[: len(slots)]
    if len(chosen) < len(slots):
        raise SystemExit(
            "ERROR: registry has %d tiles besides the hero but the grid needs %d. "
            "Short-filling would leave holes in the mosaic." % (len(chosen), len(slots))
        )
    dropped = rest[len(slots):]
    return hero, chosen, dropped


def render_hero(tile):
    bits = []
    if tile.get("ticker"):
        bits.append("<span>%s</span>" % esc(tile["ticker"]))
    for extra in tile.get("meta_extra", []):
        bits.append("<span>%s</span>" % esc(extra))
    if tile.get("pill_label"):
        if bits:
            bits.append("<span>&middot;</span>")
        bits.append('<span style="color:var(--accent)">%s</span>' % esc(tile["pill_label"]))
    meta = "".join(bits)

    return """    <a class="latest" data-tilt data-slug="{slug}" href="{url}">
      <div class="latest-media">
        <span class="latest-ribbon"><span class="dot"></span>Latest</span>
        <img src="{cover}" alt="{alt}">
      </div>
      <div class="latest-body">
        <span class="latest-tag">{cat}</span>
        <h2>{headline}</h2>
        <div class="latest-meta">{meta}</div>
      </div>
    </a>""".format(
        slug=esc(tile["slug"]),
        url=esc(tile["url"]),
        cover=esc(tile["cover"]),
        alt=esc(tile["alt"]),
        cat=esc(tile["cat"]),
        headline=esc(tile["headline"]),
        meta=meta,
    )


def render_block(tile, slot):
    bits = ['<span class="pill %s">%s</span>' % (esc(tile.get("pill", "watch")),
                                                 esc(tile.get("pill_label", "Analysis")))]
    for extra in tile.get("meta_extra", []):
        bits.append("<span>%s</span>" % esc(extra))
    if tile.get("ticker") and not tile.get("meta_extra"):
        bits.append("<span>%s</span>" % esc(tile["ticker"]))
    meta = "".join(bits)

    return """    <a class="block a-{slot}" data-tilt data-slug="{slug}" href="{url}">
      <img class="photo" alt="{alt}" src="{cover}">
      <div class="overlay"></div><div class="spotlight"></div>
      <div class="content">
        <span class="cat"><i class="swatch"></i>{cat}</span>
        <h4>{headline}</h4>
        <div class="meta">{meta}</div>
      </div>
    </a>""".format(
        slot=esc(slot),
        slug=esc(tile["slug"]),
        url=esc(tile["url"]),
        alt=esc(tile["alt"]),
        cover=esc(tile["cover"]),
        cat=esc(tile["cat"]),
        headline=esc(tile["headline"]),
        meta=meta,
    )


def splice(source, start_marker, end_marker, body):
    start = source.find(start_marker)
    end = source.find(end_marker)
    if start == -1 or end == -1:
        raise SystemExit("ERROR: markers %s / %s not found in front-page.html"
                         % (start_marker, end_marker))
    line_end = source.find("\n", start)
    # Resume from the start of the end-marker's line so its indentation survives.
    end_line_start = source.rfind("\n", 0, end) + 1
    return source[: line_end + 1] + body + "\n" + source[end_line_start:]


def build():
    _, slots, tiles = load_registry()
    hero, chosen, dropped = select(tiles, slots)

    with open(FRONT_PAGE, encoding="utf-8") as fh:
        page = fh.read()

    page = splice(page, HERO_START, HERO_END, render_hero(hero))
    blocks = "\n\n".join(render_block(t, s) for t, s in zip(chosen, slots))
    page = splice(page, TILES_START, TILES_END, "\n" + blocks + "\n")

    return page, hero, chosen, dropped


def report(hero, chosen, dropped, slots):
    print("  hero  <- %s  (%s)" % (hero["slug"], hero["date"]))
    for tile, slot in zip(chosen, slots):
        print("  %-5s <- %s  (%s)" % (slot, tile["slug"], tile["date"]))
    for tile in dropped:
        print("  DROPPED off the front page: %s  (%s)" % (tile["slug"], tile["date"]))


META_MAP = {
    "tile-kind": "kind",
    "tile-cat": "cat",
    "tile-headline": "headline",
    "tile-pill": "pill",
    "tile-pill-label": "pill_label",
    "tile-ticker": "ticker",
}


def add_article(path, date):
    with open(path, encoding="utf-8") as fh:
        source = fh.read()

    def meta(name):
        match = re.search(r'<meta\s+name="%s"\s+content="([^"]*)"' % re.escape(name), source)
        return match.group(1) if match else None

    slug = os.path.splitext(os.path.basename(path))[0]

    # A hand-curated entry always wins: if the slug is already registered, leave it
    # alone. Only genuinely new articles get a tile derived from their meta tags.
    data, _, tiles = load_registry()
    if any(t["slug"] == slug for t in tiles):
        print("  registry: '%s' is already registered - keeping the existing entry." % slug)
        return

    title = re.search(r"<title>([^|<]+)", source)

    tile = {
        "slug": slug,
        "url": os.path.basename(path),
        "date": date,
        "kind": meta("tile-kind") or "analysis",
        "cover": meta("tile-cover") or "%s-cover.png" % slug,
        "alt": meta("tile-alt") or (title.group(1).strip() if title else slug),
        "cat": meta("tile-cat") or "Analysis",
        "headline": meta("tile-headline") or (title.group(1).strip() if title else slug),
        "pill": meta("tile-pill") or "watch",
        "pill_label": meta("tile-pill-label") or "Analysis",
    }
    if meta("tile-ticker"):
        tile["ticker"] = meta("tile-ticker")
    if meta("tile-meta"):
        tile["meta_extra"] = [meta("tile-meta")]

    cover_path = os.path.join(HERE, tile["cover"])
    if not os.path.exists(cover_path):
        print("  WARNING: cover '%s' is not in the repo. The tile will show a broken image."
              % tile["cover"])

    data["tiles"].insert(0, tile)
    with open(REGISTRY, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
        fh.write("\n")
    print("  registry <- %s (%s, %s)" % (slug, date, tile["kind"]))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--add", metavar="ARTICLE.html")
    parser.add_argument("--date", metavar="YYYY-MM-DD")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    if args.add:
        if not args.date:
            raise SystemExit("ERROR: --add requires --date YYYY-MM-DD")
        add_article(args.add, args.date)

    page, hero, chosen, dropped = build()
    _, slots, _ = load_registry()

    if args.check:
        with open(FRONT_PAGE, encoding="utf-8") as fh:
            current = fh.read()
        if current != page:
            print("front-page.html is OUT OF DATE. Run: python build-front-page.py")
            return 1
        print("front-page.html is up to date.")
        return 0

    with open(FRONT_PAGE, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(page)
    print("front-page.html rebuilt:")
    report(hero, chosen, dropped, slots)
    return 0


if __name__ == "__main__":
    sys.exit(main())
