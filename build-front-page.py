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
import datetime
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


MONTHS = ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")


def nice_date(iso):
    """'2026-09-15' -> '15 Sep 2026' for the tile caption's date stamp.

    The month names are spelled out here rather than taken from strftime so the
    output cannot drift with the locale of whatever account runs the build. A
    date the registry cannot parse is passed through untouched: a caption that
    reads oddly is better than a crashed nightly publish.
    """
    try:
        day = datetime.date.fromisoformat(str(iso))
    except ValueError:
        return iso
    return "%d %s %d" % (day.day, MONTHS[day.month - 1], day.year)


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


SITE = "https://equityandmarketsinsight.com"


def branded_href(tile):
    """The branded-site route a tile opens, never the raw Pages file.

    The front page is iframed as the branded site's Home, so a relative link to
    the article file would open the full article inside that frame and skip the
    signed-out article meter (fixed 2 Nov 2026). Every tile therefore opens the
    metered route at top level:
      longreads/<stem>.html -> /LongReads/<stem>  (LongReadArticle, metered)
      <stem>.html           -> /blog/<stem>       (BlogPost, metered)
    The BlogPost/LongRead slug is the filename stem for every registered tile.
    A tile whose record has a different slug (or lives on another route, e.g.
    /TradeDetail?slug=...) can carry an explicit "href" in the registry.
    """
    if tile.get("href"):
        href = tile["href"]
        return href if href.startswith("http") else SITE + "/" + href.lstrip("/")
    url = tile["url"]
    stem = os.path.splitext(os.path.basename(url))[0]
    if url.startswith("longreads/"):
        return "%s/LongReads/%s" % (SITE, stem)
    return "%s/blog/%s" % (SITE, stem)


def link_target(tile):
    """Every tile breaks out of the Home iframe so the reader lands on the branded,
    metered route with the address bar showing it."""
    del tile
    return ' target="_top"'


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

    return """    <a class="latest" data-tilt data-slug="{slug}" href="{url}"{target}>
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
        url=esc(branded_href(tile)),
        target=link_target(tile),
        cover=esc(tile["cover"]),
        alt=esc(tile["alt"]),
        cat=esc(tile["cat"]),
        headline=esc(tile["headline"]),
        meta=meta,
    )


def render_block(tile, slot):
    """One front-page tile: a 3:2 logo plate with the caption band below it.

    `slot` is no longer used for layout - the grid is a uniform three columns,
    so there are no named grid areas to key a class to - but it stays in the
    signature because build() still pairs each tile with its registry slot.

    Inside .plate the same cover is drawn twice: .bg is a blurred, cropped copy
    that supplies the plate's colour field, and .photo is the real cover with
    object-fit:contain over it, so a logo is never cropped. No text is drawn
    over the image.
    """
    del slot  # layout no longer depends on the slot name

    bits = ['<span class="pill %s">%s</span>' % (esc(tile.get("pill", "watch")),
                                                 esc(tile.get("pill_label", "Analysis")))]
    for extra in tile.get("meta_extra", []):
        bits.append("<span>%s</span>" % esc(extra))
    if tile.get("ticker") and not tile.get("meta_extra"):
        bits.append("<span>%s</span>" % esc(tile["ticker"]))
    bits.append("<span>%s</span>" % esc(nice_date(tile["date"])))
    meta = "".join(bits)

    return """    <a class="block" data-slug="{slug}" href="{url}"{target}>
      <div class="plate">
        <img class="bg" src="{cover}" alt="" aria-hidden="true">
        <img class="photo" alt="{alt}" src="{cover}">
      </div>
      <div class="caption">
        <span class="cat"><i class="swatch"></i>{cat}</span>
        <h4>{headline}</h4>
        <div class="meta">{meta}</div>
      </div>
    </a>""".format(
        slug=esc(tile["slug"]),
        url=esc(branded_href(tile)),
        target=link_target(tile),
        alt=esc(tile["alt"]),
        cover=esc(tile["cover"]),
        cat=esc(tile["cat"]),
        headline=esc(tile["headline"]),
        meta=meta,
    )


# --------------------------------------------------------------------------- redesign (Nov 2026)
# The redesigned front page has no hero or bento grid. It carries three generated
# blocks instead - ASSESSMENTS, LONGREADS and GREATS - and build() fills whichever
# markers the target page contains, so the same script serves both layouts.

ASSESS_START, ASSESS_END = "<!-- ASSESSMENTS:START", "<!-- ASSESSMENTS:END -->"
LONG_START, LONG_END = "<!-- LONGREADS:START", "<!-- LONGREADS:END -->"
GREATS_START, GREATS_END = "<!-- GREATS:START", "<!-- GREATS:END -->"
GREATS_REGISTRY = os.path.join(HERE, "greats", "greats.json")

VERDICT_PILLS = {"buy", "hold", "sell", "noact", "oval"}
HEADLINE_RE = re.compile(
    r"^(?P<company>.+?)\s*\((?P<ticker>[A-Z0-9.]{1,6})\)\s*[—–-]\s*"
    r"(?P<verdict>[A-Za-z ]{2,20}?):\s*(?P<title>.+)$")
BUY_WORDS = {"buy", "strong buy"}
SELL_WORDS = {"sell", "short", "avoid", "overvalued", "reduce"}
HOLD_WORDS = {"hold", "fair value", "no action", "pass"}


def assessment(tile):
    """Return the tile's verdict details if it is a company assessment, else None.

    Newer articles carry everything in the headline - "Company (TICK) - Verdict:
    Title". Older ones carry the verdict in the pill and the ticker in its own
    field or in the filename (bp-BP-blog-page.html).
    """
    if tile.get("kind", "analysis") != "analysis" or tile["url"].startswith("longreads/"):
        return None
    match = HEADLINE_RE.match(tile["headline"])
    if match:
        return {"ticker": match.group("ticker"), "verdict": match.group("verdict").strip(),
                "title": match.group("title").strip()}
    in_name = re.search(r"-([A-Z0-9]{2,5})-", tile["url"])
    has_verdict = tile.get("pill") in VERDICT_PILLS
    if has_verdict or (in_name and tile["url"].endswith("blog-page.html")):
        return {"ticker": tile.get("ticker") or (in_name.group(1) if in_name else ""),
                "verdict": tile.get("pill_label") if has_verdict else "",
                "title": tile["headline"]}
    return None


def verdict_class(verdict):
    word = (verdict or "").lower()
    if word in BUY_WORDS:
        return "buy"
    if word in SELL_WORDS:
        return "sell"
    if word in HOLD_WORDS:
        return "hold"
    return ""


def article_description(tile):
    """The article's own meta description, for the featured card. Never invented."""
    path = os.path.join(HERE, tile["url"])
    try:
        with open(path, encoding="utf-8") as fh:
            head = fh.read(6000)
    except OSError:
        return ""
    match = re.search(r'<meta\s+name="description"\s+content="([^"]*)"', head)
    text = html.unescape(match.group(1)).strip() if match else ""
    if len(text) > 230:
        text = text[:230].rsplit(" ", 1)[0].rstrip(",;:") + "…"
    return text


ASSESSMENT_COUNT = 6   # one featured, four in the grid, a sixth for the phone row


def render_assessments(tiles):
    ordered = sorted(tiles, key=lambda t: t["date"], reverse=True)
    picked = [(t, assessment(t)) for t in ordered]
    picked = [(t, a) for t, a in picked if a][:ASSESSMENT_COUNT]
    if len(picked) < 5:
        raise SystemExit("ERROR: fewer than five company assessments in the registry.")
    cards = []
    for index, (tile, info) in enumerate(picked):
        classes = "ca tile" + (" feat" if index == 0 else "") + (" m-only" if index >= 5 else "")
        if info["verdict"]:
            tag = '<span class="vtag %s">%s</span>' % (verdict_class(info["verdict"]), esc(info["verdict"]))
        else:
            tag = '<span class="vtag">Analysis</span>'
        ticker = '<span class="tk">%s</span>' % esc(info["ticker"]) if info["ticker"] else ""
        extract = article_description(tile) if index == 0 else ""
        cards.append(
            '      <a class="{cls}" data-slug="{slug}" href="{url}" target="_top">\n'
            '        <div class="img"><img src="{cover}" alt="{alt}" loading="lazy">'
            '<span class="ribbon">Latest</span></div>\n'
            '        <div class="body"><div class="line">{tag}{ticker}</div><h3>{title}</h3>'
            '{extract}<span class="dt">{date}</span></div>\n'
            '      </a>'.format(
                cls=classes, slug=esc(tile["slug"]), url=esc(branded_href(tile)), cover=esc(tile["cover"]),
                alt=esc(tile["alt"]), tag=tag, ticker=ticker, title=esc(info["title"]),
                extract=('<p class="ex">%s</p>' % esc(extract)) if extract else "",
                date=esc(nice_date(tile["date"]))))
    return "\n".join(cards), [t for t, _ in picked]


LONGREAD_COUNT = 5


def longread_files():
    """Long reads that live in longreads/ but were never registered as tiles."""
    folder = os.path.join(HERE, "longreads")
    found = []
    if not os.path.isdir(folder):
        return found
    for name in sorted(os.listdir(folder)):
        match = re.match(r"(\d{4})-(\d{2})-.+\.html$", name)
        if not match:
            continue
        with open(os.path.join(folder, name), encoding="utf-8") as fh:
            head = fh.read(8000)
        title = re.search(r"<title>([^<]+)", head)
        heading = html.unescape(title.group(1)).strip() if title else name
        heading = re.split(r"\s+[—–|]\s+", heading)[0].strip()
        image = re.search(r'<meta\s+property="og:image"\s+content="[^"]*?/(longreads/images/[^"]+)"', head)
        cover = image.group(1) if image and os.path.exists(os.path.join(HERE, image.group(1))) else ""
        cat = re.search(r'<meta\s+name="tile-cat"\s+content="([^"]*)"', head)
        found.append({
            "slug": name[:-5], "url": "longreads/" + name, "date": "%s-%s-01" % match.group(1, 2),
            "month_only": True, "headline": heading, "cover": cover,
            "cat": html.unescape(cat.group(1)) if cat else "Long Read",
        })
    return found


def render_longreads(tiles):
    items = [t for t in tiles if t.get("kind", "analysis") == "analysis" and not assessment(t)]
    known = {t["url"] for t in items}
    items += [f for f in longread_files() if f["url"] not in known]
    items = sorted(items, key=lambda t: t["date"], reverse=True)[:LONGREAD_COUNT]
    cards = []
    for item in items:
        topic = item["cat"].split("·")[0].strip()
        if item.get("month_only"):
            day = datetime.date.fromisoformat(item["date"])
            when = "%s %d" % (MONTHS[day.month - 1], day.year)
        else:
            when = nice_date(item["date"])
        if item.get("cover"):
            disc = '<img src="%s" alt="" loading="lazy">' % esc(item["cover"])
        else:
            disc = esc(topic.split(" ")[0])
        cards.append(
            '      <a class="lr" data-slug="{slug}" href="{url}" target="_top">'
            '<span class="disc">{disc}</span><span class="meta">{topic} · {when}</span>'
            '<span class="ttl">{title}</span></a>'.format(
                slug=esc(item["slug"]), url=esc(branded_href(item)), disc=disc,
                topic=esc(topic), when=esc(when), title=esc(item["headline"])))
    return "\n".join(cards), items


def render_greats():
    """The Greats: names and funds from greats/greats.json. A profile with no page yet
    is drawn without a link, and a missing portrait as initials - nothing is invented."""
    if not os.path.exists(GREATS_REGISTRY):
        return "", []
    with open(GREATS_REGISTRY, encoding="utf-8") as fh:
        people = json.load(fh)["people"]
    cards = []
    for person in people:
        portrait = person.get("portrait") or ""
        if portrait and os.path.exists(os.path.join(HERE, portrait)):
            what = "photograph" if person.get("portrait_kind") == "photograph" else "illustrated portrait"
            disc = '<img src="%s" alt="%s, %s" loading="lazy">' % (esc(portrait), esc(person["name"]), what)
        else:
            disc = esc("".join(part[0] for part in person["name"].replace("-", " ").split()[:2]))
        inner = ('<span class="disc">%s</span><span class="nm">%s</span><span class="fd">%s</span>'
                 % (disc, esc(person["name"]), esc(person.get("line", ""))))
        if person.get("href"):
            href = person["href"]
            if not href.startswith("http"):
                href = SITE + "/" + href.lstrip("/")
            cards.append('      <a class="gr" href="%s" target="_top">%s</a>' % (esc(href), inner))
        else:
            cards.append('      <div class="gr">%s</div>' % inner)
    return "\n".join(cards), people


QUIZ_START, QUIZ_END = "<!-- QUIZZES:START", "<!-- QUIZZES:END -->"
QUIZ_INDEX = os.path.join(HERE, "quizzes", "index.json")
QUIZ_COUNT = 8


def render_quizzes():
    """Company data quizzes: the newest from quizzes/index.json (written by
    quizzes/build-quizzes.py). The whole section is generated, so with no quiz
    published there is no section at all rather than an empty one."""
    if not os.path.exists(QUIZ_INDEX):
        return ""
    with open(QUIZ_INDEX, encoding="utf-8") as fh:
        quizzes = json.load(fh).get("quizzes", [])[:QUIZ_COUNT]
    if not quizzes:
        return ""
    cards = []
    for quiz in quizzes:
        if quiz.get("cover"):
            image = '<img src="%s" alt="" loading="lazy">' % esc(quiz["cover"])
        else:
            image = '<span class="qz-tk">%s</span>' % esc(quiz["ticker"])
        cards.append(
            '      <a class="ca qz tile" data-slug="{slug}" href="{site}/Quizzes/{slug}" target="_top">\n'
            '        <div class="img">{image}</div>\n'
            '        <div class="body"><div class="line"><span class="tk">{company} · {ticker}</span></div>'
            '<h3>{title}</h3><span class="dt">{count} questions</span></div>\n'
            '      </a>'.format(slug=esc(quiz["slug"]), site=SITE, image=image, company=esc(quiz["company"]),
                              ticker=esc(quiz["ticker"]), title=esc(quiz["title"]), count=int(quiz["questions"])))
    return (
        '<section class="sec white" id="quizzes" aria-labelledby="qz-h">\n'
        '  <div class="wrap">\n'
        '    <div class="sec-head">\n'
        '      <div class="t">\n'
        '        <div class="kick">Investor tools</div>\n'
        '        <h2 class="h" id="qz-h"><a href="{site}/Quizzes" target="_top">Company data quizzes</a></h2>\n'
        '        <p class="lede">How well do you know the numbers? Eight questions on a company, built from its '
        'assessment. Every answer shows the working.</p>\n'
        '      </div>\n'
        '      <div class="tools">\n'
        '        <a class="all" href="{site}/Quizzes" target="_top">All<span class="vh"> quizzes</span> '
        '<span aria-hidden="true">→</span></a>\n'
        '        <button type="button" class="tog" aria-expanded="true" aria-label="Collapse or expand Company data '
        'quizzes"><svg aria-hidden="true" width="16" height="16" viewBox="0 0 24 24" fill="none" '
        'stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">'
        '<path d="M6 15l6-6 6 6"></path></svg></button>\n'
        '      </div>\n'
        '    </div>\n'
        '    <div class="ca-grid hrow body-x">\n{cards}\n    </div>\n'
        '  </div>\n'
        '</section>').format(site=SITE, cards="\n".join(cards))


VERDICT_JSON = os.path.join(HERE, "verdict-board.json")
VERDICT_ROW = re.compile(r"^\s*\['([^']+)','([^']+)','[^']+'\],?\s*$", re.M)
VERDICT_BUY = {"Buy", "Strong Buy"}
VERDICT_SELL = {"Sell", "Short", "Avoid", "Overvalued", "Reduce"}


def verdict_summary(page):
    """Group the page's own Verdict Board list the way the board groups it.

    The list stays where it is maintained (the `const verdicts=[...]` block in
    the page); this only publishes the counts and tickers so the Investor Tools
    card can show them without downloading the front page.
    """
    start = page.find("const verdicts=[")
    end = page.find("];", start)
    if start == -1 or end == -1:
        return None
    groups = {"buy": [], "hold": [], "sell": []}
    for ticker, call in VERDICT_ROW.findall(page[start:end]):
        key = "buy" if call in VERDICT_BUY else "sell" if call in VERDICT_SELL else "hold"
        groups[key].append(ticker)
    if not any(groups.values()):
        return None
    return {"groups": [
        {"key": "buy", "label": "Buy", "tickers": groups["buy"]},
        {"key": "hold", "label": "Hold", "tickers": groups["hold"]},
        {"key": "sell", "label": "Sell or avoid", "tickers": groups["sell"]},
    ]}


def verdict_json_text(page):
    summary = verdict_summary(page)
    return None if summary is None else json.dumps(summary, indent=1) + "\n"


VERDICT_ENTRY = re.compile(r"\['([^']+)','([^']+)','([^']+)'\]")


def refresh_verdicts(page, tiles):
    """Keep the Verdict Board on each company's newest assessment.

    The board's list lives in the page (`const verdicts=[...]`). For every
    company in the article registry whose newest assessment carries a call, its
    line is replaced, or added if the company was not on the board. Companies
    with no assessment in the registry keep the line they have. Decided by
    Mickey on 12 November 2026: the newest article replaces the old verdict.
    """
    start = page.find("const verdicts=[")
    end = page.find("];", start)
    if start == -1 or end == -1:
        return page
    key = lambda ticker: ticker.upper().rstrip(".")
    board = {key(t): [t, call, href] for t, call, href in VERDICT_ENTRY.findall(page[start:end])}
    newest = {}
    for tile in tiles:
        info = assessment(tile)
        if not info or not info.get("ticker") or not info.get("verdict"):
            continue
        k = key(info["ticker"])
        if k not in newest or tile["date"] > newest[k][0]:
            newest[k] = (tile["date"], info["ticker"], info["verdict"], tile["url"])
    for k, (_, ticker, call, url) in newest.items():
        shown = board[k][0] if k in board else ticker
        board[k] = [shown, call, url]
    for entry in board.values():
        if any("'" in part or "\\" in part for part in entry):
            raise SystemExit("ERROR: Verdict Board entry cannot be written safely: %r" % entry)
    lines = ",\n".join("  ['%s','%s','%s']" % tuple(board[k]) for k in sorted(board))
    return page[:start] + "const verdicts=[\n" + lines + "\n" + page[end:]


def build_redesign(page, tiles):
    page = refresh_verdicts(page, tiles)
    body, picked = render_assessments(tiles)
    page = splice(page, ASSESS_START, ASSESS_END, body)
    body, reads = render_longreads(tiles)
    page = splice(page, LONG_START, LONG_END, body)
    if GREATS_START in page:
        body, _ = render_greats()
        page = splice(page, GREATS_START, GREATS_END, body)
    if QUIZ_START in page:
        page = splice(page, QUIZ_START, QUIZ_END, render_quizzes())
    return page, picked, reads


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

    with open(FRONT_PAGE, encoding="utf-8") as fh:
        page = fh.read()

    if ASSESS_START in page:
        # Redesigned layout: no hero and no slots, so nothing "drops off".
        page, picked, reads = build_redesign(page, tiles)
        return page, picked[0], picked[1:] + reads, []

    hero, chosen, dropped = select(tiles, slots)
    page = splice(page, HERO_START, HERO_END, render_hero(hero))
    blocks = "\n\n".join(render_block(t, s) for t, s in zip(chosen, slots))
    page = splice(page, TILES_START, TILES_END, "\n" + blocks + "\n")

    return page, hero, chosen, dropped


def report(hero, chosen, dropped, slots):
    print("  hero  <- %s  (%s)" % (hero["slug"], hero["date"]))
    for tile, slot in zip(chosen, list(slots) + ["-"] * len(chosen)):
        print("  %-5s <- %s  (%s)" % (slot, tile["slug"], tile["date"]))
    for tile in dropped:
        print("  DROPPED off the front page: %s  (%s)" % (tile["slug"], tile["date"]))


def default_cover(slug):
    """Cover filename for an article that declares no explicit tile-cover.

    House convention drops a trailing '-blog-page' and appends '-cover.<ext>'
    (bp-BP-blog-page -> bp-BP-cover.png). Probe the extensions actually in use
    and fall back to .png so the missing-cover warning below still fires with a
    sensible name to create.
    """
    stem = slug[:-len("-blog-page")] if slug.endswith("-blog-page") else slug
    for ext in ("png", "jpg", "jpeg"):
        candidate = "%s-cover.%s" % (stem, ext)
        if os.path.exists(os.path.join(HERE, candidate)):
            return candidate
    return "%s-cover.png" % stem


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
        # Page sources carry HTML entities (&amp;, &mdash;). The registry stores plain
        # text and the renderer escapes on output, so decode here or entities double-escape.
        return html.unescape(match.group(1)) if match else None

    slug = os.path.splitext(os.path.basename(path))[0]

    # Special-coverage articles (tile-kind "budget") have their own hand-maintained
    # block on the front page, so they never take a tile or the hero.
    if meta("tile-kind") == "budget":
        print("  registry: '%s' is Budget coverage - it lives in the BUDGET-2026 block, not the tiles." % slug)
        return

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
        "cover": meta("tile-cover") or default_cover(slug),
        "alt": meta("tile-alt") or (html.unescape(title.group(1).strip()) if title else slug),
        "cat": meta("tile-cat") or "Analysis",
        "headline": meta("tile-headline") or (html.unescape(title.group(1).strip()) if title else slug),
        "pill": meta("tile-pill") or "watch",
        "pill_label": meta("tile-pill-label") or "Analysis",
    }
    if meta("tile-ticker"):
        tile["ticker"] = meta("tile-ticker")
    if meta("tile-meta"):
        tile["meta_extra"] = [meta("tile-meta")]

    cover_path = os.path.join(HERE, tile["cover"])
    if not os.path.exists(cover_path):
        # A cover-less tile would take a grid slot (or the hero) with a broken image.
        # Hold the article off the front page; re-run --add once its cover is pushed.
        print("  HELD OFF the front page: cover '%s' is not in the repo. "
              "Re-run --add for '%s' after the cover is pushed." % (tile["cover"], slug))
        return

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
    parser.add_argument("--page", metavar="FILE.html",
                        help="build this page instead of front-page.html (used for the redesign preview)")
    args = parser.parse_args()

    if args.page:
        global FRONT_PAGE
        FRONT_PAGE = os.path.join(HERE, args.page)

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
            print("%s is OUT OF DATE. Run: python build-front-page.py" % os.path.basename(FRONT_PAGE))
            return 1
        print("%s is up to date." % os.path.basename(FRONT_PAGE))
        return 0

    with open(FRONT_PAGE, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(page)
    verdicts = verdict_json_text(page)
    if verdicts is not None:
        with open(VERDICT_JSON, "w", encoding="utf-8", newline="\n") as fh:
            fh.write(verdicts)
    print("%s rebuilt:" % os.path.basename(FRONT_PAGE))
    report(hero, chosen, dropped, slots)
    return 0


if __name__ == "__main__":
    sys.exit(main())
