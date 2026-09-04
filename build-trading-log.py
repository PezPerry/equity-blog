#!/usr/bin/env python3
"""Build the Trading Log index and the front-page teaser from trading-log/entries.json.

    python build-trading-log.py                 # rebuild trading-log/index.html and the
                                                # TRADING-LOG block inside front-page.html
    python build-trading-log.py --add FILE.html # register a new entry page (root HTML tagged
                                                # <meta name="tile-variant" content="trading-log">)
                                                # from its meta tags, then rebuild
    python build-trading-log.py --check         # exit 1 if either output is stale

Entry pages carry these meta tags (see crest-nicholson-worse-figures-same-asset-case.html):
    title, description, log-tag, log-tickers (comma separated), log-companies, log-time (HH:MM UK),
    log-why, log-read (minutes), log-related-url, log-related-label, log-date (YYYY-MM-DD,
    optional: defaults to today when --add runs).

Design: "Trading Log Revamp" handoff, option 2a (September 2026). Colours, type and spacing
are the handoff's tokens; the site nav is copied from front-page.html and hidden when the page
is framed by the Base44 shell (?embed=1 or window.top != window.self).
"""
import argparse
import datetime as dt
import html
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REGISTRY = os.path.join(HERE, "trading-log", "entries.json")
INDEX = os.path.join(HERE, "trading-log", "index.html")
FRONT_PAGE = os.path.join(HERE, "front-page.html")
TEASER_START = "<!-- TRADING-LOG:START"
TEASER_END = "<!-- TRADING-LOG:END -->"
SITE = "https://equityandmarketsinsight.com"
LOG_URL = SITE + "/TradingLog"

TAGS = ("Bought", "Sold", "Held", "Passed", "Comment", "Market view")
TAG_CLASS = {"Bought": "bought", "Sold": "sold", "Held": "held", "Passed": "passed",
             "Comment": "comment", "Market view": "view"}
PAGE_SIZE = 25


def esc(text):
    return html.escape(text or "", quote=True)


def load():
    with open(REGISTRY, encoding="utf-8") as fh:
        data = json.load(fh)
    entries = sorted(data["entries"], key=lambda e: (e["date"], e.get("time") or ""), reverse=True)
    return data, entries


def nice_date(iso):
    d = dt.date.fromisoformat(iso)
    return "%d %s %d" % (d.day, d.strftime("%b"), d.year)


def pill(tag, cls="pill"):
    if tag not in TAG_CLASS:
        raise SystemExit("ERROR: unknown tag %r (expected one of %s)" % (tag, ", ".join(TAGS)))
    return '<span class="%s %s-%s">%s</span>' % (cls, cls, TAG_CLASS[tag], esc(tag))


# --------------------------------------------------------------------------- index page

INDEX_CSS = """
:root{--navy:#112238;--gold:#A98A48;--gold-text:#81652E;--ivory:#F8F5EF;--body:#26304A;--muted:#526173;--faint:#8A8F99;--hair:rgba(17,34,56,.12);--hover:rgba(169,138,72,.06);
  --paper:#FBF9F4;--gold-deep:#A9812E;--text:#141C2B;--text-dim:#4E5567;--line:rgba(27,42,74,.15);--nav-hair:rgba(27,42,74,.07);--accent:var(--gold-deep)}
:root[data-theme="dark"]{--paper:#0E1626;--text:#F3ECDD;--text-dim:#A9B0BF;--line:rgba(212,168,83,.16);--nav-hair:rgba(255,255,255,.07);--accent:#D4A853}
:root[data-theme="dark"] .brand-logo{filter:invert(1)}
*{box-sizing:border-box;margin:0;padding:0}
html{-webkit-text-size-adjust:100%}
body{background:var(--ivory);color:var(--navy);font-family:"Inter",system-ui,sans-serif;-webkit-font-smoothing:antialiased;line-height:1.5}
a{color:inherit;text-decoration:none}
img{display:block;max-width:100%}
.wrap{max-width:1280px;margin:0 auto;padding:0 clamp(16px,4vw,44px)}
/* site nav: copied from front-page.html, unchanged by design */
.nav{border-bottom:1px solid var(--nav-hair);background:var(--paper);position:sticky;top:0;z-index:60}
.nav-inner{display:flex;align-items:center;gap:26px;height:60px}
.brand{display:flex;align-items:center;gap:12px;flex-shrink:0}
.brand-logo{height:34px;width:auto;display:block}
.brand-name{font-family:"Playfair Display",serif;font-weight:800;font-size:18px;letter-spacing:-.01em;line-height:1;color:var(--text);white-space:nowrap}
.brand-name b{font-style:italic;font-weight:600;color:var(--accent)}
.nav-links{display:flex;gap:26px;margin-left:auto;align-items:center}
.nav-links a{font-family:"DM Sans",sans-serif;font-size:14px;font-weight:500;color:var(--text-dim);letter-spacing:.01em;position:relative;padding:4px 0;transition:color .2s;white-space:nowrap}
.nav-links a::after{content:"";position:absolute;left:0;right:100%%;bottom:-2px;height:2px;background:var(--accent);transition:right .25s ease}
.nav-links a:hover{color:var(--text)}
.nav-links a:hover::after{right:0}
.nav-links a.here{color:var(--text)}
.nav-links a.here::after{right:0}
.theme-btn{border:1px solid var(--line);background:transparent;color:var(--text-dim);font-family:"DM Mono",monospace;font-size:13px;padding:6px 11px;border-radius:999px;cursor:pointer;line-height:1}
.theme-btn:hover{border-color:var(--accent);color:var(--accent)}
html.embedded .nav{display:none}
/* log */
main{max-width:1280px;margin:0 auto}
.log-head{padding:56px 48px 20px}
.log-head h1{font-family:"Cormorant Garamond",Georgia,serif;font-weight:500;font-size:60px;line-height:1;letter-spacing:-.01em;color:var(--navy)}
.log-head p{margin-top:16px;font-size:16px;line-height:1.6;color:var(--muted)}
.cols,.row{display:grid;grid-template-columns:150px 120px 1fr 190px;gap:24px}
.cols{margin:0 48px;padding:28px 0 12px;border-bottom:1px solid var(--navy);font-size:11px;letter-spacing:.14em;text-transform:uppercase;color:var(--faint);font-weight:600}
.cols .r{text-align:right}
.rows{padding:0 48px 48px}
.row{padding:22px 0;border-bottom:1px solid var(--hair);align-items:start;color:var(--navy);transition:background .15s}
.row:hover{background:var(--hover)}
.row .when{font-size:13px;font-variant-numeric:tabular-nums;color:var(--muted);padding-top:6px;line-height:1.5}
.row .when span{color:var(--faint)}
.row .act{padding-top:5px}
.row .ttl{font-family:"Cormorant Garamond",Georgia,serif;font-size:27px;font-weight:500;line-height:1.15;color:var(--navy)}
.row .why{font-size:14.5px;line-height:1.55;color:var(--muted);margin-top:6px;max-width:64ch}
.row .cos{text-align:right;font-size:12.5px;letter-spacing:.06em;text-transform:uppercase;color:var(--navy);font-weight:600;padding-top:8px;line-height:1.7}
.pill{display:inline-block;font-family:"Inter",system-ui,sans-serif;font-size:11px;letter-spacing:.12em;text-transform:uppercase;font-weight:600;padding:4px 9px;border-radius:2px;line-height:1.3;white-space:nowrap}
.pill-bought{background:var(--navy);color:var(--ivory)}
.pill-sold{background:var(--gold);color:var(--navy)}
.pill-held{background:transparent;color:var(--navy);border:1px solid var(--navy)}
.pill-passed{background:transparent;color:var(--gold-text);border:1px solid var(--gold)}
.pill-comment{background:transparent;color:var(--muted);border:1px solid rgba(17,34,56,.3)}
.pill-view{background:rgba(17,34,56,.08);color:var(--navy)}
.log-foot{display:flex;justify-content:space-between;padding-top:22px;font-size:13px;color:var(--muted)}
.log-foot button{color:var(--gold-text);background:none;border:0;font:inherit;font-weight:600;cursor:pointer;padding:0}
.log-foot button[hidden]{display:none}
.row[hidden]{display:none}
@media (max-width:920px){.nav-links{display:none}}
@media (max-width:760px){
  .log-head{padding:40px 20px 12px}
  .log-head h1{font-size:44px}
  .cols{display:none}
  .rows{padding:0 20px 40px;border-top:1px solid var(--navy);margin-top:20px}
  .row{grid-template-columns:1fr;gap:8px;padding:18px 0}
  .row .when{display:flex;gap:10px;padding-top:0}
  .row .act{padding-top:0}
  .row .cos{text-align:left;padding-top:2px}
}
"""

NAV_HTML = """<div class="nav">
  <div class="wrap nav-inner">
    <a class="brand" href="../front-page.html">
      <img class="brand-logo" src="../emi-logo.png" alt="Equity &amp; Markets Insight logo">
      <span class="brand-name">Equity <b>&amp;</b> Markets Insight</span>
    </a>
    <nav class="nav-links">
      <a target="_top" href="https://equityandmarketsinsight.com/Companies">Company Analysis</a>
      <a target="_top" href="https://equityandmarketsinsight.com/Articles">Macro</a>
      <a target="_top" href="https://equityandmarketsinsight.com/InterestRates">Rates</a>
      <a target="_top" class="here" href="https://equityandmarketsinsight.com/TradingLog">Investor Tools</a>
    </nav>
    <button class="theme-btn" id="themeBtn" aria-label="Toggle theme">&#9685;</button>
  </div>
</div>"""

EMBED_JS = """<script>
/* Framed by the Base44 shell (?embed=1, or inside an iframe): the shell supplies the site
   chrome, so hide this page's own nav. The theme button only ever affects the nav. */
(function(){
  var framed=false;try{framed=window.self!==window.top}catch(e){framed=true}
  if(framed||new URLSearchParams(location.search).has('embed')){document.documentElement.classList.add('embedded')}
  var tb=document.getElementById('themeBtn');if(tb){tb.addEventListener('click',function(){
    var r=document.documentElement;r.setAttribute('data-theme',r.getAttribute('data-theme')==='dark'?'light':'dark')})}
  /* Tell a framing page how tall we are, so the shell can size the iframe. */
  function tall(){try{window.parent.postMessage({type:'emi-height',page:'trading-log',height:document.documentElement.scrollHeight},'*')}catch(e){}}
  if(framed){tall();window.addEventListener('load',tall);window.addEventListener('resize',tall);
    if(window.ResizeObserver){new ResizeObserver(tall).observe(document.body)}}
})();
</script>"""


def render_row(e, hidden=False):
    tickers = " · ".join(e.get("tickers") or [])
    if not tickers:
        tickers = e.get("companies") or ""
    return (
        '    <a class="row"%s href="%s" target="_top" data-slug="%s">'
        '<div class="when">%s<br><span>%s</span></div>'
        '<div class="act">%s</div>'
        '<div><div class="ttl">%s</div><div class="why">%s</div></div>'
        '<div class="cos">%s</div>'
        '</a>'
    ) % (" hidden" if hidden else "", esc(e["url"]), esc(e["slug"]), nice_date(e["date"]),
         esc(e.get("time") or ""), pill(e["tag"]), esc(e["title"]), esc(e["why"]), esc(tickers))


def render_index(entries):
    rows = "\n".join(render_row(e, hidden=i >= PAGE_SIZE) for i, e in enumerate(entries))
    total = len(entries)
    shown = min(PAGE_SIZE, total)
    older = "" if total <= PAGE_SIZE else '<button type="button" id="older">Older entries →</button>'
    return """<!DOCTYPE html>
<html lang="en-GB" data-theme="light">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Trading Log | Equity &amp; Markets Insight</title>
<meta name="description" content="Commentary on the most recent news hitting the markets: short dated notes on what I bought, sold, held or passed on, and why.">
<meta name="robots" content="index,follow">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;0,600;0,700;1,300;1,400;1,500;1,600;1,700&family=Inter:wght@100..900&family=Playfair+Display:wght@600;800&family=DM+Sans:wght@400;500&family=DM+Mono:wght@400&display=swap" rel="stylesheet">
<style>%(css)s</style>
</head>
<body>
%(nav)s
<main>
  <div class="log-head">
    <h1>Trading Log</h1>
    <p>Commentary on the most recent news hitting the markets</p>
  </div>
  <div class="cols"><div>Date</div><div>Action</div><div>Note</div><div class="r">Companies</div></div>
  <div class="rows" id="rows">
%(rows)s
    <div class="log-foot"><span id="count">Showing %(shown)d of %(total)d entries</span>%(older)s</div>
  </div>
</main>
%(embed)s
<script>
(function(){
  var btn=document.getElementById('older');if(!btn)return;
  btn.addEventListener('click',function(){
    var hidden=document.querySelectorAll('.row[hidden]');
    for(var i=0;i<hidden.length&&i<%(page)d;i++){hidden[i].hidden=false}
    var left=document.querySelectorAll('.row[hidden]').length,all=document.querySelectorAll('.row').length;
    document.getElementById('count').textContent='Showing '+(all-left)+' of '+all+' entries';
    if(!left){btn.hidden=true}
  });
})();
</script>
</body>
</html>
""" % {"css": INDEX_CSS, "nav": NAV_HTML, "rows": rows, "shown": shown, "total": total,
       "older": older, "embed": EMBED_JS, "page": PAGE_SIZE}


# --------------------------------------------------------------------------- front-page teaser

def render_teaser(entries):
    rows = []
    for e in entries[:3]:
        rows.append('      <a class="tl-row" target="_top" href="%s"><span class="tl-date">%s</span>%s<span class="tl-title">%s</span></a>'
                    % (esc(e["url"]), nice_date(e["date"]), pill(e["tag"], "tl-pill"), esc(e["title"])))
    return """  <section class="tlog" id="tlog">
    <div class="tl-lead">
      <div class="tl-kicker">Trading Log</div>
      <h3>The latest notes.</h3>
      <a target="_top" href="%s">Open the log →</a>
    </div>
    <div class="tl-list">
%s
    </div>
  </section>""" % (LOG_URL, "\n".join(rows))


def splice(source, start_marker, end_marker, body):
    start = source.find(start_marker)
    end = source.find(end_marker)
    if start == -1 or end == -1:
        raise SystemExit("ERROR: markers %s / %s not found in front-page.html" % (start_marker, end_marker))
    line_end = source.find("\n", start)
    end_line_start = source.rfind("\n", 0, end) + 1
    return source[: line_end + 1] + body + "\n" + source[end_line_start:]


def build():
    _, entries = load()
    index = render_index(entries)
    with open(FRONT_PAGE, encoding="utf-8") as fh:
        page = fh.read()
    page = splice(page, TEASER_START, TEASER_END, render_teaser(entries))
    return index, page


# --------------------------------------------------------------------------- --add

def add_entry(path):
    with open(path, encoding="utf-8") as fh:
        source = fh.read()

    def meta(name):
        m = re.search(r'<meta\s+name="%s"\s+content="([^"]*)"' % re.escape(name), source)
        return html.unescape(m.group(1)) if m else None

    if meta("tile-variant") != "trading-log":
        print("  trading log: '%s' is not tagged tile-variant=trading-log - nothing to register." % path)
        return False
    slug = os.path.splitext(os.path.basename(path))[0]
    data, _ = load()
    if any(e["slug"] == slug for e in data["entries"]):
        print("  trading log: '%s' is already registered - keeping the existing entry." % slug)
        return False
    title = re.search(r"<title>([^|<]+)", source)
    tickers = [t.strip() for t in (meta("log-tickers") or "").split(",") if t.strip()]
    words = len(re.sub(r"<[^>]+>", " ", source).split())
    entry = {
        "slug": slug,
        "date": meta("log-date") or dt.date.today().isoformat(),
        "time": meta("log-time") or "",
        "tag": meta("log-tag") or "Comment",
        "tickers": tickers,
        "companies": meta("log-companies") or ", ".join(tickers),
        "title": html.unescape(title.group(1).strip()) if title else slug,
        "why": meta("log-why") or meta("description") or "",
        "url": "%s/TradeDetail?slug=%s" % (SITE, slug),
        "source": os.path.basename(path),
        "readMinutes": int(meta("log-read") or max(1, words // 200)),
        "related": ({"url": meta("log-related-url"), "label": meta("log-related-label") or "Full valuation"}
                    if meta("log-related-url") else None),
    }
    pill(entry["tag"])  # validates the tag
    data["entries"].insert(0, entry)
    with open(REGISTRY, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)
        fh.write("\n")
    print("  trading log <- %s (%s %s, %s)" % (slug, entry["date"], entry["time"], entry["tag"]))
    return True


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--add", metavar="ENTRY.html")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    if args.add:
        add_entry(args.add)

    index, page = build()
    if args.check:
        stale = []
        with open(INDEX, encoding="utf-8") as fh:
            if fh.read() != index:
                stale.append("trading-log/index.html")
        with open(FRONT_PAGE, encoding="utf-8") as fh:
            if fh.read() != page:
                stale.append("front-page.html (TRADING-LOG block)")
        if stale:
            print("OUT OF DATE: %s. Run: python build-trading-log.py" % ", ".join(stale))
            return 1
        print("Trading log outputs are up to date.")
        return 0

    with open(INDEX, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(index)
    with open(FRONT_PAGE, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(page)
    _, entries = load()
    print("trading-log/index.html rebuilt with %d entries; front-page teaser shows the newest %d."
          % (len(entries), min(3, len(entries))))
    return 0


if __name__ == "__main__":
    sys.exit(main())
