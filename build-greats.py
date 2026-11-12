#!/usr/bin/env python3
"""Build The Greats profile pages from the Obsidian notes.

Source:  <vault>/Equity & Markets Insight/The Greats/*.md  (plus Images/)
Output:  greats/<slug>.html, greats/images/*, and greats/greats.json (the
         registry build-front-page.py and the app's /Greats pages read).

Nothing is written by hand: edit the note in Obsidian and run this again.
The article text is carried over as written. The only things this script
changes are (1) vault-only lines are dropped (the "Part of [[The Greats]]"
note and the link to the vault's image register), and (2) every image is
drawn from Images/image-manifest.json, so the caption, alt text and credit
are the manifest's.

    python build-greats.py            build everything
    python build-greats.py --check    build in memory, report, write nothing
"""
import argparse
import html
import json
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_SOURCE = r"C:\Master Control\Equity & Markets Insight\The Greats"
OUT = os.path.join(HERE, "greats")
SITE = "https://equityandmarketsinsight.com"
PAGES = "https://pezperry.github.io/equity-blog"
REGISTRY = os.path.join(OUT, "greats.json")
# Every card and page header uses a drawn portrait (Mickey, 12 November 2026). Where a
# profile's first image in the vault is a photograph, the drawn portrait from the original
# commission is used for the card and header instead; the photograph stays in the article.
DRAWN_PORTRAITS = r"C:\Master Control\output\2026-09-29-the-greats\images"
LEGAL = ("© 2026 Equity and Markets Insight. This site is for information and education only and is not "
         "financial advice. I am not a financial adviser. Investing involves risk, including loss of capital. "
         "The author may hold positions in companies mentioned. Do your own research and consider seeking "
         "independent advice.")


def esc(text):
    return html.escape(str(text), quote=True)


def slugify(name):
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def inline(text):
    """Bold, italics and links. Everything else is escaped."""
    out = esc(text)
    out = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)",
                 lambda m: '<a href="%s" target="_blank" rel="noopener">%s</a>' % (m.group(2), m.group(1)), out)
    out = re.sub(r"\*\*([^*]+)\*\*", r"<strong>\1</strong>", out)
    out = re.sub(r"(?<![*\w])\*([^*\n]+)\*(?![*\w])", r"<em>\1</em>", out)
    return out


def front_matter(text):
    match = re.match(r"---\n(.*?)\n---\n", text, re.S)
    meta = {}
    if match:
        for line in match.group(1).splitlines():
            pair = re.match(r'([a-z_]+):\s*"?(.*?)"?\s*$', line)
            if pair and pair.group(2):
                meta[pair.group(1)] = pair.group(2)
        text = text[match.end():]
    return meta, text


def figure(record):
    stem = record["id"]
    kind = record["kind"]
    credit = record["credit_markdown"].replace(" and the linked image register", "")
    wide = int(record["width"]) >= int(record["height"])
    cls = "fig" + ("" if wide else " tall") + (" chart" if kind == "original chart" else "")
    return (
        '<figure class="{cls}">'
        '<img src="images/{stem}.webp" srcset="images/{stem}-800.webp 800w, images/{stem}.webp {w}w" '
        'sizes="(max-width: 760px) 100vw, 720px" width="{w}" height="{h}" alt="{alt}" loading="lazy" decoding="async">'
        '<figcaption>{cap} <span class="cr">{credit}</span></figcaption></figure>'
    ).format(cls=cls, stem=stem, w=record["width"], h=record["height"], alt=esc(record["alt"]),
             cap=inline(record["caption"]), credit=inline(credit))


def table(rows):
    cells = [[c.strip() for c in row.strip().strip("|").split("|")] for row in rows]
    align = ["r" if c.endswith(":") else "" for c in cells[1]]
    out = ['<div class="tblw"><table><thead><tr>']
    for i, cell in enumerate(cells[0]):
        out.append('<th%s>%s</th>' % (' class="r"' if align[i] else "", inline(cell)))
    out.append("</tr></thead><tbody>")
    for row in cells[2:]:
        out.append("<tr>")
        for i, cell in enumerate(row):
            out.append('<td%s>%s</td>' % (' class="r"' if i < len(align) and align[i] else "", inline(cell)))
        out.append("</tr>")
    out.append("</tbody></table></div>")
    return "".join(out)


def convert(text, number, images, problems):
    """Markdown body -> (title, standfirst, html). Only the syntax the notes use."""
    lines = text.splitlines()
    title = standfirst = ""
    out = []
    used = []
    placeholder = 0
    section = ""
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
            continue
        marker = re.match(r"<!-- greats-image:(\d\d-\d\d) -->", stripped)
        if marker:
            ident = marker.group(1)
            while i < len(lines) and "<!-- /greats-image:%s -->" % ident not in lines[i]:
                i += 1
            i += 1
            if ident in images:
                out.append(figure(images[ident]))
                used.append(ident)
            else:
                problems.append("image %s is not in the manifest" % ident)
            continue
        if stripped.startswith(">"):
            block = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                block.append(lines[i])
                i += 1
            joined = "\n".join(block)
            if "IMAGE SPACE" in joined:
                # A note whose image spaces were never filled in: use the manifest's
                # image for that slot, in order.
                placeholder += 1
                ident = "%02d-%02d" % (number, placeholder)
                if ident in images:
                    out.append(figure(images[ident]))
                    used.append(ident)
                else:
                    problems.append("no manifest image for unfilled space %s" % ident)
            elif "[[The Greats]]" not in joined:
                problems.append("unexpected blockquote: %s" % joined[:80])
            continue
        if stripped.startswith("# "):
            title = stripped[2:].strip()
            i += 1
            continue
        if stripped.startswith("## "):
            section = stripped[3:].strip()
            out.append('<h2 id="%s">%s</h2>' % (slugify(section), inline(section)))
            i += 1
            continue
        if stripped.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                rows.append(lines[i])
                i += 1
            out.append(table(rows))
            continue
        if stripped.startswith("- "):
            items = []
            while i < len(lines) and lines[i].strip().startswith("- "):
                items.append(lines[i].strip()[2:])
                i += 1
            if section == "Life in dates":
                rows = []
                for item in items:
                    when = re.match(r"\*\*(.+?):\*\*\s*(.*)$", item)
                    if when:
                        rows.append('<li><span class="yr">%s</span><span>%s</span></li>'
                                    % (inline(when.group(1)), inline(when.group(2))))
                    else:
                        rows.append("<li><span></span><span>%s</span></li>" % inline(item))
                out.append('<ol class="dates">%s</ol>' % "".join(rows))
            else:
                cls = ' class="src"' if section == "Sources" else ""
                out.append("<ul%s>%s</ul>" % (cls, "".join("<li>%s</li>" % inline(x) for x in items)))
            continue
        if not standfirst and not section and re.match(r"^\*[^*].*\*$", stripped):
            standfirst = stripped.strip("*").strip()
            i += 1
            continue
        out.append("<p>%s</p>" % inline(stripped))
        i += 1
    return title, standfirst, "\n".join(out), used


PAGE = """<!DOCTYPE html>
<html lang="en-GB">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title} | The Greats | Equity &amp; Markets Insight</title>
<meta name="description" content="{desc}">
<link rel="canonical" href="{site}/Greats/{slug}">
<meta property="og:type" content="article">
<meta property="og:title" content="{title}">
<meta property="og:description" content="{desc}">
<meta property="og:image" content="{pages}/greats/images/{portrait}.webp">
<meta property="og:url" content="{site}/Greats/{slug}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;1,600&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root{{--navy:#112238;--ivory:#f8f5ef;--gold:#a98a48;--gold-d:#81652e;--slate:#526173;--rule:#dadbd8;--light:#e3e2dc;--serif:'Cormorant Garamond',Georgia,serif;--sans:'Inter',system-ui,sans-serif}}
*{{box-sizing:border-box}}
html{{-webkit-text-size-adjust:100%}}
body{{margin:0;background:var(--ivory);color:var(--navy);font-family:var(--sans);font-size:18px;line-height:1.7;-webkit-font-smoothing:antialiased;overflow-x:hidden}}
a{{color:inherit}}
img{{max-width:100%;height:auto;display:block}}
.bar{{background:#fff;border-bottom:1px solid var(--rule)}}
.bar .in{{max-width:1240px;margin:0 auto;padding:14px 16px;display:flex;align-items:center;justify-content:space-between;gap:16px}}
.bar .logo img{{height:40px;width:auto}}
.bar nav{{display:flex;gap:10px;align-items:center}}
.bar nav a{{display:inline-flex;align-items:center;min-height:44px;padding:0 14px;border-radius:8px;font-size:14px;font-weight:600;text-decoration:none;border:1px solid var(--rule)}}
.bar nav a.pr{{background:var(--navy);color:var(--ivory);border-color:var(--navy)}}
.bar nav a:hover{{border-color:var(--navy)}}
.hero{{background:#fff;border-bottom:1px solid var(--rule)}}
.hero .in{{max-width:1040px;margin:0 auto;padding:48px 16px 44px;display:grid;grid-template-columns:minmax(0,1fr) 220px;gap:40px;align-items:center}}
.eyebrow{{font-size:12px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--slate)}}
.eyebrow a{{text-decoration:none}}
.eyebrow a:hover{{color:var(--gold-d)}}
h1{{font-family:var(--serif);font-weight:600;font-size:clamp(38px,6vw,62px);line-height:1.05;margin:12px 0 18px;letter-spacing:-.005em}}
.stand{{font-family:var(--serif);font-style:italic;font-weight:600;font-size:clamp(21px,2.6vw,26px);line-height:1.35;color:var(--navy);margin:0 0 20px;max-width:34em}}
.facts{{display:flex;flex-wrap:wrap;gap:8px 10px;margin:0;padding:0;list-style:none;font-size:14px;color:var(--slate)}}
.facts li{{padding:6px 12px;border:1px solid var(--rule);border-radius:999px;background:var(--ivory)}}
.disc{{width:220px;height:220px;border-radius:50%;overflow:hidden;border:1px solid var(--rule);background:var(--light)}}
.disc img{{width:100%;height:100%;object-fit:cover;object-position:50% 20%}}
main{{max-width:720px;margin:0 auto;padding:44px 16px 24px}}
main p{{margin:0 0 1.15em}}
main p a,main li a,figcaption a{{color:var(--gold-d);text-decoration:underline;text-underline-offset:2px}}
h2{{font-family:var(--serif);font-weight:600;font-size:34px;line-height:1.15;margin:56px 0 18px;padding-top:28px;border-top:1px solid var(--rule)}}
.fig{{margin:36px 0}}
.fig img{{width:100%;border-radius:8px;border:1px solid var(--rule);background:#fff}}
.fig.tall img{{max-width:420px;margin:0 auto}}
.fig.chart{{margin-left:-60px;margin-right:-60px}}
figcaption{{margin-top:10px;font-size:14px;line-height:1.55;color:var(--slate)}}
figcaption .cr{{display:block;margin-top:4px;font-size:12.5px}}
.tblw{{overflow-x:auto;margin:0 0 20px;border:1px solid var(--rule);border-radius:8px;background:#fff}}
table{{border-collapse:collapse;width:100%;font-size:15px;line-height:1.4;font-variant-numeric:tabular-nums}}
th,td{{padding:11px 14px;text-align:left;border-bottom:1px solid var(--rule);vertical-align:top}}
th{{font-size:12px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:var(--slate);background:var(--ivory)}}
tr:last-child td{{border-bottom:0}}
th.r,td.r{{text-align:right;white-space:nowrap}}
.dates{{list-style:none;margin:0;padding:0;border-left:2px solid var(--gold)}}
.dates li{{display:grid;grid-template-columns:92px minmax(0,1fr);gap:14px;padding:9px 0 9px 18px;font-size:16px;line-height:1.5}}
.dates .yr{{font-weight:700}}
.src{{margin:0;padding-left:20px;font-size:15px;line-height:1.55;color:var(--slate)}}
.src li{{margin-bottom:8px;overflow-wrap:anywhere}}
.more{{max-width:720px;margin:0 auto;padding:8px 16px 56px}}
.more a{{display:inline-flex;align-items:center;min-height:48px;padding:0 20px;border-radius:8px;background:var(--navy);color:var(--ivory);font-size:15px;font-weight:700;text-decoration:none}}
.more a:hover{{color:var(--gold)}}
footer{{border-top:1px solid var(--rule);background:#fff}}
footer p{{max-width:1040px;margin:0 auto;padding:24px 16px 32px;font-size:13px;font-style:italic;line-height:1.6;color:var(--slate)}}
@media (max-width:860px){{
  .fig.chart{{margin-left:0;margin-right:0}}
  .hero .in{{grid-template-columns:1fr;gap:24px;padding:32px 16px 30px}}
  .disc{{width:132px;height:132px;order:-1}}
}}
@media (max-width:560px){{
  body{{font-size:17px}}
  h2{{font-size:29px;margin-top:44px}}
  .bar .logo img{{height:30px}}
  .bar nav a{{padding:0 10px;font-size:13px}}
  .dates li{{grid-template-columns:74px minmax(0,1fr);gap:10px;padding-left:14px}}
  th,td{{padding:9px 10px;font-size:14px}}
}}
</style>
</head>
<body>
<div class="bar"><div class="in">
  <a class="logo" href="{site}/" target="_top" aria-label="Equity and Markets Insight home"><img src="../assets/brand/emi-horizontal-navy-gold.png" width="1324" height="258" alt="Equity and Markets Insight"></a>
  <nav aria-label="Site"><a href="{site}/Greats" target="_top">The Greats</a><a class="pr" href="{site}/Premium" target="_top">Go Premium</a></nav>
</div></div>
<header class="hero"><div class="in">
  <div>
    <div class="eyebrow"><a href="{site}/Greats" target="_top">The Greats</a> · Life and times</div>
    <h1>{title}</h1>
    <p class="stand">{stand}</p>
    <ul class="facts">{facts}</ul>
  </div>
  <div class="disc"><img src="images/{portrait}-800.webp" width="800" height="800" alt="{portrait_alt}"></div>
</div></header>
<main>
{body}
</main>
<div class="more"><a href="{site}/Greats" target="_top">More of The Greats</a></div>
<footer><p>{legal}</p></footer>
</body>
</html>
"""


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", default=DEFAULT_SOURCE)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    with open(os.path.join(args.source, "Images", "image-manifest.json"), encoding="utf-8") as fh:
        images = {r["id"]: r for r in json.load(fh)}
    with open(REGISTRY, encoding="utf-8") as fh:
        registry = json.load(fh)

    notes = {}
    for folder in (args.source, os.path.join(args.source, "Word")):
        for name in sorted(os.listdir(folder)):
            if not name.endswith(".md") or name == "The Greats.md" or "Research Notes" in name:
                continue
            with open(os.path.join(folder, name), encoding="utf-8") as fh:
                meta, text = front_matter(fh.read().replace("\r\n", "\n"))
            if meta.get("series") == "The Greats" and meta.get("investor"):
                notes.setdefault(meta["investor"], (meta, text, os.path.join(folder, name)))

    problems, pages, copies, drawings = [], {}, set(), {}
    for number, person in enumerate(registry["people"], start=1):
        name = person["name"]
        if name not in notes:
            problems.append("no note for %s" % name)
            continue
        meta, text, path = notes[name]
        title, standfirst, body, used = convert(text, number, images, problems)
        expected = ["%02d-%02d" % (number, n) for n in range(1, 5)]
        if used != expected:
            problems.append("%s: images %s, expected %s" % (name, used, expected))
        if not title or not standfirst:
            problems.append("%s: missing title or standfirst" % name)
        leftovers = re.findall(r"\[\[|\]\(|\*\*|IMAGE SPACE|Image register", body)
        if leftovers:
            problems.append("%s: unconverted markup %s" % (name, sorted(set(leftovers))))
        slug = slugify(name)
        portrait = images.get(expected[0], {})
        stem, alt = expected[0], portrait.get("alt", name)
        if portrait.get("kind") == "photograph":
            drawn = [f for f in sorted(os.listdir(DRAWN_PORTRAITS))
                     if f.startswith("%02d-" % number) and f.endswith("-cover.png")]
            if len(drawn) != 1:
                problems.append("%s: no drawn portrait in %s" % (name, DRAWN_PORTRAITS))
            else:
                stem, alt = "%02d-portrait" % number, "Graphite-style illustrated portrait of %s." % name
                drawings[stem] = os.path.join(DRAWN_PORTRAITS, drawn[0])
        facts = "".join("<li>%s</li>" % esc(x) for x in (meta.get("years"), meta.get("firm"), meta.get("focus")) if x)
        pages[slug] = PAGE.format(
            title=esc(title), desc=esc(standfirst), site=SITE, pages=PAGES, slug=slug, stand=inline(standfirst),
            facts=facts, portrait=stem, portrait_alt=esc(alt), body=body, legal=esc(LEGAL))
        copies.update(used)
        person.update({
            "slug": slug, "title": title, "focus": meta.get("focus", ""),
            "portrait": "greats/images/%s-800.webp" % stem,
            "portrait_kind": "illustration",
            "href": "Greats/" + slug,
        })
        print("  %-22s %6d words  images %s  <- %s" % (slug, len(re.sub(r"<[^>]+>", " ", body).split()),
                                                   ",".join(used), os.path.basename(os.path.dirname(path)) + "/" + os.path.basename(path)))

    if problems:
        print("\nPROBLEMS:")
        for problem in problems:
            print("  - " + problem)
        return 1
    if args.check:
        print("\ncheck only: nothing written (%d pages, %d images)" % (len(pages), len(copies)))
        return 0

    os.makedirs(os.path.join(OUT, "images"), exist_ok=True)
    for slug, page in pages.items():
        with open(os.path.join(OUT, slug + ".html"), "w", encoding="utf-8", newline="\n") as fh:
            fh.write(page)
    for ident in sorted(copies):
        for suffix in (".webp", "-800.webp"):
            shutil.copyfile(os.path.join(args.source, "Images", "web", ident + suffix),
                            os.path.join(OUT, "images", ident + suffix))
    if drawings:
        from PIL import Image
        for stem, source in sorted(drawings.items()):
            image = Image.open(source).convert("RGB")
            image.save(os.path.join(OUT, "images", stem + ".webp"), "WEBP", quality=86, method=6)
            image.resize((800, 800), Image.LANCZOS).save(
                os.path.join(OUT, "images", stem + "-800.webp"), "WEBP", quality=86, method=6)
    with open(REGISTRY, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(registry, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    print("\nwrote %d pages and %d images to %s" % (len(pages), len(copies) * 2, OUT))
    return 0


if __name__ == "__main__":
    sys.exit(main())
