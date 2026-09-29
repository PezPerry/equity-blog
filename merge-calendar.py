#!/usr/bin/env python3
"""
Financial Calendar top-up — merge fresh ShareScope Calendar exports into
financial-calendar.html without wiping anything, then rebuild calendar-index.json.

Usage:
  python merge-calendar.py FTSE350.csv NASDAQ100.csv            # merge + write
  python merge-calendar.py FTSE350.csv NASDAQ100.csv --dry-run  # report only

How it works (same process as the Sep 2026 top-ups):
  * Export the Calendar from ShareScope for the FTSE 350 and NASDAQ 100 lists
    (CSV, with or without a header row).
  * Every existing row in the page's rawData block is kept. Incoming rows are
    added only if the same event is not already there for that ticker on that
    date (ignoring an "(Est)"/"(Conf)" tag and company-name spelling), then
    everything is ordered by date and ticker - existing rows never move
    relative to each other.
  * calendar-index.json (used by the RNS News "next event" chips) is rebuilt:
    for each ticker in companies.json, the earliest event from today onwards.
    Tickers with no upcoming calendar row keep their existing entry if it is
    still in the future.

Columns are detected per row, so either export layout works: the date is the
field that parses as a date, the ticker is the short upper-case code (EPIC or
NASDAQ symbol), the company is the longest remaining field and the event is
what is left. With a header row, columns named like date / name|company /
ticker|epic|code|symbol / event|type|description are used directly.
"""

import argparse, csv, json, os, re, sys
from collections import Counter
from datetime import date, datetime

REPO_DIR = os.path.dirname(os.path.abspath(__file__))
PAGE = os.path.join(REPO_DIR, "financial-calendar.html")
INDEX = os.path.join(REPO_DIR, "calendar-index.json")
COMPANIES = os.path.join(REPO_DIR, "companies.json")
START = "const rawData = `"
DATE_FORMATS = ("%d-%b-%Y", "%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d %b %Y",
                "%d %B %Y", "%d-%b-%y", "%a %d %b %Y", "%a %d/%m/%Y")
TICKER_RE = re.compile(r"^[A-Z0-9]{1,5}(\.[A-Z]?)?$")
EVENT_RE = re.compile(r"\b(div|dividend|results?|agm|egm|gm|year end|report|statement|"
                      r"trading|update|split|consolidation|q[1-4]|interim|final|est|conf)\b", re.I)


def parse_date(s):
    s = s.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def fmt_date(d):
    return d.strftime("%d-%b-%Y")


def row_from_header(rec, cols):
    d = parse_date(rec[cols["date"]])
    if not d:
        return None
    return (rec[cols["event"]].strip(), d, rec[cols["company"]].strip(),
            rec[cols["ticker"]].strip())


def row_by_shape(rec, known_events=()):
    fields = [f.strip() for f in rec if f and f.strip()]
    d, rest = None, []
    for f in fields:
        if d is None and parse_date(f):
            d = parse_date(f)
        else:
            rest.append(f)
    if d is None or len(rest) < 3:
        return None
    # Event first, so a bare "AGM" is not mistaken for a ticker.
    events = [f for f in rest if f in known_events] or [f for f in rest if EVENT_RE.search(f)]
    if not events:
        return None
    event = events[0]
    rest.remove(event)
    tickers = [f for f in rest if TICKER_RE.match(f)]
    if not tickers:
        return None
    ticker = tickers[0]
    rest.remove(ticker)
    # Unquoted commas inside a company name split it across fields; rejoin in order.
    return (event, d, ", ".join(rest), ticker)


def header_columns(first):
    names = [h.strip().lower() for h in first]
    cols = {}
    for i, h in enumerate(names):
        if "date" in h and "date" not in cols:
            cols["date"] = i
        elif h in ("name", "company", "company name") and "company" not in cols:
            cols["company"] = i
        elif h in ("ticker", "epic", "code", "symbol", "tidm") and "ticker" not in cols:
            cols["ticker"] = i
        elif h in ("event", "type", "event type", "description", "category") and "event" not in cols:
            cols["event"] = i
    return cols if len(cols) == 4 else None


def read_export(path, known_events=()):
    with open(path, newline="", encoding="utf-8-sig", errors="replace") as f:
        records = [r for r in csv.reader(f) if any(c.strip() for c in r)]
    if not records:
        return [], 0
    cols = header_columns(records[0])
    body = records[1:] if cols or not parse_date_any(records[0]) else records
    rows, skipped = [], 0
    for rec in body:
        row = row_from_header(rec, cols) if cols else row_by_shape(rec, known_events)
        if row and all(row[i] for i in (0, 2, 3)) and "|" not in "".join((row[0], row[2], row[3])):
            rows.append(row)
        else:
            skipped += 1
    return rows, skipped


def parse_date_any(rec):
    return any(parse_date(c) for c in rec)


def base_event(cat):
    return re.sub(r"\s*\((Est|Conf)\)$", "", cat)


def load_page():
    html = open(PAGE, encoding="utf-8").read()
    s = html.index(START) + len(START)
    e = html.index("`;", s)
    lines = [l for l in html[s:e].split("\n") if l.strip()]
    return html, s, e, lines


def line_key(line):
    cat, d, company, ticker = line.split("|")
    return (datetime.strptime(d, "%d-%b-%Y").date(), ticker)


def rebuild_index(lines, today):
    companies = json.load(open(COMPANIES))
    old = {}
    if os.path.exists(INDEX):
        old = json.load(open(INDEX)).get("by_ticker", {})
    nxt = {}
    for line in lines:
        cat, d, company, ticker = line.split("|")
        dt = datetime.strptime(d, "%d-%b-%Y").date()
        if ticker in companies and dt >= today:
            if ticker not in nxt or dt < nxt[ticker][0]:
                nxt[ticker] = (dt, cat)
    by_ticker = {t: {"date": dt.isoformat(), "event": cat} for t, (dt, cat) in nxt.items()}
    for t, v in old.items():
        if t not in by_ticker and v.get("date", "") >= today.isoformat():
            by_ticker[t] = v
    return {"generated": today.isoformat(), "by_ticker": dict(sorted(by_ticker.items()))}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("exports", nargs="+", help="ShareScope Calendar CSV exports")
    ap.add_argument("--dry-run", action="store_true", help="report what would be added; write nothing")
    ap.add_argument("--today", help="YYYY-MM-DD override for the index rebuild")
    args = ap.parse_args()
    today = date.fromisoformat(args.today) if args.today else date.today()

    html, s, e, lines = load_page()
    existing = set(lines)
    # Same event on the same day, differing only by an (Est)/(Conf) tag, is not new.
    seen = {(base_event(l.split("|")[0]), l.split("|")[1], l.split("|")[3]) for l in lines}
    known_events = {l.split("|")[0] for l in lines}
    added = []
    for path in args.exports:
        rows, skipped = read_export(path, known_events)
        new = 0
        for cat, d, company, ticker in rows:
            line = f"{cat}|{fmt_date(d)}|{company}|{ticker}"
            key = (base_event(cat), fmt_date(d), ticker)
            if line not in existing and key not in seen:
                existing.add(line)
                seen.add(key)
                added.append(line)
                new += 1
        print(f"{os.path.basename(path)}: {len(rows)} rows, {new} new, {skipped} skipped")

    merged = sorted(lines + added, key=line_key)  # stable: existing order kept
    months = Counter(line_key(l)[0].strftime("%Y-%m") for l in added)
    print(f"Total {len(lines)} -> {len(merged)} (+{len(added)})")
    for m, n in sorted(months.items()):
        print(f"  {m}: +{n}")
    if added:
        last = max(line_key(l)[0] for l in merged)
        print(f"Coverage now runs to {last.isoformat()}")

    index = rebuild_index(merged, today)
    print(f"calendar-index.json: {len(index['by_ticker'])} tickers")
    if args.dry_run:
        return
    with open(PAGE, "w", encoding="utf-8") as f:
        f.write(html[:s] + "\n".join(merged) + html[e:])
    with open(INDEX, "w") as f:
        json.dump(index, f, indent=2)
        f.write("\n")


if __name__ == "__main__":
    sys.exit(main())
