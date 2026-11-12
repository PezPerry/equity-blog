#!/usr/bin/env python3
"""Write the figures behind each graphic on Images & Infographics as CSV files.

Every graphic page in this repo carries its own figures inline. This script reads
them out of the page (it never retypes a number) and writes tidy CSV files plus
an index.json to the folder given with --out.

The output is NOT published here. It goes to the private content repo
(PezPerry/equity-content, folder graphics-data/), where the app's graphicData
function reads it for Premium members:

    python build-graphic-data.py --out ../../rns/equity-content/graphics-data

Run it again whenever a graphic page changes. --check writes nothing and exits 1
if the files in --out are out of date.
"""
import argparse
import csv
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def read(name):
    with open(os.path.join(HERE, name), encoding="utf-8") as fh:
        return fh.read()


def js_array(page, name):
    """A one-line `const NAME=[...]` array of numbers or quoted strings."""
    match = re.search(r"(?:const|var|let)\s+%s\s*=\s*(\[[^\]]*\])" % re.escape(name), page)
    if not match:
        raise SystemExit("ERROR: array %s not found" % name)
    return json.loads(match.group(1).replace("'", '"'))


def to_csv(header, rows):
    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\n")
    writer.writerow(header)
    for row in rows:
        writer.writerow(["" if v is None else v for v in row])
    return out.getvalue()


def uk_house_prices():
    page = read("uk_house_prices.html")
    y10, nominal, cpi = js_array(page, "Y10"), js_array(page, "NP"), js_array(page, "CPI")
    years, ratio, txns = js_array(page, "YL"), js_array(page, "LR"), js_array(page, "TX")
    if not (len(y10) == len(nominal) == len(cpi) and len(years) == len(ratio) == len(txns)):
        raise SystemExit("ERROR: uk_house_prices arrays differ in length")
    # Same sums as the page: a CPI index starting at 100 in the first year, and
    # the nominal price deflated by it.
    index = [100.0]
    for rate in cpi[1:]:
        index.append(index[-1] * (1 + rate / 100))
    real = [round(p / (i / 100)) for p, i in zip(nominal, index)]
    recent = {y: (n, c, r) for y, n, c, r in zip(y10, nominal, cpi, real)}
    rows = []
    for year, pe, tx in zip(years, ratio, txns):
        n, c, r = recent.get(year, (None, None, None))
        rows.append([year, n, c, r, "%.2f" % pe, tx])
    header = ["year", "nominal_price_gbp", "cpi_inflation_pct", "real_price_2015_gbp",
              "price_to_earnings_ratio", "transactions_thousands"]
    return [("uk-house-prices.csv", "Prices, inflation, affordability and transactions by year",
             to_csv(header, rows))]


def greggs():
    page = read("greggs-store-map.html")
    match = re.search(r"var pts = (\[\[.*?\]\]);", page, re.S)
    if not match:
        raise SystemExit("ERROR: Greggs points not found")
    points = json.loads(match.group(1))
    return [("greggs-shops.csv", "Latitude and longitude of every mapped shop",
             to_csv(["latitude", "longitude"], points))]


def crest():
    page = read("crest-nicholson-CRST-sites-map.html")
    pattern = re.compile(
        r"L\.circleMarker\(\[([-\d.]+),([-\d.]+)\][^\n]*\n\s*\.bindPopup\(\"<strong>(.*?)</strong><br>(.*?)<br><em>(.*?) · (.*?)</em>\"\)")
    rows = [[name, town, region, status, lat, lon]
            for lat, lon, name, town, region, status in pattern.findall(page)]
    if len(rows) != page.count("L.circleMarker("):
        raise SystemExit("ERROR: Crest Nicholson: %d markers but %d read"
                         % (page.count("L.circleMarker("), len(rows)))
    return [("crest-nicholson-developments.csv", "Every development with its town, region and status",
             to_csv(["development", "town", "region", "status", "latitude", "longitude"], rows))]


def iran():
    page = read("iranian_crisis_energy_map.html")
    start = page.find("const incidents = [")
    end = page.find("\n];", start)
    if start == -1 or end == -1:
        raise SystemExit("ERROR: Iranian crisis incidents not found")
    fields = ["name", "country", "category", "date", "lat", "lon", "what", "impact",
              "source_label", "source_url"]
    rows = []
    for line in page[start:end].splitlines():
        line = line.strip().rstrip(",")
        if line.startswith("{"):
            item = json.loads(line)
            rows.append([item.get(f) for f in fields])
    if not rows:
        raise SystemExit("ERROR: Iranian crisis: no incidents read")
    header = ["asset", "country", "category", "date", "latitude", "longitude", "what_happened",
              "why_it_matters", "source", "source_url"]
    return [("iranian-crisis-energy-assets.csv", "Every incident with its date, place and source",
             to_csv(header, rows))]


def hormuz():
    page = read("strait-of-hormuz-tracking-the-ships.html")
    match = re.search(r"window\.HORMUZ = (\{.*?\});?\s*\n", page, re.S)
    if not match:
        raise SystemExit("ERROR: Hormuz data not found")
    data = json.loads(match.group(1))

    def table(key, fields):
        return [[row.get(f) for f in fields] for row in data.get(key, [])]

    return [
        ("hormuz-daily-transits.csv", "Daily transits: total, tankers and cargo ships",
         to_csv(["date", "total", "tankers", "cargo", "seven_day_average", "seven_day_average_tankers", "status"],
                table("daily", ["d", "total", "tanker", "cargo", "avg7", "avg7t", "status"]))),
        ("hormuz-direction.csv", "Daily transits by direction",
         to_csv(["date", "outbound", "inbound", "seven_day_average_outbound", "seven_day_average_inbound"],
                table("dir", ["d", "out", "inb", "avgOut", "avgIn"]))),
        ("hormuz-reported-counts.csv", "Daily counts reported in the press, with the source",
         to_csv(["date", "vessels", "inbound", "outbound", "note", "source_url"],
                table("kpler", ["d", "n", "inb", "out", "obs", "url"]))),
        ("hormuz-periods.csv", "Transits by period of the crisis",
         to_csv(["period", "start", "end", "days", "transits", "transits_per_day", "inbound", "outbound",
                 "crude_share", "dark_share", "source_url"],
                table("axs", ["name", "start", "end", "days", "n", "rate", "inb", "out", "crude", "dark", "url"]))),
    ]


# slug -> (title, builder). The slug is the `data` value on the tile in the
# app's src/lib/graphics.js.
GRAPHICS = [
    ("strait-of-hormuz", "Strait of Hormuz: Tracking the Ships", hormuz),
    ("crest-nicholson-sites", "Crest Nicholson: UK development map", crest),
    ("greggs-store-map", "Greggs: the estate on the map", greggs),
    ("uk-house-prices", "UK House Prices: The Real Picture (2015 to 2025)", uk_house_prices),
    ("iranian-crisis-energy-map", "The Iranian Crisis: Energy Assets Affected", iran),
]


def build():
    files = {}
    index = {"graphics": {}}
    for slug, title, builder in GRAPHICS:
        entries = []
        for name, label, text in builder():
            files[name] = text
            entries.append({"file": name, "label": label, "rows": text.count("\n") - 1})
        index["graphics"][slug] = {"title": title, "files": entries}
    files["index.json"] = json.dumps(index, indent=1, ensure_ascii=False) + "\n"
    return files


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--out", required=True, metavar="FOLDER")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()

    files = build()
    stale = []
    for name, text in files.items():
        path = os.path.join(args.out, name)
        current = None
        if os.path.exists(path):
            with open(path, encoding="utf-8", newline="") as fh:
                current = fh.read()
        if current != text:
            stale.append(name)
    if args.check:
        if stale:
            print("Graphic data is OUT OF DATE: %s" % ", ".join(stale))
            return 1
        print("Graphic data is up to date.")
        return 0
    os.makedirs(args.out, exist_ok=True)
    for name, text in files.items():
        with open(os.path.join(args.out, name), "w", encoding="utf-8", newline="\n") as fh:
            fh.write(text)
    for name in sorted(files):
        print("  %-40s %s" % (name, "written" if name in stale else "unchanged"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
