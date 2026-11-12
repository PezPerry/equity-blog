#!/usr/bin/env python3
"""Make a quiz tile cover: the company's approved logo on its approved ground,
over a navy "numbers quiz" band. No image is generated; this only composes
approved artwork, so every cover matches.

    python quizzes/make-cover.py --company "Watches of Switzerland Group PLC" --ticker WOSG --slug watches-of-switzerland-wosg

--company is the file name (without .json) of the entry in the logo library,
C:\\Master Control\\System\\Assets\\Logos. With no logo entry the cover shows the
ticker in the house serif instead, and says so on exit.
"""
import argparse
import json
import os
import sys

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
LOGOS = r"C:\Master Control\System\Assets\Logos"
FONTS = (r"C:\Master Control\System\Assets\E&MI Approved Graphics\2026-11-03"
         r"\Equity-and-Markets-Insight-Approved-Graphics\08-Fonts")
NAVY, IVORY, GOLD, LIGHT = "#112238", "#F8F5EF", "#A98A48", "#E3E2DC"
W, H, BAND = 1200, 800, 190


def font(name, size, weight):
    face = ImageFont.truetype(os.path.join(FONTS, name), size)
    try:
        face.set_variation_by_axes([weight] if "Cormorant" in name else [14, weight])
    except (OSError, ValueError):
        pass
    return face


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--company", required=True)
    parser.add_argument("--ticker", required=True)
    parser.add_argument("--slug", required=True)
    args = parser.parse_args()

    ground, logo, note = IVORY, None, ""
    entry = os.path.join(LOGOS, args.company + ".json")
    if os.path.exists(entry):
        with open(entry, encoding="utf-8") as fh:
            meta = json.load(fh)
        dark = meta.get("background_mode") == "dark"
        name = meta.get("reversed_file") if dark and meta.get("reversed_file") else meta.get("logo_file")
        path = os.path.join(LOGOS, name or "")
        if name and name.lower().endswith(".png") and os.path.exists(path):
            logo = Image.open(path).convert("RGBA")
            ground = meta.get("background_hex") or (NAVY if dark else IVORY)
        else:
            note = "logo entry has no PNG artwork"
    else:
        note = "no logo entry for '%s'" % args.company

    card = Image.new("RGB", (W, H), ground)
    draw = ImageDraw.Draw(card)
    top = H - BAND
    if logo:
        box_w, box_h = int(W * 0.58), int(top * 0.46)
        scale = min(box_w / logo.width, box_h / logo.height)
        logo = logo.resize((max(1, round(logo.width * scale)), max(1, round(logo.height * scale))), Image.LANCZOS)
        card.paste(logo, ((W - logo.width) // 2, (top - logo.height) // 2), logo)
    else:
        card.paste(LIGHT, (0, 0, W, top))
        face = font("CormorantGaramond[wght].ttf", 230, 600)
        draw.text((W / 2, top / 2), args.ticker, font=face, fill=NAVY, anchor="mm")

    draw.rectangle((0, top, W, H), fill=NAVY)
    draw.rectangle((0, top, W, top + 6), fill=GOLD)
    draw.text((64, top + BAND / 2 + 2), "The numbers quiz", font=font("CormorantGaramond[wght].ttf", 76, 600),
              fill=IVORY, anchor="lm")
    draw.text((W - 190, top + BAND / 2 + 4), "8 QUESTIONS", font=font("Inter[opsz,wght].ttf", 24, 600),
              fill=GOLD, anchor="rm")
    r = 52
    cx, cy = W - 64 - r, top + BAND // 2 + 3
    draw.ellipse((cx - r, cy - r, cx + r, cy + r), fill=GOLD)
    draw.text((cx, cy - 4), "?", font=font("CormorantGaramond[wght].ttf", 84, 700), fill=NAVY, anchor="mm")

    out = os.path.join(HERE, "covers", args.slug + ".webp")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    card.save(out, "WEBP", quality=88, method=6)
    print("wrote quizzes/covers/%s.webp%s" % (args.slug, (" (ticker cover: %s)" % note) if note else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
