# UK Autumn Budget 2026 — cover images (prompt for Codex)

> **Done 29 Oct 2026.** The articles shipped as `budget-2026-*.html` (not the
> `uk-budget-2026-*` names planned below); each cover is `<article>-cover.jpg`,
> e.g. `budget-2026-overlooked-winners-cover.jpg`. Kept for reference.

Paste everything in the box below into Codex, run from the root of the `equity-blog` repo.

---

```text
Generate 7 cover images for a series of UK Autumn Budget 2026 articles on the
Equity & Markets Insight website, and save each one in the repo root under the
exact file name given. They will appear together in one front-page section, so
they must read as ONE matching series.

OUTPUT
- JPG (quality ~85, keep each under ~400 KB), 1536 x 1024 px (3:2 landscape).
  Save to the repo root, exact names below.
- Main subject centred, with ~10% clear margin on every side: the site crops
  nothing, but the images are shown small (about 380 px wide) on phones, so the
  central symbol must read at thumbnail size.

HOUSE STYLE (same for all 7)
- Editorial illustration in the manner of a fine banknote / copperplate
  engraving: crisp ink linework and cross-hatching, not photographic, not 3D
  render, not cartoon.
- Background: deep navy #1B2A4A, with a soft warm glow from the top-right corner.
- Linework and highlights in antique gold #D4A853 (deeper gold #A9812E for
  shadows), cream #F3ECDD for the brightest highlights.
- ONE spot colour only: the red of the Chancellor's Budget box, #C8102E, used
  only where the brief mentions it. Everything else stays navy, gold and cream.
- A faint thin gold rule framing the image about 24 px in from the edge.
- Calm, considered, premium: think the Financial Times or The Economist, not a
  tabloid. Even lighting, generous negative space, one clear focal object.

DO NOT
- No words, letters, numbers, tickers or currency figures anywhere in the image
  (the site prints the titles; generated text comes out garbled). A plain "%"
  or "£" symbol is the only exception, and only where the brief asks for it.
- No real people or likenesses (no Chancellor, no politicians).
- No company logos, brand names, or recognisable branded shopfronts.
- No royal coat of arms; any crest on the Budget box is a plain, generic gold
  emblem.
- No stock-market chart clichés filling the frame (a faint line in the
  background is fine where asked).

THE SEVEN IMAGES

1) uk-budget-2026-ftse-winners-losers-cover.jpg
   Article: "UK Budget 2026: Winners, Losers and the Changes That Really Matter
   to Investors" (overview of the whole Budget; shown largest, as the lead).
   Image: the Chancellor's red leather Budget despatch box (the one red object),
   closed, sitting on a ledge. Rising from behind it, a large set of antique
   brass balance scales, one pan tipped high and one tipped low. Behind
   everything, a very faint gold line chart running across the navy.

2) uk-budget-2026-bank-tax-lloyds-natwest-cover.jpg
   Article: banks face a temporary increase in the Bank Corporation Tax
   Surcharge (3% to 6% for 15 months, about £1.5bn).
   Image: a generic classical bank portico (columns and pediment, no
   inscription). In front, a tall stack of gold coins with the top few coins
   sliding off into a small open red Budget box. A single large engraved "%"
   symbol carved faintly into the pediment.

3) uk-budget-2026-housebuilders-first-home-cover.jpg
   Article: the £4.7bn "Your First Home" equity-loan scheme for first-time
   buyers and what it means for housebuilders.
   Image: a short row of new-build British terraced houses, the last one still
   in scaffolding with a tower crane above. Foreground: a single front-door key
   on a plain ring with a blank paper tag, catching the light. Warm gold glow in
   one finished window.

4) uk-budget-2026-pubs-retail-business-rates-cover.jpg
   Article: bigger business-rates relief for pubs (15% to 35%), but a higher
   National Living Wage could swallow the saving.
   Image: a traditional British pub bar counter with a full pint of ale. Beside
   it, a small set of brass scales: one pan holds a neat stack of coins (the
   rates saving), the other holds a plain brown wage envelope; the scales are
   almost exactly level. A blank hanging pub sign glimpsed through the window
   behind.

5) uk-budget-2026-gambling-tax-shares-cover.jpg
   Article: Machine Games Duty rises from 20% to 25% on higher-stakes gaming
   machines.
   Image: a close view of gaming-machine reels showing simple engraved symbols
   (cherries, bell, star), with two dice and a short stack of casino chips in
   front. Across the chips, a round official-looking tax stamp impression in
   gold (blank, no wording).

6) uk-budget-2026-manufacturing-energy-investment-cover.jpg
   Article: up to 25% electricity discounts for energy-intensive manufacturers,
   new investment finance, and a proposed jet-engine programme.
   Image: in the foreground, the front of a jet-engine turbine fan, blades
   engraved in fine detail. Behind it, a brick factory with a saw-tooth roof and
   chimneys, and an electricity pylon whose cables carry a small gold spark of
   current. A few interlocking gears at the base.

7) uk-budget-2026-small-impact-ftse-shares-cover.jpg
   Article: household measures (a £100 energy payment, fuel-duty freeze, gas
   standing charges) that matter to families but barely move FTSE profits.
   Image: a single gold coin dropping into a perfectly still pond, sending out
   only a few small, neat ripples. Reflected faintly in the water, the skyline
   of the City of London (generic towers, no identifiable logos). Lots of calm,
   empty water: the point is how small the ripple is.

CONSISTENCY CHECK before saving
- Same navy, same gold, same engraving line weight, same frame rule on all 7.
- Red appears only in images 1 and 2.
- No text or numbers crept in (except the single "%" in image 2).
- Exact file names, 1536 x 1024, JPG.
```

---

## Publishing checklist for tomorrow (29 Oct)

The front-page section is already in `front-page.html` (between the
`BUDGET-2026:START` / `END` markers). For each article as it is finished:

1. **File name** — save the article in the repo root with the name the card
   already links to:

   | # | Article file | Cover file |
   |---|---|---|
   | 01 | `uk-budget-2026-ftse-winners-losers.html` | `uk-budget-2026-ftse-winners-losers-cover.jpg` |
   | 02 | `uk-budget-2026-bank-tax-lloyds-natwest.html` | `uk-budget-2026-bank-tax-lloyds-natwest-cover.jpg` |
   | 03 | `uk-budget-2026-housebuilders-first-home.html` | `uk-budget-2026-housebuilders-first-home-cover.jpg` |
   | 04 | `uk-budget-2026-pubs-retail-business-rates.html` | `uk-budget-2026-pubs-retail-business-rates-cover.jpg` |
   | 05 | `uk-budget-2026-gambling-tax-shares.html` | `uk-budget-2026-gambling-tax-shares-cover.jpg` |
   | 06 | `uk-budget-2026-manufacturing-energy-investment.html` | `uk-budget-2026-manufacturing-energy-investment-cover.jpg` |
   | 07 | `uk-budget-2026-small-impact-ftse-shares.html` | `uk-budget-2026-small-impact-ftse-shares-cover.jpg` |

2. **Tag it as Budget coverage** — put this in each article's `<head>`:

   ```html
   <meta name="tile-kind" content="budget">
   ```

   Without it, the publish workflow will also drop the article into the normal
   front-page tiles and push company analyses off the page.

3. **Switch the card on** — in `front-page.html`, delete the word `soon` from
   that article's card (`class="block soon"` → `class="block"`, or
   `class="bud-lead soon"` → `class="bud-lead"` for 01). That makes it
   clickable and turns "Publishing 29 Oct 2026" into "29 Oct 2026".

4. If a final headline changes, update the `<h4>` (or `<h3>` for 01) on its card.

Push the articles, covers and `front-page.html` together. The workflow still
publishes each article to Base44 (All Articles) as before.
