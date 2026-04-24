# Long Reads

Standalone scroll-driven editorial articles. Separate from company write-ups in `/articles/`.

## How this differs from company write-ups

| | Company write-ups (`/articles/`) | Long reads (`/longreads/`) |
|---|---|---|
| Publishing | Auto-pushed via GitHub Action to Base44 API | Manual Base44 entity entry with `external_url` |
| Display | Iframed inside Base44 page | Linked to directly, opens as standalone page |
| Design | Navy/red/gold company research design system | Editorial cream/navy/red system with scroll-driven interactions |
| Header | None (inside iframe) | Replicated site header at top |

## Creating a new long-read

1. Duplicate `_template.html` and rename using the convention `YYYY-MM-topic-slug.html`
   - e.g. `2026-04-housing-outlook.html`
2. Open the file and edit the placeholder blocks marked with `<!-- ... -->` comments
3. Commit and push to the `main` branch
4. Wait ~1–2 minutes for GitHub Pages to rebuild
5. Add a `LongRead` entity record on Base44:
   - `external_url` = `https://pezperry.github.io/equity-blog/longreads/<your-filename>.html`
   - Fill in `title`, `deck`, `published_date`, `category`, etc.
6. Surface it on the main site homepage or long-reads index as needed

## GitHub Action exclusion

The auto-publish workflow at `.github/workflows/` is configured to IGNORE this folder. Only files in `/articles/` get auto-posted to the Base44 API. Do not move long-reads into `/articles/` — they will break the iframe pipeline and create malformed records.

## Template structure

`_template.html` is a single self-contained HTML file with:
- Replicated site header at top
- Scroll progress bar
- Full-viewport hero
- Numbered sections with chapter cards
- IntersectionObserver-driven fade-and-rise animations
- Chart.js reusable chart patterns
- Stat block with count-up animation
- Pull quotes, forecast grid, components table, sources block
- Desktop-only sticky side-rail TOC
- `prefers-reduced-motion` support

All styling uses CSS variables defined at the top. Palette, fonts, and visual language match the editorial "Markets & Equities" aesthetic established in the first article (March 2026 inflation).
