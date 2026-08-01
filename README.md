# The First Job in Industry

A talk on navigating the transition from academia into an industry job: what
industry looks like, building a digital presence, getting interview
invitations, the interview process itself, and handling offers.

**View the slides:** https://project-delphi.github.io/first-job-in-industry/

Use the arrow keys (or swipe) to move between slides, `S` for speaker notes,
and `?` for the full list of keyboard shortcuts.

## About this repo

This is a [Quarto](https://quarto.org) `revealjs` presentation, converted
from the original `Getting the first job in Industry.pptx` (kept in this repo
for provenance). The conversion is scripted rather than manual:

- `scripts/extract_pptx.py` parses the `.pptx` XML directly (titles, bullet
  text, image positions/sizes, the one table, hyperlinks, speaker notes) into
  `scripts/deck.json`.
- `scripts/gen_qmd.py` turns that JSON into `index.qmd`, reproducing each
  slide's original image layout via absolutely-positioned overlays so the
  deck keeps fidelity to the source PowerPoint.
- `styles.css` matches the original theme (white background, Arial, black
  text, teal `#158158` accent).

## Re-rendering

If the deck is ever edited, regenerate and rebuild the site with:

```sh
python3 scripts/extract_pptx.py   # only needed if the .pptx itself changed
python3 scripts/gen_qmd.py        # only needed if the .pptx itself changed
quarto render                     # rebuilds docs/ for GitHub Pages
```

Then commit the updated `docs/` folder — GitHub Pages serves directly from
`main:/docs`, there's no CI/CD step.
