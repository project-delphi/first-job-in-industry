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

- `scripts/extract_pptx.py` parses the `.pptx` XML directly into
  `scripts/deck.json`. Every shape keeps its original position and size, and
  text styling (font size, bullets, indents, line spacing, alignment) is
  resolved through PowerPoint's inheritance chain: master placeholder ->
  layout placeholder -> shape -> paragraph -> run.
- `scripts/gen_qmd.py` turns that JSON into `index.qmd`. Each slide is
  rebuilt as an absolutely-positioned 1280x720 canvas, so shapes land exactly
  where they sat in the source deck instead of reflowing into each other.
  It also emits a small "shrink text on overflow" script mirroring
  PowerPoint's autofit, so no text box can spill over its neighbours.
- `styles.css` matches the original theme (white background, Arial, black
  text, teal `#158158` accent).

Two deliberate departures from the source deck, both for readability: text is
drawn above artwork (a few slides stack a picture over a text box, hiding the
words), and text that overflows its box is scaled to fit.

## Re-rendering

If the deck is ever edited, regenerate and rebuild the site with:

```sh
python3 scripts/extract_pptx.py   # only needed if the .pptx itself changed
python3 scripts/gen_qmd.py        # only needed if the .pptx itself changed
quarto render                     # rebuilds docs/ for GitHub Pages
```

Then commit the updated `docs/` folder — GitHub Pages serves directly from
`main:/docs`, there's no CI/CD step.
