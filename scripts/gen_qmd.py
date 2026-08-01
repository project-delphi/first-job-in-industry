#!/usr/bin/env python3
"""Generate index.qmd (a Quarto revealjs deck) from scripts/deck.json.

Each slide is rebuilt as an absolutely-positioned canvas so text boxes,
images and tables land where they sat in the original PowerPoint. The
markdown heading is kept (it drives slide separation and the URL fragment)
but hidden, since the visible title is drawn as a positioned text box.
"""
import html
import json

IN = "scripts/deck.json"
OUT = "index.qmd"

ANCHOR_TO_FLEX = {"t": "flex-start", "ctr": "center", "b": "flex-end"}
ALGN_TO_CSS = {"l": "left", "ctr": "center", "r": "right", "just": "justify"}


def esc(s):
    return html.escape(s, quote=True)


def render_runs(runs):
    parts = []
    for run in runs:
        if run.get("br"):
            parts.append("<br>")
            continue
        styles = []
        if run.get("b"):
            styles.append("font-weight:bold")
        if run.get("i"):
            styles.append("font-style:italic")
        if run.get("u"):
            styles.append("text-decoration:underline")
        text = esc(run["t"])
        if run.get("href"):
            parts.append(f'<a href="{esc(run["href"])}" style="{";".join(styles)}">{text}</a>')
        elif styles:
            parts.append(f'<span style="{";".join(styles)}">{text}</span>')
        else:
            parts.append(text)
    return "".join(parts)


def render_para(p):
    styles = [
        f"font-size:{p['px']}px",
        f"line-height:{p['lineHeight']}",
        f"text-align:{ALGN_TO_CSS.get(p['algn'], 'left')}",
    ]
    if p["marL"]:
        styles.append(f"margin-left:{p['marL']}px")
    if p["spcBef"]:
        styles.append(f"margin-top:{p['spcBef']}px")
    if p["spcAft"]:
        styles.append(f"margin-bottom:{p['spcAft']}px")

    bullet = ""
    if p["bullet"]:
        # Hanging indent: the bullet span fills exactly the negative indent,
        # so wrapped lines line up with the text, as PowerPoint renders it.
        hang = abs(p["indent"]) or p["px"] * 0.6
        styles.append(f"text-indent:{-hang}px")
        bullet = (
            f'<span class="bu" style="display:inline-block;width:{hang}px;'
            f'text-indent:0;font-size:{p["bulletPx"]}px">{esc(p["bullet"])}</span>'
        )

    return f'<div class="para" style="{";".join(styles)}">{bullet}{render_runs(p["runs"])}</div>'


def render_text_box(el):
    styles = [
        f"left:{el['x']}px",
        f"top:{el['y']}px",
        f"width:{el['w']}px",
        f"height:{el['h']}px",
        f"justify-content:{ANCHOR_TO_FLEX.get(el['anchor'], 'flex-start')}",
        f"padding:{el['pad']['t']}px {el['pad']['r']}px {el['pad']['b']}px {el['pad']['l']}px",
    ]
    paras = "".join(render_para(p) for p in el["paras"])
    cls = "tb title-box" if el.get("isTitle") else "tb"
    return f'<div class="{cls}" style="{";".join(styles)}"><div class="tb-inner">{paras}</div></div>'


def render_image(el):
    styles = [
        f"left:{el['x']}px",
        f"top:{el['y']}px",
        f"width:{el['w']}px",
        f"height:{el['h']}px",
    ]
    return f'<img class="pic" src="images/{el["src"]}" style="{";".join(styles)}" alt="">'


def render_table(el):
    width = sum(el["cols"])
    rows = []
    for i, row in enumerate(el["rows"]):
        tag = "th" if i == 0 else "td"
        cells = []
        for j, c in enumerate(row["cells"]):
            # Quarto's table post-processor drops <colgroup>, so column widths
            # ride on the cells themselves.
            w = el["cols"][j] if j < len(el["cols"]) else None
            attr = f' style="width:{w}px"' if w else ""
            cells.append(f"<{tag}{attr}>{esc(c)}</{tag}>")
        rows.append(f"<tr>{''.join(cells)}</tr>")
    styles = [f"left:{el['x']}px", f"top:{el['y']}px", f"width:{width}px"]
    return f'<table class="deck-table" style="{";".join(styles)}">{"".join(rows)}</table>'


AUTOFIT_SCRIPT = """
<script>
// PowerPoint's "shrink text on overflow": scale any text box whose text is
// taller than its original box, so it can never spill over the shapes around
// it. Boxes are measured only while their slide is on screen, because
// reveal.js hides the others and they would measure as zero.
(function () {
  function fitSlide(section) {
    if (!section) return;
    section.querySelectorAll('.tb').forEach(function (tb) {
      var inner = tb.querySelector('.tb-inner');
      if (!inner) return;
      inner.style.transform = '';
      inner.style.height = '';
      if (!tb.clientHeight) return;
      var cs = getComputedStyle(tb);
      var avail = tb.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
      var needed = inner.scrollHeight;
      if (avail > 0 && needed > avail + 1) {
        inner.style.transform = 'scale(' + (avail / needed) + ')';
        // A transform does not change layout size, so a vertically centred
        // box would still be positioned on the unscaled height and spill.
        // Pinning the height to what is actually drawn keeps it anchored.
        inner.style.height = avail + 'px';
      }
    });
  }
  function hook() {
    if (typeof Reveal === 'undefined' || !Reveal.on) return false;
    var run = function () {
      var go = function () { fitSlide(Reveal.getCurrentSlide()); };
      go();
      // Re-measure once webfonts settle, since metrics can shift.
      if (document.fonts && document.fonts.ready) document.fonts.ready.then(go);
    };
    Reveal.on('ready', run);
    Reveal.on('slidechanged', run);
    Reveal.on('resize', run);
    if (Reveal.isReady && Reveal.isReady()) run();
    return true;
  }
  if (!hook()) {
    window.addEventListener('load', function () {
      if (!hook()) setTimeout(hook, 500);
    });
  }
})();
</script>
"""


def main():
    data = json.load(open(IN))
    slides = data["slides"]
    width, height = data["canvas"]
    deck_title = slides[0]["title"] or "Presentation"

    out = [
        "---",
        f'pagetitle: "{deck_title}"',
        "format:",
        "  revealjs:",
        "    css: styles.css",
        "    theme: simple",
        "    center: false",
        "    slide-number: true",
        "    transition: fade",
        f"    width: {width}",
        f"    height: {height}",
        "    margin: 0",
        "    min-scale: 0.1",
        "    max-scale: 3.0",
        "    embed-resources: true",
        "    include-in-header:",
        "      text: '<meta name=\"color-scheme\" content=\"only light\">'",
        "---",
        "",
    ]

    for s in slides:
        heading = s["title"] or f"Slide {s['index']}"
        out.append(f"## {heading}")
        out.append("")
        out.append("```{=html}")
        for el in s["elements"]:
            if el["kind"] == "text":
                out.append(render_text_box(el))
            elif el["kind"] == "image":
                out.append(render_image(el))
            elif el["kind"] == "table":
                out.append(render_table(el))
        out.append("```")
        out.append("")
        if s["notes"].strip():
            out.append("::: {.notes}")
            out.extend(s["notes"].splitlines())
            out.append(":::")
            out.append("")

    out.append("```{=html}")
    out.append(AUTOFIT_SCRIPT.strip())
    out.append("```")
    out.append("")

    with open(OUT, "w") as f:
        f.write("\n".join(out))
    print(f"Wrote {OUT}: {len(slides)} slides")


if __name__ == "__main__":
    main()
