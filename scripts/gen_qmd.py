#!/usr/bin/env python3
"""Generate index.qmd (a Quarto revealjs deck) from scripts/deck.json."""
import html
import json

IN = "scripts/deck.json"
OUT = "index.qmd"


def esc(s):
    return html.escape(s, quote=True)


def render_bullets(body):
    lines = []
    for lvl, text in body:
        indent = "    " * lvl
        lines.append(f"{indent}- {text}")
    return "\n".join(lines)


def render_media(images):
    if not images:
        return ""
    parts = ['<div class="slide-media">']
    for im in images:
        style = (
            f"left:{im['left_pct']}%;top:{im['top_pct']}%;"
            f"width:{im['width_pct']}%;height:{im['height_pct']}%;"
        )
        parts.append(f'<img src="images/{im["src"]}" style="{style}" alt="">')
    parts.append("</div>")
    return "\n".join(parts)


def render_table(table):
    if not table:
        return ""
    rows = []
    header, *body = table
    rows.append("<tr>" + "".join(f"<th>{esc(c)}</th>" for c in header) + "</tr>")
    for r in body:
        rows.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in r) + "</tr>")
    return '<table class="fidelity-table">\n' + "\n".join(rows) + "\n</table>"


def render_notes(notes):
    if not notes.strip():
        return ""
    body = "\n".join(f"{line}\n" for line in notes.splitlines())
    return f"::: {{.notes}}\n{body}\n:::"


def main():
    data = json.load(open(IN))
    slides = data["slides"]
    deck_title = slides[0]["title"] or "Presentation"

    out = []
    out.append("---")
    out.append(f'pagetitle: "{deck_title}"')
    out.append("format:")
    out.append("  revealjs:")
    out.append("    css: styles.css")
    out.append("    theme: simple")
    out.append("    center: false")
    out.append("    slide-number: true")
    out.append("    transition: fade")
    out.append("    width: 1280")
    out.append("    height: 720")
    out.append("    margin: 0.06")
    out.append("    embed-resources: true")
    out.append('    include-in-header:')
    out.append('      text: \'<meta name="color-scheme" content="only light">\'')
    out.append("---")
    out.append("")

    for i, s in enumerate(slides):
        title = s["title"] or ""
        out.append(f"## {title}")
        out.append("")

        bullets = render_bullets(s["body"])
        if bullets:
            out.append(bullets)
            out.append("")

        table_html = render_table(s["table"])
        if table_html:
            out.append(table_html)
            out.append("")

        media_html = render_media(s["images"])
        if media_html:
            out.append(media_html)
            out.append("")

        notes_md = render_notes(s["notes"])
        if notes_md:
            out.append(notes_md)
            out.append("")

    with open(OUT, "w") as f:
        f.write("\n".join(out))
    print(f"Wrote {OUT}: {len(slides)} slides")


if __name__ == "__main__":
    main()
