#!/usr/bin/env python3
"""Extract slide content from the source .pptx into structured JSON for gen_qmd.py.

Every shape keeps its original position, size and text styling, resolved
through the PowerPoint inheritance chain (master placeholder -> layout
placeholder -> shape -> paragraph -> run). Geometry is emitted in pixels on a
1280x720 canvas, which is the reveal.js slide size the deck renders at.
"""
import json
import re
import zipfile
import xml.etree.ElementTree as ET

PPTX = "Getting the first job in Industry.pptx"
OUT = "scripts/deck.json"

CANVAS_W = 1280
CANVAS_H = 720

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

# Placeholder categories map onto the master's three text style blocks.
TITLE_TYPES = {"title", "ctrTitle"}
BODY_TYPES = {"body", "subTitle"}


def qn(tag):
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def emu_px(v):
    """EMU -> px on the 1280x720 canvas (uniform: the deck is exactly 16:9)."""
    return round(v / 7143.75, 2)


def pt_px(hundredths):
    """Font size in hundredths of a point -> px on the canvas."""
    return round(hundredths / 100 * (CANVAS_W / 720), 2)


def parse_lvl_props(el):
    """Read the paragraph-level properties we care about off an a:*pPr element."""
    if el is None:
        return {}
    props = {}
    for attr in ("marL", "indent", "algn"):
        if el.get(attr) is not None:
            props[attr] = el.get(attr)
    lnspc = el.find("a:lnSpc/a:spcPct", NS)
    if lnspc is not None:
        props["lnSpc"] = int(lnspc.get("val"))
    for name, tag in (("spcBef", "a:spcBef/a:spcPts"), ("spcAft", "a:spcAft/a:spcPts")):
        sp = el.find(tag, NS)
        if sp is not None:
            props[name] = int(sp.get("val"))
    if el.find("a:buNone", NS) is not None:
        props["buChar"] = None
    buchar = el.find("a:buChar", NS)
    if buchar is not None:
        props["buChar"] = buchar.get("char")
    busz = el.find("a:buSzPts", NS)
    if busz is not None:
        props["buSzPts"] = int(busz.get("val"))
    defrpr = el.find("a:defRPr", NS)
    if defrpr is not None:
        props.update(parse_run_props(defrpr))
    return props


def parse_run_props(rpr):
    """Read run properties (size / bold / italic / underline) off an a:rPr."""
    if rpr is None:
        return {}
    props = {}
    if rpr.get("sz"):
        props["sz"] = int(rpr.get("sz"))
    if rpr.get("b") is not None:
        props["b"] = rpr.get("b") == "1"
    if rpr.get("i") is not None:
        props["i"] = rpr.get("i") == "1"
    if rpr.get("u") is not None and rpr.get("u") != "none":
        props["u"] = True
    return props


def parse_lst_style(lst):
    """Map level index -> properties for an a:lstStyle element."""
    out = {}
    if lst is None:
        return out
    for lvl in range(9):
        el = lst.find(f"a:lvl{lvl + 1}pPr", NS)
        if el is not None:
            out[lvl] = parse_lvl_props(el)
    return out


def ph_key(ph):
    """Normalised placeholder identity. Title variants share one slot."""
    if ph is None:
        return None
    t = ph.get("type") or "body"
    if t in TITLE_TYPES:
        return ("title", None)
    return (t, ph.get("idx"))


def category_for(ph):
    if ph is None:
        return "other"
    t = ph.get("type") or "body"
    if t in TITLE_TYPES:
        return "title"
    if t in BODY_TYPES:
        return "body"
    return "other"


def shape_xfrm(sp, path="./p:spPr/a:xfrm"):
    xfrm = sp.find(path, NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))


class Transform:
    """Affine map (no rotation) from a shape's local EMU space to slide EMU space."""

    def __init__(self, ox=0, oy=0, sx=1.0, sy=1.0):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def rect(self, x, y, cx, cy):
        return (
            self.ox + x * self.sx,
            self.oy + y * self.sy,
            cx * self.sx,
            cy * self.sy,
        )

    def child(self, off, ext, ch_off, ch_ext):
        ax, ay, acx, acy = self.rect(*off, *ext)
        sx = acx / ch_ext[0] if ch_ext[0] else 1.0
        sy = acy / ch_ext[1] if ch_ext[1] else 1.0
        return Transform(ax - ch_off[0] * sx, ay - ch_off[1] * sy, sx, sy)


class Deck:
    def __init__(self, path):
        self.z = zipfile.ZipFile(path)
        self._load_master()
        self._layout_cache = {}

    def _load_master(self):
        root = ET.fromstring(self.z.read("ppt/slideMasters/slideMaster1.xml"))
        # The three master-wide text style blocks.
        self.txstyles = {}
        for cat, tag in (("title", "titleStyle"), ("body", "bodyStyle"), ("other", "otherStyle")):
            self.txstyles[cat] = parse_lst_style(root.find(f"./p:txStyles/p:{tag}", NS))
        # Master placeholders carry the real per-level sizes in Google exports.
        self.master_ph = {}
        for sp in root.findall("./p:cSld/p:spTree/p:sp", NS):
            ph = sp.find(".//p:nvPr/p:ph", NS)
            key = ph_key(ph)
            if key is None:
                continue
            self.master_ph[key] = {
                "lvls": parse_lst_style(sp.find("./p:txBody/a:lstStyle", NS)),
                "xfrm": shape_xfrm(sp),
                "bodyPr": sp.find("./p:txBody/a:bodyPr", NS),
            }

    def layout_for(self, slide_n):
        rels = self.z.read(f"ppt/slides/_rels/slide{slide_n}.xml.rels").decode()
        m = re.search(r'Target="\.\./slideLayouts/(slideLayout\d+\.xml)"', rels)
        if not m:
            return {}
        name = m.group(1)
        if name in self._layout_cache:
            return self._layout_cache[name]
        root = ET.fromstring(self.z.read(f"ppt/slideLayouts/{name}"))
        out = {}
        for sp in root.findall("./p:cSld/p:spTree/p:sp", NS):
            ph = sp.find(".//p:nvPr/p:ph", NS)
            key = ph_key(ph)
            if key is None:
                continue
            out[key] = {
                "lvls": parse_lst_style(sp.find("./p:txBody/a:lstStyle", NS)),
                "xfrm": shape_xfrm(sp),
                "bodyPr": sp.find("./p:txBody/a:bodyPr", NS),
            }
        self._layout_cache[name] = out
        return out

    def slide_rels(self, n):
        x = self.z.read(f"ppt/slides/_rels/slide{n}.xml.rels").decode()
        return dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', x))


def autofit_scale(bodyPr):
    """PowerPoint's stored 'shrink text on overflow' factors."""
    if bodyPr is None:
        return 1.0, 0
    fit = bodyPr.find("a:normAutofit", NS)
    if fit is None:
        return 1.0, 0
    scale = int(fit.get("fontScale", "100000")) / 100000
    reduction = int(fit.get("lnSpcReduction", "0")) / 100000
    return scale, reduction


def build_paragraphs(txbody, lvl_chain, rels, font_scale, ln_reduction, default_sz):
    """Turn a p:txBody into positioned-ready paragraph dicts."""
    paras = []
    for para in txbody.findall("a:p", NS):
        pPr = para.find("a:pPr", NS)
        lvl = int(pPr.get("lvl", "0")) if pPr is not None else 0

        props = {}
        for source in lvl_chain:
            props.update(source.get(lvl, {}))
        props.update(parse_lvl_props(pPr))

        runs = []
        for child in para:
            tag = child.tag.split("}")[-1]
            if tag == "br":
                runs.append({"br": True})
                continue
            if tag != "r":
                continue
            t_el = child.find("a:t", NS)
            text = t_el.text or "" if t_el is not None else ""
            if not text:
                continue
            rpr = child.find("a:rPr", NS)
            rprops = parse_run_props(rpr)
            href = None
            if rpr is not None:
                hl = rpr.find("a:hlinkClick", NS)
                if hl is not None:
                    href = rels.get(hl.get(qn("r:id")), "")
            sz = rprops.get("sz", props.get("sz", default_sz))
            runs.append({
                "t": text,
                "px": pt_px(sz * font_scale),
                "b": rprops.get("b", props.get("b", False)),
                "i": rprops.get("i", props.get("i", False)),
                "u": rprops.get("u", props.get("u", False)),
                "href": href,
            })

        if not any(r.get("t") for r in runs):
            continue

        sz = props.get("sz", default_sz)
        bu = props.get("buChar")
        lnspc = props.get("lnSpc", 100000) / 100000
        paras.append({
            "px": pt_px(sz * font_scale),
            "marL": emu_px(int(props.get("marL", 0))),
            "indent": emu_px(int(props.get("indent", 0))),
            "algn": props.get("algn", "l"),
            "lineHeight": round(max(lnspc - ln_reduction, 0.6), 3),
            "spcBef": pt_px(props.get("spcBef", 0)),
            "spcAft": pt_px(props.get("spcAft", 0)),
            "bullet": bu,
            "bulletPx": pt_px(props.get("buSzPts", sz) * font_scale),
            "runs": runs,
        })
    return paras


def extract_table(gf, transform):
    tbl = gf.find(".//a:tbl", NS)
    xfrm = shape_xfrm(gf, "./p:xfrm")
    cols = [int(g.get("w")) for g in tbl.findall("./a:tblGrid/a:gridCol", NS)]
    rows = []
    for tr in tbl.findall("a:tr", NS):
        cells = []
        for tc in tr.findall("a:tc", NS):
            lines = []
            for para in tc.findall(".//a:p", NS):
                text = "".join(t.text or "" for t in para.findall(".//a:t", NS))
                if text.strip():
                    lines.append(text)
            cells.append(" ".join(lines))
        rows.append({"h": int(tr.get("h", "0")), "cells": cells})
    x, y, _, _ = transform.rect(xfrm[0], xfrm[1], xfrm[2], xfrm[3])
    return {
        "kind": "table",
        "x": emu_px(x),
        "y": emu_px(y),
        "cols": [emu_px(c * transform.sx) for c in cols],
        "rows": [{"h": emu_px(r["h"] * transform.sy), "cells": r["cells"]} for r in rows],
    }


def collect(shapes, transform, deck, layout, rels, elements):
    for sh in shapes:
        tag = sh.tag.split("}")[-1]

        if tag == "pic":
            blip = sh.find(".//a:blip", NS)
            rid = blip.get(qn("r:embed")) if blip is not None else None
            box = shape_xfrm(sh)
            if not rid or rid not in rels or box is None:
                continue
            x, y, cx, cy = transform.rect(*box)
            elements.append({
                "kind": "image",
                "src": rels[rid].split("/")[-1],
                "x": emu_px(x), "y": emu_px(y),
                "w": emu_px(cx), "h": emu_px(cy),
            })

        elif tag == "graphicFrame":
            if sh.find(".//a:tbl", NS) is not None:
                elements.append(extract_table(sh, transform))

        elif tag == "grpSp":
            xfrm = sh.find("./p:grpSpPr/a:xfrm", NS)
            off = xfrm.find("a:off", NS)
            ext = xfrm.find("a:ext", NS)
            ch_off = xfrm.find("a:chOff", NS)
            ch_ext = xfrm.find("a:chExt", NS)
            child = transform.child(
                (int(off.get("x")), int(off.get("y"))),
                (int(ext.get("cx")), int(ext.get("cy"))),
                (int(ch_off.get("x")), int(ch_off.get("y"))),
                (int(ch_ext.get("cx")), int(ch_ext.get("cy"))),
            )
            collect(list(sh), child, deck, layout, rels, elements)

        elif tag == "sp":
            txbody = sh.find("./p:txBody", NS)
            if txbody is None:
                continue
            ph = sh.find(".//p:nvPr/p:ph", NS)
            key = ph_key(ph)
            cat = category_for(ph)

            master = deck.master_ph.get(key, {}) if key else {}
            lay = layout.get(key, {}) if key else {}

            # Geometry: shape wins, then layout placeholder, then master placeholder.
            box = shape_xfrm(sh) or lay.get("xfrm") or master.get("xfrm")
            if box is None:
                continue

            bodyPr = txbody.find("a:bodyPr", NS)
            scale, reduction = autofit_scale(bodyPr)
            if scale == 1.0 and reduction == 0:
                scale, reduction = autofit_scale(lay.get("bodyPr"))

            lvl_chain = [
                deck.txstyles.get(cat, {}),
                master.get("lvls", {}),
                lay.get("lvls", {}),
                parse_lst_style(txbody.find("a:lstStyle", NS)),
            ]
            default_sz = 2800 if cat == "title" else 1800 if cat == "body" else 1400
            paras = build_paragraphs(txbody, lvl_chain, rels, scale, reduction, default_sz)
            if not paras:
                continue

            anchor = None
            for source in (bodyPr, lay.get("bodyPr"), master.get("bodyPr")):
                if source is not None and source.get("anchor"):
                    anchor = source.get("anchor")
                    break

            insets = {"l": 91425, "t": 91425, "r": 91425, "b": 91425}
            if bodyPr is not None:
                for k, attr in (("l", "lIns"), ("t", "tIns"), ("r", "rIns"), ("b", "bIns")):
                    if bodyPr.get(attr) is not None:
                        insets[k] = int(bodyPr.get(attr))

            x, y, cx, cy = transform.rect(*box)
            elements.append({
                "kind": "text",
                "x": emu_px(x), "y": emu_px(y),
                "w": emu_px(cx), "h": emu_px(cy),
                "anchor": anchor or "t",
                "pad": {k: emu_px(v) for k, v in insets.items()},
                "paras": paras,
                "isTitle": cat == "title",
            })


def main():
    deck = Deck(PPTX)
    z = deck.z

    pres = z.read("ppt/presentation.xml").decode()
    pres_rels = z.read("ppt/_rels/presentation.xml.rels").decode()
    rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="slides/(slide\d+\.xml)"', pres_rels))
    order_ids = re.findall(r'<p:sldId id="\d+" r:id="([^"]+)"/>', pres)
    slide_files = [rel_map[i] for i in order_ids]

    slides = []
    for idx, fname in enumerate(slide_files, start=1):
        n = int(re.search(r"slide(\d+)\.xml", fname).group(1))
        root = ET.fromstring(z.read(f"ppt/slides/{fname}"))
        rels = deck.slide_rels(n)
        layout = deck.layout_for(n)

        elements = []
        tree = root.find("./p:cSld/p:spTree", NS)
        collect(list(tree), Transform(), deck, layout, rels, elements)

        # Title text drives the slide's heading in the outline / URL fragment.
        title = ""
        for el in elements:
            if el["kind"] == "text" and el.get("isTitle"):
                title = " ".join(
                    r["t"] for p in el["paras"] for r in p["runs"] if r.get("t")
                ).strip()
                break

        notes = ""
        for rid, tgt in rels.items():
            if "notesSlide" in tgt:
                path = "ppt/" + tgt.replace("../", "")
                if path in z.namelist():
                    nroot = ET.fromstring(z.read(path))
                    lines = []
                    for para in nroot.findall(".//a:p", NS):
                        s = "".join(t.text or "" for t in para.findall(".//a:t", NS))
                        if s.strip():
                            lines.append(s)
                    notes = "\n".join(lines)

        slides.append({
            "index": idx,
            "source_slide": n,
            "title": title,
            "elements": elements,
            "notes": notes,
        })

    with open(OUT, "w") as f:
        json.dump({"canvas": [CANVAS_W, CANVAS_H], "slides": slides}, f, indent=2)
    print(f"Wrote {OUT}: {len(slides)} slides")


if __name__ == "__main__":
    main()
