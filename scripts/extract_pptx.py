#!/usr/bin/env python3
"""Extract slide content from the source .pptx into structured JSON for gen_qmd.py."""
import json
import re
import zipfile
import xml.etree.ElementTree as ET

PPTX = "Getting the first job in Industry.pptx"
OUT = "scripts/deck.json"

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def qn(tag):
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


def text_of(el):
    return "".join(t.text or "" for t in el.findall(".//a:t", NS))


def para_runs_markdown(para, rels):
    parts = []
    for r in para.findall("a:r", NS):
        t = "".join(x.text or "" for x in r.findall("a:t", NS))
        rpr = r.find("a:rPr", NS)
        hlink = rpr.find("a:hlinkClick", NS) if rpr is not None else None
        if hlink is not None:
            rid = hlink.get(qn("r:id"))
            url = rels.get(rid, "")
            parts.append(f"[{t}]({url})")
        else:
            parts.append(t)
    return "".join(parts)


def get_slide_rels(z, n):
    path = f"ppt/slides/_rels/slide{n}.xml.rels"
    rels = {}
    if path in z.namelist():
        x = z.read(path).decode()
        for rid, tgt in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', x):
            rels[rid] = tgt
    return rels


def get_notes(z, n, rels):
    notes_target = None
    for rid, tgt in rels.items():
        if "notesSlide" in tgt:
            notes_target = tgt
    if not notes_target:
        return ""
    path = "ppt/" + notes_target.replace("../", "")
    if path not in z.namelist():
        return ""
    root = ET.fromstring(z.read(path))
    lines = []
    for para in root.findall(".//a:p", NS):
        s = text_of(para)
        if s.strip():
            lines.append(s)
    return "\n".join(lines)


def extract_table(gf):
    tbl = gf.find(".//a:tbl", NS)
    rows = []
    for tr in tbl.findall("a:tr", NS):
        row = []
        for tc in tr.findall("a:tc", NS):
            cell_lines = [text_of(p) for p in tc.findall(".//a:p", NS) if text_of(p).strip()]
            row.append(" ".join(cell_lines))
        rows.append(row)
    return rows


class Transform:
    """Affine map (no rotation) from a shape's local EMU space to absolute slide EMU space."""

    def __init__(self, ox=0, oy=0, sx=1.0, sy=1.0):
        self.ox, self.oy, self.sx, self.sy = ox, oy, sx, sy

    def point(self, x, y):
        return self.ox + x * self.sx, self.oy + y * self.sy

    def ext(self, cx, cy):
        return cx * self.sx, cy * self.sy

    def child_transform(self, off, ext, ch_off, ch_ext):
        abs_off_x, abs_off_y = self.point(*off)
        abs_ext_cx, abs_ext_cy = self.ext(*ext)
        sx = abs_ext_cx / ch_ext[0] if ch_ext[0] else 1.0
        sy = abs_ext_cy / ch_ext[1] if ch_ext[1] else 1.0
        # local (x,y) -> abs_off + (local - ch_off) * s
        return Transform(abs_off_x - ch_off[0] * sx, abs_off_y - ch_off[1] * sy, sx, sy)


def collect_images(shapes, transform, rels, W, H, images):
    for sh in shapes:
        tag = sh.tag.split("}")[-1]
        if tag == "pic":
            blip = sh.find(".//a:blip", NS)
            rid = blip.get(qn("r:embed")) if blip is not None else None
            off_el = sh.find("./p:spPr/a:xfrm/a:off", NS)
            ext_el = sh.find("./p:spPr/a:xfrm/a:ext", NS)
            if rid and rid in rels and off_el is not None and ext_el is not None:
                x, y = int(off_el.get("x")), int(off_el.get("y"))
                cx, cy = int(ext_el.get("cx")), int(ext_el.get("cy"))
                ax, ay = transform.point(x, y)
                acx, acy = transform.ext(cx, cy)
                src = rels[rid].split("/")[-1]
                images.append({
                    "src": src,
                    "left_pct": round(ax / W * 100, 3),
                    "top_pct": round(ay / H * 100, 3),
                    "width_pct": round(acx / W * 100, 3),
                    "height_pct": round(acy / H * 100, 3),
                })
        elif tag == "grpSp":
            xfrm = sh.find("./p:grpSpPr/a:xfrm", NS)
            off_el = xfrm.find("a:off", NS)
            ext_el = xfrm.find("a:ext", NS)
            choff_el = xfrm.find("a:chOff", NS)
            chext_el = xfrm.find("a:chExt", NS)
            off = (int(off_el.get("x")), int(off_el.get("y")))
            ext = (int(ext_el.get("cx")), int(ext_el.get("cy")))
            ch_off = (int(choff_el.get("x")), int(choff_el.get("y")))
            ch_ext = (int(chext_el.get("cx")), int(chext_el.get("cy")))
            new_t = transform.child_transform(off, ext, ch_off, ch_ext)
            collect_images(list(sh), new_t, rels, W, H, images)


def extract_shape_text(sp):
    ph = sp.find(".//p:nvSpPr/p:nvPr/p:ph", NS)
    phtype = ph.get("type") if ph is not None else None
    paras = sp.findall(".//a:p", NS)
    return phtype, paras


def main():
    z = zipfile.ZipFile(PPTX)

    pres = z.read("ppt/presentation.xml").decode()
    pres_rels = z.read("ppt/_rels/presentation.xml.rels").decode()
    rel_map = dict(re.findall(r'Id="([^"]+)"[^>]*Target="slides/(slide\d+\.xml)"', pres_rels))
    order_ids = re.findall(r'<p:sldId id="\d+" r:id="([^"]+)"/>', pres)
    slide_files = [rel_map[i] for i in order_ids]

    sz = re.search(r'<p:sldSz[^/]*/>', pres).group(0)
    W = int(re.search(r'cx="(\d+)"', sz).group(1))
    H = int(re.search(r'cy="(\d+)"', sz).group(1))

    slides = []
    for idx, fname in enumerate(slide_files, start=1):
        n = int(re.search(r"slide(\d+)\.xml", fname).group(1))
        root = ET.fromstring(z.read(f"ppt/slides/{fname}"))
        rels = get_slide_rels(z, n)

        title = None
        body_items = []
        images = []
        table = None

        tree = root.find(".//p:cSld/p:spTree", NS)
        for sh in tree:
            tag = sh.tag.split("}")[-1]
            if tag == "sp":
                phtype, paras = extract_shape_text(sh)
                if phtype in ("title", "ctrTitle"):
                    title = text_of(sh).strip()
                elif paras:
                    for para in paras:
                        pPr = para.find("a:pPr", NS)
                        lvl = int(pPr.get("lvl", "0")) if pPr is not None else 0
                        md = para_runs_markdown(para, rels).strip()
                        if md:
                            body_items.append((lvl, md))
            elif tag == "graphicFrame":
                if sh.find(".//a:tbl", NS) is not None:
                    table = extract_table(sh)

        collect_images(list(tree), Transform(), rels, W, H, images)

        notes = get_notes(z, n, rels)

        slides.append({
            "index": idx,
            "source_slide": n,
            "title": title,
            "body": body_items,
            "images": images,
            "table": table,
            "notes": notes,
        })

    with open(OUT, "w") as f:
        json.dump({"width_emu": W, "height_emu": H, "slides": slides}, f, indent=2)
    print(f"Wrote {OUT}: {len(slides)} slides")


if __name__ == "__main__":
    main()
