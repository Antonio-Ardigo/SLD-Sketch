#!/usr/bin/env python3
"""SLD-Sketch — DXF export.

    python sld_dxf.py <workbook.xlsx> [-o output.dxf]
    python sld_sketch.py <workbook.xlsx> --dxf          # SVG and DXF together

Writes an R12 (AC1009) DXF, the dialect every CAD package and viewer reads:
the sketch exactly as the SVG draws it (the DXF canvas reuses the engine's
own symbol primitives), and the equipment table that produced it, laid out
under the sheet.  One drawing unit is one sketch pixel; take it as 1 mm.

Layers: SLD_DRAWING (conductors, symbols), SLD_BUSBAR (thick bars),
SLD_TEXT (labels), SLD_ENCLOSURE (RMU boxes), SLD_FRAME (title, title
block), SLD_LEGEND, SLD_TABLE.
"""

import argparse
import math
import re

import sld_sketch as S

LAYERS = [                      # name, colour index
    ("SLD_DRAWING", 7), ("SLD_BUSBAR", 7), ("SLD_TEXT", 7),
    ("SLD_ENCLOSURE", 8), ("SLD_FRAME", 8), ("SLD_LEGEND", 8),
    ("SLD_TABLE", 7),
]
TEXT_H = 0.72                   # DXF text height per px of SVG font size
CHAR_W = 0.6                    # em width of a character, for the table
SUBST = {"—": "-", "–": "-", "·": "-", "×": "x", "→": "->", "←": "<-",
         "±": "+/-", "…": "..."}


def num(v):
    """A DXF number: two decimals, no trailing zeros, no negative zero."""
    t = f"{v:.2f}".rstrip("0").rstrip(".")
    return "0" if t in ("-0", "") else t


def clean(s):
    return "".join(SUBST.get(c, c) for c in str(s))


class DXF(S.SVG):
    """The same drawing surface as the SVG, writing DXF entities.  The
    sketch's y grows downward; the DXF's grows upward, so y is negated."""

    def __init__(self, table=None):
        super().__init__()
        self.ents = []
        self.table = table          # (info, items, order) to draw below
        self.count = 0

    # -- helpers ------------------------------------------------------------
    def _layer(self, kind="drawing"):
        if self.layer == "frame":
            return "SLD_FRAME"
        if self.layer == "legend":
            return "SLD_LEGEND"
        if self.layer == "table":
            return "SLD_TABLE"
        return {"text": "SLD_TEXT", "busbar": "SLD_BUSBAR",
                "enclosure": "SLD_ENCLOSURE"}.get(kind, "SLD_DRAWING")

    def _e(self, *pairs):
        self.ents.append("\n".join(f"{c}\n{v}" for c, v in pairs))
        self.count += 1

    # -- primitives (the engine's symbols call only these) ------------------
    def line(self, x1, y1, x2, y2, w=2, dash=None):
        if w >= 3:                  # a bar: a polyline with width
            self._e((0, "POLYLINE"), (8, self._layer("busbar")), (66, 1),
                    (70, 0), (40, num(w)), (41, num(w)),
                    (10, 0), (20, 0), (30, 0))
            self._e((0, "VERTEX"), (8, self._layer("busbar")),
                    (10, num(x1)), (20, num(-y1)), (30, 0))
            self._e((0, "VERTEX"), (8, self._layer("busbar")),
                    (10, num(x2)), (20, num(-y2)), (30, 0))
            self._e((0, "SEQEND"), (8, self._layer("busbar")))
            return
        pairs = [(0, "LINE"), (8, self._layer())]
        if dash:
            pairs.append((6, "DASHED"))
        pairs += [(10, num(x1)), (20, num(-y1)), (30, 0),
                  (11, num(x2)), (21, num(-y2)), (31, 0)]
        self._e(*pairs)

    def rect(self, x, y, w, h, sw=2, dash=None, fill="none"):
        lay = self._layer("enclosure" if dash else "drawing")
        pairs = [(0, "POLYLINE"), (8, lay)]
        if dash:
            pairs.append((6, "DASHED"))
        pairs += [(66, 1), (70, 1), (10, 0), (20, 0), (30, 0)]
        self._e(*pairs)
        for px, py in ((x, y), (x + w, y), (x + w, y + h), (x, y + h)):
            self._e((0, "VERTEX"), (8, lay),
                    (10, num(px)), (20, num(-py)), (30, 0))
        self._e((0, "SEQEND"), (8, lay))

    def circle(self, x, y, r, sw=2):
        self._e((0, "CIRCLE"), (8, self._layer()),
                (10, num(x)), (20, num(-y)), (30, 0), (40, num(r)))

    def dot(self, x, y, r=3.2):
        # a filled dot: the classic donut, a closed polyline of two bulged
        # vertices whose width equals the radius
        lay = self._layer()
        self._e((0, "POLYLINE"), (8, lay), (66, 1), (70, 1),
                (40, num(r)), (41, num(r)), (10, 0), (20, 0), (30, 0))
        for px in (x - r / 2, x + r / 2):
            self._e((0, "VERTEX"), (8, lay), (10, num(px)), (20, num(-y)),
                    (30, 0), (42, 1))
        self._e((0, "SEQEND"), (8, lay))

    def poly(self, pts, fill="#111"):
        pts = list(pts)
        if len(pts) < 3:
            return
        while len(pts) < 4:
            pts.append(pts[-1])
        pairs = [(0, "SOLID"), (8, self._layer())]
        # a SOLID's corners run 1-2-4-3 (a bow-tie otherwise)
        for k, (px, py) in zip((0, 1, 3, 2), pts[:4]):
            pairs += [(10 + k, num(px)), (20 + k, num(-py)), (30 + k, 0)]
        self._e(*pairs)

    def text(self, x, y, s, size=12, anchor="middle", bold=False,
             rotate=None, color="#111"):
        if not s:
            return
        pairs = [(0, "TEXT"), (8, self._layer("text")), (7, "STANDARD"),
                 (10, num(x)), (20, num(-y)), (30, 0),
                 (40, num(size * TEXT_H)), (1, clean(s))]
        if rotate:
            pairs.append((50, num(-rotate)))   # SVG rotates clockwise
        if anchor != "start":
            pairs += [(72, 1 if anchor == "middle" else 2),
                      (11, num(x)), (21, num(-y)), (31, 0)]
        self._e(*pairs)

    def path(self, d, sw=2):
        # the engine draws one kind of path: a semicircle
        # "M x1,y1 A r,r 0 large sweep x2,y2" (the contactor's hinge arc)
        m = re.match(r"M\s*([-\d.]+),([-\d.]+)\s+A\s*([-\d.]+),[-\d.]+\s+"
                     r"[-\d.]+\s+(\d)\s+(\d)\s+([-\d.]+),([-\d.]+)", d)
        if not m:
            return
        x1, y1, r, _, sweep, x2, y2 = (float(v) for v in m.groups())
        cx, cy = (x1 + x2) / 2, -(y1 + y2) / 2      # DXF coordinates
        a1 = math.degrees(math.atan2(-y1 - cy, x1 - cx)) % 360
        a2 = math.degrees(math.atan2(-y2 - cy, x2 - cx)) % 360
        # SVG sweep=1 is clockwise on screen; a DXF arc runs anticlockwise
        start, end = (a2, a1) if sweep == 1 else (a1, a2)
        self._e((0, "ARC"), (8, self._layer()),
                (10, num(cx)), (20, num(cy)), (30, 0), (40, num(r)),
                (50, num(start)), (51, num(end)))

    # -- the equipment table under the sheet --------------------------------
    def draw_table(self, y_top, width):
        info, items, order = self.table
        self.layer = "table"
        size, row_h, pad = 11, 18, 8
        x0, y = 24, y_top
        site = info.get("site", "")
        self.text(x0, y + 14, ("EQUIPMENT TABLE - " + site) if site
                  else "EQUIPMENT TABLE", size=14, anchor="start", bold=True)
        y += 24
        for k, v in (("Site", info.get("site", "")), ("Date", info.get("date", "")),
                     ("By", info.get("surveyed by", info.get("by", ""))),
                     ("Notes", info.get("notes", ""))):
            if v:
                self.text(x0, y + 12, f"{k}: {v}", size=11, anchor="start")
                y += 15
        y += 6
        heads = ["ID", "Type", "Description", "Rating", "Voltage",
                 "Protection", "Feeds From", "Notes"]
        labels = {S.CAPACITOR: "Capacitor Bank", S.EARTHING: "Earthing/NER",
                  S.ARRESTER: "Surge Arrester"}
        rows = []
        for oid in order:
            it = items[oid]
            typ = labels.get(it.type, it.type.title().replace("Mv ", "MV ")
                             .replace("Lv ", "LV ").replace("Rmu", "RMU")
                             .replace("Mcc", "MCC"))
            rows.append([it.id, typ, it.desc, it.rating, it.voltage,
                         ", ".join(it.prots), ", ".join(it.parents), it.notes])
        cols = [max([len(h)] + [len(r[i]) for r in rows]) * size * CHAR_W
                + 2 * pad for i, h in enumerate(heads)]
        total = sum(cols)
        x1 = x0 + total
        self.line(x0, y, x1, y, w=1.2)
        cx = x0
        for h, cw in zip(heads, cols):
            self.text(cx + pad, y + 13, h, size=size, anchor="start", bold=True)
            cx += cw
        y += row_h
        self.line(x0, y, x1, y, w=1.2)
        for r in rows:
            cx = x0
            for v, cw in zip(r, cols):
                if v:
                    self.text(cx + pad, y + 13, v, size=size, anchor="start")
                cx += cw
            y += row_h
            self.line(x0, y, x1, y, w=1)
        cx = x0
        y_head = y - row_h * len(rows) - row_h
        for cw in cols + [0]:
            self.line(cx, y_head, cx, y, w=1)
            cx += cw
        self.layer = "drawing"
        return y

    # -- the file -------------------------------------------------------------
    def document(self, width, height):
        h = height
        if self.table:
            h = self.draw_table(height + 40, width) + 24
        out = ["0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1009",
               f"9\n$EXTMIN\n10\n0\n20\n{num(-h)}\n30\n0",
               f"9\n$EXTMAX\n10\n{num(width)}\n20\n0\n30\n0",
               "0\nENDSEC", "0\nSECTION\n2\nTABLES",
               "0\nTABLE\n2\nLTYPE\n70\n2",
               "0\nLTYPE\n2\nCONTINUOUS\n70\n0\n3\nSolid line\n72\n65\n73\n0\n40\n0",
               "0\nLTYPE\n2\nDASHED\n70\n0\n3\n__ __ __\n72\n65\n73\n2\n40\n9\n49\n6\n49\n-3",
               "0\nENDTAB", f"0\nTABLE\n2\nLAYER\n70\n{len(LAYERS)}"]
        for name, col in LAYERS:
            out.append(f"0\nLAYER\n2\n{name}\n70\n0\n62\n{col}\n6\nCONTINUOUS")
        out += ["0\nENDTAB", "0\nTABLE\n2\nSTYLE\n70\n1",
                "0\nSTYLE\n2\nSTANDARD\n70\n0\n40\n0\n41\n1\n50\n0\n71\n0\n42\n2.5\n3\ntxt\n4\n",
                "0\nENDTAB", "0\nENDSEC", "0\nSECTION\n2\nBLOCKS\n0\nENDSEC",
                "0\nSECTION\n2\nENTITIES"]
        out += self.ents
        out += ["0\nENDSEC", "0\nEOF", ""]
        return "\n".join(out)


def render_dxf(info, items, order, width):
    """The DXF text of a laid-out sheet, table included."""
    canvas = DXF(table=(info, items, order))
    text = S.render(info, items, order, width, canvas=canvas)
    return text, canvas.count


def write_dxf(path, info, items, order, width):
    text, n = render_dxf(info, items, order, width)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(text)
    return n


def main():
    ap = argparse.ArgumentParser(
        description="Export a site-survey workbook's single-line diagram "
                    "and its equipment table to DXF.")
    ap.add_argument("workbook", help="input .xlsx file")
    ap.add_argument("-o", "--output",
                    help="output .dxf file (default: <workbook>.dxf)")
    args = ap.parse_args()
    out = args.output or (args.workbook.rsplit(".", 1)[0] + ".dxf")
    info, items, order = S.read_workbook(args.workbook)
    width = S.layout(items, order)
    n = write_dxf(out, info, items, order, width)
    print(f"wrote {out}  ({len(items)} items, {n} entities)")


if __name__ == "__main__":
    main()
