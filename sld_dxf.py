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
WIDTH_F = 0.8                   # STANDARD style width factor: txt at this
                                # factor is no wider than the browser's Arial
CHAR_W = 0.9 * WIDTH_F          # advance of one txt character, per unit height
WRAP_AT = 40                    # table cells longer than this wrap
SUBST = {"—": "-", "–": "-", "·": "-", "×": "x", "→": "->", "←": "<-",
         "±": "+/-", "…": "..."}


def num(v):
    """A DXF number: two decimals, no trailing zeros, no negative zero."""
    t = f"{v:.2f}".rstrip("0").rstrip(".")
    return "0" if t in ("-0", "") else t


def clean(s):
    return "".join(SUBST.get(c, c) for c in str(s))


def text_w(s, size):
    """Width of a DXF text of SVG size `size`, in drawing units."""
    return len(clean(s)) * size * TEXT_H * CHAR_W


def wrap(s, n=WRAP_AT):
    """Fold a long cell at word boundaries."""
    out, cur = [], ""
    for w in str(s).split():
        if cur and len(cur) + 1 + len(w) > n:
            out.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur:
        out.append(cur)
    return out or [""]


class DXF(S.SVG):
    """The same drawing surface as the SVG, writing DXF entities.  The
    sketch's y grows downward; the DXF's grows upward, so y is negated."""

    def __init__(self, table=None):
        super().__init__()
        self.ents = []
        self.table = table          # (info, items, order) to draw beside
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
        """Queue an entity.  Coordinates (codes 10-13 / 20-23) stay raw
        numbers in DXF space until document() knows the extents and can
        centre the whole drawing on the origin."""
        self.ents.append(pairs)
        self.count += 1

    @staticmethod
    def _emit(pairs, dx, dy):
        out = []
        for c, v in pairs:
            if c in (10, 11, 12, 13):
                v = num(v + dx)
            elif c in (20, 21, 22, 23):
                v = num(v + dy)
            out.append(f"{c}\n{v}")
        return "\n".join(out)

    # -- primitives (the engine's symbols call only these) ------------------
    def line(self, x1, y1, x2, y2, w=2, dash=None):
        self._track(x1, y1, x2, y2)
        if w >= 3:                  # a bar: a polyline with width
            self._e((0, "POLYLINE"), (8, self._layer("busbar")), (66, 1),
                    (70, 0), (40, num(w)), (41, num(w)),
                    (10, 0), (20, 0), (30, 0))
            self._e((0, "VERTEX"), (8, self._layer("busbar")),
                    (10, x1), (20, -y1), (30, 0))
            self._e((0, "VERTEX"), (8, self._layer("busbar")),
                    (10, x2), (20, -y2), (30, 0))
            self._e((0, "SEQEND"), (8, self._layer("busbar")))
            return
        pairs = [(0, "LINE"), (8, self._layer())]
        if dash:
            pairs.append((6, "DASHED"))
        pairs += [(10, x1), (20, -y1), (30, 0),
                  (11, x2), (21, -y2), (31, 0)]
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
                    (10, px), (20, -py), (30, 0))
        self._e((0, "SEQEND"), (8, lay))

    def circle(self, x, y, r, sw=2):
        self._e((0, "CIRCLE"), (8, self._layer()),
                (10, x), (20, -y), (30, 0), (40, num(r)))

    def dot(self, x, y, r=3.2):
        # a filled dot: the classic donut, a closed polyline of two bulged
        # vertices whose width equals the radius
        lay = self._layer()
        self._e((0, "POLYLINE"), (8, lay), (66, 1), (70, 1),
                (40, num(r)), (41, num(r)), (10, 0), (20, 0), (30, 0))
        for px in (x - r / 2, x + r / 2):
            self._e((0, "VERTEX"), (8, lay), (10, px), (20, -y),
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
            pairs += [(10 + k, px), (20 + k, -py), (30 + k, 0)]
        self._e(*pairs)

    def text(self, x, y, s, size=12, anchor="middle", bold=False,
             rotate=None, color="#111"):
        if not s:
            return
        pairs = [(0, "TEXT"), (8, self._layer("text")), (7, "STANDARD"),
                 (10, x), (20, -y), (30, 0),
                 (40, num(size * TEXT_H)), (1, clean(s))]
        if rotate:
            pairs.append((50, num(-rotate)))   # SVG rotates clockwise
        if anchor != "start":
            pairs += [(72, 1 if anchor == "middle" else 2),
                      (11, x), (21, -y), (31, 0)]
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
                (10, cx), (20, cy), (30, 0), (40, num(r)),
                (50, num(start)), (51, num(end)))

    # -- the equipment table under the sheet --------------------------------
    def draw_table(self, x_left, y_top):
        """The equipment table beside the sheet; returns (right, bottom)."""
        info, items, order = self.table
        self.layer = "table"
        size, row_h, pad = 11, 18, 8
        x0, y = x_left, y_top
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
            rows.append([wrap(v) for v in
                         (it.id, typ, it.desc, it.rating, it.voltage,
                          ", ".join(it.prots), ", ".join(it.parents),
                          it.notes)])
        cols = [max([text_w(h, size)] + [text_w(l, size) for r in rows
                                          for l in r[i]]) * 1.15 + 2 * pad
                for i, h in enumerate(heads)]
        x1 = x0 + sum(cols)
        y_head = y
        self.line(x0, y, x1, y, w=1.2)
        cx = x0
        for h, cw in zip(heads, cols):
            self.text(cx + pad, y + 13, h, size=size, anchor="start", bold=True)
            cx += cw
        y += row_h
        self.line(x0, y, x1, y, w=1.2)
        for r in rows:
            cx = x0
            lines = max(len(c) for c in r)
            for cell, cw in zip(r, cols):
                for k, v in enumerate(cell):
                    if v:
                        self.text(cx + pad, y + 13 + 15 * k, v, size=size,
                                  anchor="start")
                cx += cw
            y += row_h + 15 * (lines - 1)
            self.line(x0, y, x1, y, w=1)
        cx = x0
        for cw in cols + [0]:
            self.line(cx, y_head, cx, y, w=1)
            cx += cw
        self.layer = "drawing"
        return x1, y

    # -- the file -------------------------------------------------------------
    def document(self, width, height):
        # the table stands to the right of the sheet, 40 units clear
        right, bottom = width, height
        if self.table:
            x1, y1 = self.draw_table(width + 40, 24)
            right, bottom = max(right, x1 + 24), max(bottom, y1 + 24)
        # centre the whole drawing on the origin
        dx, dy = -right / 2, bottom / 2
        xmin, xmax = -right / 2, right / 2
        ymin, ymax = -bottom / 2, bottom / 2
        out = ["0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\nAC1009",
               f"9\n$EXTMIN\n10\n{num(xmin)}\n20\n{num(ymin)}\n30\n0",
               f"9\n$EXTMAX\n10\n{num(xmax)}\n20\n{num(ymax)}\n30\n0",
               f"9\n$LIMMIN\n10\n{num(xmin)}\n20\n{num(ymin)}",
               f"9\n$LIMMAX\n10\n{num(xmax)}\n20\n{num(ymax)}",
               "0\nENDSEC", "0\nSECTION\n2\nTABLES",
               # the opening view: centred on the drawing, fitted with a margin
               "0\nTABLE\n2\nVPORT\n70\n1",
               "0\nVPORT\n2\n*ACTIVE\n70\n0\n10\n0\n20\n0\n11\n1\n21\n1"
               "\n12\n0\n22\n0\n13\n0\n23\n0\n14\n10\n24\n10\n15\n10\n25\n10"
               "\n16\n0\n26\n0\n36\n1\n17\n0\n27\n0\n37\n0"
               f"\n40\n{num(bottom * 1.08)}\n41\n{num(right / bottom)}"
               "\n42\n50\n43\n0\n44\n0\n50\n0\n51\n0\n71\n0\n72\n100"
               "\n73\n1\n74\n3\n75\n0\n76\n0\n77\n0\n78\n0",
               "0\nENDTAB",
               "0\nTABLE\n2\nLTYPE\n70\n2",
               "0\nLTYPE\n2\nCONTINUOUS\n70\n0\n3\nSolid line\n72\n65\n73\n0\n40\n0",
               "0\nLTYPE\n2\nDASHED\n70\n0\n3\n__ __ __\n72\n65\n73\n2\n40\n9\n49\n6\n49\n-3",
               "0\nENDTAB", f"0\nTABLE\n2\nLAYER\n70\n{len(LAYERS)}"]
        for name, col in LAYERS:
            out.append(f"0\nLAYER\n2\n{name}\n70\n0\n62\n{col}\n6\nCONTINUOUS")
        out += ["0\nENDTAB", "0\nTABLE\n2\nSTYLE\n70\n1",
                f"0\nSTYLE\n2\nSTANDARD\n70\n0\n40\n0\n41\n{num(WIDTH_F)}\n50\n0\n71\n0\n42\n2.5\n3\ntxt\n4\n",
                "0\nENDTAB", "0\nENDSEC", "0\nSECTION\n2\nBLOCKS\n0\nENDSEC",
                "0\nSECTION\n2\nENTITIES"]
        out += [self._emit(e, dx, dy) for e in self.ents]
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


def check_dxf(text):
    """Read the TEXT and LINE entities back and report texts that overlap
    another text, table text crossing a table rule, and anything outside
    the declared extents.  Returns a list of messages (empty = clean)."""
    lines = text.split("\n")
    ents, cur = [], None
    for i in range(0, len(lines) - 1, 2):
        code, val = lines[i].strip(), lines[i + 1]
        if code == "0":
            cur = {"type": val}
            ents.append(cur)
        elif cur is not None:
            cur[code] = val
    ext = {}
    for i in range(0, len(lines) - 1, 2):
        if lines[i].strip() == "9" and lines[i + 1] in ("$EXTMIN", "$EXTMAX"):
            ext[lines[i + 1]] = (float(lines[i + 3]), float(lines[i + 5]))
    boxes = []
    for e in ents:
        if e["type"] != "TEXT":
            continue
        h = float(e["40"])
        w = len(e["1"]) * h * CHAR_W
        x, y = float(e["10"]), float(e["20"])
        rot = float(e.get("50", 0)) % 360
        al = int(e.get("72", 0))
        if abs(rot - 270) < 1:                  # runs downward
            box = (x - h, y - w, x, y)
        else:
            x0 = x - (w / 2 if al == 1 else w if al == 2 else 0)
            box = (x0, y, x0 + w, y + h)
        boxes.append((box, e["1"], e["8"]))
    out = []
    for i, (a, ta, la) in enumerate(boxes):
        for b, tb, lb in boxes[i + 1:]:
            if a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]:
                out.append(f"text overlaps text: '{ta}' / '{tb}' ({la})")
    rules = [(float(e["10"]), float(e["20"]), float(e["11"]), float(e["21"]))
             for e in ents if e["type"] == "LINE" and e["8"] == "SLD_TABLE"]
    for box, t, lay in boxes:
        if lay != "SLD_TABLE":
            continue
        for x1, y1, x2, y2 in rules:
            if abs(x1 - x2) < 0.01 and box[0] < x1 < box[2] \
                    and min(y1, y2) < box[3] and max(y1, y2) > box[1]:
                out.append(f"table text crosses a column rule: '{t}'")
            if abs(y1 - y2) < 0.01 and box[1] < y1 < box[3] \
                    and min(x1, x2) < box[2] and max(x1, x2) > box[0]:
                out.append(f"table text crosses a row rule: '{t}'")
    if ext:
        (xmin, ymin), (xmax, ymax) = ext["$EXTMIN"], ext["$EXTMAX"]
        for box, t, lay in boxes:
            if box[0] < xmin - 1 or box[2] > xmax + 1 \
                    or box[1] < ymin - 1 or box[3] > ymax + 1:
                out.append(f"text outside the sheet: '{t}' ({lay})")
        vp = next((e for e in ents if e["type"] == "VPORT"
                   and e.get("2") == "*ACTIVE"), None)
        if vp is None:
            out.append("no *ACTIVE viewport: the file opens wherever the CAD "
                       "program likes")
        else:
            cx, cy = float(vp["12"]), float(vp["22"])
            vh, asp = float(vp["40"]), float(vp["41"])
            if abs(cx - (xmin + xmax) / 2) > 1 or abs(cy - (ymin + ymax) / 2) > 1:
                out.append(f"opening view centred at ({cx}, {cy}), not on the "
                           f"drawing ({(xmin + xmax) / 2}, {(ymin + ymax) / 2})")
            if vh < ymax - ymin or vh * asp < xmax - xmin:
                out.append("opening view smaller than the drawing")
    return out


def main():
    ap = argparse.ArgumentParser(
        description="Export a site-survey workbook's single-line diagram "
                    "and its equipment table to DXF.")
    ap.add_argument("workbook", help="input .xlsx file")
    ap.add_argument("-o", "--output",
                    help="output .dxf file (default: <workbook>.dxf)")
    ap.add_argument("--check", action="store_true",
                    help="read the file back and report overlapping text")
    args = ap.parse_args()
    out = args.output or (args.workbook.rsplit(".", 1)[0] + ".dxf")
    info, items, order = S.read_workbook(args.workbook)
    width = S.layout(items, order)
    n = write_dxf(out, info, items, order, width)
    print(f"wrote {out}  ({len(items)} items, {n} entities)")
    if args.check:
        probs = check_dxf(open(out, encoding="utf-8").read())
        for m in probs:
            print("  " + m)
        print(f"  {len(probs)} text problems" if probs else "  text clean")
        if probs:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
