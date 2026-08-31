#!/usr/bin/env python3
"""SLD-Sketch — turn a simple site-survey spreadsheet into a single-line diagram.

Usage:
    python sld_sketch.py <workbook.xlsx> [-o output.svg]

The workbook needs two sheets:
  * "Info"       — key/value rows (Site, Date, Surveyed by, Notes)
  * "Equipment"  — one row per item:
        ID | Type | Description | Rating | Voltage | Feeds From | Notes

Recognised types: MV Incomer, RMU, Transformer, LV Busbar, Feeder, Bus Coupler.
"Feeds From" holds the parent item's ID (comma-separated when an item has two
parents, e.g. a bus coupler between two busbars or an RMU on a ring).

Only dependency: openpyxl.
"""

import argparse
import sys
from xml.sax.saxutils import escape

from openpyxl import load_workbook

# ---------------------------------------------------------------- data model

MV_INCOMER = "mv incomer"
RMU = "rmu"
MV_BUSBAR = "mv busbar"
TRANSFORMER = "transformer"
PUMP = "pump"
LV_BUSBAR = "lv busbar"
FEEDER = "feeder"
MCC = "mcc"
BUS_COUPLER = "bus coupler"

TYPE_ALIASES = {
    MV_INCOMER: MV_INCOMER,
    "incomer": MV_INCOMER,
    "mv": MV_INCOMER,
    RMU: RMU,
    "ring main unit": RMU,
    MV_BUSBAR: MV_BUSBAR,
    "mv board": MV_BUSBAR,
    "mv switchboard": MV_BUSBAR,
    "mv distribution board": MV_BUSBAR,
    PUMP: PUMP,
    "motor": PUMP,
    "load": PUMP,
    MCC: MCC,
    "motor control centre": MCC,
    "motor control center": MCC,
    TRANSFORMER: TRANSFORMER,
    "trafo": TRANSFORMER,
    "tx": TRANSFORMER,
    LV_BUSBAR: LV_BUSBAR,
    "busbar": LV_BUSBAR,
    "lv board": LV_BUSBAR,
    "lv switchboard": LV_BUSBAR,
    FEEDER: FEEDER,
    "lv feeder": FEEDER,
    "outgoing": FEEDER,
    BUS_COUPLER: BUS_COUPLER,
    "coupler": BUS_COUPLER,
}


class Item:
    def __init__(self, id_, type_, desc, rating, voltage, parents, notes):
        self.id = id_
        self.type = type_
        self.desc = desc
        self.rating = rating
        self.voltage = voltage
        self.parents = parents  # list of parent IDs
        self.notes = notes
        self.x = None  # centre x, assigned during layout
        # busbar geometry
        self.x_left = None
        self.x_right = None

    def __repr__(self):
        return f"<{self.type} {self.id}>"


def norm_type(raw):
    key = " ".join(str(raw).lower().split())
    return TYPE_ALIASES.get(key)


# ---------------------------------------------------------------- xlsx input

def read_workbook(path):
    wb = load_workbook(path, data_only=True)

    info = {}
    if "Info" in wb.sheetnames:
        for row in wb["Info"].iter_rows(values_only=True):
            if row and row[0]:
                key = str(row[0]).rstrip(":").strip()
                val = "" if len(row) < 2 or row[1] is None else str(row[1]).strip()
                info[key.lower()] = val

    if "Equipment" not in wb.sheetnames:
        sys.exit(f"error: no 'Equipment' sheet in {path} "
                 f"(found: {', '.join(wb.sheetnames)})")

    ws = wb["Equipment"]
    rows = list(ws.iter_rows(values_only=True))
    # locate the header row (the one containing "ID" and "Type")
    header_idx = None
    for i, row in enumerate(rows):
        cells = [str(c).strip().lower() if c else "" for c in row]
        if "id" in cells and "type" in cells:
            header_idx = i
            headers = cells
            break
    if header_idx is None:
        sys.exit("error: could not find a header row with 'ID' and 'Type' "
                 "columns in the Equipment sheet")

    def col(*names):
        for n in names:
            for j, h in enumerate(headers):
                if n in h:
                    return j
        return None

    c_id = col("id")
    c_type = col("type")
    c_desc = col("desc")
    c_rating = col("rating")
    c_volt = col("volt")
    c_parent = col("feeds from", "parent", "from")
    c_notes = col("note")

    def cell(row, j):
        if j is None or j >= len(row) or row[j] is None:
            return ""
        return str(row[j]).strip()

    items, order = {}, []
    for row in rows[header_idx + 1:]:
        id_ = cell(row, c_id)
        if not id_:
            continue
        raw_type = cell(row, c_type)
        type_ = norm_type(raw_type)
        if type_ is None:
            print(f"warning: row '{id_}' has unrecognised type '{raw_type}' "
                  f"- drawing it as a generic feeder", file=sys.stderr)
            type_ = FEEDER
        parents = [p.strip() for p in cell(row, c_parent).split(",") if p.strip()]
        if id_ in items:
            sys.exit(f"error: duplicate ID '{id_}' in Equipment sheet")
        items[id_] = Item(id_, type_, cell(row, c_desc), cell(row, c_rating),
                          cell(row, c_volt), parents, cell(row, c_notes))
        order.append(id_)

    if not items:
        sys.exit("error: the Equipment sheet has no data rows")

    for it in items.values():
        for p in it.parents:
            if p not in items:
                sys.exit(f"error: '{it.id}' feeds from unknown ID '{p}' "
                         f"- check the 'Feeds From' column")

    return info, items, order


# ---------------------------------------------------------------- layout

MARGIN = 90
FEEDER_SPACING = 95
BUS_GAP = 110          # horizontal gap between adjacent busbars
MIN_BUS_WIDTH = 170

SLOT_GAP = 30          # gap between slots on an MV switchboard
PUMP_SLOT = 115        # slot width of a pump/motor way

Y_LABEL = 34           # MV incomer labels
Y_MV_TOP = 62          # top of the MV incomer stub
Y_RMU_TOP = 150
Y_RMU_BOT = 268
Y_MVBUS = 208          # MV switchboard busbar (same tier as the RMU)
Y_PUMP = 352           # pump/motor circle centre
PUMP_R = 20
Y_TX_C1 = 342          # centre of upper transformer circle
TX_R = 19
Y_TX_C2 = Y_TX_C1 + 27
Y_BUS = 486
Y_FEED_BRK = 516       # feeder breaker centre
Y_ARROW = 574          # arrow tip
Y_FEED_LBL = 592
DIAG_H = 780


def children_of(items, order, pid, types=None):
    out = []
    for oid in order:
        it = items[oid]
        if pid in it.parents and (types is None or it.type in types):
            out.append(it)
    return out


def place_lv_board(items, order, bb, center_x):
    """Place an LV busbar (and its feeders/MCCs) centred on center_x."""
    kids = children_of(items, order, bb.id, {FEEDER, MCC})
    width = max(MIN_BUS_WIDTH, len(kids) * FEEDER_SPACING)
    bb.x_left, bb.x_right = center_x - width / 2, center_x + width / 2
    bb.x = center_x
    pad = (width - len(kids) * FEEDER_SPACING) / 2
    for i, k in enumerate(kids):
        k.x = bb.x_left + pad + FEEDER_SPACING * (i + 0.5)
    return width


def slot_width(items, order, item):
    """Width needed under one way of an MV switchboard."""
    if item.type == PUMP:
        return PUMP_SLOT
    if item.type == TRANSFORMER:
        boards = children_of(items, order, item.id, {LV_BUSBAR})
        w = 130
        for bb in boards:
            kids = children_of(items, order, bb.id, {FEEDER, MCC})
            w = max(w, MIN_BUS_WIDTH, len(kids) * FEEDER_SPACING)
        return w
    return 130


def layout_mv_boards(items, order):
    """Layout when the site has MV switchboards (MV Busbar rows)."""
    mvbs = [items[i] for i in order if items[i].type == MV_BUSBAR]
    mvs = [items[i] for i in order if items[i].type == MV_INCOMER]

    x = MARGIN
    for mvb in mvbs:
        kids = children_of(items, order, mvb.id, {TRANSFORMER, PUMP})
        widths = [slot_width(items, order, k) for k in kids]
        need = sum(widths) + SLOT_GAP * max(0, len(kids) - 1)
        total = max(MIN_BUS_WIDTH, need)
        mvb.x_left, mvb.x_right = x, x + total
        mvb.x = x + total / 2
        cursor = x + (total - need) / 2
        for k, w in zip(kids, widths):
            k.x = cursor + w / 2
            cursor += w + SLOT_GAP
            if k.type == TRANSFORMER:
                for bb in children_of(items, order, k.id, {LV_BUSBAR}):
                    place_lv_board(items, order, bb, k.x)
        x = mvb.x_right + BUS_GAP

    # incomers centred over the board(s) they feed
    for m in mvs:
        for mvb in mvbs:
            feeds = [i for i in mvs if any(
                mvb.id == k.id for k in children_of(items, order, i.id))]
            if m in feeds:
                n = len(feeds)
                m.x = mvb.x + (feeds.index(m) - (n - 1) / 2) * 80
                break

    # anything left over (e.g. an RMU branch mixed in) goes after the boards
    for oid in order:
        it = items[oid]
        if it.x is None:
            it.x = x + 40
            if it.type in (LV_BUSBAR, MV_BUSBAR):
                it.x_left, it.x_right = it.x - 85, it.x + 85
            x += 130

    return max(x - BUS_GAP + MARGIN, 640) + 230


def layout(items, order):
    """Assign x coordinates to every item. Returns total drawing width."""
    if any(items[i].type == MV_BUSBAR for i in order):
        return layout_mv_boards(items, order)
    busbars = [items[i] for i in order if items[i].type == LV_BUSBAR]
    rmus = [items[i] for i in order if items[i].type == RMU]
    txs = [items[i] for i in order if items[i].type == TRANSFORMER]
    mvs = [items[i] for i in order if items[i].type == MV_INCOMER]

    # 1. busbars left-to-right, width driven by feeder count
    x = MARGIN
    for bb in busbars:
        feeders = children_of(items, order, bb.id, {FEEDER, MCC})
        width = max(MIN_BUS_WIDTH, len(feeders) * FEEDER_SPACING)
        bb.x_left, bb.x_right = x, x + width
        bb.x = x + width / 2
        for i, f in enumerate(feeders):
            f.x = bb.x_left + FEEDER_SPACING * (i + 0.5) if feeders else bb.x
            if len(feeders) * FEEDER_SPACING < width:  # centre them
                pad = (width - len(feeders) * FEEDER_SPACING) / 2
                f.x = bb.x_left + pad + FEEDER_SPACING * (i + 0.5)
        x = bb.x_right + BUS_GAP

    # 2. transformers centred over the busbar they feed
    for tx in txs:
        fed = children_of(items, order, tx.id, {LV_BUSBAR})
        if fed:
            siblings = [t for t in txs
                        if children_of(items, order, t.id, {LV_BUSBAR}) == fed]
            if len(siblings) > 1:  # several transformers on one busbar
                k = siblings.index(tx)
                spread = 90
                tx.x = fed[0].x + (k - (len(siblings) - 1) / 2) * spread
            else:
                tx.x = fed[0].x

    # 3. RMUs centred over their transformer children
    for rmu in rmus:
        kids = children_of(items, order, rmu.id, {TRANSFORMER})
        if kids and all(k.x is not None for k in kids):
            rmu.x = sum(k.x for k in kids) / len(kids)

    # 4. MV incomers spread over the RMU (or over the transformer they feed)
    for rmu in rmus:
        if rmu.x is None:
            continue
        feeds = [m for m in mvs if rmu.id in
                 [c.id for c in children_of(items, order, m.id)]]
        n = len(feeds)
        for i, m in enumerate(feeds):
            m.x = rmu.x + (i - (n - 1) / 2) * 80
    for m in mvs:
        if m.x is None:
            kids = children_of(items, order, m.id)
            placed = [k.x for k in kids if k.x is not None]
            if placed:
                m.x = sum(placed) / len(placed)

    # 5. anything still unplaced goes in a row after the busbars
    for oid in order:
        it = items[oid]
        if it.x is None:
            it.x = x + 40
            x += 130

    width = max(x - BUS_GAP + MARGIN, 640)
    width += 230  # room for side labels + title block
    return width


# ---------------------------------------------------------------- SVG symbols

class SVG:
    def __init__(self):
        self.parts = []

    def line(self, x1, y1, x2, y2, w=2, dash=None):
        d = f' stroke-dasharray="{dash}"' if dash else ""
        self.parts.append(
            f'<line x1="{x1:.1f}" y1="{y1:.1f}" x2="{x2:.1f}" y2="{y2:.1f}" '
            f'stroke="#111" stroke-width="{w}"{d} stroke-linecap="round"/>')

    def rect(self, x, y, w, h, sw=2, dash=None, fill="none"):
        d = f' stroke-dasharray="{dash}"' if dash else ""
        self.parts.append(
            f'<rect x="{x:.1f}" y="{y:.1f}" width="{w:.1f}" height="{h:.1f}" '
            f'fill="{fill}" stroke="#111" stroke-width="{sw}"{d}/>')

    def circle(self, x, y, r, sw=2):
        self.parts.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{r}" fill="none" '
            f'stroke="#111" stroke-width="{sw}"/>')

    def dot(self, x, y, r=3.2):
        self.parts.append(
            f'<circle cx="{x:.1f}" cy="{y:.1f}" r="{r}" fill="#111"/>')

    def poly(self, pts, fill="#111"):
        p = " ".join(f"{a:.1f},{b:.1f}" for a, b in pts)
        self.parts.append(f'<polygon points="{p}" fill="{fill}"/>')

    def text(self, x, y, s, size=12, anchor="middle", bold=False,
             rotate=None, color="#111"):
        if not s:
            return
        wgt = ' font-weight="bold"' if bold else ""
        tr = (f' transform="translate({x:.1f},{y:.1f}) rotate({rotate})"'
              if rotate is not None else "")
        xy = 'x="0" y="0"' if rotate is not None else f'x="{x:.1f}" y="{y:.1f}"'
        self.parts.append(
            f'<text {xy}{tr} font-family="Arial, Helvetica, sans-serif" '
            f'font-size="{size}" fill="{color}" text-anchor="{anchor}"{wgt}>'
            f'{escape(str(s))}</text>')

    # -- composite symbols ------------------------------------------------

    def load_break_switch(self, x, ytop, ybot):
        """IEC-style switch: short stub, diagonal blade, contact tick."""
        self.line(x, ytop, x, ytop + 5)
        self.line(x, ytop + 5, x + 9, ybot - 9)          # blade
        self.line(x - 6, ybot - 9, x + 6, ybot - 9)      # contact bar
        self.line(x, ybot - 9, x, ybot)

    def fuse_switch(self, x, ytop, ybot):
        """Switch with a fuse rectangle on the lower run (RMU tee-off)."""
        mid = (ytop + ybot) / 2
        self.load_break_switch(x, ytop, mid + 4)
        self.rect(x - 4, mid + 6, 8, ybot - mid - 8)
        self.line(x, mid + 4, x, mid + 6)
        self.line(x, ybot - 2, x, ybot)

    def breaker(self, x, y, size=13):
        """LV circuit breaker: open square on the conductor."""
        h = size / 2
        self.rect(x - h, y - h, size, size, sw=2, fill="white")

    def transformer(self, x, label_lines):
        self.circle(x, Y_TX_C1, TX_R, sw=2.2)
        self.circle(x, Y_TX_C2, TX_R, sw=2.2)
        ty = Y_TX_C1 - 6
        for s in label_lines:
            self.text(x + TX_R + 10, ty, s, anchor="start")
            ty += 15

    def arrow_down(self, x, ytip):
        self.poly([(x - 6, ytip - 11), (x + 6, ytip - 11), (x, ytip)])


# ---------------------------------------------------------------- rendering

def render(info, items, order, width):
    svg = SVG()
    busbars = [items[i] for i in order if items[i].type == LV_BUSBAR]
    mvbs = [items[i] for i in order if items[i].type == MV_BUSBAR]
    rmus = [items[i] for i in order if items[i].type == RMU]
    txs = [items[i] for i in order if items[i].type == TRANSFORMER]
    pumps = [items[i] for i in order if items[i].type == PUMP]
    mvs = [items[i] for i in order if items[i].type == MV_INCOMER]
    couplers = [items[i] for i in order if items[i].type == BUS_COUPLER]
    feeders = [items[i] for i in order if items[i].type in (FEEDER, MCC)]

    site = info.get("site", "")
    title = f"{site} — Single Line Diagram (sketch)" if site \
        else "Single Line Diagram (sketch)"
    svg.text(24, DIAG_H - 26, title, size=16, anchor="start", bold=True)

    # --- RMU enclosures --------------------------------------------------
    rmu_box = {}
    for rmu in rmus:
        ways_in = [m for m in mvs if any(
            rmu.id == k.id for k in children_of(items, order, m.id))]
        ways_out = children_of(items, order, rmu.id, {TRANSFORMER})
        xs = [w.x for w in ways_in + ways_out if w.x is not None] or [rmu.x]
        left, right = min(xs) - 42, max(xs) + 42
        rmu_box[rmu.id] = (left, right)
        svg.rect(left, Y_RMU_TOP, right - left, Y_RMU_BOT - Y_RMU_TOP,
                 sw=1.6, dash="7 5")
        # internal bus
        ymid = (Y_RMU_TOP + Y_RMU_BOT) / 2
        svg.line(min(xs) - 18, ymid, max(xs) + 18, ymid, w=3.4)
        # incoming ways: load-break switches from the top edge to the bus
        for m in ways_in:
            svg.load_break_switch(m.x, Y_RMU_TOP + 12, ymid)
            svg.line(m.x, Y_RMU_TOP, m.x, Y_RMU_TOP + 12)
            svg.dot(m.x, ymid)
        # outgoing ways: fuse-switches from the bus to the bottom edge
        for t in ways_out:
            svg.fuse_switch(t.x, ymid + 4, Y_RMU_BOT - 8)
            svg.line(t.x, ymid, t.x, ymid + 4)
            svg.line(t.x, Y_RMU_BOT - 8, t.x, Y_RMU_BOT)
            svg.dot(t.x, ymid)
        lbl = [rmu.id, rmu.desc,
               " ".join(v for v in (rmu.rating, rmu.voltage) if v)]
        ty = Y_RMU_TOP + 16
        for i, s in enumerate(lbl):
            svg.text(right + 10, ty, s, anchor="start", bold=(i == 0))
            ty += 15

    # --- MV incomers -----------------------------------------------------
    for m in mvs:
        kids = children_of(items, order, m.id)
        svg.line(m.x - 11, Y_MV_TOP, m.x + 11, Y_MV_TOP, w=3)  # source tick
        lbl = [m.id, m.desc, m.voltage]
        ty = Y_LABEL - 16
        for i, s in enumerate(lbl):
            svg.text(m.x, ty, s, size=11.5, bold=(i == 0))
            ty += 14
        for k in kids:
            if k.type == RMU:
                svg.line(m.x, Y_MV_TOP, m.x, Y_RMU_TOP)
            elif k.type == MV_BUSBAR:  # incoming breaker onto the board
                ybrk = (Y_MV_TOP + Y_MVBUS) / 2
                svg.line(m.x, Y_MV_TOP, m.x, ybrk - 8)
                svg.breaker(m.x, ybrk)
                svg.line(m.x, ybrk + 8, m.x, Y_MVBUS)
                svg.dot(m.x, Y_MVBUS)
            elif k.type == TRANSFORMER:  # direct feed, no RMU
                svg.line(m.x, Y_MV_TOP, m.x, Y_TX_C1 - TX_R)

    # --- MV switchboards -------------------------------------------------
    for mvb in mvbs:
        svg.line(mvb.x_left, Y_MVBUS, mvb.x_right, Y_MVBUS, w=5.5)
        lbl = " ".join(v for v in (mvb.id, mvb.desc, mvb.rating, mvb.voltage)
                       if v)
        svg.text(mvb.x_left, Y_MVBUS - 12, lbl, size=11.5, anchor="start",
                 bold=True)

    # --- pumps / motor loads --------------------------------------------
    for p in pumps:
        if p.x is None:
            continue
        for par in (items[q] for q in p.parents):
            if par.type == MV_BUSBAR:
                ybrk = (Y_MVBUS + Y_PUMP - PUMP_R) / 2
                svg.line(p.x, Y_MVBUS, p.x, ybrk - 8)
                svg.dot(p.x, Y_MVBUS)
                svg.breaker(p.x, ybrk)
                svg.line(p.x, ybrk + 8, p.x, Y_PUMP - PUMP_R)
            elif par.type == RMU:
                svg.line(p.x, Y_RMU_BOT, p.x, Y_PUMP - PUMP_R)
        svg.circle(p.x, Y_PUMP, PUMP_R, sw=2.2)
        svg.text(p.x, Y_PUMP + 5, "M", size=15, bold=True)
        lbl = " · ".join(v for v in (p.id, p.desc, p.rating) if v)
        svg.text(p.x + 4, Y_PUMP + PUMP_R + 14, lbl, size=11,
                 anchor="start", rotate=90)

    # --- transformers ----------------------------------------------------
    for tx in txs:
        if tx.x is None:
            continue
        lbl = [tx.id, tx.desc,
               " ".join(v for v in (tx.rating, tx.voltage) if v)]
        svg.transformer(tx.x, [s for s in lbl if s])
        for p in tx.parents:
            par = items[p]
            if par.type == RMU:
                svg.line(tx.x, Y_RMU_BOT, tx.x, Y_TX_C1 - TX_R)
            elif par.type == MV_BUSBAR:  # feeder breaker off the board
                ybrk = (Y_MVBUS + Y_TX_C1 - TX_R) / 2
                svg.line(tx.x, Y_MVBUS, tx.x, ybrk - 8)
                svg.dot(tx.x, Y_MVBUS)
                svg.breaker(tx.x, ybrk)
                svg.line(tx.x, ybrk + 8, tx.x, Y_TX_C1 - TX_R)
        fed = children_of(items, order, tx.id, {LV_BUSBAR})
        for bb in fed:
            # drop to the busbar through the LV incomer breaker
            ybrk = (Y_TX_C2 + TX_R + Y_BUS) / 2
            svg.line(tx.x, Y_TX_C2 + TX_R, tx.x, ybrk - 8)
            svg.breaker(tx.x, ybrk)
            svg.line(tx.x, ybrk + 8, tx.x, Y_BUS)
            svg.dot(tx.x, Y_BUS)

    # --- busbars ---------------------------------------------------------
    for bb in busbars:
        svg.line(bb.x_left, Y_BUS, bb.x_right, Y_BUS, w=5.5)
        lbl = " ".join(v for v in (bb.id, bb.desc, bb.rating, bb.voltage) if v)
        svg.text(bb.x_left, Y_BUS - 12, lbl, size=11.5, anchor="start",
                 bold=True)

    # --- bus couplers / ties ---------------------------------------------
    for bc in couplers:
        ends = [items[p] for p in bc.parents
                if items[p].type in (LV_BUSBAR, MV_BUSBAR)]
        if len(ends) != 2 or ends[0].type != ends[1].type:
            print(f"warning: bus coupler '{bc.id}' should feed from exactly "
                  f"two busbars of the same kind - skipping", file=sys.stderr)
            continue
        y = Y_MVBUS if ends[0].type == MV_BUSBAR else Y_BUS
        a, b = sorted(ends, key=lambda e: e.x)
        xm = (a.x_right + b.x_left) / 2
        svg.line(a.x_right, y, xm - 8, y, w=2)
        svg.breaker(xm, y)
        svg.line(xm + 8, y, b.x_left, y, w=2)
        lbl = " ".join(v for v in (bc.id, bc.rating) if v)
        svg.text(xm, y + 24, lbl, size=11)
        svg.text(xm, y + 38, bc.notes, size=10)

    # --- feeders ---------------------------------------------------------
    for f in feeders:
        if f.x is None:
            continue
        svg.line(f.x, Y_BUS, f.x, Y_FEED_BRK - 7)
        svg.dot(f.x, Y_BUS)
        svg.breaker(f.x, Y_FEED_BRK)
        if f.type == MCC:  # motor control centre: box instead of arrow
            svg.line(f.x, Y_FEED_BRK + 7, f.x, Y_ARROW - 26)
            svg.rect(f.x - 14, Y_ARROW - 26, 28, 26, sw=2)
            svg.text(f.x, Y_ARROW - 8, "MCC", size=8)
        else:
            svg.line(f.x, Y_FEED_BRK + 7, f.x, Y_ARROW - 10)
            svg.arrow_down(f.x, Y_ARROW)
        lbl = " · ".join(v for v in (f.id, f.desc, f.rating) if v)
        svg.text(f.x + 4, Y_FEED_LBL, lbl, size=11, anchor="start", rotate=90)

    # --- title block -----------------------------------------------------
    tb_w, tb_h = 288, 96
    tb_x, tb_y = width - tb_w - 24, DIAG_H - tb_h - 20
    svg.rect(tb_x, tb_y, tb_w, tb_h, sw=1.5)
    svg.line(tb_x, tb_y + 24, tb_x + tb_w, tb_y + 24, w=1.5)
    svg.text(tb_x + 10, tb_y + 17, "SLD SKETCH — SITE SURVEY", size=12,
             anchor="start", bold=True)
    rows = [("Site", info.get("site", "")),
            ("Date", info.get("date", "")),
            ("By", info.get("surveyed by", info.get("by", ""))),
            ("Notes", info.get("notes", ""))]
    ty = tb_y + 40
    for k, v in rows:
        svg.text(tb_x + 10, ty, f"{k}:", size=11, anchor="start", bold=True)
        svg.text(tb_x + 58, ty, v[:44], size=11, anchor="start")
        ty += 16

    body = "\n".join(svg.parts)
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.0f}" '
            f'height="{DIAG_H}" viewBox="0 0 {width:.0f} {DIAG_H}">\n'
            f'<rect width="100%" height="100%" fill="white"/>\n'
            f'{body}\n</svg>\n')


# ---------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser(
        description="Sketch a single-line diagram from a site-survey workbook.")
    ap.add_argument("workbook", help="input .xlsx file")
    ap.add_argument("-o", "--output",
                    help="output .svg file (default: <workbook>.svg)")
    args = ap.parse_args()

    out = args.output or (args.workbook.rsplit(".", 1)[0] + ".svg")
    info, items, order = read_workbook(args.workbook)
    width = layout(items, order)
    svg = render(info, items, order, width)
    with open(out, "w", encoding="utf-8") as fh:
        fh.write(svg)
    print(f"wrote {out}  ({len(items)} items)")


if __name__ == "__main__":
    main()
