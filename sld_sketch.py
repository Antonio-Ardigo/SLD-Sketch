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


PROT_ALIASES = {
    "cb": "cb", "circuit breaker": "cb", "breaker": "cb", "acb": "cb",
    "mccb": "cb", "mcb": "cb", "vcb": "cb",
    "lbs": "lbs", "load break switch": "lbs", "load-break switch": "lbs",
    "switch": "lbs", "isolator": "lbs", "disconnector": "lbs",
    "fuse": "fuse", "fuses": "fuse",
    "fuse-switch": "fuse-switch", "fuse switch": "fuse-switch",
    "switch-fuse": "fuse-switch", "switch fuse": "fuse-switch",
    "sfu": "fuse-switch",
    "contactor": "contactor", "vacuum contactor": "contactor",
    "fuse-contactor": "fuse-contactor", "fuse contactor": "fuse-contactor",
    "fused contactor": "fuse-contactor", "contactor-fuse": "fuse-contactor",
    "motor starter": "fuse-contactor", "starter": "fuse-contactor",
}


def prot_for(item, parent_id=None):
    """The Protection entry for this item's supply from parent_id.

    Returns (raw_text, normalised_kind_or_None). One value applies to
    every supply; a comma list matches the Feeds From order.
    """
    if not item.prots:
        return "", None
    raw = item.prots[0]
    if parent_id is not None and len(item.prots) > 1 \
            and parent_id in item.parents:
        i = item.parents.index(parent_id)
        if i < len(item.prots):
            raw = item.prots[i]
    return raw, PROT_ALIASES.get(" ".join(raw.lower().split()))


class Item:
    def __init__(self, id_, type_, desc, rating, voltage, parents, notes,
                 prots=None):
        self.id = id_
        self.type = type_
        self.desc = desc
        self.rating = rating
        self.voltage = voltage
        self.parents = parents  # list of parent IDs
        self.notes = notes
        self.prots = prots or []  # protection, one entry per supply
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
    c_prot = col("protection", "prot")
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
        prots = [p.strip() for p in cell(row, c_prot).split(",") if p.strip()]
        if id_ in items:
            sys.exit(f"error: duplicate ID '{id_}' in Equipment sheet")
        items[id_] = Item(id_, type_, cell(row, c_desc), cell(row, c_rating),
                          cell(row, c_volt), parents, cell(row, c_notes),
                          prots)
        order.append(id_)

    if not items:
        sys.exit("error: the Equipment sheet has no data rows")

    for it in items.values():
        for p in it.parents:
            if p not in items:
                sys.exit(f"error: '{it.id}' feeds from unknown ID '{p}' "
                         f"- check the 'Feeds From' column")
        if it.prots and it.type == MV_INCOMER:
            print(f"warning: '{it.id}' - protection on an MV incomer is on "
                  f"the utility side and is not drawn", file=sys.stderr)
        elif it.type not in (LV_BUSBAR, MV_BUSBAR, BUS_COUPLER):
            for raw in it.prots:
                if PROT_ALIASES.get(" ".join(raw.lower().split())) is None:
                    print(f"warning: '{it.id}' - unknown protection "
                          f"'{raw}', the default symbol is drawn",
                          file=sys.stderr)

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
LEGEND_H = 100         # symbol legend strip below the drawing


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
        x = mvb.x_right + BUS_GAP

    # LV boards centred under their supply transformer(s) - the mean of
    # the supplies when a board has two incomers
    for oid in order:
        bb = items[oid]
        if bb.type != LV_BUSBAR or bb.x is not None:
            continue
        pxs = [items[p].x for p in bb.parents
               if items[p].type == TRANSFORMER and items[p].x is not None]
        if pxs:
            place_lv_board(items, order, bb, sum(pxs) / len(pxs))

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
            else:  # centred over its panel(s) - a TX may feed two boards
                tx.x = sum(bb.x for bb in fed) / len(fed)

    # 3. RMUs centred over their transformer children
    for rmu in rmus:
        kids = children_of(items, order, rmu.id, {TRANSFORMER})
        if kids and all(k.x is not None for k in kids):
            rmu.x = sum(k.x for k in kids) / len(kids)
    for rmu in rmus:  # cascaded RMU: centre over the RMUs it feeds
        if rmu.x is None:
            kids = children_of(items, order, rmu.id, {RMU})
            placed = [k.x for k in kids if k.x is not None]
            if placed:
                rmu.x = sum(placed) / len(placed)

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

    def path(self, d, sw=2):
        self.parts.append(
            f'<path d="{d}" fill="none" stroke="#111" stroke-width="{sw}" '
            f'stroke-linecap="round"/>')

    # -- composite symbols ------------------------------------------------

    def load_break_switch(self, x, ytop, ybot):
        """IEC switch-disconnector: hinge circle, blade, contact bar."""
        self.line(x, ytop, x, ytop + 3)
        self.circle(x, ytop + 5.5, 2.5, sw=1.5)          # hinge = switch
        self.line(x + 2, ytop + 8, x + 9, ybot - 9)      # blade
        self.line(x - 6, ybot - 9, x + 6, ybot - 9)      # contact bar
        self.line(x, ybot - 9, x, ybot)

    def fuse_switch(self, x, ytop, ybot):
        """Switch with a fuse rectangle on the lower run (RMU tee-off)."""
        mid = (ytop + ybot) / 2
        self.load_break_switch(x, ytop, mid + 4)
        self.rect(x - 4, mid + 6, 8, ybot - mid - 8)
        self.line(x, mid + 4, x, mid + 6)
        self.line(x, ybot - 2, x, ybot)

    def device(self, kind, x, y):
        """Protection device centred at y on a vertical conductor.
        Returns the half-height the conductor must leave clear."""
        if kind == "fuse":
            self.rect(x - 4, y - 11, 8, 22)
            self.line(x, y - 11, x, y + 11)
            return 11
        if kind == "lbs":
            self.load_break_switch(x, y - 13, y + 13)
            return 13
        if kind == "fuse-switch":
            self.fuse_switch(x, y - 16, y + 16)
            return 16
        if kind == "contactor":
            # IEC: switch with the arc function symbol at the hinge,
            # open side facing the blade
            self.line(x, y - 13, x, y - 11)
            self.path(f"M {x-4:.1f},{y-7:.1f} A 4,4 0 0 1 "
                      f"{x+4:.1f},{y-7:.1f}")                    # hinge arc
            self.line(x + 2, y - 5, x + 8, y + 7)                # blade
            self.line(x, y + 9, x, y + 13)
            return 13
        if kind == "fuse-contactor":
            # MV motor starter: back-up fuse in series with the contactor
            self.rect(x - 4, y - 16, 8, 12)                      # fuse
            self.line(x, y - 16, x, y - 4)
            self.path(f"M {x-4:.1f},{y:.1f} A 4,4 0 0 1 "
                      f"{x+4:.1f},{y:.1f}")                      # hinge arc
            self.line(x + 2, y + 2, x + 7, y + 11)               # blade
            self.line(x, y + 12, x, y + 16)
            return 16
        # 'cb' and anything unknown - circuit breaker: the switch
        # symbol with an X at its hinge edge (IEC 60617)
        self.line(x, y - 13, x, y - 11)
        self.line(x - 3.5, y - 11, x + 3.5, y - 4)               # X
        self.line(x - 3.5, y - 4, x + 3.5, y - 11)
        self.line(x + 2, y - 4.5, x + 9, y + 4)                  # blade
        self.line(x - 6, y + 4, x + 6, y + 4)                    # contact bar
        self.line(x, y + 4, x, y + 13)
        return 13

    def device_h(self, kind, x, y):
        """Protection device centred at x on a horizontal conductor.
        Returns the half-width the conductor must leave clear."""
        if kind == "fuse":
            self.rect(x - 11, y - 4, 22, 8)
            self.line(x - 11, y, x + 11, y)
            return 11
        if kind == "fuse-switch":
            self.line(x - 16, y, x - 14.5, y)
            self.circle(x - 12, y, 2.5, sw=1.5)     # hinge = switch
            self.line(x - 10, y - 1.5, x + 1, y - 9)  # blade
            self.line(x + 1, y - 6, x + 1, y + 6)   # contact bar
            self.rect(x + 3, y - 4, 12, 8)          # fuse
            self.line(x + 1, y, x + 3, y)
            self.line(x + 15, y, x + 16, y)
            return 16
        if kind == "contactor":
            # IEC: switch with the arc function symbol at the hinge
            self.line(x - 13, y, x - 11, y)
            self.path(f"M {x-7:.1f},{y-4:.1f} A 4,4 0 0 0 "
                      f"{x-7:.1f},{y+4:.1f}")                    # hinge arc
            self.line(x - 5, y - 2, x + 7, y - 8)                # blade
            self.line(x + 9, y, x + 13, y)
            return 13
        if kind == "fuse-contactor":
            self.rect(x - 16, y - 4, 12, 8)                      # fuse
            self.line(x - 16, y, x - 4, y)
            self.path(f"M {x:.1f},{y-4:.1f} A 4,4 0 0 0 "
                      f"{x:.1f},{y+4:.1f}")                      # hinge arc
            self.line(x + 2, y - 2, x + 11, y - 7)               # blade
            self.line(x + 12, y, x + 16, y)
            return 16
        if kind == "lbs":
            self.line(x - 13, y, x - 10.5, y)
            self.circle(x - 8, y, 2.5, sw=1.5)      # hinge = switch
            self.line(x - 6, y - 1.5, x + 4, y - 9)  # blade
            self.line(x + 4, y - 6, x + 4, y + 6)   # contact bar
            self.line(x + 4, y, x + 13, y)
            return 13
        # 'cb' and anything unknown - circuit breaker: the switch
        # symbol with an X at its hinge edge (IEC 60617)
        self.line(x - 13, y, x - 11, y)
        self.line(x - 11, y - 3.5, x - 4, y + 3.5)               # X
        self.line(x - 11, y + 3.5, x - 4, y - 3.5)
        self.line(x - 4.5, y - 2, x + 4, y - 9)                  # blade
        self.line(x + 4, y - 6, x + 4, y + 6)                    # contact bar
        self.line(x + 4, y, x + 13, y)
        return 13

    def drop(self, x, ytop, ybot, kind, ydev=None):
        """Vertical conductor with a protection device on it."""
        y = ydev if ydev is not None else (ytop + ybot) / 2
        gap = self.device(kind, x, y)
        self.line(x, ytop, x, y - gap)
        self.line(x, y + gap, x, ybot)

    def transformer(self, x, label_lines, side="right"):
        self.circle(x, Y_TX_C1, TX_R, sw=2.2)
        self.circle(x, Y_TX_C2, TX_R, sw=2.2)
        ty = Y_TX_C1 - 6
        for s in label_lines:
            if side == "left":
                self.text(x - TX_R - 10, ty, s, anchor="end")
            else:
                self.text(x + TX_R + 10, ty, s, anchor="start")
            ty += 15

    def arrow_down(self, x, ytip):
        self.poly([(x - 6, ytip - 11), (x + 6, ytip - 11), (x, ytip)])


# ---------------------------------------------------------------- rendering

LEGEND_ITEMS = [
    ("cb", "Circuit breaker"), ("lbs", "Load-break switch"),
    ("fuse", "Fuse"), ("fuse-switch", "Fuse-switch"),
    ("contactor", "Contactor"), ("fuse-contactor", "Fused contactor"),
    ("tx", "Transformer"), ("pump", "Pump/motor"), ("bus", "Busbar"),
    ("mcc", "MCC"), ("feeder", "Feeder"), ("rmu", "RMU enclosure"),
]


def draw_legend(svg):
    cell = 68
    x0, y0 = 24, DIAG_H + 6
    svg.rect(x0, y0, 16 + cell * len(LEGEND_ITEMS), 82, sw=1.2)
    svg.text(x0 + 8, y0 + 14, "LEGEND", size=10, anchor="start", bold=True)
    ytop, ybot = y0 + 22, y0 + 52
    yc = (ytop + ybot) / 2
    for i, (kind, label) in enumerate(LEGEND_ITEMS):
        cx = x0 + 8 + cell * i + cell / 2
        if kind in ("cb", "fuse", "contactor", "fuse-contactor"):
            svg.drop(cx, ytop, ybot, kind)
        elif kind == "lbs":
            svg.load_break_switch(cx, ytop, ybot)
        elif kind == "fuse-switch":
            svg.fuse_switch(cx, ytop, ybot)
        elif kind == "tx":
            svg.circle(cx, yc - 5, 8)
            svg.circle(cx, yc + 5, 8)
        elif kind == "pump":
            svg.circle(cx, yc, 9)
            svg.text(cx, yc + 3.5, "M", size=9, bold=True)
        elif kind == "bus":
            svg.line(cx - 14, yc, cx + 14, yc, w=5)
        elif kind == "mcc":
            svg.rect(cx - 11, yc - 8, 22, 16, sw=1.5)
            svg.text(cx, yc + 3, "MCC", size=6.5)
        elif kind == "feeder":
            svg.line(cx, ytop + 2, cx, ybot - 11)
            svg.arrow_down(cx, ybot)
        elif kind == "rmu":
            svg.rect(cx - 13, yc - 10, 26, 20, sw=1.2, dash="4 3")
        ty = y0 + 64
        for s in (label.split(" ", 1) if len(label) > 11 else [label]):
            svg.text(cx, ty, s, size=9)
            ty += 10


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

    # --- RMU-to-RMU link topology ----------------------------------------
    # straight cable between neighbouring boxes; when another RMU sits in
    # between (a ring closure), the cable loops over the top instead and
    # enters each box through an extra load-break-switch way
    side_links, ring_links = [], []
    ring_entries = {}  # rmu id -> extra top-entry x positions
    for rmu in rmus:
        for p in rmu.parents:
            if items[p].type != RMU:
                continue
            if items[p].x is None or rmu.x is None:
                continue
            a, b = sorted((items[p], rmu), key=lambda r: r.x)
            between = any(r is not a and r is not b and a.x < r.x < b.x
                          for r in rmus)
            if between:
                xa, xb = a.x - 28, b.x + 28
                ring_entries.setdefault(a.id, []).append((xa, b.id))
                ring_entries.setdefault(b.id, []).append((xb, a.id))
                ring_links.append((xa, xb))
            else:
                side_links.append((a, b))

    # linked sides get a wider enclosure so the way switch fits inside
    pad_l = {r.id: 42 for r in rmus}
    pad_r = {r.id: 42 for r in rmus}
    for a, b in side_links:
        pad_r[a.id] = 64
        pad_l[b.id] = 64

    # --- RMU enclosures --------------------------------------------------
    rmu_box = {}
    for rmu in rmus:
        ways_in = [m for m in mvs if any(
            rmu.id == k.id for k in children_of(items, order, m.id))]
        ways_out = children_of(items, order, rmu.id, {TRANSFORMER})
        xs = [w.x for w in ways_in + ways_out if w.x is not None] or [rmu.x]
        xs = xs + [e[0] for e in ring_entries.get(rmu.id, [])]
        left = min(xs) - pad_l[rmu.id]
        right = max(xs) + pad_r[rmu.id]
        bus_l, bus_r = min(xs) - 18, max(xs) + 18
        rmu_box[rmu.id] = (left, right, bus_l, bus_r)
        svg.rect(left, Y_RMU_TOP, right - left, Y_RMU_BOT - Y_RMU_TOP,
                 sw=1.6, dash="7 5")
        # internal bus
        ymid = (Y_RMU_TOP + Y_RMU_BOT) / 2
        svg.line(bus_l, ymid, bus_r, ymid, w=3.4)
        # incoming ways from the top edge to the bus (default LBS,
        # overridden by the RMU row's Protection entry for that supply)
        def in_way(x, kind):
            if kind and kind != "lbs":
                svg.drop(x, Y_RMU_TOP, ymid, kind)
            else:
                svg.load_break_switch(x, Y_RMU_TOP + 12, ymid)
                svg.line(x, Y_RMU_TOP, x, Y_RMU_TOP + 12)
            svg.dot(x, ymid)
        for m in ways_in:
            in_way(m.x, prot_for(rmu, m.id)[1])
        # ring-closure entries come in through the top the same way
        for x_e, other in ring_entries.get(rmu.id, []):
            in_way(x_e, prot_for(rmu, other)[1]
                   if other in rmu.parents else None)
        # outgoing ways from the bus to the bottom edge (default
        # fuse-switch, overridden by the fed item's Protection)
        for t in ways_out:
            kind = prot_for(t, rmu.id)[1]
            if kind and kind != "fuse-switch":
                svg.drop(t.x, ymid, Y_RMU_BOT, kind)
            else:
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

    # --- RMU-to-RMU interconnecting cables -------------------------------
    y_link = (Y_RMU_TOP + Y_RMU_BOT) / 2
    for a, b in side_links:  # a is the left box
        if a.id not in rmu_box or b.id not in rmu_box:
            continue
        _, a_right, _, a_bus_r = rmu_box[a.id]
        b_left, _, b_bus_l, _ = rmu_box[b.id]
        svg.line(a_right, y_link, b_left, y_link, w=2)  # cable between boxes
        # the way switch inside each box, between the wall and the bus
        # (default LBS; the fed RMU's Protection entry can override it)
        for xe, xc, owner, other in ((a_right, a_bus_r, a, b),
                                     (b_left, b_bus_l, b, a)):
            kind = (prot_for(owner, other.id)[1]
                    if other.id in owner.parents else None) or "lbs"
            xm = (xe + xc) / 2
            gap = svg.device_h(kind, xm, y_link)
            lo, hi = min(xe, xc), max(xe, xc)
            svg.line(lo, y_link, xm - gap, y_link, w=2)
            svg.line(xm + gap, y_link, hi, y_link, w=2)
    y_ring = Y_RMU_TOP - 26
    for xa, xb in ring_links:  # loop over the top of the boxes in between
        svg.line(xa, Y_RMU_TOP, xa, y_ring)
        svg.line(xa, y_ring, xb, y_ring)
        svg.line(xb, y_ring, xb, Y_RMU_TOP)

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
            elif k.type == MV_BUSBAR:  # incoming device onto the board
                svg.drop(m.x, Y_MV_TOP, Y_MVBUS,
                         prot_for(k, m.id)[1] or "cb")
                svg.dot(m.x, Y_MVBUS)
            elif k.type == TRANSFORMER:  # direct feed, no RMU
                svg.line(m.x, Y_MV_TOP, m.x, Y_TX_C1 - TX_R)

    # --- MV switchboards -------------------------------------------------
    for mvb in mvbs:
        svg.line(mvb.x_left, Y_MVBUS, mvb.x_right, Y_MVBUS, w=5.5)
        raw, kind = prot_for(mvb)
        zone = raw if raw and kind is None else ""  # e.g. 87B differential
        lbl = " ".join(v for v in (mvb.id, mvb.desc, mvb.rating, mvb.voltage)
                       if v)
        if zone:
            lbl += " · " + zone
        svg.text(mvb.x_left, Y_MVBUS - 12, lbl, size=11.5, anchor="start",
                 bold=True)

    # --- pumps / motor loads --------------------------------------------
    for p in pumps:
        if p.x is None:
            continue
        for par in (items[q] for q in p.parents):
            if par.type == MV_BUSBAR:
                svg.drop(p.x, Y_MVBUS, Y_PUMP - PUMP_R,
                         prot_for(p, par.id)[1] or "cb")
                svg.dot(p.x, Y_MVBUS)
            elif par.type == RMU:
                kind = prot_for(p, par.id)[1]
                if kind:
                    svg.drop(p.x, Y_RMU_BOT, Y_PUMP - PUMP_R, kind)
                else:
                    svg.line(p.x, Y_RMU_BOT, p.x, Y_PUMP - PUMP_R)
        svg.circle(p.x, Y_PUMP, PUMP_R, sw=2.2)
        svg.text(p.x, Y_PUMP + 1, "M", size=13, bold=True)
        svg.text(p.x, Y_PUMP + 13, "3~", size=8.5)
        lbl = " · ".join(v for v in (p.id, p.desc, p.rating) if v)
        svg.text(p.x + 4, Y_PUMP + PUMP_R + 14, lbl, size=11,
                 anchor="start", rotate=90)

    # --- transformers ----------------------------------------------------
    for tx in txs:
        if tx.x is None:
            continue
        fed = children_of(items, order, tx.id, {LV_BUSBAR})
        lbl = [tx.id, tx.desc,
               " ".join(v for v in (tx.rating, tx.voltage) if v)]
        # transformers sharing a board label away from each other
        side = ("left" if any(len(bb.parents) > 1 and tx.x < bb.x
                              for bb in fed) else "right")
        svg.transformer(tx.x, [s for s in lbl if s], side)
        for p in tx.parents:
            par = items[p]
            if par.type == RMU:
                svg.line(tx.x, Y_RMU_BOT, tx.x, Y_TX_C1 - TX_R)
            elif par.type == MV_BUSBAR:  # feeder device off the board
                svg.drop(tx.x, Y_MVBUS, Y_TX_C1 - TX_R,
                         prot_for(tx, par.id)[1] or "cb")
                svg.dot(tx.x, Y_MVBUS)
        ytop = Y_TX_C2 + TX_R
        dual = [bb for bb in fed if len(bb.parents) > 1]
        rest = [bb for bb in fed if len(bb.parents) == 1]
        for bb in dual:
            # a board with two incomers: each supply gets its own drop
            # at its own position, never at the shared board centre
            kind = prot_for(bb, tx.id)[1] or "cb"
            xc = min(max(tx.x, bb.x_left + 25), bb.x_right - 25)
            if abs(xc - tx.x) < 1:
                svg.drop(tx.x, ytop, Y_BUS, kind)
            else:  # supply sits outside the board span: elbow over
                ysplit = ytop + 32
                svg.line(tx.x, ytop, tx.x, ysplit)
                svg.line(tx.x, ysplit, xc, ysplit)
                svg.drop(xc, ysplit, Y_BUS, kind)
            svg.dot(xc, Y_BUS)
        if len(rest) == 1 and abs(rest[0].x - tx.x) < 1:
            # straight drop to the busbar through the LV incomer device
            svg.drop(tx.x, ytop, Y_BUS,
                     prot_for(rest[0], tx.id)[1] or "cb")
            svg.dot(tx.x, Y_BUS)
        elif rest:
            # one transformer feeding several panels: split, then one
            # incomer device per panel
            ysplit = ytop + 32
            svg.line(tx.x, ytop, tx.x, ysplit)
            xs = [bb.x for bb in rest] + [tx.x]
            svg.line(min(xs), ysplit, max(xs), ysplit)
            for bb in rest:
                svg.dot(bb.x, ysplit)
                svg.drop(bb.x, ysplit, Y_BUS,
                         prot_for(bb, tx.id)[1] or "cb")
                svg.dot(bb.x, Y_BUS)

    # --- busbars ---------------------------------------------------------
    for bb in busbars:
        svg.line(bb.x_left, Y_BUS, bb.x_right, Y_BUS, w=5.5)
        raw, kind = prot_for(bb)
        zone = raw if raw and kind is None else ""  # e.g. 87B differential
        lbl = " ".join(v for v in (bb.id, bb.desc, bb.rating, bb.voltage) if v)
        if zone:
            lbl += " · " + zone
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
        raw, kind = prot_for(bc)
        gap = svg.device_h(kind or "cb", xm, y)
        svg.line(a.x_right, y, xm - gap, y, w=2)
        svg.line(xm + gap, y, b.x_left, y, w=2)
        extra = raw if raw and kind is None else ""
        lbl = " ".join(v for v in (bc.id, bc.rating, extra) if v)
        svg.text(xm, y + 24, lbl, size=11)
        svg.text(xm, y + 38, bc.notes, size=10)

    # --- feeders ---------------------------------------------------------
    for f in feeders:
        if f.x is None:
            continue
        kind = prot_for(f, f.parents[0] if f.parents else None)[1] or "cb"
        svg.dot(f.x, Y_BUS)
        if f.type == MCC:  # motor control centre: box instead of arrow
            svg.drop(f.x, Y_BUS, Y_ARROW - 26, kind, ydev=Y_FEED_BRK)
            svg.rect(f.x - 14, Y_ARROW - 26, 28, 26, sw=2)
            svg.text(f.x, Y_ARROW - 8, "MCC", size=8)
        else:
            svg.drop(f.x, Y_BUS, Y_ARROW - 10, kind, ydev=Y_FEED_BRK)
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

    draw_legend(svg)

    body = "\n".join(svg.parts)
    h = DIAG_H + LEGEND_H
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{width:.0f}" '
            f'height="{h}" viewBox="0 0 {width:.0f} {h}">\n'
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
