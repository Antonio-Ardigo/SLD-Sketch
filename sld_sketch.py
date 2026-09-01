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
GENERATOR = "generator"
SU_TRANSFORMER = "su transformer"
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
    GENERATOR: GENERATOR,
    "gen": GENERATOR,
    "genset": GENERATOR,
    "gen set": GENERATOR,
    "dg": GENERATOR,
    "alternator": GENERATOR,
    MCC: MCC,
    "motor control centre": MCC,
    "motor control center": MCC,
    TRANSFORMER: TRANSFORMER,
    "trafo": TRANSFORMER,
    "tx": TRANSFORMER,
    SU_TRANSFORMER: SU_TRANSFORMER,
    "su tx": SU_TRANSFORMER,
    "su trafo": SU_TRANSFORMER,
    "step-up transformer": SU_TRANSFORMER,
    "step up transformer": SU_TRANSFORMER,
    "step-up": SU_TRANSFORMER,
    "step up": SU_TRANSFORMER,
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

    for it in items.values():
        if it.type == TRANSFORMER:
            up = [c for c in items.values()
                  if it.id in c.parents and c.type in (MV_BUSBAR, RMU)]
            dn = [c for c in items.values()
                  if it.id in c.parents and c.type == LV_BUSBAR]
            if up and dn:
                print(f"warning: '{it.id}' feeds both an MV and an LV board "
                      f"- drawn as a step-up", file=sys.stderr)
        if it.type == GENERATOR and not any(it.id in c.parents
                                            for c in items.values()) \
                and not any(items[p].type == SU_TRANSFORMER
                            for p in it.parents if p in items):
            print(f"warning: generator '{it.id}' feeds nothing",
                  file=sys.stderr)

    return info, items, order


# ---------------------------------------------------------------- layout

MARGIN = 90
FEEDER_SPACING = 95
BUS_GAP = 110          # horizontal gap between adjacent busbars
MIN_BUS_WIDTH = 170

SLOT_GAP = 30          # gap between slots on an MV switchboard
PUMP_SLOT = 115        # slot width of a pump/motor way
TX_LABEL_W = 150       # room a transformer's label block needs

Y_LABEL = 34           # MV incomer labels
Y_MV_TOP = 62          # top of the MV incomer stub
Y_RMU_TOP = 150
Y_RMU_BOT = 268
Y_MVBUS = 208          # MV switchboard busbar (same tier as the RMU)
PUMP_R = 20
TX_R = 19
LEGEND_H = 100         # symbol legend strip below the drawing
TIER_H = 150           # vertical pitch of a cascaded MV level
STEPUP_H = 170         # headroom above the MV rows for a step-up column
STEPUP_SHIFT = 0       # set per drawing by set_tiers()
Y_GEN = 96             # generation source symbol, step-up column
Y_SU_C1 = 196          # step-up transformer, upper circle
Y_SU_C2 = Y_SU_C1 + 27

# rows below the MV distribution - re-based by set_tiers() when MV boards
# or RMUs are fed from another MV board
Y_PUMP = 352           # pump/motor circle centre
Y_TX_C1 = 342          # centre of upper transformer circle
Y_TX_C2 = Y_TX_C1 + 27
Y_BUS = 486
Y_FEED_BRK = 516       # feeder breaker centre
Y_ARROW = 574          # arrow tip
Y_FEED_LBL = 592
DIAG_H = 780


def set_tiers(extra, top=0):
    """Push the transformer/LV rows down by `extra` px; `top` is extra
    headroom above the MV rows for a step-up column."""
    global Y_PUMP, Y_TX_C1, Y_TX_C2, Y_BUS, Y_FEED_BRK, Y_ARROW
    global Y_FEED_LBL, DIAG_H, STEPUP_SHIFT
    STEPUP_SHIFT = top
    extra = extra + top
    Y_PUMP = 352 + extra
    Y_TX_C1 = 342 + extra
    Y_TX_C2 = Y_TX_C1 + 27
    Y_BUS = 486 + extra
    Y_FEED_BRK = 516 + extra
    Y_ARROW = 574 + extra
    Y_FEED_LBL = 592 + extra
    DIAG_H = 780 + extra


def step_ups(items, order):
    """Transformers that feed an MV busbar or RMU (step-up), mapped to
    the source row that feeds them."""
    out = {}
    for oid in order:
        tx = items[oid]
        if tx.type != TRANSFORMER:
            continue
        if not children_of(items, order, tx.id, {MV_BUSBAR, RMU}):
            continue
        src = next((items[p] for p in tx.parents), None)
        out[tx.id] = src
    return out


def mv_depth(items, order):
    """MV distribution level of each busbar/RMU.

    A board or RMU fed from an MV busbar sits one level below it; RMUs
    linked to each other (a ring) stay on their parent's level.
    """
    depth = {i: 0 for i in order if items[i].type in (MV_BUSBAR, RMU)}
    for _ in range(len(depth) + 1):          # relax along the chains
        changed = False
        for oid in depth:
            for p in items[oid].parents:
                par = items[p]
                if par.id not in depth:
                    continue
                d = depth[par.id] + (1 if par.type == MV_BUSBAR else 0)
                if d > depth[oid]:
                    depth[oid] = d
                    changed = True
        if not changed:
            break
    return depth


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
    if item.type in (TRANSFORMER, GENERATOR, SU_TRANSFORMER):
        boards = children_of(items, order, item.id, {LV_BUSBAR})
        w = 130
        for bb in boards:
            kids = children_of(items, order, bb.id, {FEEDER, MCC})
            w = max(w, MIN_BUS_WIDTH, len(kids) * FEEDER_SPACING)
        return w
    return 130


def ring_group(items, order, head, depth):
    """RMUs linked to `head` through RMU-to-RMU cables on the same tier,
    in chain order, with the head centred when it has two neighbours."""
    def links(r):
        out = [items[p] for p in r.parents
               if items[p].type == RMU and depth.get(p) == depth.get(r.id)]
        out += [k for k in children_of(items, order, r.id, {RMU})
                if depth.get(k.id) == depth.get(r.id)]
        return out

    seen, chain = {head.id}, []
    for k in links(head):               # walk each branch away from the head
        branch, node = [], k
        while node is not None and node.id not in seen:
            seen.add(node.id)
            branch.append(node)
            node = next((n for n in links(node) if n.id not in seen), None)
        chain.append(branch)
    if len(chain) >= 2:                 # head sits between its two branches
        return list(reversed(chain[0])) + [head] + [n for b in chain[1:]
                                                    for n in b]
    return [head] + [n for b in chain for n in b]


def mv_children(items, order, node):
    """The ways of an MV board / RMU that occupy a slot beneath it."""
    types = ({TRANSFORMER, SU_TRANSFORMER, PUMP, MV_BUSBAR, RMU}
             if node.type == MV_BUSBAR else {TRANSFORMER, SU_TRANSFORMER, PUMP})
    return children_of(items, order, node.id, types)


def mv_own_width(items, order, node, depth=None):
    """Width one board / RMU needs for its own ways (no ring members)."""
    if node.type in (TRANSFORMER, SU_TRANSFORMER, PUMP):
        return slot_width(items, order, node)
    kids = mv_children(items, order, node)
    if not kids:
        return MIN_BUS_WIDTH
    need = (sum(mv_width(items, order, k, depth) for k in kids)
            + SLOT_GAP * (len(kids) - 1))
    return max(MIN_BUS_WIDTH, need)


def mv_width(items, order, node, depth=None):
    """Width the whole subtree under an MV board / RMU needs, including
    the ring an RMU belongs to."""
    if node.type == RMU and depth is not None:
        group = ring_group(items, order, node, depth)
        if len(group) > 1:
            return (sum(mv_own_width(items, order, m, depth) for m in group)
                    + SLOT_GAP * (len(group) - 1))
    return mv_own_width(items, order, node, depth)


def place_mv_node(items, order, node, left, width, depth=None):
    """Place an MV board / RMU across [left, left+width] and its ways.
    A ring head lays its whole ring out side by side in that band."""
    if node.type == RMU and depth is not None:
        group = ring_group(items, order, node, depth)
        if len(group) > 1:
            widths = [mv_own_width(items, order, m, depth) for m in group]
            need = sum(widths) + SLOT_GAP * (len(group) - 1)
            cursor = left + (width - need) / 2
            for m, w in zip(group, widths):
                place_own(items, order, m, cursor, w, depth)
                cursor += w + SLOT_GAP
            return
    place_own(items, order, node, left, width, depth)


def place_own(items, order, node, left, width, depth=None):
    """Place one board / RMU and the ways directly beneath it."""
    node.x = left + width / 2
    if node.type == MV_BUSBAR:
        node.x_left, node.x_right = left, left + width
    kids = mv_children(items, order, node)
    widths = [mv_width(items, order, k, depth) for k in kids]
    need = sum(widths) + SLOT_GAP * max(0, len(kids) - 1)
    cursor = left + (width - need) / 2
    for k, w in zip(kids, widths):
        if k.type in (MV_BUSBAR, RMU):
            place_mv_node(items, order, k, cursor, w, depth)
        else:
            k.x = cursor + w / 2
        cursor += w + SLOT_GAP


def supplies_of(items, order, board, sus):
    """Everything feeding a board: utility incomers and step-up columns."""
    return [items[i] for i in order
            if (items[i].type == MV_INCOMER or i in sus)
            and board.id in [k.id for k in children_of(items, order, i)]]


def place_step_ups(items, order):
    """A step-up transformer (and its source) sits above the MV busbar or
    RMU it feeds, beside any utility incomers on that board."""
    sus = step_ups(items, order)
    for tx_id, src in sus.items():
        tx = items[tx_id]
        fed = children_of(items, order, tx_id, {MV_BUSBAR, RMU})
        anchor = next((f for f in fed if f.x is not None), None)
        if anchor is None:
            continue
        peers = supplies_of(items, order, anchor, sus)
        n = max(1, len(peers))
        k = peers.index(tx) if tx in peers else 0
        tx.x = anchor.x + (k - (n - 1) / 2) * 90
        if src is not None:
            src.x = tx.x
            if src.type == LV_BUSBAR:
                src.x_left, src.x_right = tx.x - 85, tx.x + 85


def place_su_sources(items, order):
    """The generation source drawn under a reversed step-up follows it."""
    for oid in order:
        tx = items[oid]
        if tx.type != SU_TRANSFORMER or tx.x is None:
            continue
        for src in children_of(items, order,
                               tx.id, {GENERATOR, LV_BUSBAR, MV_INCOMER}):
            src.x = tx.x
            if src.type == LV_BUSBAR:
                src.x_left, src.x_right = tx.x - 85, tx.x + 85


def layout_mv_boards(items, order):
    """Layout when the site has MV switchboards (MV Busbar rows)."""
    mvbs = [items[i] for i in order if items[i].type == MV_BUSBAR]
    mvs = [items[i] for i in order if items[i].type == MV_INCOMER]

    # roots are the boards not fed from another MV board
    depth = mv_depth(items, order)
    roots = [b for b in mvbs
             if not any(items[p].type == MV_BUSBAR for p in b.parents)]
    x = MARGIN
    for mvb in roots:
        w = mv_width(items, order, mvb, depth)
        place_mv_node(items, order, mvb, x, w, depth)
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

    # step-up chains take an incomer position over the board they feed
    place_step_ups(items, order)
    place_su_sources(items, order)

    # incomers share the supply spread with any step-up columns
    sus = step_ups(items, order)
    for m in mvs:
        for mvb in mvbs:
            feeds = supplies_of(items, order, mvb, sus)
            if m in feeds:
                n = len(feeds)
                m.x = mvb.x + (feeds.index(m) - (n - 1) / 2) * 90
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
        # pump ways sit beside the transformer ways, left to right
        pumps_r = children_of(items, order, rmu.id, {PUMP})
        txs_r = children_of(items, order, rmu.id, {TRANSFORMER})
        placed = [k.x for k in txs_r if k.x is not None]
        if pumps_r and placed:
            # clear the last transformer's label block before the first pump
            cursor = max(placed) + TX_R + TX_LABEL_W + PUMP_SLOT / 2
            for k in pumps_r:
                k.x = cursor
                cursor += PUMP_SLOT + SLOT_GAP
        kids = [k for k in txs_r + pumps_r if k.x is not None]
        if kids:
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
    ("tx", "Transformer"), ("gen", "Generator"), ("pump", "Pump/motor"), ("bus", "Busbar"),
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
        elif kind == "gen":
            svg.circle(cx, yc, 9)
            svg.text(cx, yc + 3.5, "G", size=9, bold=True)
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
    depth = mv_depth(items, order)
    sus = step_ups(items, order)
    set_tiers(max(depth.values(), default=0) * TIER_H,
              STEPUP_H if sus else 0)

    def dy(it):                       # vertical offset of an MV level
        return depth.get(it.id, 0) * TIER_H

    def y_rmu(r):                     # (top, bottom, bus) of an RMU box
        o = dy(r) + STEPUP_SHIFT
        return Y_RMU_TOP + o, Y_RMU_BOT + o, (Y_RMU_TOP + Y_RMU_BOT) / 2 + o

    def y_bus(b):                     # busbar level of an MV board
        return Y_MVBUS + dy(b) + STEPUP_SHIFT
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
            tier = depth.get(a.id)
            between = any(r is not a and r is not b and a.x < r.x < b.x
                          and depth.get(r.id) == tier for r in rmus)
            if between:
                xa, xb = a.x - 28, b.x + 28
                ring_entries.setdefault(a.id, []).append((xa, b.id))
                ring_entries.setdefault(b.id, []).append((xb, a.id))
                ring_links.append((xa, xb, Y_RMU_TOP + dy(a)))
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
        ways_out = children_of(items, order, rmu.id, {TRANSFORMER, PUMP})
        xs = [w.x for w in ways_in + ways_out if w.x is not None] or [rmu.x]
        xs = xs + [e[0] for e in ring_entries.get(rmu.id, [])]
        left = min(xs) - pad_l[rmu.id]
        right = max(xs) + pad_r[rmu.id]
        bus_l, bus_r = min(xs) - 18, max(xs) + 18
        rmu_box[rmu.id] = (left, right, bus_l, bus_r)
        rt, rb, ymid = y_rmu(rmu)
        svg.rect(left, rt, right - left, rb - rt, sw=1.6, dash="7 5")
        # internal bus
        svg.line(bus_l, ymid, bus_r, ymid, w=3.4)
        # incoming ways from the top edge to the bus (default LBS,
        # overridden by the RMU row's Protection entry for that supply)
        def in_way(x, kind):
            if kind and kind != "lbs":
                svg.drop(x, rt, ymid, kind)
            else:
                svg.load_break_switch(x, rt + 12, ymid)
                svg.line(x, rt, x, rt + 12)
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
                svg.drop(t.x, ymid, rb, kind)
            else:
                svg.fuse_switch(t.x, ymid + 4, rb - 8)
                svg.line(t.x, ymid, t.x, ymid + 4)
                svg.line(t.x, rb - 8, t.x, rb)
            svg.dot(t.x, ymid)
        lbl = [rmu.id, rmu.desc,
               " ".join(v for v in (rmu.rating, rmu.voltage) if v)]
        # keep the label block clear of the next item on the same tier
        tier = depth.get(rmu.id)
        edges_r = [o.x_left if o.type == MV_BUSBAR else o.x - 60
                   for o in (items[q] for q in order)
                   if o.type in (RMU, MV_BUSBAR) and o is not rmu
                   and o.x is not None and depth.get(o.id) == tier
                   and o.x > rmu.x]
        edges_l = [o.x_right if o.type == MV_BUSBAR else o.x + 60
                   for o in (items[q] for q in order)
                   if o.type in (RMU, MV_BUSBAR) and o is not rmu
                   and o.x is not None and depth.get(o.id) == tier
                   and o.x < rmu.x]
        room_r = (min(edges_r) - right) if edges_r else 1e9
        room_l = (left - max(edges_l)) if edges_l else 1e9
        if room_r >= 130:
            lx, anc, ty = right + 10, "start", rt + 16
        elif room_l >= 130:
            lx, anc, ty = left - 10, "end", rt + 16
        else:                       # crowded tier: stack it above the box,
            # clear of the ring-closure lane at rt - 26
            lx, anc, ty = rmu.x, "middle", rt - 34 - 15 * (len(lbl) - 1)
        for i, t in enumerate(lbl):
            svg.text(lx, ty, t, anchor=anc, bold=(i == 0))
            ty += 15

    # --- RMU-to-RMU interconnecting cables -------------------------------
    for a, b in side_links:  # a is the left box
        if a.id not in rmu_box or b.id not in rmu_box:
            continue
        y_link = y_rmu(a)[2]
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
    for xa, xb, r_top in ring_links:  # loop over the boxes in between
        y_ring = r_top - 26
        svg.line(xa, r_top, xa, y_ring)
        svg.line(xa, y_ring, xb, y_ring)
        svg.line(xb, y_ring, xb, r_top)

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
                svg.line(m.x, Y_MV_TOP, m.x, y_rmu(k)[0])
            elif k.type == MV_BUSBAR:  # incoming device onto the board
                svg.drop(m.x, Y_MV_TOP, y_bus(k),
                         prot_for(k, m.id)[1] or "cb")
                svg.dot(m.x, y_bus(k))
            elif k.type == TRANSFORMER:  # direct feed, no RMU
                svg.line(m.x, Y_MV_TOP, m.x, Y_TX_C1 - TX_R)

    # --- MV switchboards -------------------------------------------------
    for mvb in mvbs:
        yb = y_bus(mvb)
        svg.line(mvb.x_left, yb, mvb.x_right, yb, w=5.5)
        raw, kind = prot_for(mvb)
        zone = raw if raw and kind is None else ""  # e.g. 87B differential
        lbl = " ".join(v for v in (mvb.id, mvb.desc, mvb.rating, mvb.voltage)
                       if v)
        if zone:
            lbl += " · " + zone
        svg.text(mvb.x_left, yb - 12, lbl, size=11.5, anchor="start",
                 bold=True)

    # --- feeds from an MV board down to a sub-board or an RMU -----------
    for oid in order:
        it = items[oid]
        if it.type not in (MV_BUSBAR, RMU) or it.x is None:
            continue
        for p in it.parents:
            par = items[p]
            if par.type != MV_BUSBAR:
                continue
            kind = prot_for(it, par.id)[1] or "cb"
            y_from = y_bus(par)
            y_to = y_bus(it) if it.type == MV_BUSBAR else y_rmu(it)[0]
            x_from = min(max(it.x, par.x_left), par.x_right)
            svg.dot(x_from, y_from)
            if abs(x_from - it.x) < 1:
                svg.drop(it.x, y_from, y_to, kind)
            else:                       # sub-board offset from the way
                y_mid = (y_from + y_to) / 2
                svg.drop(x_from, y_from, y_mid, kind)
                svg.line(x_from, y_mid, it.x, y_mid)
                svg.line(it.x, y_mid, it.x, y_to)

    # --- pumps / motor loads --------------------------------------------
    for p in pumps:
        if p.x is None:
            continue
        for par in (items[q] for q in p.parents):
            if par.type == MV_BUSBAR:
                svg.drop(p.x, y_bus(par), Y_PUMP - PUMP_R,
                         prot_for(p, par.id)[1] or "cb")
                svg.dot(p.x, y_bus(par))
            elif par.type == RMU:   # the way device is inside the enclosure
                svg.line(p.x, y_rmu(par)[1], p.x, Y_PUMP - PUMP_R)
        svg.circle(p.x, Y_PUMP, PUMP_R, sw=2.2)
        svg.text(p.x, Y_PUMP + 1, "M", size=13, bold=True)
        svg.text(p.x, Y_PUMP + 13, "3~", size=8.5)
        lbl = " · ".join(v for v in (p.id, p.desc, p.rating) if v)
        svg.text(p.x + 4, Y_PUMP + PUMP_R + 14, lbl, size=11,
                 anchor="start", rotate=90)

    # --- LV supply routes ------------------------------------------------
    # A board fed from several transformers gets one landing point per
    # supply, and every sideways run its own horizontal lane, so
    # cross-feeds between different MV boards never sit on one line.
    ytop_tx = Y_TX_C2 + TX_R
    routes, elbows = {}, []
    for tx in txs:
        if tx.x is None:
            continue
        for bb in children_of(items, order, tx.id, {LV_BUSBAR}):
            sup = [items[p] for p in bb.parents
                   if items[p].type == TRANSFORMER and items[p].x is not None]
            if len(sup) < 2:
                continue                      # single supply: handled below
            sup.sort(key=lambda t: t.x)
            i = [t.id for t in sup].index(tx.id)
            w = bb.x_right - bb.x_left
            x_land = bb.x_left + w * (i + 0.5) / len(sup)
            routes[(tx.id, bb.id)] = [x_land, None]
            if abs(x_land - tx.x) > 1:
                elbows.append((tx.id, bb.id))
    elbows.sort(key=lambda k: -abs(routes[k][0] - items[k[0]].x))
    if elbows:
        top = ytop_tx + 14
        step = (min(13.0, ((Y_BUS - 42) - top) / (len(elbows) - 1))
                if len(elbows) > 1 else 0)
        for i, k in enumerate(elbows):
            routes[k][1] = top + step * i

    # --- generation sources feeding an LV board directly -----------------
    for g in (items[i] for i in order if items[i].type == GENERATOR):
        su_src = [k.id for i in order if items[i].type == SU_TRANSFORMER
                  for k in children_of(items, order, i, {GENERATOR})]
        if g.x is None or g.id in [s.id for s in sus.values() if s] \
                or g.id in su_src:
            continue                  # step-up sources are drawn in-column
        svg.circle(g.x, Y_TX_C1 + 13, 20, sw=2.2)
        svg.text(g.x, Y_TX_C1 + 17, "G", size=13, bold=True)
        lbl = [g.id, g.desc, " ".join(v for v in (g.rating, g.voltage) if v)]
        ty = Y_TX_C1 - 6
        for t in [t for t in lbl if t]:
            svg.text(g.x + 30, ty, t, anchor="start")
            ty += 15
        for bb in children_of(items, order, g.id, {LV_BUSBAR}):
            svg.drop(g.x, Y_TX_C1 + 33, Y_BUS, prot_for(bb, g.id)[1] or "cb")
            svg.dot(g.x, Y_BUS)

    # --- step-up columns: source -> transformer -> MV busbar / RMU -------
    for tx_id, src in sus.items():
        tx = items[tx_id]
        if tx.x is None:
            continue
        # the source symbol at the top of the column
        y_src_bot = Y_GEN + 20
        if src is None:
            y_src_bot = Y_GEN
        elif src.type == GENERATOR:
            svg.circle(tx.x, Y_GEN, 20, sw=2.2)
            svg.text(tx.x, Y_GEN + 4, "G", size=13, bold=True)
            lbl = [src.id, src.desc,
                   " ".join(v for v in (src.rating, src.voltage) if v)]
            ty = Y_GEN - 26
            for t in [t for t in lbl if t]:
                svg.text(tx.x + 30, ty, t, size=11, anchor="start")
                ty += 14
        elif src.type == LV_BUSBAR:   # a generation board
            svg.line(tx.x - 60, Y_GEN, tx.x + 60, Y_GEN, w=5.5)
            svg.text(tx.x - 60, Y_GEN - 12,
                     " ".join(v for v in (src.id, src.desc, src.rating,
                                          src.voltage) if v),
                     size=11.5, anchor="start", bold=True)
            y_src_bot = Y_GEN
        else:                          # an incomer row used as the source
            svg.line(tx.x - 11, Y_GEN, tx.x + 11, Y_GEN, w=3)
            lbl = [src.id, src.desc, src.voltage]
            ty = Y_GEN - 40
            for i, t in enumerate([t for t in lbl if t]):
                svg.text(tx.x, ty, t, size=11.5, bold=(i == 0))
                ty += 14
            y_src_bot = Y_GEN
        svg.line(tx.x, y_src_bot, tx.x, Y_SU_C1 - TX_R)
        svg.circle(tx.x, Y_SU_C1, TX_R, sw=2.2)
        svg.circle(tx.x, Y_SU_C2, TX_R, sw=2.2)
        lbl = [tx.id, tx.desc,
               " ".join(v for v in (tx.rating, tx.voltage) if v)]
        ty = Y_SU_C1 - 6
        for t in [t for t in lbl if t]:
            svg.text(tx.x + TX_R + 10, ty, t, anchor="start")
            ty += 15
        for fed in children_of(items, order, tx_id, {MV_BUSBAR, RMU}):
            y_to = y_bus(fed) if fed.type == MV_BUSBAR else y_rmu(fed)[0]
            svg.drop(tx.x, Y_SU_C2 + TX_R, y_to,
                     prot_for(fed, tx_id)[1] or "cb")
            if fed.type == MV_BUSBAR:
                svg.dot(tx.x, y_to)

    # --- reversed step-ups (SU Transformer): board on top, source below --
    su_below = [items[i] for i in order if items[i].type == SU_TRANSFORMER]
    for tx in su_below:
        if tx.x is None:
            continue
        fed = [items[p] for p in tx.parents
               if items[p].type in (MV_BUSBAR, RMU)]
        for f in fed:
            y_from = y_bus(f) if f.type == MV_BUSBAR else y_rmu(f)[1]
            svg.drop(tx.x, y_from, Y_TX_C1 - TX_R,
                     prot_for(tx, f.id)[1] or "cb")
            if f.type == MV_BUSBAR:
                svg.dot(tx.x, y_from)
        svg.circle(tx.x, Y_TX_C1, TX_R, sw=2.2)
        svg.circle(tx.x, Y_TX_C2, TX_R, sw=2.2)
        lbl = [tx.id, tx.desc,
               " ".join(v for v in (tx.rating, tx.voltage) if v)]
        ty = Y_TX_C1 - 6
        for t in [t for t in lbl if t]:
            svg.text(tx.x + TX_R + 10, ty, t, anchor="start")
            ty += 15
        # the generation source hangs below the transformer
        kids = children_of(items, order, tx.id,
                           {GENERATOR, LV_BUSBAR, MV_INCOMER})
        src = kids[0] if kids else None
        y_src = Y_BUS
        if src is None:
            continue
        if src.type == GENERATOR:
            svg.line(tx.x, Y_TX_C2 + TX_R, tx.x, y_src - 20)
            svg.circle(tx.x, y_src, 20, sw=2.2)
            svg.text(tx.x, y_src + 4, "G", size=13, bold=True)
            lbl = [src.id, src.desc,
                   " ".join(v for v in (src.rating, src.voltage) if v)]
            ty = y_src - 6
            for t in [t for t in lbl if t]:
                svg.text(tx.x + 30, ty, t, size=11, anchor="start")
                ty += 14
        elif src.type == LV_BUSBAR:
            svg.line(tx.x, Y_TX_C2 + TX_R, tx.x, y_src)
            svg.dot(tx.x, y_src)     # the board itself is drawn with the
        else:                        # other LV busbars, on its own row
            svg.line(tx.x, Y_TX_C2 + TX_R, tx.x, y_src)
            svg.line(tx.x - 11, y_src, tx.x + 11, y_src, w=3)
            lbl = [src.id, src.desc, src.voltage]
            ty = y_src + 18
            for i, t in enumerate([t for t in lbl if t]):
                svg.text(tx.x, ty, t, size=11.5, bold=(i == 0))
                ty += 14

    # --- transformers ----------------------------------------------------
    for tx in txs:
        if tx.x is None or tx.id in sus:
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
                svg.line(tx.x, y_rmu(par)[1], tx.x, Y_TX_C1 - TX_R)
            elif par.type == MV_BUSBAR:  # feeder device off the board
                svg.drop(tx.x, y_bus(par), Y_TX_C1 - TX_R,
                         prot_for(tx, par.id)[1] or "cb")
                svg.dot(tx.x, y_bus(par))
        ytop = ytop_tx
        dual = [bb for bb in fed if (tx.id, bb.id) in routes]
        rest = [bb for bb in fed if (tx.id, bb.id) not in routes]
        for bb in dual:
            # a board with several incomers: each supply drops at its own
            # landing point, sideways runs on their own lanes
            kind = prot_for(bb, tx.id)[1] or "cb"
            x_land, y_lane = routes[(tx.id, bb.id)]
            if y_lane is None:
                svg.drop(tx.x, ytop, Y_BUS, kind)
            else:
                svg.line(tx.x, ytop, tx.x, y_lane)
                svg.line(tx.x, y_lane, x_land, y_lane)
                svg.drop(x_land, y_lane, Y_BUS, kind)
            svg.dot(x_land, Y_BUS)
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
        y = y_bus(ends[0]) if ends[0].type == MV_BUSBAR else Y_BUS
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
