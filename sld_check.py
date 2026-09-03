#!/usr/bin/env python3
"""Topology checker for SLD-Sketch drawings.

Reads a survey workbook, renders it with the normal engine, then reads the
SVG *back* as raw geometry - lines, circles, rectangles, text - with no help
from the drawing code, and compares what is drawn with what the table says:

  * every item has exactly one symbol;
  * every `Feeds From` edge is a continuous conductor between the two
    symbols that passes through no other item;
  * no two different connections share a piece of conductor, and no drawn
    conductor joins two items the table does not connect;
  * crossings, label collisions and anything outside the sheet are counted.

Each failure is tagged with the modification that would address it, so a
change to the engine can be measured by re-running the same workbooks.

    python sld_check.py site.xlsx [more.xlsx ...] [--json] [--quiet]

Exit status is 1 when any item is missing or any edge is disconnected.
"""
import argparse
import json
import math
import os
import re
import sys
from collections import defaultdict, deque

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import sld_sketch as S  # noqa: E402

TOL = 1.0            # snapping tolerance, px
GAP_MIN, GAP_MAX = 3.0, 48.0   # a protection device interrupts a conductor
LONG = 20.0          # a conductor worth counting for overlaps / crossings

# --------------------------------------------------------------- SVG parse

R_LINE = re.compile(r'<line x1="([\d.-]+)" y1="([\d.-]+)" x2="([\d.-]+)" '
                    r'y2="([\d.-]+)" stroke="#\w+" stroke-width="([\d.]+)"'
                    r'( stroke-dasharray="[^"]*")?')
R_CIRC = re.compile(r'<circle cx="([\d.-]+)" cy="([\d.-]+)" r="([\d.]+)" '
                    r'fill="(\w+|#\w+)"')
R_RECT = re.compile(r'<rect x="([\d.-]+)" y="([\d.-]+)" width="([\d.]+)" '
                    r'height="([\d.]+)" fill="(\w+)" stroke="#\w+" '
                    r'stroke-width="([\d.]+)"( stroke-dasharray="[^"]*")?')
R_POLY = re.compile(r'<polygon points="([^"]+)"')
R_TEXT = re.compile(r'<text x="([\d.-]+)" y="([\d.-]+)"'
                    r'(?: transform="translate\(([\d.-]+),([\d.-]+)\) '
                    r'rotate\(([\d.-]+)\)")? font-family="[^"]*" '
                    r'font-size="([\d.]+)" fill="#\w+" text-anchor="(\w+)"'
                    r'( font-weight="bold")?>(.*?)</text>')
R_PATH = re.compile(r'<path d="M ([\d.-]+),([\d.-]+)')
R_SIZE = re.compile(r'<svg [^>]*width="(\d+)" height="(\d+)"')


def unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&amp;", "&"))


def parse_svg(svg):
    w, h = (int(v) for v in R_SIZE.search(svg).groups())
    lines, circles, rects, polys, texts, marks = [], [], [], [], [], []
    for raw in svg.splitlines():
        m = R_LINE.match(raw)
        if m:
            x1, y1, x2, y2, sw = (float(v) for v in m.groups()[:5])
            lines.append(dict(x1=x1, y1=y1, x2=x2, y2=y2, w=sw,
                              dash=bool(m.group(6))))
            continue
        m = R_CIRC.match(raw)
        if m:
            circles.append(dict(cx=float(m.group(1)), cy=float(m.group(2)),
                                r=float(m.group(3)),
                                filled=m.group(4) != "none"))
            continue
        m = R_RECT.match(raw)
        if m:
            rects.append(dict(x=float(m.group(1)), y=float(m.group(2)),
                              w=float(m.group(3)), h=float(m.group(4)),
                              filled=m.group(5) != "none",
                              dash=bool(m.group(7))))
            continue
        m = R_POLY.match(raw)
        if m:
            pts = [tuple(float(c) for c in p.split(","))
                   for p in m.group(1).split()]
            polys.append(pts)
            continue
        m = R_TEXT.match(raw)
        if m:
            x, y, tx, ty, rot, size, anchor, bold, content = m.groups()
            if tx is not None:
                x, y = tx, ty
            texts.append(dict(x=float(x), y=float(y),
                              rot=float(rot) if rot else 0.0,
                              size=float(size), anchor=anchor,
                              bold=bool(bold), s=unescape(content)))
            continue
        m = R_PATH.match(raw)
        if m:
            marks.append((float(m.group(1)), float(m.group(2))))
    return dict(w=w, h=h, lines=lines, circles=circles, rects=rects,
                polys=polys, texts=texts, paths=marks)


# --------------------------------------------------------------- geometry

def seg_len(s):
    return math.hypot(s["x2"] - s["x1"], s["y2"] - s["y1"])


def pt_seg_dist(px, py, s):
    """Distance from a point to a segment, and the parameter t along it."""
    dx, dy = s["x2"] - s["x1"], s["y2"] - s["y1"]
    L2 = dx * dx + dy * dy
    if L2 == 0:
        return math.hypot(px - s["x1"], py - s["y1"]), 0.0
    t = ((px - s["x1"]) * dx + (py - s["y1"]) * dy) / L2
    t = max(0.0, min(1.0, t))
    qx, qy = s["x1"] + t * dx, s["y1"] + t * dy
    return math.hypot(px - qx, py - qy), t


def segs_cross(a, b):
    """True when the two segments intersect strictly inside both."""
    def orient(ax, ay, bx, by, cx, cy):
        return (bx - ax) * (cy - ay) - (by - ay) * (cx - ax)
    o1 = orient(a["x1"], a["y1"], a["x2"], a["y2"], b["x1"], b["y1"])
    o2 = orient(a["x1"], a["y1"], a["x2"], a["y2"], b["x2"], b["y2"])
    o3 = orient(b["x1"], b["y1"], b["x2"], b["y2"], a["x1"], a["y1"])
    o4 = orient(b["x1"], b["y1"], b["x2"], b["y2"], a["x2"], a["y2"])
    eps = 1e-6
    return (o1 * o2 < -eps) and (o3 * o4 < -eps)


def collinear_overlap(a, b):
    """Length of overlap between two collinear axis-aligned segments."""
    va = abs(a["x1"] - a["x2"]) < 0.6
    vb = abs(b["x1"] - b["x2"]) < 0.6
    ha = abs(a["y1"] - a["y2"]) < 0.6
    hb = abs(b["y1"] - b["y2"]) < 0.6
    if va and vb and abs(a["x1"] - b["x1"]) < 0.6:
        lo = max(min(a["y1"], a["y2"]), min(b["y1"], b["y2"]))
        hi = min(max(a["y1"], a["y2"]), max(b["y1"], b["y2"]))
        return hi - lo
    if ha and hb and abs(a["y1"] - b["y1"]) < 0.6:
        lo = max(min(a["x1"], a["x2"]), min(b["x1"], b["x2"]))
        hi = min(max(a["x1"], a["x2"]), max(b["x1"], b["x2"]))
        return hi - lo
    return 0.0


# --------------------------------------------------------------- the drawing

class Drawing:
    """Geometry of one rendered sheet, with items bound to symbols and a
    conductor graph rebuilt from the strokes."""

    def __init__(self, svg, items, order):
        g = parse_svg(svg)
        self.w, self.h = g["w"], g["h"]
        self.rects = g["rects"]
        self.items, self.order = items, order
        self.off_sheet = 0
        self._drop_furniture(g)
        self._classify(g)
        self._bind_labels()
        self._bind_symbols()
        self._build_graph()

    # -- furniture: legend strip, title block, drawing title ---------------
    def _drop_furniture(self, g):
        top_of_legend = self.h - S.LEGEND_H
        tb = next((r for r in g["rects"]
                   if abs(r["w"] - 288) < 1 and abs(r["h"] - 96) < 1), None)

        def in_tb(x, y):
            return tb and tb["x"] - 2 <= x <= tb["x"] + tb["w"] + 2 \
                and tb["y"] - 2 <= y <= tb["y"] + tb["h"] + 2

        def keep(x, y):
            return y < top_of_legend - 2 and not in_tb(x, y)

        for k in ("lines", "circles", "rects", "polys", "texts", "paths"):
            out = []
            for e in g[k]:
                if k == "lines":
                    x, y = e["x1"], e["y1"]
                elif k == "circles":
                    x, y = e["cx"], e["cy"]
                elif k == "rects":
                    x, y = e["x"], e["y"]
                    if abs(e["w"] - 288) < 1 and abs(e["h"] - 96) < 1:
                        continue
                elif k == "polys":
                    x, y = e[0]
                elif k == "texts":
                    x, y = e["x"], e["y"]
                    if e["size"] == 16 and "Single Line Diagram" in e["s"]:
                        continue
                else:
                    x, y = e
                # anything reaching outside the sheet is a defect in its
                # own right, and a label that starts inside but runs past
                # the edge counts too: measure the extent, not the anchor
                if k == "lines":
                    x1, y1 = max(e["x1"], e["x2"]), max(e["y1"], e["y2"])
                    x0, y0 = min(e["x1"], e["x2"]), min(e["y1"], e["y2"])
                elif k == "circles":
                    x0, x1 = e["cx"] - e["r"], e["cx"] + e["r"]
                    y0, y1 = e["cy"] - e["r"], e["cy"] + e["r"]
                elif k == "rects":
                    x0, y0 = e["x"], e["y"]
                    x1, y1 = e["x"] + e["w"], e["y"] + e["h"]
                elif k == "texts":
                    tw = 0.6 * e["size"] * len(e["s"])
                    if e["rot"]:
                        x0, x1 = x - e["size"], x + e["size"] * 0.3
                        y0, y1 = y, y + tw
                    else:
                        x0 = {"start": x, "middle": x - tw / 2,
                              "end": x - tw}[e["anchor"]]
                        x1, y0, y1 = x0 + tw, y - e["size"], y
                else:
                    x0, y0, x1, y1 = x, y, x, y
                if x0 < -1 or x1 > self.w + 1 or y0 < -1 or y1 > self.h + 1:
                    self.off_sheet += 1
                if keep(x, y):
                    out.append(e)
            g[k] = out
        self.g = g

    # -- symbols by geometry ----------------------------------------------
    def _classify(self, g):
        self.busbars, self.ticks, self.rmubars = [], [], []
        self.conductors, self.marks = [], []
        plates = []
        for ln in g["lines"]:
            L = seg_len(ln)
            axis = (abs(ln["x1"] - ln["x2"]) < 0.6
                    or abs(ln["y1"] - ln["y2"]) < 0.6)
            if abs(ln["w"] - 5.5) < 0.1:
                self.busbars.append(ln)
            elif abs(ln["w"] - 2.5) < 0.1 and abs(L - 18) < 1:
                plates.append(ln)          # capacitor plates
            elif abs(ln["w"] - 3.4) < 0.1:
                self.rmubars.append(ln)
            elif abs(ln["w"] - 3) < 0.1 and abs(L - 22) < 1:
                self.ticks.append(ln)
            elif not axis:                 # x arms and switch blades
                self.marks.append(ln)
            else:                          # every axis-aligned stroke is wire
                self.conductors.append(ln)
        for c in g["circles"]:             # switch hinges
            if not c["filled"] and c["r"] <= 3:
                self.marks.append(dict(x1=c["cx"] - 2, y1=c["cy"],
                                       x2=c["cx"] + 2, y2=c["cy"], w=1))
        for r in g["rects"]:
            if r["w"] <= 10 and r["h"] <= 24:      # fuse bodies
                self.marks.append(dict(x1=r["x"], y1=r["y"] + r["h"] / 2,
                                       x2=r["x"] + r["w"],
                                       y2=r["y"] + r["h"] / 2, w=2))
        for (px, py) in g["paths"]:               # contactor arcs
            self.marks.append(dict(x1=px, y1=py, x2=px + 8, y2=py, w=2))

        # transformers: two hollow circles r=19, same x, 27 apart
        hollow = [c for c in g["circles"] if not c["filled"]]
        used = set()
        self.transformers = []
        by_x = defaultdict(list)
        for i, c in enumerate(hollow):
            if abs(c["r"] - S.TX_R) < 0.1:
                by_x[round(c["cx"], 1)].append(i)
        for _, idxs in by_x.items():
            idxs.sort(key=lambda i: hollow[i]["cy"])
            for a, b in zip(idxs, idxs[1:]):
                if a in used or b in used:
                    continue
                if abs(hollow[b]["cy"] - hollow[a]["cy"] - 27) < 0.6:
                    ca, cb = hollow[a], hollow[b]
                    self.transformers.append(dict(
                        x=ca["cx"], top=ca["cy"] - S.TX_R,
                        bot=cb["cy"] + S.TX_R, mid=(ca["cy"] + cb["cy"]) / 2))
                    used.update((a, b))
        # motors / generators: r=20 or 14 with an M or G inside
        self.machines = []
        for i, c in enumerate(hollow):
            if i in used or c["r"] not in (20.0, 14.0):
                continue
            tag = next((t["s"] for t in g["texts"]
                        if t["s"] in ("M", "G")
                        and math.hypot(t["x"] - c["cx"], t["y"] - c["cy"])
                        < c["r"]), None)
            if tag:
                self.machines.append(dict(x=c["cx"], y=c["cy"], r=c["r"],
                                          tag=tag))
        # capacitor banks: two plates 18 wide, 6 apart, one above the other
        self.caps = []
        plates.sort(key=lambda p: (round((p["x1"] + p["x2"]) / 2), p["y1"]))
        for a, b in zip(plates, plates[1:]):
            if abs((a["x1"] + a["x2"]) - (b["x1"] + b["x2"])) < 1 \
                    and abs(b["y1"] - a["y1"] - 6) < 0.6:
                self.caps.append(dict(x=(a["x1"] + a["x2"]) / 2, top=a["y1"]))
        # earthing resistors 12 x 30, surge arresters 14 x 28
        self.ners = [r for r in g["rects"]
                     if abs(r["w"] - 12) < 0.1 and abs(r["h"] - 30) < 0.1]
        self.arresters = [r for r in g["rects"]
                          if abs(r["w"] - 14) < 0.1 and abs(r["h"] - 28) < 0.1]
        # MCC boxes
        self.mccs = [r for r in g["rects"]
                     if abs(r["w"] - 28) < 0.1 and abs(r["h"] - 26) < 0.1]

        def holds_mcc(r):        # a dashed outline round an MCC box + bus
            return any(r["x"] <= m["x"] and m["x"] + m["w"] <= r["x"] + r["w"]
                       and r["y"] <= m["y"] and m["y"] + m["h"] <= r["y"] + r["h"]
                       for m in self.mccs)
        dashed = [r for r in g["rects"] if r["dash"] and r["w"] > 40]
        self.mcc_encl = [r for r in dashed if holds_mcc(r)]
        # RMU enclosures
        self.rmus = [r for r in dashed if not holds_mcc(r)]
        # feeder arrows
        self.arrows = [dict(x=p[2][0], top=p[0][1]) for p in g["polys"]
                       if len(p) == 3]
        self.texts = g["texts"]

    # -- labels: which text belongs to which item -------------------------
    def _bind_labels(self):
        self.labels = defaultdict(list)          # id -> [(x, y, text)]
        ids = set(self.order)

        def inside_symbol(t):              # the M / G / MCC glyphs
            return any(math.hypot(t["x"] - m["x"], t["y"] - m["y"]) < m["r"]
                       for m in self.machines) or \
                any(r["x"] <= t["x"] <= r["x"] + r["w"]
                    and r["y"] <= t["y"] <= r["y"] + r["h"] for r in self.mccs)

        by_len = sorted(ids, key=len, reverse=True)
        for t in self.texts:
            if inside_symbol(t):
                continue
            txt = t["s"].strip()
            # the label starts with the ID (IDs may contain spaces:
            # longest match first), followed by the end or a separator
            first = next((i for i in by_len if txt == i
                          or txt.startswith(i + " ")), None)
            if first is not None:
                self.labels[first].append((t["x"], t["y"], t))

    # -- bind each symbol to the nearest compatible label ------------------
    def _bind_symbols(self):
        it = self.items
        self.sym = {}          # id -> list of symbol dicts
        self.duplicates = []
        cands = []             # (symbol, allowed types, anchor x, y)
        for tx in self.transformers:
            cands.append((dict(kind="tx", **tx), {S.TRANSFORMER},
                          tx["x"], tx["mid"]))
        for m in self.machines:
            typ = {S.PUMP} if m["tag"] == "M" else {S.GENERATOR}
            cands.append((dict(kind="machine", **m), typ, m["x"], m["y"]))
        for b in self.busbars:
            cands.append((dict(kind="bar", seg=b),
                          {S.LV_BUSBAR, S.MV_BUSBAR, S.MCC},
                          min(b["x1"], b["x2"]), b["y1"]))
        for r in self.rmus:
            bar = next((s for s in self.rmubars
                        if r["x"] <= s["x1"] <= r["x"] + r["w"]
                        and r["y"] <= s["y1"] <= r["y"] + r["h"]), None)
            cands.append((dict(kind="rmu", rect=r, seg=bar), {S.RMU},
                          [r["x"], r["x"] + r["w"] / 2, r["x"] + r["w"]],
                          r["y"]))
        for c in self.caps:
            cands.append((dict(kind="terminal", x=c["x"], top=c["top"]),
                          {S.CAPACITOR}, c["x"], c["top"]))
        for r in self.ners:
            cands.append((dict(kind="terminal", x=r["x"] + 6, top=r["y"]),
                          {S.EARTHING}, r["x"] + 6, r["y"]))
        for r in self.arresters:
            cands.append((dict(kind="terminal", x=r["x"] + 7, top=r["y"]),
                          {S.ARRESTER}, r["x"] + 7, r["y"]))
        for t in self.ticks:
            cands.append((dict(kind="tick", seg=t),
                          {S.MV_INCOMER, S.MV_BUSBAR, S.LV_BUSBAR},
                          (t["x1"] + t["x2"]) / 2, t["y1"]))
        for a in self.arrows:
            cands.append((dict(kind="arrow", **a), {S.FEEDER},
                          a["x"], a["top"] + 11))
        for r in self.mccs:
            anchors = [(r["x"] + 14, r["y"] + 26)]
            for e in self.mcc_encl:      # its bus label sits at the
                if e["x"] <= r["x"] and r["x"] + r["w"] <= e["x"] + e["w"] \
                        and e["y"] <= r["y"] <= e["y"] + e["h"]:
                    anchors.append((e["x"] + 10, e["y"] + e["h"] - 22))
            cands.append((dict(kind="mcc", rect=r), {S.MCC}, anchors, None))
        for sym, types, ax, ay in cands:
            best, bd = None, 1e9
            for oid in self.order:
                if it[oid].type not in types:
                    continue
                for lx, ly, t in self.labels.get(oid, []):
                    pts = ax if isinstance(ax, list) else [ax]
                    d = min(math.hypot(lx - q[0], ly - q[1])
                            if isinstance(q, tuple)
                            else math.hypot(lx - q, ly - ay) for q in pts)
                    if d < bd:
                        best, bd = oid, d
            if best is not None and bd < 260:
                self.sym.setdefault(best, []).append(sym)
        for oid, syms in self.sym.items():
            if len(syms) > 1 and it[oid].type not in (S.LV_BUSBAR,):
                kinds = sorted(sy["kind"] for sy in syms)
                if it[oid].type == S.MCC and kinds == ["bar", "mcc"]:
                    continue               # an MCC with its own bus below
                self.duplicates.append(oid)

    # -- conductor graph ----------------------------------------------------
    def _node(self, x, y):
        key = (round(x), round(y))
        for k in (key, (key[0] - 1, key[1]), (key[0] + 1, key[1]),
                  (key[0], key[1] - 1), (key[0], key[1] + 1)):
            n = self.nodes.get(k)
            if n is not None and math.hypot(n[0] - x, n[1] - y) <= TOL:
                return k
        self.nodes[key] = (x, y)
        return key

    def _link(self, a, b):
        if a != b:
            self.adj[a].add(b)
            self.adj[b].add(a)

    def _build_graph(self):
        self.nodes, self.adj = {}, defaultdict(set)
        wires = ([dict(s, kind="cond") for s in self.conductors]
                 + [dict(s, kind="bus") for s in self.busbars]
                 + [dict(s, kind="rmubar") for s in self.rmubars]
                 + [dict(s, kind="tick") for s in self.ticks])
        self.wires = wires
        for s in wires:
            s["n1"] = self._node(s["x1"], s["y1"])
            s["n2"] = self._node(s["x2"], s["y2"])
            self._link(s["n1"], s["n2"])
        # a node sitting on another wire's interior is a landing
        for key, (x, y) in list(self.nodes.items()):
            for s in wires:
                if key in (s["n1"], s["n2"]):
                    continue
                d, t = pt_seg_dist(x, y, s)
                if d <= TOL and 0.0 < t < 1.0:
                    self._link(key, s["n1"])
                    self._link(key, s["n2"])
        # a protection device interrupts a conductor: bridge the gap when
        # two conductor ends face each other on one axis with marks between
        ends = [(s[k], self.nodes[s[k]], s) for s in wires
                if s["kind"] == "cond" for k in ("n1", "n2")]

        def runs(w, vertical):
            return (abs(w["x1"] - w["x2"]) < 0.6 if vertical
                    else abs(w["y1"] - w["y2"]) < 0.6)

        def outside(w, vertical, lo, hi):
            """The wire lies beyond the gap, not inside it."""
            a, b = ((w["y1"], w["y2"]) if vertical else (w["x1"], w["x2"]))
            return max(a, b) <= lo + 0.5 or min(a, b) >= hi - 0.5

        def blocked(vertical, c, lo, hi):
            """Another wire lies on the axis strictly inside the gap."""
            for w in wires:
                if vertical:
                    if abs(w["x1"] - w["x2"]) < 0.6 and abs(w["x1"] - c) < 0.6:
                        if min(w["y1"], w["y2"]) < hi - 0.5 \
                                and max(w["y1"], w["y2"]) > lo + 0.5:
                            return True
                    elif abs(w["y1"] - w["y2"]) < 0.6 \
                            and lo + 0.5 < w["y1"] < hi - 0.5 \
                            and min(w["x1"], w["x2"]) - 0.5 <= c \
                            <= max(w["x1"], w["x2"]) + 0.5:
                        return True
                else:
                    if abs(w["y1"] - w["y2"]) < 0.6 and abs(w["y1"] - c) < 0.6:
                        if min(w["x1"], w["x2"]) < hi - 0.5 \
                                and max(w["x1"], w["x2"]) > lo + 0.5:
                            return True
                    elif abs(w["x1"] - w["x2"]) < 0.6 \
                            and lo + 0.5 < w["x1"] < hi - 0.5 \
                            and min(w["y1"], w["y2"]) - 0.5 <= c \
                            <= max(w["y1"], w["y2"]) + 0.5:
                        return True
            return False

        for i, (ka, (xa, ya), wa) in enumerate(ends):
            for kb, (xb, yb), wb in ends[i + 1:]:
                if ka == kb or wa is wb:
                    continue
                if abs(xa - xb) < 0.6 and GAP_MIN <= abs(ya - yb) <= GAP_MAX:
                    lo, hi = min(ya, yb), max(ya, yb)
                    if runs(wa, True) and runs(wb, True) \
                            and outside(wa, True, lo, hi) \
                            and outside(wb, True, lo, hi) \
                            and any(lo - 1 <= (m["y1"] + m["y2"]) / 2 <= hi + 1
                                    and abs((m["x1"] + m["x2"]) / 2 - xa) <= 12
                                    for m in self.marks) \
                            and not blocked(True, xa, lo, hi):
                        self._link(ka, kb)
                elif abs(ya - yb) < 0.6 and GAP_MIN <= abs(xa - xb) <= GAP_MAX:
                    lo, hi = min(xa, xb), max(xa, xb)
                    if runs(wa, False) and runs(wb, False) \
                            and outside(wa, False, lo, hi) \
                            and outside(wb, False, lo, hi) \
                            and any(lo - 1 <= (m["x1"] + m["x2"]) / 2 <= hi + 1
                                    and abs((m["y1"] + m["y2"]) / 2 - ya) <= 12
                                    for m in self.marks) \
                            and not blocked(False, ya, lo, hi):
                        self._link(ka, kb)
        # item -> its graph nodes
        self.item_nodes = defaultdict(set)
        for oid, syms in self.sym.items():
            for sym in syms:
                for n in self._terminals(sym):
                    self.item_nodes[oid].add(n)
        self.owner = {}
        for oid, ns in self.item_nodes.items():
            for n in ns:
                self.owner[n] = oid

    def _terminals(self, sym):
        k = sym["kind"]
        pts = []
        if k == "tx":
            pts = [(sym["x"], sym["top"]), (sym["x"], sym["bot"])]
        elif k == "machine":
            pts = [(sym["x"], sym["y"] - sym["r"]),
                   (sym["x"], sym["y"] + sym["r"])]
        elif k in ("bar", "tick", "rmu"):
            s = sym.get("seg")
            if s:
                pts = [(s["x1"], s["y1"]), (s["x2"], s["y2"])]
                for (nx, ny) in self.nodes.values():
                    d, t = pt_seg_dist(nx, ny, s)
                    if d <= TOL and 0.0 < t < 1.0:
                        pts.append((nx, ny))
        elif k in ("arrow", "terminal"):
            pts = [(sym["x"], sym["top"])]
        elif k == "mcc":
            r = sym["rect"]
            pts = [(r["x"] + 14, r["y"]), (r["x"] + 14, r["y"] + 26)]
        out = set()
        for x, y in pts:
            key = (round(x), round(y))
            hit = None
            for kk, (nx, ny) in self.nodes.items():
                if abs(kk[0] - key[0]) <= 2 and abs(kk[1] - key[1]) <= 2 \
                        and math.hypot(nx - x, ny - y) <= 1.5:
                    hit = kk
                    break
            if hit is not None:
                out.add(hit)
        return out

    # -- path search ------------------------------------------------------
    def path(self, a, b, avoid_others=True):
        """Nodes of a path between items a and b, or None. With
        avoid_others the path may not touch any third item's nodes; without
        it the set of third items touched is returned as the second value."""
        src, dst = self.item_nodes.get(a, set()), self.item_nodes.get(b, set())
        if not src or not dst:
            return None, set()
        seen, q, prev = set(src), deque(src), {}
        while q:
            n = q.popleft()
            if n in dst:
                via = set()
                cur = n
                while cur in prev:
                    o = self.owner.get(cur)
                    if o and o not in (a, b):
                        via.add(o)
                    cur = prev[cur]
                return n, via
            for m in self.adj[n]:
                if m in seen:
                    continue
                o = self.owner.get(m)
                if o and o not in (a, b) and avoid_others and m not in dst:
                    continue
                seen.add(m)
                prev[m] = n
                q.append(m)
        return None, set()


# --------------------------------------------------------------- checks

TAGS = {
    "multi-voltage": "tiers by voltage level",
    "lv-subboard": "LV cascades (sub-boards below their supply)",
    "source": "sources as first-class supplies / changeover",
    "lane-overlap": "one lane allocator for every sideways run",
    "mv-feeder": "MV outgoing ways and terminal items",
    "ring-group": "ring groups headed by an RMU",
    "rmu-entry": "board-fed RMU: draw the incoming way through to its bar",
    "no-load-ok": "terminal item types (NER, capacitor bank, arrester)",
    "tx-bypass": "a transformer with supply and load on one terminal",
    "other": "unclassified",
}
NO_LOAD_WORDS = ("ner", "earthing", "neutral", "capacitor", "arrester",
                 "surge", "pfc", "kvar")


def mv_chain_set(items):
    """Boards, transformers and incomers on an MV -> transformer -> MV
    chain: the third-voltage-level structure the row layout cannot hold."""
    out = set()
    for tx in items.values():
        if tx.type != S.TRANSFORMER:
            continue
        ups = [items[p] for p in tx.parents
               if p in items and items[p].type in (S.MV_BUSBAR, S.RMU)]
        downs = [c for c in items.values()
                 if tx.id in c.parents and c.type in (S.MV_BUSBAR, S.RMU)]
        if ups and downs:
            out.add(tx.id)
            out.update(u.id for u in ups)
            out.update(d.id for d in downs)
            for d in downs:                 # incomers on the upper board
                for inc in items.values():
                    if inc.type == S.MV_INCOMER and d.id in [
                            k for k in items if inc.id in items[k].parents]:
                        out.add(inc.id)
            for u in ups:
                for inc in items.values():
                    if inc.type == S.MV_INCOMER and inc.id in u.parents:
                        out.add(inc.id)
    return out


def classify_edge(items, parent, child, via, mvset=frozenset(),
                  broken=frozenset()):
    p, c = items[parent], items[child]
    if via and c.type == S.RMU and any(items[v].type == S.RMU for v in via):
        return "ring-group"
    if p.id in mvset or c.id in mvset:
        return "multi-voltage"       # a third level breaks whatever it touches
    if p.type == S.RMU and c.type == S.RMU:
        return "ring-group"
    if (p.type == S.RMU and p.id in broken) or \
            (c.type == S.RMU and c.id in broken):
        return "ring-group"
    if c.type == S.RMU and p.type == S.MV_BUSBAR and not via:
        return "rmu-entry"
    if c.type == S.BUS_COUPLER or p.type == S.BUS_COUPLER:
        ends = [items[q] for q in c.parents] if c.type == S.BUS_COUPLER else []
        if any(e.type == S.GENERATOR for e in ends):
            return "source"
    if p.type == S.GENERATOR or c.type == S.GENERATOR:
        return "source"
    if c.type == S.LV_BUSBAR and p.type in (S.FEEDER, S.LV_BUSBAR):
        return "lv-subboard"
    if c.type in (S.FEEDER, S.MCC) + S.TERMINALS \
            and p.type in (S.MV_BUSBAR, S.RMU):
        return "mv-feeder"
    # MV board -> transformer -> MV board, seen from either edge
    if c.type == S.TRANSFORMER and p.type in (S.MV_BUSBAR, S.RMU):
        kids = [items[o] for o in items if c.id in items[o].parents]
        if any(k.type == S.MV_BUSBAR for k in kids):
            return "multi-voltage"
    if c.type == S.MV_BUSBAR and p.type == S.TRANSFORMER:
        if any(items[q].type in (S.MV_BUSBAR, S.RMU) for q in p.parents):
            return "multi-voltage"
        return "multi-voltage" if p.parents else "source"
    if c.type == S.BUS_COUPLER or p.type == S.BUS_COUPLER:
        return "multi-voltage"
    return "other"


def link_feeders(items):
    """Feeders that carry on to an LV sub-board: drawn as the cable
    between the two boards, with no symbol of their own."""
    return {f.id: next((p for p in f.parents
                        if p in items and items[p].type == S.LV_BUSBAR), None)
            for f in items.values() if f.type == S.FEEDER
            and any(f.id in c.parents and c.type == S.LV_BUSBAR
                    for c in items.values())}


def check(path):
    info, items, order = S.read_workbook(path)
    width = S.layout(items, order)
    svg = S.render(info, items, order, width)
    # a load named on a transformer that feeds a board is drawn as a way
    # of that board: judge the drawing against that reading
    for it in items.values():
        if it.type in (S.PUMP, S.MCC):
            it.parents = [S.tx_board(items, items[q]).id
                          if q in items and items[q].type == S.TRANSFORMER
                          and S.tx_board(items, items[q]) is not None else q
                          for q in it.parents]
    d = Drawing(svg, items, order)
    lfeed = link_feeders(items)

    rep = dict(file=os.path.basename(path), items=len(order),
               drawn=0, missing=[], duplicates=list(d.duplicates),
               edges=0, connected=0, disconnected=[], via=[],
               overlaps=[], unexpected=[], false_nets=[], crossings=0,
               labels=0, bypassed=[],
               off_sheet=d.off_sheet, tags=defaultdict(int))

    # items
    for oid in order:
        it = items[oid]
        if it.type == S.BUS_COUPLER or oid in lfeed:
            continue                       # a coupler is an edge, not a symbol
        if oid in d.sym and d.item_nodes.get(oid):
            rep["drawn"] += 1
        elif oid in d.sym:
            rep["drawn"] += 1              # drawn but nothing touches it
        else:
            rep["missing"].append(oid)
    rep["items"] = sum(1 for o in order if items[o].type != S.BUS_COUPLER
                       and o not in lfeed)
    mvset = mv_chain_set(items)
    broken = set(rep["missing"]) | set(rep["duplicates"])
    for oid in broken:
        t = items[oid].type
        tag = ("multi-voltage" if oid in mvset else
               "ring-group" if t == S.RMU else "other")
        rep["tags"][tag] += 1

    # edges
    expected = []
    for oid in order:
        it = items[oid]
        if it.type == S.BUS_COUPLER:
            ends = [p for p in it.parents if p in items]
            if len(ends) == 2:
                expected.append((ends[0], ends[1], oid))
            continue
        if oid in lfeed:
            continue                       # counted as board -> sub-board
        for p in it.parents:
            if p in lfeed and lfeed[p] is not None:
                expected.append((lfeed[p], oid, p))
            elif p in items:
                expected.append((p, oid, None))
    rep["edges"] = len(expected)
    edge_paths = {}
    for p, c, cpl in expected:
        n, _ = d.path(p, c, avoid_others=True)
        name = f"{p}>{c}" if cpl is None else f"{p}={c} ({cpl})"
        if n is not None:
            rep["connected"] += 1
            edge_paths[name] = n
            continue
        n2, via = d.path(p, c, avoid_others=False)
        child = c if cpl is None else cpl
        if n2 is not None and via:
            tag = classify_edge(items, p, child, via, mvset, broken)
            rep["via"].append((name, sorted(via), tag))
        else:
            tag = classify_edge(items, p, child, set(), mvset, broken)
            rep["disconnected"].append((name, tag))
        rep["tags"][tag] += 1

    # a transformer must be crossed: its supply lands on one terminal and
    # its loads leave from the other.  Supply and a load on the same
    # terminal means the wires reach round it and it changes nothing.
    for oid in order:
        it = items[oid]
        if it.type != S.TRANSFORMER or oid not in d.item_nodes:
            continue
        sup = {edge_paths.get(f"{p}>{oid}") for p in it.parents} - {None}
        loads = set()
        for c in order:
            if oid in items[c].parents and items[c].type != S.BUS_COUPLER:
                n, _ = d.path(c, oid, avoid_others=True)
                if n is not None:
                    loads.add(n)
        if sup & loads:
            rep["bypassed"].append(oid)
            rep["tags"]["tx-bypass"] += 1

    # superimposed conductors: collinear overlap between different wires
    conds = [s for s in d.wires if s["kind"] == "cond" and seg_len(s) > LONG]
    for i, a in enumerate(conds):
        for b in conds[i + 1:]:
            ov = collinear_overlap(a, b)
            if ov > 4.0:
                rep["overlaps"].append((round(a["x1"]), round(a["y1"]),
                                        round(ov)))
    # overlaps above the MV bar row on a multi-voltage site are that bug
    mv_bar_y = min([sym["seg"]["y1"] for o, syms in d.sym.items()
                    for sym in syms if sym["kind"] == "bar"
                    and items[o].type == S.MV_BUSBAR] or [0])
    for x, y, ov in rep["overlaps"]:
        in_rmu = any(r["x"] - 2 <= x <= r["x"] + r["w"] + 2
                     and r["y"] - 2 <= y <= r["y"] + r["h"] + 2
                     for r in d.rmus)
        if in_rmu:
            rep["tags"]["ring-group"] += 1
        elif mvset and y < mv_bar_y:
            rep["tags"]["multi-voltage"] += 1
        else:
            rep["tags"]["lane-overlap"] += 1

    # unexpected direct connections between items the table does not join
    exp_pairs = {frozenset((p, c)) for p, c, _ in expected}
    # ways off one supply share its split: two panels on a transformer
    # secondary, two transformers tee'd off one incomer
    for oid in order:
        it = items[oid]
        if it.type in (S.TRANSFORMER, S.GENERATOR, S.FEEDER, S.MV_INCOMER):
            kids = [o for o in order if oid in items[o].parents]
            for i, a in enumerate(kids):
                for b in kids[i + 1:]:
                    exp_pairs.add(frozenset((a, b)))
    ids = [o for o in order if o in d.item_nodes]
    for i, a in enumerate(ids):
        for b in ids[i + 1:]:
            if frozenset((a, b)) in exp_pairs:
                continue
            n, _ = d.path(a, b, avoid_others=True)
            if n is not None:
                rep["unexpected"].append(f"{a}~{b}")
    # one merged lane joins many items: count nets, not pairs
    nets, seen = [], set()
    adj = defaultdict(set)
    for pair in rep["unexpected"]:
        a, b = pair.split("~")
        adj[a].add(b)
        adj[b].add(a)
    for a in adj:
        if a in seen:
            continue
        comp, q = set(), [a]
        while q:
            n = q.pop()
            if n in comp:
                continue
            comp.add(n)
            q.extend(adj[n])
        seen |= comp
        nets.append(sorted(comp))
    rep["false_nets"] = nets
    for comp in nets:
        if any(o in mvset for o in comp):
            rep["tags"]["multi-voltage"] += 1
        elif any(items[o].type == S.RMU for o in comp):
            rep["tags"]["ring-group"] += 1
        else:
            rep["tags"]["lane-overlap"] += 1

    # crossings between long strokes (cosmetic)
    longs = [s for s in d.wires if seg_len(s) > LONG]
    for i, a in enumerate(longs):
        for b in longs[i + 1:]:
            if segs_cross(a, b):
                rep["crossings"] += 1

    # label collisions: text boxes hit by a long stroke or another text.
    # A drawn box with a fill (the VSD drive box) masks the conductor
    # behind it, so its own text is not a collision.
    masks = [r for r in d.rects if r.get("filled")]

    def masked(t):
        return any(r["x"] <= t["x"] <= r["x"] + r["w"]
                   and r["y"] <= t["y"] <= r["y"] + r["h"] for r in masks)
    boxes = []
    for t in d.texts:
        if masked(t):
            continue
        wdt = 0.55 * t["size"] * len(t["s"])
        if t["rot"]:
            boxes.append((t["x"] - t["size"] * 0.8, t["y"],
                          t["x"] + t["size"] * 0.2, t["y"] + wdt))
        else:
            x0 = {"start": t["x"], "middle": t["x"] - wdt / 2,
                  "end": t["x"] - wdt}[t["anchor"]]
            boxes.append((x0, t["y"] - t["size"], x0 + wdt, t["y"]))
    for i, (ax0, ay0, ax1, ay1) in enumerate(boxes):
        for (bx0, by0, bx1, by1) in boxes[i + 1:]:
            if ax0 < bx1 and bx0 < ax1 and ay0 < by1 and by0 < ay1:
                rep["labels"] += 1
        for s in longs:
            lx0, lx1 = sorted((s["x1"], s["x2"]))
            ly0, ly1 = sorted((s["y1"], s["y2"]))
            if lx0 <= ax1 and ax0 <= lx1 and ly0 <= ay1 and ay0 <= ly1:
                # axis-aligned strokes only: box vs. degenerate rect
                if (lx1 - lx0 < 0.6 and ax0 < lx0 < ax1) or \
                   (ly1 - ly0 < 0.6 and ay0 < ly0 < ay1):
                    rep["labels"] += 1

    # a false "outgoing not defined" on an item that has no load by design
    for t in d.texts:
        if "outgoing not defined" in t["s"]:
            owner = min(((math.hypot(t["x"] - tx["x"], t["y"] - tx["bot"]), o)
                         for o, syms in d.sym.items() for tx in syms
                         if tx["kind"] == "tx"), default=(1e9, None))[1]
            if owner:
                blob = (items[owner].desc + " " + items[owner].notes).lower()
                if any(w in blob for w in NO_LOAD_WORDS):
                    rep["tags"]["no-load-ok"] += 1
    rep["tags"] = dict(rep["tags"])
    return rep


# --------------------------------------------------------------- report

def fmt(rep):
    line = (f"{rep['file']:26s} items {rep['drawn']}/{rep['items']}  "
            f"edges {rep['connected']}/{rep['edges']}")
    bad = len(rep["disconnected"]) + len(rep["via"])
    line += (f" ({len(rep['disconnected'])} disconnected, "
             f"{len(rep['via'])} via-other)" if bad else "")
    line += (f"  overlaps {len(rep['overlaps'])}  false nets "
             f"{len(rep['false_nets'])}  crossings {rep['crossings']}  "
             f"labels {rep['labels']}  off-sheet {rep['off_sheet']}")
    if rep["bypassed"]:
        line += f"  bypassed {len(rep['bypassed'])}"
    out = [line]
    if rep["missing"]:
        out.append(f"    missing symbols: {', '.join(rep['missing'])}")
    if rep["duplicates"]:
        out.append(f"    drawn twice: {', '.join(rep['duplicates'])}")
    for name, tag in rep["disconnected"]:
        out.append(f"    disconnected  {name:22s} [{tag}]")
    for name, via, tag in rep["via"]:
        out.append(f"    via {','.join(via):10s} {name:22s} [{tag}]")
    for comp in rep["false_nets"]:
        out.append(f"    drawn as one net, table says no: {' ~ '.join(comp)}")
    for oid in rep["bypassed"]:
        out.append(f"    transformer {oid} bypassed: supply and load on the "
                   f"same terminal")
    for x, y, ov in rep["overlaps"][:6]:
        out.append(f"    superimposed conductors near ({x},{y}), {ov} px")
    if len(rep["overlaps"]) > 6:
        out.append(f"    ... {len(rep['overlaps']) - 6} more overlaps")
    return "\n".join(out)


def matrix(reps):
    tags = [t for t in TAGS if any(r["tags"].get(t) for r in reps)]
    if not tags:
        return "no failures to attribute"
    files = [r["file"].replace(".xlsx", "") for r in reps]
    head = f"{'modification':46s}" + "".join(f"{f[:10]:>11s}" for f in files) \
        + f"{'total':>8s}"
    rows = [head, "-" * len(head)]
    for t in tags:
        cells = [r["tags"].get(t, 0) for r in reps]
        rows.append(f"{TAGS[t]:46s}" + "".join(f"{c:>11d}" for c in cells)
                    + f"{sum(cells):>8d}")
    return "\n".join(rows)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("workbooks", nargs="+")
    ap.add_argument("--json", action="store_true", help="machine output")
    ap.add_argument("--quiet", action="store_true",
                    help="scorecard lines only, no failure detail")
    args = ap.parse_args()
    reps, failed = [], False
    for wb in args.workbooks:
        rep = check(wb)
        reps.append(rep)
        if rep["missing"] or rep["disconnected"] or rep["via"]:
            failed = True
        if not args.json:
            print(fmt(rep) if not args.quiet else fmt(rep).splitlines()[0])
    if args.json:
        print(json.dumps(reps, indent=1))
    elif len(reps) > 1:
        print()
        print(matrix(reps))
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
