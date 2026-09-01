# SLD-Sketch

Turn a very simple site-survey spreadsheet into a single-line diagram sketch.

On a site visit you fill in a small Excel workbook (MV incomers, RMUs,
transformers, LV busbars, feeders). Back at the office you run one command and
get an SVG single-line diagram for future reference.

## Quick start

```bash
pip install -r requirements.txt          # just openpyxl

python sld_sketch.py examples/config1_single_tx.xlsx -o output/config1.svg
```

Open the SVG in any browser. To start a new survey, copy
`examples/template.xlsx` and fill in the yellow cells.

**No Python at hand?** Open `sld_sketchpad.html` in a browser — the same layout
engine ported to JavaScript, with an editable equipment table instead of Excel
and the drawing rebuilding live as you type. The six example configurations are
built in, and your table is kept in the browser between visits.

## The spreadsheet

One workbook per site, two sheets (plus a "How to fill" sheet with these same
instructions):

**`Info`** — key/value rows: Site, Date, Surveyed by, Notes. Shown in the
diagram's title block.

**`Equipment`** — one row per item:

| Column | Meaning | Example |
|---|---|---|
| ID | Short unique tag you invent | `MV1`, `RMU1`, `TX1`, `BB1`, `F1` |
| Type | Dropdown: MV Incomer, MV Busbar, RMU, Transformer, Pump, LV Busbar, Feeder, MCC, Bus Coupler | `Transformer` |
| Description | Free text | `Oil-immersed, Dyn11` |
| Rating | From the nameplate | `1000 kVA`, `630 A` |
| Voltage | From the nameplate | `11/0.4 kV`, `400 V` |
| Protection | Device on **this item's supply side**: CB, LBS, Fuse, Fuse-switch, Contactor. Blank = the usual default. Comma list matches Feeds From order. Free text on a busbar (e.g. `87B differential`) is printed as a label annotation | `CB` or `LBS, CB` |
| Feeds From | ID of the item supplying this one; comma for two supplies | `RMU1` or `BB1, BB2` |
| Notes | Anything else | `Normally open` |

An MV Busbar or RMU can itself feed from another MV Busbar: the fed board or
RMU is then drawn on its own tier below its source, with the feed through its
Protection device. Cascades can be any depth (main MV board -> sub-board ->
sub-sub-board), and the transformer and LV rows move down to suit.

That's all the connectivity the tool needs: everything hangs off `Feeds From`.
A bus coupler feeds from its two busbars (`BB1, BB2`); an RMU on a ring feeds
from both ring incomers (`MV1, MV2`).

## Example configurations

| Workbook | Configuration | Sketch |
|---|---|---|
| `examples/config1_single_tx.xlsx` | MV incomer → RMU → 1000 kVA transformer → LV busbar, 4 feeders | `output/config1.svg` |
| `examples/config2_twin_tx.xlsx` | MV incomer → RMU → 2× 1600 kVA transformers → two LV busbars + bus coupler, 6 feeders | `output/config2.svg` |
| `examples/config3_ring_main.xlsx` | Ring main: 2 MV incomers → RMU → 800 kVA transformer → LV busbar, 3 feeders | `output/config3.svg` |
| `examples/config4_dual_mv_boards.xlsx` | 2 utility incomers → 2 MV switchboards (6 riser feeders each, N.O. bus tie) → 3 transformers + 3 MV pumps per board; each transformer → LV board with 2–3 MCCs | `output/config4.svg` |
| `examples/config5_cascaded_rmus.xlsx` | Utility incomer → RMU1, which feeds RMU2 and RMU3 by interconnecting cables; each of those feeds a 1000 kVA transformer, and each transformer feeds two LV panels | `output/config5.svg` |
| `examples/config6_closed_ring.xlsx` | Same as config 5 but with the RMU2–RMU3 cable in place, closing the ring RMU1–RMU2–RMU3–RMU1 (RMU3 feeds from `RMU1, RMU2`) | `output/config6.svg` |

Regenerate the workbooks with `python make_examples.py`, and the sketches with
`python sld_sketch.py <workbook> -o <out.svg>`.

## What the symbols mean

IEC-style sketch symbols: MV incomers come in from the top (source tick),
the RMU is a dashed enclosure with load-break switches on the incoming ways and
a fuse-switch on each transformer tee-off, MV and LV busbars are the thick
horizontal bars, transformers are the two overlapping circles, pumps/motors are
a circle with an "M 3~", MCCs are small labelled boxes, and the switch symbol
with an × at its hinge is a circuit breaker (every way on an MV switchboard
gets one). Feeders drop off a busbar
and end in an arrow; a Bus Coupler between two busbars is drawn as a breaker in
the gap, with its Notes text (e.g. "Normally open") underneath.

The Protection column swaps the device drawn on an item's supply side, using
IEC 60617 notation, where the function mark sits at the hinge of the switch
blade: an × at the hinge is a circuit breaker, a circle at the hinge (blade
onto a contact bar) is a load-break switch (switch-disconnector), an arc at
the hinge is a contactor, a small rectangle with the conductor through it is
a fuse (switch + rectangle = fuse-switch), and a Fuse-contactor (an MV motor
starter) is the fuse in series with the contactor. Protection never changes the topology — only `Feeds
From` does — it only changes which symbol sits on the connection. Protection
on an MV Incomer is the utility's device and is not drawn (you get a warning).
RMU-to-RMU interconnecting cables draw a load-break switch inside each
enclosure by default; the fed RMU's Protection entry can override the symbol
on its incoming way.

This is a *sketch* tool for survey records — not a protection study or a
CAD-grade drawing.
