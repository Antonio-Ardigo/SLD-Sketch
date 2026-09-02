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
and the drawing rebuilding live as you type. The seven example configurations are
built in, and your table is kept in the browser between visits. The drawing
sits in a window of its own with scrollbars on both axes: drag it to pan,
Fit / 100 % / − / + (or ctrl+wheel, the keys 0, 1, −, +) to zoom, so a
2500 px site stays reachable end to end.

## The spreadsheet

One workbook per site, two sheets (plus a "How to fill" sheet with these same
instructions):

**`Info`** — key/value rows: Site, Date, Surveyed by, Notes. Shown in the
diagram's title block.

**`Equipment`** — one row per item:

| Column | Meaning | Example |
|---|---|---|
| ID | Short unique tag you invent | `MV1`, `RMU1`, `TX1`, `BB1`, `F1` |
| Type | Dropdown: MV Incomer, Generator, MV Busbar, RMU, Transformer, Pump, LV Busbar, Feeder, MCC, Bus Coupler, Capacitor Bank, Earthing/NER, Surge Arrester | `Transformer` |
| Description | Free text | `Oil-immersed, Dyn11` |
| Rating | From the nameplate | `1000 kVA`, `630 A` |
| Voltage | From the nameplate | `11/0.4 kV`, `400 V` |
| Protection | Device on **this item's supply side**: CB, LBS, Fuse, Fuse-switch, Contactor. Blank = the usual default. Comma list matches Feeds From order. Free text on a busbar (e.g. `87B differential`) is printed as a label annotation | `CB` or `LBS, CB` |
| Feeds From | ID of the item supplying this one; comma for two supplies | `RMU1` or `BB1, BB2` |
| Notes | Anything else | `Normally open` |

**Generation and step-up transformers.** A `Generator` feeding an LV board is
drawn as a G-circle in the transformer row. A transformer that feeds an MV
Busbar or RMU is a **step-up**: its source (a Generator, an LV Busbar acting as
a generation board, or an MV Incomer row) is drawn at the top, the transformer
below it, then down into the MV board beside any utility incomers.

There is **one `Transformer` type** — step-up and step-down are the same
IEC symbol, and which way round it is drawn follows `Feeds From`. Write
"step-up" in Description if you like; the Voltage field (`0.4/11 kV` against
`11/0.4 kV`) says it too. Sheets that used the old `SU Transformer` type still
load — the name is kept as an alias of `Transformer`.

| Wiring | Drawn as |
|---|---|
| TX `Feeds From` MV gear, an LV board feeds from TX | ordinary step-down |
| TX `Feeds From` a Generator or generation board, MV gear feeds from TX | source-on-top column |
| TX `Feeds From` a live LV board, MV gear feeds from TX | step-up in the transformer row |
| TX `Feeds From` MV gear, a Generator feeds from TX | the same column drawn upside down |
| TX `Feeds From` an LV board, only Pumps feed from TX (or nothing yet) | step-down under the board: a way of it, the transformer on the row below, the motor (or an open terminal) under it |
| TX `Feeds From` MV gear, no board but a Pump or MCC feeds from TX | dedicated transformer: the load under it (an MCC becomes the board: incomer device, box, bus and motors in the dashed outline) |
| TX feeds a board, and a Pump or MCC also names TX | a way of that board, not a tap on the secondary (a note says so) |

The last one works because a generator is never a load: a `Generator` whose
`Feeds From` names a transformer can only be feeding *up* through it.

Because `Feeds From` only ever points *upstream*, a step-up needs to be named
on the row of the board it supplies. Until you do that it still draws — as a
way under its LV board with an **open terminal** marked "outgoing not defined"
— and a warning tells you which row to add it to.

**A half-filled row still draws.** A transformer whose supply or whose load
you have not entered yet is drawn where it belongs, with an **open terminal**
on the missing side ("supply not defined" / "outgoing not defined") and a
warning saying so — rather than floating unconnected or vanishing. The board
under such a transformer still gets a bar sized from its own feeders.
Any other row with an empty `Feeds From`, a row whose supply cannot feed it
(a feeder off a pump, a pump off an MV incomer, anything off a bus coupler or
a terminal item), and rows that feed from each other round a loop no supply
reaches, are drawn floating and warned about by ID, so a survey sheet never
loses a row silently.

A transformer can also sit between **two LV boards**: a 400/690 V unit for
drives, say. Fed from one LV board and feeding another, it is drawn in the
transformer row between the two, its supply taken from the parent board's bar
and its output dropped into the fed board, which stands beside its parent with
its own feeders. These chain, and one unit can feed several boards.

**Generators as supplies.** A `Generator` can feed any board directly: name
it in the board's `Feeds From` (`MV, DG1, DG2` on an 11 kV generation board
with two gensets). On MV gear it stands above the bar like an incomer, on an
LV board over the bar; several supplies share one spread. A standby set on a
**changeover** is a `Bus Coupler` whose two ends are the board and the
generator (`MSB3, G1`, Notes `ATS`): the generator drops onto the board
through the coupler's device, with the coupler's ID and Notes beside it.

**Sub-boards.** An `LV Busbar` fed from a **Feeder** (`DBL1` feeds from `F1`)
or straight from another **LV Busbar** hangs on a row below its supply board,
under the way that feeds it, with its own feeders under it; the feeder's
device sits by the upper bar and the sub-board's own Protection, when given,
by its bar as its incomer. Cascades can be any depth, two boards can share
one feeder, and a sub-board can carry a tie, a motor, an MCC or a generator
like any other board.

**Outgoing ways on MV gear and terminal items.** A `Feeder` on an MV Busbar
or RMU is an outgoing cable way with its device and an arrow in the
transformer row. Three types need no load by design: `Capacitor Bank`,
`Earthing/NER` (neutral earthing resistor) and `Surge Arrester`, each drawn
with its IEC symbol to earth, on MV gear or an LV board. You need not even
change the Type: a `Feeder` whose Description or Notes says *capacitor*,
*PFC*, *kvar*, *NER*, *earthing* or *arrester* is drawn as that item, and a
`Transformer` with no load whose row says *earthing*, *NER* or *zig-zag* is
an earthing transformer, ending in a resistor to earth instead of an "outgoing
not defined" stub.

**Notes that change the symbol.** *VSD* (or *VFD*, *drive*) in a motor's
Description or Notes puts a drive box on its drop; Notes starting with
*spare*, *future* or *out of service* dash the way's conductor; *N.O.* or
*normally open* in an RMU's Notes marks the open point of a ring with an
"N.O." on the cable to the RMU it names (`N.O. towards RMU2`), or under the
box when no way is named. Nothing else in Notes is read.

**Motors.** A `Pump` on an MV Busbar or RMU is an MV motor drawn in the
transformer row. A `Pump` on an **LV Busbar** is an LV motor drawn in the
feeder band below the board, and a `Pump` fed straight from a **Transformer**
hangs under that transformer (a dedicated motor supply). When that transformer
itself feeds from an LV board (a 400/300 V motor supply, say) it is a way of
the board: dot on the bar, the board's device, the transformer on the row
below, the motor under its secondary.

A transformer's secondary goes to one place: the incomer of the board it
feeds. So a `Pump` or `MCC` whose `Feeds From` names a transformer that also
feeds a board is drawn as a **way of that board**, taking its protection from
the bar, and the reader prints a note suggesting the board in `Feeds From`. A
transformer that feeds **no board** keeps its loads under it: a dedicated
motor transformer, or a pump-station transformer feeding an `MCC` directly,
which is then drawn as the board itself: transformer, incomer device, MCC box,
its bus and the motors inside the dashed outline. An `MCC` belongs on an
LV Busbar; putting one on MV gear warns. A `Pump` or `Feeder` can feed from
an **MCC**: the MCC then gets a bus of its own on the row below the board,
and its motors hang off that bus with their starters (a contactor unless the
row's Protection says otherwise). Its incomer, box and bus sit inside a
dashed outline, like an RMU, so the MCC reads as one piece of switchgear.

Either type can also hang off an **RMU**: it takes a proper way inside the
enclosure, with its device on the tee-off. A generation **LV board** under an
SU is sized from its own feeders like any other board.

An MV Busbar or RMU can itself feed from another MV Busbar: the fed board or
RMU is then drawn on its own tier below its source, with the feed through its
Protection device. Cascades can be any depth (main MV board -> sub-board ->
sub-sub-board), and the transformer and LV rows move down to suit.

**Spurs and sub-rings.** RMUs that feed from each other draw side by side as
a ring. Where one RMU of a ring feeds a branch that has no supply of its own
(a spur RMU, or a sub-ring fed at both ends from the same RMU), that branch
hangs a tier below it: a tee-off way in the enclosure, a dog-leg cable down
to the branch, and the branch's own ring beside the enclosure's other ways.
A link written on both rows (`R1` feeds from `R2` and `R2` from `R1`) is
drawn once.

**Several voltage levels.** Whatever feeds a board is drawn above it. A board
fed from another board, directly or through a transformer, sits one tier
below it, so a 33 kV board, the 11 kV board its grid transformer feeds and a
3.3 kV pump board under that draw as three rows, with each transformer drawn
between the two tiers it joins: the upper board's breaker above it, the fed
board's incomer below. The Voltage column is printed but never read for
layout, so two boards of the same voltage can sit at different heights, and
the row direction is the one lever: write `MV ← ET ← HV ← U` to put the grid
on top, or `HV ← ET ← MV` to show an export board below the collector with
its utility incomer beside it.

Every sideways run (a board fed from two transformers, a step-up taking
supply from an LV board, a transformer fed from a board with no load entered
yet, a sub-board or RMU offset from its feeder) gets a lane of its own, so no
two connections ever share a line; the protection device always sits at the
board end of the run, where the cubicle is.

That's all the connectivity the tool needs: everything hangs off `Feeds From`.
A bus coupler feeds from its two busbars (`BB1, BB2`); an RMU on a ring feeds
from both ring incomers (`MV1, MV2`). A coupler between boards on **different
levels** is routed clear of both bars, and one that reaches **past an
intervening board** runs on its own lane above the bar row. Couplers must join
two busbars of the same kind — anything else (an RMU, one board, three boards)
warns instead of drawing, and a duplicate coupler on a pair warns too.

## Example configurations

| Workbook | Configuration | Sketch |
|---|---|---|
| `examples/config1_single_tx.xlsx` | MV incomer → RMU → 1000 kVA transformer → LV busbar, 4 feeders | `output/config1.svg` |
| `examples/config2_twin_tx.xlsx` | MV incomer → RMU → 2× 1600 kVA transformers → two LV busbars + bus coupler, 6 feeders | `output/config2.svg` |
| `examples/config3_ring_main.xlsx` | Ring main: 2 MV incomers → RMU → 800 kVA transformer → LV busbar, 3 feeders | `output/config3.svg` |
| `examples/config4_dual_mv_boards.xlsx` | 2 utility incomers → 2 MV switchboards (6 riser feeders each, N.O. bus tie) → 3 transformers + 3 MV pumps per board; each transformer → LV board with 2–3 MCCs | `output/config4.svg` |
| `examples/config5_cascaded_rmus.xlsx` | Utility incomer → RMU1, which feeds RMU2 and RMU3 by interconnecting cables; each of those feeds a 1000 kVA transformer, and each transformer feeds two LV panels | `output/config5.svg` |
| `examples/config6_closed_ring.xlsx` | Same as config 5 but with the RMU2–RMU3 cable in place, closing the ring RMU1–RMU2–RMU3–RMU1 (RMU3 feeds from `RMU1, RMU2`) | `output/config6.svg` |
| `examples/config7_mcc_motors.xlsx` | Pump station: utility incomer → RMU → 1000 kVA transformer → LV board with three MCCs; the pump MCC feeds four motors (one on a VSD) and an auxiliaries feeder, the blower MCC two VSD motors, each MCC in its own dashed enclosure | `output/config7.svg` |

Regenerate the workbooks with `python make_examples.py`, and the sketches with
`python sld_sketch.py <workbook> -o <out.svg>`.

**DXF export.** `python sld_sketch.py <workbook> --dxf` writes the SVG and
a `.dxf` beside it (or `python sld_dxf.py <workbook> -o out.dxf` for the DXF
alone). The file is R12 DXF, the dialect every CAD package and viewer opens:
the sketch exactly as the SVG draws it, built from the same symbol
primitives, with the equipment table that produced it laid out beside the
sheet, to its right. The drawing is centred on the origin and the file's
opening view is fitted to it, so it comes up in the middle of the window in
any CAD program. One drawing unit is one sketch pixel, meant as 1 mm. Entities are
sorted onto layers so a CAD user can switch parts off: `SLD_DRAWING`
(conductors and symbols), `SLD_BUSBAR` (the thick bars, as polylines with
width), `SLD_TEXT`, `SLD_ENCLOSURE` (RMU boxes, dashed), `SLD_FRAME` (title
and title block), `SLD_LEGEND` and `SLD_TABLE`. The Sketchpad page has the
same exporter behind its **Download DXF** button, and **Copy DXF** puts the
file text on the clipboard for places where downloads are blocked. CAD text
is set with a width factor that keeps it no wider than the browser's, long
table cells wrap, and `python sld_dxf.py <workbook> --check` reads the file
back and reports any text that overlaps another text or crosses a table
rule (none on any workbook in the repository).

**Checking a drawing against its table.** `python sld_check.py <workbook>`
renders the sheet, reads the SVG back as raw geometry and verifies that every
item is drawn once and every `Feeds From` edge is a continuous conductor
between the two symbols; it also reports conductors drawn on top of each
other and joints the table does not contain. `tests/` holds five demanding
sites, ten multi-level arrangements, five feature sheets and the baseline
scores.

## What the symbols mean

IEC-style sketch symbols: MV incomers come in from the top (source tick),
the RMU is a dashed enclosure with load-break switches on the incoming ways and
a fuse-switch on each transformer tee-off, MV and LV busbars are the thick
horizontal bars, transformers are the two overlapping circles, pumps/motors are
a circle with an "M 3~", MCCs are small labelled boxes, and the switch symbol
with an × at its hinge is a circuit breaker (every way on an MV switchboard
gets one). Feeders drop off a busbar
and end in an arrow; a Bus Coupler between two busbars is drawn as a breaker in
the gap, with its Notes text (e.g. "Normally open") underneath. The legend
folds into as many rows as the sheet width allows, and the sheet grows so
the longest feeder label never reaches the title line. A capacitor
bank is the two plates to earth, a neutral earthing resistor the box to
earth, a surge arrester the box with the arrow inside, to earth; the legend
gains these entries only on sheets that use them.

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
