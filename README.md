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

## The spreadsheet

One workbook per site, two sheets (plus a "How to fill" sheet with these same
instructions):

**`Info`** — key/value rows: Site, Date, Surveyed by, Notes. Shown in the
diagram's title block.

**`Equipment`** — one row per item:

| Column | Meaning | Example |
|---|---|---|
| ID | Short unique tag you invent | `MV1`, `RMU1`, `TX1`, `BB1`, `F1` |
| Type | Dropdown: MV Incomer, RMU, Transformer, LV Busbar, Feeder, Bus Coupler | `Transformer` |
| Description | Free text | `Oil-immersed, Dyn11` |
| Rating | From the nameplate | `1000 kVA`, `630 A` |
| Voltage | From the nameplate | `11/0.4 kV`, `400 V` |
| Feeds From | ID of the item supplying this one; comma for two supplies | `RMU1` or `BB1, BB2` |
| Notes | Anything else | `Normally open` |

That's all the connectivity the tool needs: everything hangs off `Feeds From`.
A bus coupler feeds from its two busbars (`BB1, BB2`); an RMU on a ring feeds
from both ring incomers (`MV1, MV2`).

## Example configurations

| Workbook | Configuration | Sketch |
|---|---|---|
| `examples/config1_single_tx.xlsx` | MV incomer → RMU → 1000 kVA transformer → LV busbar, 4 feeders | `output/config1.svg` |
| `examples/config2_twin_tx.xlsx` | MV incomer → RMU → 2× 1600 kVA transformers → two LV busbars + bus coupler, 6 feeders | `output/config2.svg` |
| `examples/config3_ring_main.xlsx` | Ring main: 2 MV incomers → RMU → 800 kVA transformer → LV busbar, 3 feeders | `output/config3.svg` |

Regenerate the workbooks with `python make_examples.py`, and the sketches with
`python sld_sketch.py <workbook> -o <out.svg>`.

## What the symbols mean

IEC-style sketch symbols: MV incomers come in from the top (source tick),
the RMU is a dashed enclosure with load-break switches on the incoming ways and
a fuse-switch on each transformer tee-off, transformers are the two overlapping
circles, LV busbars are the thick horizontal bars, and the small open squares
are LV breakers. Feeders drop off the busbar and end in an arrow.

This is a *sketch* tool for survey records — not a protection study or a
CAD-grade drawing.
