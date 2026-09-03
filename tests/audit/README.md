# Audit workbooks

Ten sites written by an independent tester working only from `README.md`, not
from the engine. They are the only workbooks here nobody wrote to match the
code, which is what makes them worth keeping: every other fixture was added
after its bug was found.

| workbook | what it exercises |
|---|---|
| `w01_waterworks.xlsx` | two MV boards with an N.O. tie, six transformers, LV tie, MCCs with VSD and spare motors, sub-boards, capacitor bank and NER on MV gear |
| `w02_ring.xlsx` | four RMUs closed in a ring, a spur RMU, a sub-ring, a genset through a changeover |
| `w03_pumpstation.xlsx` | a dedicated transformer whose MCC is the board, two MCCs on an MSB, a 400/300 V motor transformer, an LV/LV transformer to a 230 V board |
| `w04_generation.xlsx` | a generation board feeding a step-up beside a utility incomer, and a reversed step-up to a generator |
| `w05_deepcascade.xlsx` | four board levels with 40-character descriptions |
| `w06_wideboard.xlsx` | one bar with 24 ways of every load type |
| `w07_multisupply.xlsx` | a board fed from two transformers on different MV boards, a sub-board fed from two feeders |
| `w08_wrongloads.xlsx` | **rows that are wrong on purpose**: a pump on an MV incomer, a feeder on a motor, an MCC on a feeder and on an RMU. Five edges must stay disconnected and every bad row must draw a message |
| `w10_everything.xlsx` | 95 rows, the three sites above merged |
| `w11_couplers_levels.xlsx` | couplers between boards on different levels and past an intervening board, three voltage tiers |

`w08_wrongloads.xlsx` is the one sheet that never scores clean, by design. All
the others must: every item drawn once, every edge connected, no superimposed
conductors, no false nets.

```bash
python sld_check.py tests/audit/*.xlsx
```

The findings that came out of this set, and the fixes, are the last four rows
of `../BASELINE.md`.
