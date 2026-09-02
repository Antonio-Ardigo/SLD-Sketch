# Multi-level boards with step-up and step-down transformers

Ten workbooks probing every arrangement of boards at different voltage
levels joined by transformers. Built with the ordinary vocabulary (`MV Busbar`
for every MV level, `Transformer` for both directions, `Generator` for
PV and gensets), voltages filled in on every row.

`python sld_check.py tests/levels/*.xlsx`

## Result by arrangement

| # | arrangement | result |
|---|---|---|
| L1 | 33 kV → 33/11 → 11 kV → 11/0.4 → LV | **fails** — 33 kV board an orphan bar, its incomer has no symbol, grid tx drawn as a column with the 33 kV board's label as its source |
| L2 | the same with A/B sections and ties at every level | **fails** twice over — both 33 kV sections orphaned, their tie joins two disconnected bars |
| L3 | 11 kV → 11/3.3 → 3.3 kV pump board, and 11/0.4 → LV | **fails** — the 3.3 kV board's transformer becomes a column with a phantom "11 kV board" stub as its source; the 11 kV board is drawn twice |
| L4 | 11 kV board fed from 33 kV above and a genset step-up below | **fails** on the 33 kV side; the step-up side draws correctly |
| L5 | PV → 400 V boards → 0.4/11 → 11 kV → 11/33 → 33 kV export board | **connects** (all 12 edges), but the export tx repeats the 11 kV board's label as a stub, and the two step-ups share one lane (one false net) |
| L6 | 11 kV A → 11/0.4 → LV → 0.4/11 → 11 kV B, A–B tie | **clean** |
| L7 | 33 kV → 11 kV → 3.3 kV, LV boards under the 11 kV and 3.3 kV boards | **fails** — three islands on one row: the 33 kV bar, the 11 kV bar, and the 3.3 kV bar fed from a phantom "11 kV board" stub |
| L8 | 11 kV main → 11 kV sub-board (same voltage) and → 11/3.3 → 3.3 kV board | **fails** on the 3.3 kV side only; the same-voltage cascade draws correctly one tier down |
| L9 | RMU ring → RMU 1 → 11/3.3 → 3.3 kV board with pumps | **fails** worst — the 3.3 kV bar is drawn on the RMU row, straight through both enclosures, superimposed on their inner buses; RMU 1 loses its symbol, RMU 2 is bound twice |
| L10 | one 11 kV board with a step-down, a reversed step-up and a step-up column | **clean** |

## Result after the voltage-tier layout

| # | before | after |
|---|---|---|
| L1 | 7 of 8 items, 6 of 7 edges | clean |
| L2 | 12 of 14, 13 of 15 | clean |
| L3 | 9 of 10, 8 of 9 | clean |
| L4 | 9 of 10, 8 of 9 | clean |
| L5 | connected, phantom stub, shared lane | connected, no stub; the two step-ups still share one lane (`lane-overlap`) |
| L6 | clean | clean, byte-identical |
| L7 | 13 of 14, 12 of 13 | clean |
| L8 | 11 of 12, 10 of 11 | clean |
| L9 | 9 of 10, 3 of 9 | clean |
| L10 | clean | clean, byte-identical |

`multi-voltage` failures on this suite: 53 → 0.

```
l1_chain.xlsx              items 8/8  edges 7/7  overlaps 0  false nets 0  crossings 0  labels 3  off-sheet 2
l2_sections.xlsx           items 14/14  edges 15/15  overlaps 0  false nets 0  crossings 0  labels 6  off-sheet 0
l3_two_down.xlsx           items 10/10  edges 9/9  overlaps 0  false nets 0  crossings 0  labels 2  off-sheet 0
l4_mid_two_supplies.xlsx   items 10/10  edges 9/9  overlaps 0  false nets 0  crossings 0  labels 4  off-sheet 0
l5_two_up.xlsx             items 13/13  edges 12/12  overlaps 1  false nets 1  crossings 1  labels 5  off-sheet 0
    drawn as one net, table says no: PVB1 ~ PVB2 ~ SU1 ~ SU2
l6_loop.xlsx               items 8/8  edges 8/8  overlaps 0  false nets 0  crossings 0  labels 4  off-sheet 0
l7_three_tier.xlsx         items 14/14  edges 13/13  overlaps 0  false nets 0  crossings 0  labels 2  off-sheet 0
l8_mixed_cascade.xlsx      items 12/12  edges 11/11  overlaps 0  false nets 0  crossings 0  labels 4  off-sheet 0
l9_rmu_down.xlsx           items 10/10  edges 9/9  overlaps 0  false nets 0  crossings 0  labels 1  off-sheet 2
l10_both_under_board.xlsx  items 10/10  edges 9/9  overlaps 0  false nets 0  crossings 0  labels 3  off-sheet 0

modification                                     l1_chain l2_section l3_two_dow l4_mid_two  l5_two_up    l6_loop l7_three_t l8_mixed_c l9_rmu_dow l10_both_u   total
--------------------------------------------------------------------------------------------------------------------------------------------------------------------
one lane allocator for every sideways run               0          0          0          0          2          0          0          0          0          0       2
```

## How the layout works now

Every MV busbar and RMU gets a **tier**: one per voltage level present on
the sheet (from the Voltage column, highest on top; a board with no voltage
takes its supplier's level, or the far side of the transformer that feeds
it), and within a level one sub-tier per same-voltage cascade, as before.
A transformer that joins two boards is drawn **between the two tiers**,
with the board's outgoing breaker above it and the fed board's incomer
below, whichever way the rows were written — so 33 → 11 kV, 11 → 3.3 kV
and an 11 → 33 kV export step-up all draw the same way. Tiers crossed by a
transformer are 200 px apart, plain cascades 150 px. The lower board takes
a slot beneath the upper one exactly as a same-voltage sub-board does, and
an RMU tree fed by an incomer is placed by the same machinery, so a 3.3 kV
board can hang off an RMU.
