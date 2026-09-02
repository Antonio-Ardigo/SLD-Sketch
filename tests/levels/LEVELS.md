# Multi-level boards with step-up and step-down transformers

Ten workbooks probing every arrangement of boards at different voltage
levels joined by transformers. Built with today's vocabulary (`MV Busbar`
for every MV level, `Transformer` for both directions, `Generator` for
PV and gensets), voltages filled in on every row. Engine unchanged.

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

## Checker output

```
l1_chain.xlsx              items 7/8  edges 6/7 (1 disconnected, 0 via-other)  overlaps 3  false nets 1  crossings 1  labels 7  off-sheet 2
    missing symbols: U
    drawn twice: HV
    disconnected  U>HV                   [multi-voltage]
    drawn as one net, table says no: HV ~ MV
l2_sections.xlsx           items 12/14  edges 13/15 (2 disconnected, 0 via-other)  overlaps 6  false nets 2  crossings 2  labels 14  off-sheet 0
    missing symbols: U1, U2
    drawn twice: HVA, HVB
    disconnected  U1>HVA                 [multi-voltage]
    disconnected  U2>HVB                 [multi-voltage]
    drawn as one net, table says no: HVA ~ MVA
    drawn as one net, table says no: HVB ~ MVB
l3_two_down.xlsx           items 9/10  edges 8/9 (1 disconnected, 0 via-other)  overlaps 3  false nets 1  crossings 1  labels 7  off-sheet 0
    missing symbols: U
    drawn twice: MV
    disconnected  U>MV                   [multi-voltage]
    drawn as one net, table says no: MV ~ PB
l4_mid_two_supplies.xlsx   items 9/10  edges 8/9 (1 disconnected, 0 via-other)  overlaps 3  false nets 1  crossings 1  labels 9  off-sheet 0
    missing symbols: U
    drawn twice: HV
    disconnected  U>HV                   [multi-voltage]
    drawn as one net, table says no: HV ~ MV
l5_two_up.xlsx             items 13/13  edges 12/12  overlaps 1  false nets 1  crossings 1  labels 7  off-sheet 0
    drawn twice: MV
    drawn as one net, table says no: PVB1 ~ PVB2 ~ SU1 ~ SU2
l6_loop.xlsx               items 8/8  edges 8/8  overlaps 0  false nets 0  crossings 0  labels 4  off-sheet 0
l7_three_tier.xlsx         items 13/14  edges 12/13 (1 disconnected, 0 via-other)  overlaps 3  false nets 1  crossings 1  labels 8  off-sheet 0
    missing symbols: U
    drawn twice: HV, MV
    disconnected  U>HV                   [multi-voltage]
    drawn as one net, table says no: HV ~ MV
l8_mixed_cascade.xlsx      items 11/12  edges 10/11 (1 disconnected, 0 via-other)  overlaps 3  false nets 1  crossings 1  labels 9  off-sheet 0
    missing symbols: U
    drawn twice: MV
    disconnected  U>MV                   [multi-voltage]
    drawn as one net, table says no: MV ~ PB
l9_rmu_down.xlsx           items 9/10  edges 6/9 (3 disconnected, 0 via-other)  overlaps 0  false nets 1  crossings 1  labels 3  off-sheet 0
    missing symbols: R1
    drawn twice: R2
    disconnected  U>R1                   [multi-voltage]
    disconnected  R1>R2                  [multi-voltage]
    disconnected  R1>PT                  [multi-voltage]
    drawn as one net, table says no: P2 ~ PB ~ PT ~ R2 ~ U
l10_both_under_board.xlsx  items 10/10  edges 9/9  overlaps 0  false nets 0  crossings 0  labels 3  off-sheet 0

modification                                     l1_chain l2_section l3_two_dow l4_mid_two  l5_two_up    l6_loop l7_three_t l8_mixed_c l9_rmu_dow l10_both_u   total
--------------------------------------------------------------------------------------------------------------------------------------------------------------------
tiers by voltage level                                  7         14          7          7          1          0          8          7          5          0      56
one lane allocator for every sideways run               0          0          0          0          2          0          0          0          0          0       2
ring groups headed by an RMU                            0          0          0          0          0          0          0          0          1          0       1
```

## The pattern

Every failure is one mechanism. A transformer whose *child* is an MV
Busbar is treated as a step-up column: its parent is drawn above it as a
"source" and the fed board is placed on the single MV row like any other
board. When that parent is itself an MV Busbar rather than a generator, the
column paints the parent's *label* on a stub and the parent's real bar is
left as an orphan somewhere else on the row. So:

- a step-up **into** a board that also has a utility supply works (L4's
  genset side, L5, L10, the SU matrix);
- a same-voltage cascade works (L8's sub-board, `mv_depth` tiers);
- any board reached **only through a transformer from another MV board**
  is broken, regardless of direction (33 → 11 in L1/L2/L4/L7, 11 → 3.3 in
  L3/L7/L8/L9), and the breakage compounds per level (L7 has two).

The two clean sheets are exactly the two with a single MV voltage.

Score on this suite: 56 failures, 53 of them `multi-voltage`; the rest are
the shared step-up lane (L5) and one RMU consequence in L9.

## What this says about the fix

The row model has one MV row and one LV row. Anything else needs a tier
per voltage level, with a transformer drawn between the two tiers it
bridges; `mv_depth()` already does this for same-voltage cascades and is
the mechanism to generalise. These ten workbooks are the acceptance test:
after the change, L1–L9 must score every item drawn once and every edge
connected, and L6/L10 must stay byte-identical.
