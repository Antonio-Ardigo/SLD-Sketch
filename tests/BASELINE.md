# Baseline — topology checker, before any engine change

Generated with `python sld_check.py tests/sites/*.xlsx` at commit
`75ec1a5` + checker (engine unchanged since "Collapse SU Transformer").
Re-run after a change and diff against this file.

## What the columns mean

- **items a/b** — items with exactly one drawn symbol / items in the table
  (bus couplers are edges, not symbols, and are not counted here)
- **edges a/b** — `Feeds From` edges drawn as a continuous conductor between
  the two symbols passing through no other item / edges in the table.
  *disconnected*: no conductor path at all. *via-other*: a path exists but
  only through a third item's symbol.
- **overlaps** — pairs of conductor segments lying on top of each other
- **false nets** — groups of items the drawing joins that the table does not
  (one merged lane joining six boards counts once)
- **crossings**, **labels**, **off-sheet** — cosmetic; reported, not scored

## The six shipped examples

All clean: every item drawn once, every edge connected, no overlaps, no
false nets. (`off-sheet 2` on the narrow sheets is the fixed-width legend
overflowing an 870 px drawing.)

## The five demanding sites

```
c1_wtw.xlsx                items 27/29  edges 26/31 (5 disconnected, 0 via-other)  overlaps 7  false nets 2  crossings 2  labels 12  off-sheet 0
    missing symbols: U1, U2
    drawn twice: HV1, HV2, MV1
    disconnected  U1>HV1                 [multi-voltage]
    disconnected  U2>HV2                 [multi-voltage]
    disconnected  MV1>RMU1               [rmu-entry]
    disconnected  MV2>RMU3               [rmu-entry]
    disconnected  MSB3=G1 (ATS)          [source]
    drawn as one net, table says no: HV1 ~ MV1
    drawn as one net, table says no: HV2 ~ MV2
c2_building.xlsx           items 29/29  edges 25/29 (4 disconnected, 0 via-other)  overlaps 1  false nets 1  crossings 1  labels 7  off-sheet 0
    disconnected  MSBB=G1 (CO)           [source]
    disconnected  F1>DBL1                [lv-subboard]
    disconnected  F13>DBE                [lv-subboard]
    disconnected  MSBB>DBL2              [lv-subboard]
    drawn as one net, table says no: MSBA ~ MSBB ~ TD ~ UPS
c3_pumps.xlsx              items 26/26  edges 25/26 (1 disconnected, 0 via-other)  overlaps 0  false nets 0  crossings 5  labels 7  off-sheet 0
    disconnected  B1>CAP                 [mv-feeder]
c4_ring.xlsx               items 29/30  edges 31/37 (6 disconnected, 0 via-other)  overlaps 7  false nets 1  crossings 2  labels 7  off-sheet 0
    missing symbols: R41
    drawn twice: R42
    disconnected  PB>R1                  [rmu-entry]
    disconnected  PB>R5                  [rmu-entry]
    disconnected  R4>R41                 [ring-group]
    disconnected  R42>R41                [ring-group]
    disconnected  R41>R42                [ring-group]
    disconnected  R41>T41                [ring-group]
    drawn as one net, table says no: R42 ~ T41
c5_hybrid.xlsx             items 21/22  edges 18/21 (3 disconnected, 0 via-other)  overlaps 6  false nets 2  crossings 13  labels 16  off-sheet 0
    missing symbols: U
    drawn twice: HV
    disconnected  U>HV                   [multi-voltage]
    disconnected  DG1>GB                 [source]
    disconnected  DG2>GB                 [source]
    drawn as one net, table says no: HV ~ MV
    drawn as one net, table says no: BB ~ PVB1 ~ PVB2 ~ SU1 ~ SU2 ~ SU3

modification                                       c1_wtw c2_buildin   c3_pumps    c4_ring  c5_hybrid   total
-------------------------------------------------------------------------------------------------------------
tiers by voltage level                                 15          0          0          0          7      22
LV cascades (sub-boards below their supply)             0          3          0          0          0       3
sources as first-class supplies / changeover            1          1          0          0          2       4
one lane allocator for every sideways run               1          2          0          1          4       8
MV outgoing ways and terminal items                     0          0          1          0          0       1
ring groups headed by an RMU                            0          0          0         13          0      13
board-fed RMU: draw the incoming way through to its bar          2          0          0          2          0       4
terminal item types (NER, capacitor bank, arrester)          0          0          1          0          0       1
```

## Probe workbooks (scratchpad regression set, 31 files)

Failures, all of them:

- `subdbA`, `subdbB` — sub-board fed from a feeder / from the main board:
  1 disconnected edge each `[lv-subboard]`
- `userA`, `userB`, `userC`, `suD`, `sumatrix`, `sketch2`, `mvtier`,
  `ringtier` (×2) — an RMU fed from an MV busbar: the incoming way stops at
  the enclosure border `[rmu-entry]`
- `break_test` — a deliberately broken sheet (feeder claiming two boards):
  exactly the one expected failure, reported as *via* the board it was drawn
  under

Everything else in the set scores clean.

## Reading the matrix

Numbers are failures attributed to the engine change that would remove
them, summed over missing/duplicate items, disconnected and via-other edges,
false nets and overlaps.

| change | failures removed (5 sites) | sites affected |
|---|---|---|
| tiers by voltage level | 22 | C1, C5 |
| ring groups headed by an RMU | 13 | C4 |
| one lane allocator for every sideways run | 8 | C1, C2, C4, C5 |
| sources as first-class supplies / changeover | 4 | C1, C2, C5 |
| board-fed RMU: incoming way drawn through to its bar | 4 (+9 probes) | C1, C4 |
| LV cascades | 3 (+2 probes) | C2 |
| MV outgoing ways and terminal items | 1 | C3 |
| terminal item types (NER, capacitor bank, arrester) | 1 | C3 |

The `rmu-entry` finding was not visible by eye: the enclosure makes the
reader assume the connection. It is the cheapest fix on the list and
appears in 11 workbooks.
