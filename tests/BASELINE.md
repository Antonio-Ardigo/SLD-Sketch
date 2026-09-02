# Baseline — topology checker scores

Generated with `python sld_check.py tests/sites/*.xlsx`. Re-run after a
change and diff against this file; the History table records each step.

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

## History

| engine state | C1 edges | C2 | C3 | C4 | C5 | rmu-entry failures |
|---|---|---|---|---|---|---|
| before any change (`75ec1a5`) | 26/31 | 25/29 | 25/26 | 31/37 | 18/21 | 4 sites + 9 probes |
| board-fed RMU entry drawn through | 27/31 | 25/29 | 25/26 | 33/37 | 18/21 | 0 |
| voltage tiers: a transformer between two boards draws between their tiers | 29/31 (items 29/29) | 25/29 | 25/26 | 33/37 | 19/21 (items 22/22) | 0 |
| one lane allocator: every sideways run its own lane, device at the board | 30/31, false nets 0 | 25/29, false nets 0 | 25/26 | 33/37 | 19/21, false nets 0 | 0 |

The first change removed every `rmu-entry` failure (C1 2, C4 2, nine
probes) and changed nothing else: the six examples are byte-identical and
every other row of the matrix is unchanged. In C1 the fix exposed the next
defect on the same ring: MV2's feed to RMU3 is now drawn, but on the same
lane as MV1's feed to RMU1, so the checker reports it *via* the ring
(counted under "ring groups" because the path runs through RMUs; the cause
is the shared lane).

## The five demanding sites (current)

```
c1_wtw.xlsx                items 27/29  edges 27/31 (3 disconnected, 1 via-other)  overlaps 7  false nets 2  crossings 2  labels 12  off-sheet 0
    missing symbols: U1, U2
    drawn twice: HV1, HV2, MV1
    disconnected  U1>HV1                 [multi-voltage]
    disconnected  U2>HV2                 [multi-voltage]
    disconnected  MSB3=G1 (ATS)          [source]
    via MV1,RMU1,RMU2 MV2>RMU3               [ring-group]
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
c4_ring.xlsx               items 29/30  edges 33/37 (4 disconnected, 0 via-other)  overlaps 7  false nets 1  crossings 2  labels 7  off-sheet 0
    missing symbols: R41
    drawn twice: R42
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
```

## Probe workbooks (scratchpad regression set, 31 files)

Failures, all of them:

- `subdbA`, `subdbB` — sub-board fed from a feeder / from the main board:
  1 disconnected edge each `[lv-subboard]`
- `break_test` — a deliberately broken sheet (feeder claiming two boards):
  exactly the one expected failure, reported as *via* the board it was drawn
  under

Everything else in the set scores clean, including the nine board-fed RMU
cases that failed before this change.

## Reading the matrix

Numbers are failures attributed to the engine change that would remove
them, summed over missing/duplicate items, disconnected and via-other edges,
false nets and overlaps.

| change | failures removed (5 sites) | sites affected |
|---|---|---|
| ring groups headed by an RMU | 14 | C1, C4 |
| sources as first-class supplies / changeover | 4 | C1, C2, C5 |
| LV cascades | 3 (+2 probes) | C2 |
| MV outgoing ways and terminal items | 1 | C3 |
| terminal item types (NER, capacitor bank, arrester) | 1 | C3 |
| ~~board-fed RMU: incoming way drawn through to its bar~~ | done | — |
| ~~tiers by voltage level~~ | done (22 removed, plus 53 on `tests/levels`) | — |
| ~~one lane allocator for every sideways run~~ | done (7 removed; C4's remaining overlap is inside the sub-ring) | — |
