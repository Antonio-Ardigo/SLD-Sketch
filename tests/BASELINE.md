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

## The seven shipped examples

All clean: every item drawn once, every edge connected, no overlaps, no
false nets, nothing off the sheet (the legend now wraps on narrow sheets,
and the sheet grows so downward feeder labels clear the title line).
Config 7 (a pump station with MCCs feeding the motors) was added with the
MCC enclosure and scores clean from the start: 16/16 items, 15/15 edges.

## History

| engine state | C1 edges | C2 | C3 | C4 | C5 | rmu-entry failures |
|---|---|---|---|---|---|---|
| before any change (`75ec1a5`) | 26/31 | 25/29 | 25/26 | 31/37 | 18/21 | 4 sites + 9 probes |
| board-fed RMU entry drawn through | 27/31 | 25/29 | 25/26 | 33/37 | 18/21 | 0 |
| voltage tiers: a transformer between two boards draws between their tiers | 29/31 (items 29/29) | 25/29 | 25/26 | 33/37 | 19/21 (items 22/22) | 0 |
| one lane allocator: every sideways run its own lane, device at the board | 30/31, false nets 0 | 25/29, false nets 0 | 25/26 | 33/37 | 19/21, false nets 0 | 0 |
| tiers from Feeds From alone (voltage a label) | 30/31 | 25/29 | 25/26 | 33/37 | 19/21 | 0 |
| generators as supplies and changeovers, LV sub-boards, MV outgoing ways and terminal items, spurs and sub-rings off an RMU, Notes keywords | **31/31** | **27/27** | **26/26** | **37/37** | **21/21** | 0 |
| motors and feeders under an MCC: its own bus on the row below, incomer, box and bus in a dashed outline (`tests/features/f6_mcc.xlsx`) | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |
| a transformer off an LV board feeding a motor hangs under the board (`tests/features/f7_board_tx.xlsx`); the checker now reports a transformer whose supply and load meet on one terminal (`bypassed`) | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |
| loads on a transformer: an MCC or motor fed straight from a transformer with no board hangs under it (the MCC as the board); named on a transformer that feeds a board, it is a way of that board (`tests/features/f8_tx_loads.xlsx`); board-fed transformers each land on their own point of the bar, farthest outermost, and a no-load board-fed transformer hangs under its board like a motor one (F7 gains an LV/LV transformer with a sub-board and a motor) | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |
| tester pass, batch A: `Earthing/NER` alias added to the Python reader (engines' alias tables now identical); checker binds IDs containing spaces and MCC boxes inside wide enclosures (two false positives gone); the reader warns on empty Feeds From, impossible supplies and unsupplied loops | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |
| tester pass, batch B: a coupler between levels runs in the gap beside the boards (it bridged the neighbour's bar); a coupler past an intervening board keeps its device at its own end (it sat on that board's incomer); a sub-board fed from two feeders gets a landing and a device per feeder (`tests/features/f9_ties.xlsx`, `f10_dual_feeds.xlsx`) | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |
| tester pass, batch C: a bar is at least wide enough for its label to sit on one side of a centred incomer, and the label moves past the landing conductors when it would still be crossed (label collisions on the tester set: 9/8/15 → 0/0/0); a changeover's Notes are not printed when they repeat its ID | 31/31 | 27/27 | 26/26 | 37/37 | 21/21 | 0 |

The first change removed every `rmu-entry` failure (C1 2, C4 2, nine
probes) and changed nothing else: the six examples are byte-identical and
every other row of the matrix is unchanged. In C1 the fix exposed the next
defect on the same ring: MV2's feed to RMU3 is now drawn, but on the same
lane as MV1's feed to RMU1, so the checker reports it *via* the ring
(counted under "ring groups" because the path runs through RMUs; the cause
is the shared lane).

## The five demanding sites (current)

Every item drawn once, every edge connected, no overlaps, no false nets.
The last change removed all 22 remaining failures: the sources / changeover
cases (C1 ATS, C2 changeover, C5 two MV gensets), the three sub-boards of C2,
C3's MV capacitor bank and earthing transformer, and C4's sub-ring, which is
now a compact tee-off pair in RMU 4 with the two sub-ring RMUs hung below.
C2 counts 27 edges instead of 29 because a feeder that carries on to a
sub-board is a cable, not a symbol: the checker scores the board-to-board
edge through it. (The `RCBO` protection warning is the sheet's, not the
engine's.)

```
c1_wtw.xlsx                items 29/29  edges 31/31  overlaps 0  false nets 0  crossings 1  labels 7  off-sheet 0
c2_building.xlsx           items 27/27  edges 27/27  overlaps 0  false nets 0  crossings 3  labels 6  off-sheet 0
c3_pumps.xlsx              items 26/26  edges 26/26  overlaps 0  false nets 0  crossings 5  labels 9  off-sheet 0
c4_ring.xlsx               items 30/30  edges 37/37  overlaps 0  false nets 0  crossings 1  labels 7  off-sheet 0
c5_hybrid.xlsx             items 22/22  edges 21/21  overlaps 0  false nets 0  crossings 11  labels 12  off-sheet 0

no failures to attribute
```

## The five feature sheets (`tests/features/`)

Built to exercise what the sites do not: gensets beside a utility incomer,
on an RMU and through two MV changeovers (F1); a spur RMU off a ring with
the links written on both rows and a spare feeder (F2); every terminal type
on MV gear, an RMU and an LV board, MV outgoing cable ways, an earthing
transformer by keyword (F3); three levels of sub-boards, two boards on one
feeder, a tie between sub-boards, a genset on a sub-board through an ATS,
a VSD motor (F4); an RMU-only sheet with a spur (F5).

```
f1_mv_sources.xlsx         items 14/14  edges 13/13  overlaps 0  false nets 0  crossings 0  labels 5  off-sheet 0
f2_spur.xlsx               items 18/18  edges 20/20  overlaps 0  false nets 0  crossings 0  labels 4  off-sheet 0
f3_terminals.xlsx          items 15/15  edges 14/14  overlaps 0  false nets 0  crossings 0  labels 0  off-sheet 0
f4_subboards.xlsx          items 15/15  edges 15/15  overlaps 0  false nets 0  crossings 0  labels 3  off-sheet 0
f5_rmu_spur.xlsx           items 17/17  edges 17/17  overlaps 0  false nets 0  crossings 1  labels 4  off-sheet 0

no failures to attribute
```

## Probe workbooks (scratchpad regression set, 31 files)

Failures, all of them:

- `break_test` — a deliberately broken sheet (feeder claiming two boards):
  exactly the one expected failure, reported as *via* the board it was drawn
  under
- `lvB`, `suB`, `userD` — deliberately wrong-way sheets (a generator row
  that *feeds from* its generation board, a step-up with no supply): the
  expected disconnected edge on each, unchanged since the first baseline

Everything else in the set scores clean, `subdbA` and `subdbB` (sub-board
fed from a feeder / from the main board) included since the sub-board
change.

## Reading the matrix

Numbers are failures attributed to the engine change that would remove
them, summed over missing/duplicate items, disconnected and via-other edges,
false nets and overlaps.

| change | failures removed (5 sites) | sites affected |
|---|---|---|
| ~~ring groups headed by an RMU~~ | done (14) | C1, C4 |
| ~~sources as first-class supplies / changeover~~ | done (4) | C1, C2, C5 |
| ~~LV cascades~~ | done (3, +2 probes) | C2 |
| ~~MV outgoing ways and terminal items~~ | done (1) | C3 |
| ~~terminal item types (NER, capacitor bank, arrester)~~ | done (1) | C3 |
| ~~board-fed RMU: incoming way drawn through to its bar~~ | done | — |
| ~~tiers by voltage level~~ | done (22 removed, plus 53 on `tests/levels`) | — |
| ~~one lane allocator for every sideways run~~ | done (7 removed; C4's remaining overlap is inside the sub-ring) | — |

The matrix is empty on every workbook in the repository; the tags stay in
the checker so a future regression is attributed the same way.

## RMU configurations

Four and five RMUs in every arrangement the reader accepts (chain, closed
ring, ring fed at both ends, ring plus spur, closed sub-ring, two substations
per RMU, long descriptions with MCCs, six / seven / eight RMUs) all draw
inside the sheet. Two arrangements did not, and are now fixtures:

| workbook | was |
|---|---|
| `tests/features/f11_rmu_cascade.xlsx` | an RMU feeding two RMUs put 18 items off the right edge: the sheet width followed the row cursor, which never saw a branch placed in a slot of its own |
| `tests/features/f12_rmu_mv_loads.xlsx` | a capacitor and an arrester on ring RMUs, on a site with no MV busbar, got no slot and fell into the leftover row at the far right, dragging their RMU's enclosure across the sheet: one 476 px superimposed bus and a false net over five items |

## Audit against the engine 15 commits earlier

Both engines drawing the same 42 workbooks, read back by the same checker
(`20bc766` versus `136f2ce`, plus this change):

| counted over the whole corpus | before | after |
|---|---|---|
| edges drawn disconnected (of 817) | 89 | 5 |
| superimposed conductors | 22 | 0 |
| false nets | 7 | 0 |
| items off the sheet | 8 | 0 |
| label collisions | 194 | 12 |
| crossings (not treated as faults) | 41 | 29 |

The five remaining disconnected edges are `tests/audit/w08_wrongloads.xlsx`,
whose rows are impossible on purpose. Of the twelve workbooks the older engine
mis-drew, seven printed no message at all; every one of those rows now names
itself. The ten independent sites the audit was written against are kept in
`tests/audit/`.
