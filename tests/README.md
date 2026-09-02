# Topology test suite

`sld_check.py` renders a workbook with the normal engine, reads the SVG back
as raw geometry (no help from the drawing code), and compares it with the
table: every item drawn once, every `Feeds From` edge a continuous conductor
between its two symbols passing through no other item, no two connections
sharing a piece of conductor, no drawn joint the table does not contain.

```bash
python sld_check.py examples/config*.xlsx tests/sites/*.xlsx
python sld_check.py tests/sites/c2_building.xlsx        # one site, full detail
python sld_check.py tests/sites/*.xlsx --json > out.json
```

Exit status is 1 when any item is missing or any edge is disconnected, so
the command can gate a commit. Each failure is tagged with the engine change
that would address it, and a run over several workbooks ends with a matrix
of tag × workbook.

`tests/sites/` holds five deliberately demanding sites (water works with
four voltage levels, building with a three-deep LV cascade, pumping station,
five-RMU ring with a sub-ring, hybrid PV/BESS/genset plant). `tests/levels/`
holds ten multi-level board arrangements and `tests/features/` five sheets
exercising generators as supplies and changeovers, spurs off a ring,
terminal items and MV outgoing ways, sub-boards, and an RMU-only chain with
a spur. `BASELINE.md` records the scores at each engine change; re-run and
diff to measure a change.

```bash
python sld_check.py examples/config*.xlsx tests/levels/*.xlsx tests/features/*.xlsx tests/sites/*.xlsx
```

The seven `examples/` workbooks must always score clean: all items drawn, all
edges connected, no overlaps, no false nets.
