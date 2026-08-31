#!/usr/bin/env python3
"""Generate the example site-survey workbooks (and a blank template).

Run:  python make_examples.py
Writes examples/config1_single_tx.xlsx, config2_twin_tx.xlsx,
config3_ring_main.xlsx and template.xlsx.
"""

import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ARIAL = "Arial"
HEADERS = ["ID", "Type", "Description", "Rating", "Voltage",
           "Feeds From", "Notes"]
COL_WIDTHS = [10, 16, 26, 12, 12, 14, 30]
TYPES = ["MV Incomer", "MV Busbar", "RMU", "Transformer", "Pump",
         "LV Busbar", "Feeder", "MCC", "Bus Coupler"]

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")


def build_workbook(path, info, rows, legend_note=None):
    wb = Workbook()

    # ---- Info sheet ----------------------------------------------------
    ws = wb.active
    ws.title = "Info"
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 46
    for r, (key, val) in enumerate(info, start=1):
        kc = ws.cell(row=r, column=1, value=key)
        kc.font = Font(name=ARIAL, bold=True, size=11)
        vc = ws.cell(row=r, column=2, value=val)
        vc.font = Font(name=ARIAL, size=11)
        vc.fill = INPUT_FILL

    # ---- Equipment sheet -----------------------------------------------
    ws = wb.create_sheet("Equipment")
    for j, (h, w) in enumerate(zip(HEADERS, COL_WIDTHS), start=1):
        ws.column_dimensions[get_column_letter(j)].width = w
        c = ws.cell(row=1, column=j, value=h)
        c.font = Font(name=ARIAL, bold=True, size=11, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    for i, row in enumerate(rows, start=2):
        for j, val in enumerate(row, start=1):
            c = ws.cell(row=i, column=j, value=val)
            c.font = Font(name=ARIAL, size=11)
            c.fill = INPUT_FILL

    # Type column dropdown (rows 2..60 so extra rows added on site get it)
    dv = DataValidation(type="list",
                        formula1='"' + ",".join(TYPES) + '"',
                        allow_blank=True, showErrorMessage=True,
                        errorTitle="Unknown type",
                        error="Pick one of: " + ", ".join(TYPES))
    ws.add_data_validation(dv)
    dv.add("B2:B60")
    # extend yellow input fill to the blank rows people will fill on site
    for i in range(len(rows) + 2, 31):
        for j in range(1, len(HEADERS) + 1):
            ws.cell(row=i, column=j).fill = INPUT_FILL
            ws.cell(row=i, column=j).font = Font(name=ARIAL, size=11)

    # ---- legend / how-to sheet -----------------------------------------
    ws = wb.create_sheet("How to fill")
    ws.column_dimensions["A"].width = 100
    lines = [
        "HOW TO FILL THIS WORKBOOK (on site)",
        "",
        "Yellow cells are the ones to fill in. Everything else is fixed.",
        "1. Info sheet: site name, date, surveyor, free notes.",
        "2. Equipment sheet: one row per item, top of the network first.",
        "   - ID: short unique tag you invent (MV1, RMU1, TX1, BB1, F1...).",
        "   - Type: pick from the dropdown (MV Incomer, MV Busbar, RMU,",
        "     Transformer, Pump, LV Busbar, Feeder, MCC, Bus Coupler).",
        "     MV Busbar = an MV switchboard; Pump = an MV motor load;",
        "     MCC = motor control centre on an LV board. A Bus Coupler",
        "     between two MV Busbars is the bus tie.",
        "   - Rating / Voltage: as read from the nameplate (e.g. 1000 kVA,",
        "     11/0.4 kV, 630 A, 400 V).",
        "   - Feeds From: the ID of the item supplying this one. Use a comma",
        "     for two supplies (bus coupler between BB1, BB2 - or an RMU fed",
        "     by ring incomers MV1, MV2).",
        "3. Back at the office:  python sld_sketch.py thisfile.xlsx",
        "   and you get the single-line diagram as an SVG.",
    ]
    if legend_note:
        lines.insert(2, legend_note)
    for r, s in enumerate(lines, start=1):
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(name=ARIAL, size=11, bold=(r == 1))

    wb.save(path)
    print(f"wrote {path}")


# ------------------------------------------------------------------ data

CONFIG1 = dict(
    path="examples/config1_single_tx.xlsx",
    info=[("Site", "Example Site A"),
          ("Date", "2026-08-31"),
          ("Surveyed by", "A. Ardigo"),
          ("Notes", "Config 1 - single transformer substation")],
    rows=[
        ("MV1", "MV Incomer", "Utility supply", "", "11 kV", "",
         "Cable from utility"),
        ("RMU1", "RMU", "3-way ring main unit", "630 A", "11 kV", "MV1", ""),
        ("TX1", "Transformer", "Oil-immersed, Dyn11", "1000 kVA",
         "11/0.4 kV", "RMU1", ""),
        ("BB1", "LV Busbar", "Main LV board", "1600 A", "400 V", "TX1", ""),
        ("F1", "Feeder", "Lighting", "100 A", "400 V", "BB1", ""),
        ("F2", "Feeder", "Small power", "160 A", "400 V", "BB1", ""),
        ("F3", "Feeder", "HVAC", "250 A", "400 V", "BB1", ""),
        ("F4", "Feeder", "Spare", "100 A", "400 V", "BB1", ""),
    ])

CONFIG2 = dict(
    path="examples/config2_twin_tx.xlsx",
    info=[("Site", "Example Site B"),
          ("Date", "2026-08-31"),
          ("Surveyed by", "A. Ardigo"),
          ("Notes", "Config 2 - twin transformers with bus coupler")],
    rows=[
        ("MV1", "MV Incomer", "Utility supply", "", "11 kV", "", ""),
        ("RMU1", "RMU", "4-way ring main unit", "630 A", "11 kV", "MV1", ""),
        ("TX1", "Transformer", "Cast resin, Dyn11", "1600 kVA",
         "11/0.4 kV", "RMU1", ""),
        ("TX2", "Transformer", "Cast resin, Dyn11", "1600 kVA",
         "11/0.4 kV", "RMU1", ""),
        ("BB1", "LV Busbar", "LV board A", "2500 A", "400 V", "TX1", ""),
        ("BB2", "LV Busbar", "LV board B", "2500 A", "400 V", "TX2", ""),
        ("BC1", "Bus Coupler", "Bus section coupler", "2500 A", "400 V",
         "BB1, BB2", "Normally open"),
        ("F1", "Feeder", "Chillers", "630 A", "400 V", "BB1", ""),
        ("F2", "Feeder", "Riser 1", "400 A", "400 V", "BB1", ""),
        ("F3", "Feeder", "Lighting", "160 A", "400 V", "BB1", ""),
        ("F4", "Feeder", "Pumps", "400 A", "400 V", "BB2", ""),
        ("F5", "Feeder", "Riser 2", "400 A", "400 V", "BB2", ""),
        ("F6", "Feeder", "Spare", "250 A", "400 V", "BB2", ""),
    ])

CONFIG3 = dict(
    path="examples/config3_ring_main.xlsx",
    info=[("Site", "Example Site C"),
          ("Date", "2026-08-31"),
          ("Surveyed by", "A. Ardigo"),
          ("Notes", "Config 3 - ring main, two MV incomers")],
    rows=[
        ("MV1", "MV Incomer", "Ring in", "", "11 kV", "",
         "From substation North"),
        ("MV2", "MV Incomer", "Ring out", "", "11 kV", "",
         "To substation South"),
        ("RMU1", "RMU", "3-way ring main unit", "630 A", "11 kV",
         "MV1, MV2", ""),
        ("TX1", "Transformer", "Oil-immersed, Dyn11", "800 kVA",
         "11/0.4 kV", "RMU1", ""),
        ("BB1", "LV Busbar", "Main LV board", "1250 A", "400 V", "TX1", ""),
        ("F1", "Feeder", "Distribution board 1", "250 A", "400 V", "BB1", ""),
        ("F2", "Feeder", "Distribution board 2", "250 A", "400 V", "BB1", ""),
        ("F3", "Feeder", "Spare", "160 A", "400 V", "BB1", ""),
    ])

CONFIG4 = dict(
    path="examples/config4_dual_mv_boards.xlsx",
    info=[("Site", "Example Site D"),
          ("Date", "2026-08-31"),
          ("Surveyed by", "A. Ardigo"),
          ("Notes", "Config 4 - dual MV boards, riser feeders, N.O. tie")],
    rows=[
        ("MV1", "MV Incomer", "Utility incomer A", "", "11 kV", "", ""),
        ("MV2", "MV Incomer", "Utility incomer B", "", "11 kV", "", ""),
        ("MVB1", "MV Busbar", "MV switchboard A", "630 A", "11 kV",
         "MV1", "6 feeders to risers"),
        ("MVB2", "MV Busbar", "MV switchboard B", "630 A", "11 kV",
         "MV2", "6 feeders to risers"),
        ("TIE1", "Bus Coupler", "MV bus tie", "630 A", "11 kV",
         "MVB1, MVB2", "Normally open"),
        # ---- board A ways (row order = left-to-right on the board) ----
        ("TX1", "Transformer", "Riser R1, Dyn11", "1600 kVA",
         "11/0.4 kV", "MVB1", ""),
        ("TX2", "Transformer", "Riser R2, Dyn11", "1250 kVA",
         "11/0.4 kV", "MVB1", ""),
        ("TX3", "Transformer", "Riser R3, Dyn11", "1000 kVA",
         "11/0.4 kV", "MVB1", ""),
        ("P1", "Pump", "CHW pump 1", "315 kW", "11 kV", "MVB1", ""),
        ("P2", "Pump", "CHW pump 2", "315 kW", "11 kV", "MVB1", ""),
        ("P3", "Pump", "CW pump 1", "160 kW", "11 kV", "MVB1", ""),
        ("LVB1", "LV Busbar", "LV board R1", "2500 A", "400 V", "TX1", ""),
        ("MCC1", "MCC", "AHU plant", "400 A", "400 V", "LVB1", ""),
        ("MCC2", "MCC", "Pump room", "400 A", "400 V", "LVB1", ""),
        ("MCC3", "MCC", "Ventilation", "250 A", "400 V", "LVB1", ""),
        ("LVB2", "LV Busbar", "LV board R2", "2000 A", "400 V", "TX2", ""),
        ("MCC4", "MCC", "AHU plant", "400 A", "400 V", "LVB2", ""),
        ("MCC5", "MCC", "Ventilation", "250 A", "400 V", "LVB2", ""),
        ("LVB3", "LV Busbar", "LV board R3", "1600 A", "400 V", "TX3", ""),
        ("MCC6", "MCC", "AHU plant", "400 A", "400 V", "LVB3", ""),
        ("MCC7", "MCC", "Ventilation", "250 A", "400 V", "LVB3", ""),
        # ---- board B ways ----
        ("TX4", "Transformer", "Riser R4, Dyn11", "1600 kVA",
         "11/0.4 kV", "MVB2", ""),
        ("TX5", "Transformer", "Riser R5, Dyn11", "1250 kVA",
         "11/0.4 kV", "MVB2", ""),
        ("TX6", "Transformer", "Riser R6, Dyn11", "1000 kVA",
         "11/0.4 kV", "MVB2", ""),
        ("P4", "Pump", "CHW pump 3", "315 kW", "11 kV", "MVB2", ""),
        ("P5", "Pump", "CHW pump 4", "315 kW", "11 kV", "MVB2", ""),
        ("P6", "Pump", "CW pump 2", "160 kW", "11 kV", "MVB2", ""),
        ("LVB4", "LV Busbar", "LV board R4", "2500 A", "400 V", "TX4", ""),
        ("MCC8", "MCC", "AHU plant", "400 A", "400 V", "LVB4", ""),
        ("MCC9", "MCC", "Pump room", "400 A", "400 V", "LVB4", ""),
        ("MCC10", "MCC", "Ventilation", "250 A", "400 V", "LVB4", ""),
        ("LVB5", "LV Busbar", "LV board R5", "2000 A", "400 V", "TX5", ""),
        ("MCC11", "MCC", "AHU plant", "400 A", "400 V", "LVB5", ""),
        ("MCC12", "MCC", "Ventilation", "250 A", "400 V", "LVB5", ""),
        ("LVB6", "LV Busbar", "LV board R6", "1600 A", "400 V", "TX6", ""),
        ("MCC13", "MCC", "AHU plant", "400 A", "400 V", "LVB6", ""),
        ("MCC14", "MCC", "Ventilation", "250 A", "400 V", "LVB6", ""),
    ])

TEMPLATE = dict(
    path="examples/template.xlsx",
    info=[("Site", ""), ("Date", ""), ("Surveyed by", ""), ("Notes", "")],
    rows=[
        ("MV1", "MV Incomer", "Utility supply", "", "11 kV", "",
         "EXAMPLE ROW - overwrite with real data"),
    ],
    legend_note="The Equipment sheet contains ONE example row - "
                "overwrite it with the first real item.")


def main():
    os.makedirs("examples", exist_ok=True)
    for cfg in (CONFIG1, CONFIG2, CONFIG3, CONFIG4, TEMPLATE):
        build_workbook(**cfg)


if __name__ == "__main__":
    main()
