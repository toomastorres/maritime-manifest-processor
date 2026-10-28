"""
ens.py - Debit note de ENS (sólo EXPO).

Un renglón por B/L con ENS, a monto fijo (config.ENS_FIXED_AMOUNT).
"""

from openpyxl import Workbook

from .. import config
from ._style import (
    ALIGN_CENTER, BORDER_BOTTOM_THIN, BORDER_DOUBLE, FONT_CALIBRI_11,
    FONT_CALIBRI_12BI, FONT_CALIBRI_14B, FONT_CALIBRI_16B,
    set_column_widths, sheet_title,
)

COLUMN_WIDTHS = {"A": 6, "B": 21, "C": 34, "D": 21, "E": 10, "F": 18}
FIRST_DATA_ROW = 18


def _voyage_label(voyage):
    """'0426' -> '04/26'."""
    v = str(voyage or "")
    return f"{v[:2]}/{v[2:]}" if len(v) == 4 and v.isdigit() else v


def generar_ens(bls_expo, ship_name, voyage, output_path, cfg=None):
    cfg = cfg or config.EngineConfig()
    wb = Workbook()
    ws = wb.active

    ens_bls = sorted((bl for bl in bls_expo if bl.charge("ENS").monto),
                     key=lambda b: (b.port_of_discharge, b.bl_no))

    if not ens_bls:
        ws.title = "ENS"
        ws["B2"] = "Este viaje no tiene B/L con ENS."
        wb.save(output_path)
        return output_path

    ws.title = sheet_title(str(ens_bls[0].port_of_loading).title(), "ENS")
    set_column_widths(ws, COLUMN_WIDTHS)

    ws["B1"] = config.ENS_ISSUER_ADDRESS
    ws["B1"].font = FONT_CALIBRI_11

    ws["B3"] = "Messrs"
    ws["B3"].font = FONT_CALIBRI_12BI
    ws["C3"] = config.ENS_PRINCIPAL_NAME
    ws["C3"].font = FONT_CALIBRI_16B
    ws["C3"].border = BORDER_BOTTOM_THIN

    ws.merge_cells("B13:E15")
    ws["B13"] = "DEBIT NOTE"
    ws["B13"].font = FONT_CALIBRI_16B
    ws["B13"].alignment = ALIGN_CENTER
    ws.merge_cells("F13:F15")
    ws["F13"] = "u$s"
    ws["F13"].font = FONT_CALIBRI_16B
    ws["F13"].alignment = ALIGN_CENTER
    for r in range(13, 16):
        for c in range(2, 7):
            ws.cell(row=r, column=c).border = BORDER_DOUBLE

    for r in (1, 3, 13, 14, 15):
        ws.row_dimensions[r].height = 18

    ws["B17"] = f"{ship_name} {_voyage_label(voyage)}".strip()
    ws["B17"].font = FONT_CALIBRI_14B
    ws["D17"] = "BL"
    ws["D17"].font = FONT_CALIBRI_14B

    row = FIRST_DATA_ROW
    for bl in ens_bls:
        ws.cell(row=row, column=3,
                value=f"{bl.port_of_loading} / {bl.port_of_discharge}"
                ).font = FONT_CALIBRI_11
        ws.cell(row=row, column=4, value=bl.bl_clean).font = FONT_CALIBRI_14B
        ws.cell(row=row, column=5, value="ENS").font = FONT_CALIBRI_11
        ws.cell(row=row, column=6,
                value=config.ENS_FIXED_AMOUNT).font = FONT_CALIBRI_14B
        ws.row_dimensions[row].height = 18.75
        row += 1

    row += 1
    ws.cell(row=row, column=5, value="TOTAL").font = FONT_CALIBRI_14B
    total = ws.cell(row=row, column=6,
                    value=f"=SUM(F{FIRST_DATA_ROW}:F{row - 2})")
    total.font = FONT_CALIBRI_16B
    total.border = BORDER_DOUBLE

    wb.save(output_path)
    return output_path
