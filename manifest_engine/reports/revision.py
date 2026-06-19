"""
revision.py - Agrega la hoja 'REVISIÓN' a un workbook con la lista de alertas.
"""

from .styles import (
    apply_style_range, set_column_widths,
    BORDER_THIN, BORDER_MEDIUM, FILL_HEADER, FILL_TOTAL,
    FONT_ARIAL_10, FONT_ARIAL_10B, FONT_ARIAL_12B, ALIGN_LEFT,
)


def add_revision_sheet(wb, alertas):
    """Crea/rellena la hoja 'REVISIÓN' con las alertas (lista simple)."""
    ws = wb.create_sheet(title="REVISIÓN")
    set_column_widths(ws, {"A": 16, "B": 16, "C": 80})

    ws.cell(row=1, column=1, value="REVISIÓN — puntos a controlar").font = FONT_ARIAL_12B
    apply_style_range(ws, 1, 1, 3, fill=FILL_TOTAL)

    ws.cell(row=3, column=1, value="B/L")
    ws.cell(row=3, column=2, value="TIPO")
    ws.cell(row=3, column=3, value="DETALLE")
    apply_style_range(ws, 3, 1, 3, font=FONT_ARIAL_10B, border=BORDER_MEDIUM, fill=FILL_HEADER)

    row = 4
    if not alertas:
        ws.cell(row=row, column=1, value="Sin alertas.").font = FONT_ARIAL_10
        return ws

    for a in alertas:
        ws.cell(row=row, column=1, value=a.bl)
        ws.cell(row=row, column=2, value=a.tipo)
        c = ws.cell(row=row, column=3, value=a.detalle)
        c.alignment = ALIGN_LEFT
        apply_style_range(ws, row, 1, 3, font=FONT_ARIAL_10, border=BORDER_THIN)
        row += 1

    # Resumen por tipo
    row += 1
    ws.cell(row=row, column=1, value=f"TOTAL: {len(alertas)} alertas").font = FONT_ARIAL_10B
    return ws
