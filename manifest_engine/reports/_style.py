"""
_style.py - Estilos y utilidades comunes a todos los reportes Excel.
"""

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

_thin = Side(style="thin")
_medium = Side(style="medium")
_double = Side(style="double")

BORDER_THIN = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
BORDER_MEDIUM = Border(left=_medium, right=_medium, top=_medium, bottom=_medium)
BORDER_DOUBLE = Border(left=_double, right=_double, top=_double, bottom=_double)
BORDER_BOTTOM_THIN = Border(bottom=_thin)

FONT_ARIAL_10 = Font(name="Arial", size=10)
FONT_ARIAL_10B = Font(name="Arial", size=10, bold=True)
FONT_ARIAL_12B = Font(name="Arial", size=12, bold=True)
FONT_ARIAL_14BI = Font(name="Arial", size=14, bold=True, italic=True)
FONT_ARIAL_16B = Font(name="Arial", size=16, bold=True)
FONT_CALIBRI_11 = Font(name="Calibri", size=11)
FONT_CALIBRI_12BI = Font(name="Calibri", size=12, bold=True, italic=True)
FONT_CALIBRI_14B = Font(name="Calibri", size=14, bold=True)
FONT_CALIBRI_16B = Font(name="Calibri", size=16, bold=True)

FILL_HEADER = PatternFill("solid", start_color="D9E1F2", end_color="D9E1F2")
FILL_SECTION = PatternFill("solid", start_color="E2EFDA", end_color="E2EFDA")
FILL_TOTAL = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")
FILL_YELLOW = PatternFill("solid", start_color="FFFFFF00", end_color="FFFFFF00")
FILL_ALERT = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
FILL_INPUT = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)

CURRENCY_FORMATS = {
    "USD": '#,##0.00',
    "EUR": '"EUR" #,##0.00',
    "GBP": '"GBP" #,##0.00',
    "BRL": '"BRL" #,##0.00',
}

REVISION_SHEET = "REVISIÓN"

_INVALID_TITLE_CHARS = str.maketrans({c: "-" for c in "[]:*?/\\"})


def sheet_title(name, default="HOJA"):
    """Nombre de hoja válido para Excel: sin caracteres prohibidos, <= 31."""
    limpio = str(name or "").translate(_INVALID_TITLE_CHARS).strip()
    return limpio[:31] or default


def apply_style_range(ws, row, col_start, col_end,
                      font=None, border=None, fill=None, alignment=None):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        if font:
            cell.font = font
        if border:
            cell.border = border
        if fill:
            cell.fill = fill
        if alignment:
            cell.alignment = alignment


def set_column_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def money_cell(ws, row, col, value, currency=None):
    """Escribe un importe como número, mostrando la moneda si no es USD."""
    cell = ws.cell(row=row, column=col, value=value)
    if currency and currency != "USD":
        cell.number_format = CURRENCY_FORMATS.get(currency, f'"{currency}" #,##0.00')
    return cell


def format_val(val, condition=None, currency=None, item_type=None):
    """
    Formateo de celda: los gastos de línea prepagos se muestran como texto
    'Prepaid' y los importes en moneda extranjera llevan el código adelante.
    """
    if not val or val == 0:
        return ""
    if condition == "P" and item_type in ("THC", "SWP", "SC"):
        return "Prepaid"
    if currency and currency in ("EUR", "BRL", "GBP"):
        return f"{currency} {val:,.2f}" if isinstance(val, (int, float)) else str(val)
    if item_type in ("BAF", "COMM") and currency:
        return f"{val:,.2f} {currency}" if isinstance(val, (int, float)) else str(val)
    if isinstance(val, (int, float)):
        return float(val)
    return str(val)


def add_revision_sheet(wb, alertas):
    """
    Agrega la hoja REVISIÓN con las alertas accionables del viaje.

    Siempre se agrega, aunque no haya alertas: que la hoja esté y diga
    'sin alertas' es la señal de que el control corrió.
    """
    ws = wb.create_sheet(title=REVISION_SHEET)
    set_column_widths(ws, {"A": 20, "B": 18, "C": 80})

    ws["A1"] = "REVISIÓN - puntos a controlar"
    ws["A1"].font = FONT_ARIAL_12B

    ws["A3"] = "B/L"
    ws["B3"] = "TIPO"
    ws["C3"] = "DETALLE"
    apply_style_range(ws, 3, 1, 3, font=FONT_ARIAL_10B,
                      border=BORDER_THIN, fill=FILL_HEADER)

    row = 4
    for a in alertas:
        ws.cell(row=row, column=1, value=a.bl)
        ws.cell(row=row, column=2, value=a.tipo)
        ws.cell(row=row, column=3, value=a.detalle)
        apply_style_range(ws, row, 1, 3, font=FONT_ARIAL_10, border=BORDER_THIN)
        ws.cell(row=row, column=3).alignment = ALIGN_WRAP
        row += 1

    row += 1
    total = ws.cell(row=row, column=1,
                    value=f"TOTAL: {len(alertas)} alertas" if alertas
                    else "Sin alertas: no hay nada para revisar.")
    total.font = FONT_ARIAL_10B
    if alertas:
        total.fill = FILL_ALERT
    return ws
