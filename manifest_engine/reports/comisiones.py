"""
comisiones.py - Planilla de comisiones con formato final ('Comm. Inward' / 'Comm. Outward').

Diseño de 'ATL 0426 - Comm. Inward.xlsx' (anchos, colores, negritas, bordes,
alturas, fórmulas y ceros ocultos). La EXPO ('Comm. Outward') usa el mismo
diseño, con columnas para la tasa de Brasil.

  1. VEHICLES BOOKED BY AMA     rodados de la lista de Comercial: 1% del flete
  2. VEHICLES                   rodados NO AMA: tarifa fija USD por Car / Van / RoRo
  3. GENERAL CARGO COMMISSION   Bbulk: IMPO 2% | EXPO 4% (Brasil 3,5%) [ilustrativo]
  4. CONTAINERS COMMISSION      IMPO 2% | EXPO 4% (Brasil 3,5%), mínimo 25 USD/ctr
  5. TOTAL COMMISSION + CONTROL EQUIPMENT FEES PER CNTR = GRAND TOTAL

Los importes del manifiesto van como valores y los cálculos como fórmulas. Un
B/L en revisión aparece en rojo, sin importes, y el motivo está en la hoja
REVISIÓN; no suma a los totales.
"""

from collections import OrderedDict

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .. import config
from ..catalog import load_catalog
from ..domain.classification import format_cargo
from ..rules.commission import G_BBULK, G_CTR, G_ROD, MODE_AGENCY, evaluate_all
from ..validation import Alerta, revisar_bls
from ._style import add_revision_sheet

# ── estilos ───────────────────────────────────────────────────────────
F_TITLE = Font(name="Arial", size=16, bold=True)
F_ROE = Font(name="Arial", size=10, bold=True)
F_SECTION = Font(name="Arial", size=14, bold=True, italic=True)
F_HDR = Font(name="Arial", size=10, bold=True)
F_DATA = Font(name="Arial", size=10)
F_TOTAL = Font(name="Arial", size=12, bold=True)
F_TC = Font(name="Arial", size=11, bold=True)
F_BIG = Font(name="Arial", size=16, bold=True)
F_TOTCOMM = Font(name="Arial", size=14, bold=True, italic=True)
F_CAL = Font(name="Calibri", size=11)
F_CALB = Font(name="Calibri", size=11, bold=True)
F_CAL16B = Font(name="Calibri", size=16, bold=True)
F_PEND = Font(name="Arial", size=10, bold=True, color="9C0006")
F_BANNER = Font(name="Arial", size=11, bold=True, color="9C0006")

FILL_SECTION = PatternFill("solid", start_color="FFE2EFDA", end_color="FFE2EFDA")
FILL_PORT = PatternFill("solid", start_color="FFD9E1F2", end_color="FFD9E1F2")
FILL_TOTAL = PatternFill("solid", start_color="FFFCE4D6", end_color="FFFCE4D6")
FILL_PEND = PatternFill("solid", start_color="FFFFC7CE", end_color="FFFFC7CE")

_t = Side(style="thin")
_m = Side(style="medium")
B_THIN = Border(left=_t, right=_t, top=_t, bottom=_t)
B_MED = Border(left=_m, right=_m, top=_m, bottom=_m)
B_MED_TB = Border(top=_m, bottom=_m)
B_MED_RTB = Border(right=_m, top=_m, bottom=_m)
B_MED_LTB = Border(left=_m, top=_m, bottom=_m)
B_LRB = Border(left=_t, right=_t, bottom=_t)
B_LRT = Border(left=_t, right=_t, top=_t)
B_LR = Border(left=_t, right=_t)

CENTER = Alignment(horizontal="center", vertical="center")
HCENTER = Alignment(horizontal="center")

MONEY = "#,##0.00"
PCT = "0.00%"
INT = "#,##0_ ;\\-#,##0\\ "
ROE_FMT = "0.0000"
CEF_FMT = "[$USD]\\ #,##0.00"

WIDTHS = {"A": 25.7109375, "B": 15.7109375, "C": 15.7109375, "D": 15.7109375,
          "E": 15.7109375, "F": 15.7109375, "G": 15.7109375, "H": 15.0, "I": 13.28515625}

SHEET_TITLE = {"IMPO": "Comm Inward", "EXPO": "Comm Outward"}
PENDIENTE = "PENDIENTE DE REVISIÓN (ver hoja REVISIÓN)"


def _L(c):
    return "ABCDEFGHI"[c - 1]


class _W:
    """Escritor de la hoja con puntero de fila."""

    def __init__(self, ws):
        self.ws = ws
        self.row = 1

    def cell(self, col, value=None, font=None, fill=None, border=None, align=None, fmt=None, row=None):
        c = self.ws.cell(row=row or self.row, column=col)
        if value is not None:
            c.value = value
        if font:
            c.font = font
        if fill:
            c.fill = fill
        if border:
            c.border = border
        if align:
            c.alignment = align
        if fmt:
            c.number_format = fmt
        return c

    def height(self, h):
        self.ws.row_dimensions[self.row].height = h

    def merge(self, c1, c2):
        self.ws.merge_cells(start_row=self.row, start_column=c1, end_row=self.row, end_column=c2)

    def skip(self, n=1, last_height=None):
        self.row += n
        if last_height:
            self.ws.row_dimensions[self.row - 1].height = last_height

    def section_title(self, text, last_col):
        self.cell(1, text, F_SECTION, FILL_SECTION, B_MED)
        for c in range(2, last_col):
            self.cell(c, None, F_CAL, border=B_MED_TB)
        self.cell(last_col, None, F_CAL, border=B_MED_RTB)
        self.merge(1, last_col)
        self.height(19.5)
        self.row += 2

    def header(self, values, last_col):
        for c in range(1, last_col + 1):
            self.cell(c, values.get(c), F_HDR, border=B_THIN, align=CENTER)
        self.row += 1

    def port(self, name, last_col, centered=()):
        self.cell(1, name, F_HDR, FILL_PORT, B_THIN)
        for c in range(2, last_col + 1):
            self.cell(c, None, F_CAL, FILL_PORT, B_THIN, align=HCENTER if c in centered else None)
        self.row += 1

    def pending(self, bl_no, last_col):
        self.cell(1, bl_no, F_PEND, FILL_PEND, B_THIN)
        for c in range(2, last_col + 1):
            self.cell(c, None, F_DATA, FILL_PEND, B_THIN)
        self.cell(2, PENDIENTE, F_PEND)
        self.row += 1

    def total_row(self, label, formulas, last_col, fmts=None, aligns=None):
        self.cell(1, label, F_TOTAL, FILL_TOTAL, B_MED)
        for c in range(2, last_col + 1):
            self.cell(c, formulas.get(c), F_TOTAL, FILL_TOTAL, B_MED,
                      align=(aligns or {}).get(c), fmt=(fmts or {}).get(c))
        self.height(16.5)
        t = self.row
        self.row += 1
        return t

    def total_commission(self, label, formula, label_font=F_TC, merge=True, value_border=B_MED_RTB,
                         label_border=B_MED):
        self.cell(1, label, label_font, border=label_border, align=HCENTER if merge else None)
        self.cell(2, None, F_CAL, border=B_MED_RTB)
        self.cell(3, formula, F_BIG, border=value_border, fmt=MONEY)
        if merge:
            self.merge(1, 2)
        self.height(21)
        ref = f"C{self.row}"
        self.row += 1
        return ref


def _by_port(results):
    out = OrderedDict()
    for r in sorted(results, key=lambda r: (r.port, r.bl.bl_no)):
        out.setdefault(r.port, []).append(r)
    return out


def generar_comisiones(bls, op_type, ship_name, voyage, output_path, cfg=None, catalog=None,
                       extra_alerts=None):
    cfg = cfg or config.EngineConfig()
    catalog = catalog or load_catalog()
    op = op_type.upper()
    results = evaluate_all(bls, op, cfg, catalog)
    wb = _build(op, f"{ship_name} {voyage}", cfg, results)

    _detail_sheet(wb, results)
    alertas = []
    for r in results:
        if not r.ok:
            detalle = " | ".join(r.reasons) + (f" — {r.suggestion}" if r.suggestion else "")
            alertas.append(Alerta(r.bl.bl_no, "COMISIÓN - REVISAR", detalle))
    alertas += list(extra_alerts or [])
    alertas += revisar_bls(bls, op, cfg, catalog=catalog)
    add_revision_sheet(wb, alertas)
    wb.save(output_path)
    return results


def generar_modelo_comisiones(op_type, output_path, filas=8):
    """Planilla vacía con el formato final, para usar como modelo a mano."""
    op = op_type.upper()
    cfg = config.EngineConfig(roe=None, agency_bls=frozenset())
    wb = _build(op, "[BUQUE] [VIAJE]", cfg, [], filas_modelo=filas)
    wb.save(output_path)
    return output_path


def _build(op, buque_viaje, cfg, results, filas_modelo=0):
    rates = cfg.rates_for(op)
    expo = op == "EXPO"
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_TITLE[op]
    for col, w in WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.sheet_view.showZeros = False
    ws.page_setup.paperSize = 9
    ws.page_setup.orientation = "portrait"
    ws.page_setup.scale = 65
    ws.page_margins.left = ws.page_margins.right = 0.7086614173228347
    ws.page_margins.top = ws.page_margins.bottom = 0.7480314960629921
    w = _W(ws)

    # ── encabezado ────────────────────────────────────────────────────
    w.cell(1, f"COMMISSIONS {op} — {buque_viaje}", F_TITLE)
    w.height(20.25)
    w.row = 2
    w.cell(1, "ROE: EU 1,00 =", F_ROE)
    w.cell(2, cfg.roe, F_ROE, fmt=ROE_FMT)
    roe = "$B$2"
    w.row = 3
    pend = [r for r in results if not r.ok]
    avisos = []
    if cfg.agency_bls is None:
        avisos.append("ESTIMACIÓN: sin lista de Comercial, todos los rodados por unidad")
    if pend:
        avisos.append(f"{len(pend)} B/L PENDIENTES DE REVISIÓN (en rojo), no incluidos en los totales")
    if avisos:
        w.cell(1, " | ".join(avisos), F_BANNER)
    w.height(15.75)
    w.row = 4

    def grupos(fn):
        return _by_port([r for r in results if fn(r)])

    def modelo_rows(last_col, data_fn, centered=()):
        """Filas vacías con formato (sólo en el modelo)."""
        w.port(None, last_col, centered)
        for _ in range(filas_modelo):
            data_fn(None)
            w.row += 1

    tot = {}

    # ── 1. VEHICLES BOOKED BY AMA ─────────────────────────────────────
    w.section_title("VEHICLES BOOKED BY AMA", 7)
    w.header({2: "UNITS", 3: "AMOUNTS", 4: "AMOUNTS"}, 4)
    w.header({1: "B/L", 2: "CAR-VAN-RORO", 3: "USD", 4: "EURO"}, 4)
    w.cell(1, None, F_CAL, border=B_THIN, align=CENTER)
    w.cell(2, None, F_CAL, border=B_THIN, align=CENTER)
    w.cell(3, rates.rolling, F_CAL, border=B_THIN, align=CENTER, fmt=PCT)
    w.cell(4, rates.rolling, F_CAL, border=B_THIN, align=CENTER, fmt=PCT)
    rr = w.row
    w.row += 1
    first = w.row

    def ama_row(r):
        w.cell(1, r.bl.bl_clean if r else None, F_HDR, border=B_THIN)
        p = r.part(G_ROD) if r else None
        w.cell(2, sum(p.units.values()) if p else None, F_DATA, border=B_THIN)
        w.cell(3, round(p.base, 2) if p and p.currency != "EUR" else None, F_DATA, border=B_THIN,
               fmt=MONEY if p and p.currency != "EUR" else None)
        w.cell(4, round(p.base, 2) if p and p.currency == "EUR" else None, F_DATA, border=B_THIN,
               fmt=MONEY if p and p.currency == "EUR" else None)

    n_pend = 0
    if filas_modelo:
        modelo_rows(4, ama_row)
    for port, rs in grupos(lambda r: G_ROD in r.groups and r.rodado_mode == MODE_AGENCY).items():
        w.port(port, 4)
        for r in rs:
            if not r.ok:
                n_pend += 1
                w.pending(r.bl.bl_clean, 4)
                continue
            ama_row(r)
            w.row += 1
    last = w.row - 1
    w.skip(1, last_height=15.75)
    t = w.total_row("TOTAL", {c: f"=SUM({_L(c)}{first}:{_L(c)}{last})" for c in (2, 3, 4)}, 4,
                    fmts={3: MONEY, 4: MONEY})
    w.cell(1, "COMMISSION", F_CALB, border=B_LRB)
    w.cell(2, None, F_CALB, border=B_LRB)
    w.cell(3, f"=C{t}*$C${rr}", F_CALB, border=B_LRB, fmt=MONEY)
    w.cell(4, f"=D{t}*$D${rr}", F_CALB, border=B_LRB, fmt=MONEY)
    cm = w.row
    w.row += 1
    w.cell(1, "US DOLLARS", F_CALB, border=B_LRT)
    w.cell(2, None, F_CALB, border=B_LRT)
    w.cell(3, f"=C{cm}", F_CALB, border=B_LRT, fmt=MONEY)
    w.cell(4, f"=D{cm}*{roe}", F_CALB, border=B_THIN, fmt=MONEY)
    w.height(15.75)
    us = w.row
    w.row += 1
    tot["AMA"] = w.total_commission("TOTAL COMMISSION AMA", f"=C{us}+D{us}", value_border=B_MED)
    _pend_line(w, n_pend)
    w.skip(2, last_height=15.75)

    # ── 2. VEHICLES (por unidad) ──────────────────────────────────────
    estimado = cfg.agency_bls is None
    w.section_title("VEHICLES" + (" (ESTIMADO: SIN LISTA DE COMERCIAL)" if estimado else ""), 7)
    w.header({2: "CAR", 3: "VAN", 4: "RORO", 5: "CAR", 6: "VAN", 7: "RORO"}, 7)
    w.header({1: "B/L", 2: "UNITS", 3: "UNITS", 4: "UNITS", 5: "USD", 6: "USD", 7: "USD"}, 7)
    w.cell(1, "USD por unidad", F_CALB, border=B_THIN, align=CENTER)
    for c, fam in ((2, "CAR"), (3, "VAN"), (4, "RORO")):
        w.cell(c, config.PER_UNIT_TARIFF_USD[fam], F_CAL, border=B_THIN, align=CENTER, fmt=MONEY)
    for c in (5, 6, 7):
        w.cell(c, None, F_CAL, border=B_THIN, align=CENTER)
    vr = w.row
    w.row += 1
    first = w.row

    def unit_formulas():
        for c in (5, 6, 7):
            w.cell(c, f"=+{_L(c - 3)}{w.row}*{_L(c - 3)}${vr}", F_DATA, border=B_THIN, fmt=MONEY)

    def unit_row(r):
        w.cell(1, r.bl.bl_clean if r else None, F_HDR, border=B_THIN)
        p = r.part(G_ROD) if r else None
        for c, fam in ((2, "CAR"), (3, "VAN"), (4, "RORO")):
            w.cell(c, (p.units.get(fam) or None) if p else None, F_DATA, border=B_THIN, align=HCENTER)
        unit_formulas()

    n_pend = 0
    if filas_modelo:
        modelo_rows(7, unit_row, (2, 3, 4))
    for port, rs in grupos(lambda r: G_ROD in r.groups and r.rodado_mode != MODE_AGENCY).items():
        w.port(port, 7, (2, 3, 4))
        for r in rs:
            if not r.ok:
                n_pend += 1
                w.pending(r.bl.bl_clean, 7)
                continue
            unit_row(r)
            w.row += 1
    last = w.row - 1
    for c in (2, 3, 4):
        w.cell(c, None, F_CAL, align=HCENTER)
    unit_formulas()
    w.height(15.75)
    w.row += 1
    t = w.total_row("TOTAL", {c: f"=SUM({_L(c)}{first}:{_L(c)}{last})" for c in range(2, 8)}, 7,
                    fmts={2: INT, 3: INT, 4: INT, 5: MONEY, 6: MONEY, 7: MONEY},
                    aligns={2: HCENTER, 3: HCENTER, 4: HCENTER})
    w.cell(1, "TOTAL UNIDADES", F_CALB, border=B_LR)
    w.cell(2, f"=B{t}+C{t}+D{t}", F_CAL16B, border=B_LR, align=HCENTER, fmt=INT)
    w.height(21.75)
    w.row += 1
    tot["VEH"] = w.total_commission("TOTAL COMMISSION VEHICLES", f"=E{t}+F{t}+G{t}")
    _pend_line(w, n_pend)
    w.skip(2, last_height=15.75)

    # ── 3. GENERAL CARGO ──────────────────────────────────────────────
    cols = [(2, "USD", False), (3, "EURO", False)]
    if expo:
        cols += [(4, "USD", True), (5, "EURO", True)]
    lc = cols[-1][0]
    w.section_title("GENERAL CARGO COMMISSION", 7)
    w.header({c: ("AMOUNTS BRASIL" if b else "AMOUNTS") for c, _, b in cols}, lc)
    w.header({1: "B/L"} | {c: n for c, n, _ in cols}, lc)
    w.cell(1, None, F_CAL, border=B_THIN, align=CENTER)
    for c, _, b in cols:
        w.cell(c, rates.brasil_general_cargo if b else rates.general_cargo, F_CAL, border=B_THIN,
               align=CENTER, fmt=PCT)
    gr = w.row
    w.row += 1
    first = w.row

    def gc_row(r):
        w.cell(1, r.bl.bl_clean if r else None, F_HDR, border=B_THIN)
        p = r.part(G_BBULK) if r else None
        col = None
        if p:
            col = (4 if r.brasil else 2) + (1 if p.currency == "EUR" else 0)
        for c, _, _ in cols:
            v = round(p.base, 2) if p and c == col else None
            w.cell(c, v, F_DATA, border=B_THIN, fmt=MONEY if v is not None else None)

    n_pend = 0
    if filas_modelo:
        modelo_rows(lc, gc_row)
    for port, rs in grupos(lambda r: G_BBULK in r.groups).items():
        w.port(port, lc)
        for r in rs:
            if not r.ok:
                n_pend += 1
                w.pending(r.bl.bl_clean, lc)
                continue
            gc_row(r)
            w.row += 1
    last = w.row - 1
    w.skip(1, last_height=15.75)
    t = w.total_row("TOTAL GRAL CARGO", {c: f"=SUM({_L(c)}{first}:{_L(c)}{last})" for c, _, _ in cols}, lc,
                    fmts={c: MONEY for c, _, _ in cols})
    w.cell(1, "COMMISSION", F_CALB, border=B_LRB)
    for c, _, _ in cols:
        w.cell(c, f"={_L(c)}{t}*${_L(c)}${gr}", F_CALB, border=B_LRB, fmt=MONEY)
    cm = w.row
    w.row += 1
    w.cell(1, "US DOLLARS", F_CALB, border=B_LRT)
    for c, n, _ in cols:
        w.cell(c, f"={_L(c)}{cm}*{roe}" if n == "EURO" else f"={_L(c)}{cm}", F_CALB, border=B_LRT, fmt=MONEY)
    w.height(15.75)
    us = w.row
    w.row += 1
    tot["GC"] = w.total_commission("TOTAL COMMISSION GENERAL CARGO",
                                   "=" + "+".join(f"{_L(c)}{us}" for c, _, _ in cols))
    _pend_line(w, n_pend)
    w.skip(2, last_height=15.75)

    # ── 4. CONTAINERS ─────────────────────────────────────────────────
    amt = [(2, "USD", False), (3, "EURO", False)]
    if expo:
        amt += [(4, "USD", True), (5, "EURO", True)]
    c_ctr = amt[-1][0] + 1
    c_calc, c_min, c_fin = c_ctr + 1, c_ctr + 2, c_ctr + 3
    w.section_title("CONTAINERS COMMISSION", max(7, c_fin))
    w.header({c: ("AMOUNTS BRASIL" if b else "AMOUNTS") for c, _, b in amt}
             | {c_ctr: "CTRS", c_calc: "COMM", c_min: "MINIMUM", c_fin: "FINAL"}, c_fin)
    w.header({1: "B/L"} | {c: n for c, n, _ in amt}
             | {c_calc: "CALCULATION", c_min: "COMM X CONT.", c_fin: "COMMISSION"}, c_fin)
    w.cell(1, None, F_CAL, border=B_THIN, align=CENTER)
    for c, _, b in amt:
        w.cell(c, rates.brasil_containers if b else rates.containers, F_CAL, border=B_THIN,
               align=CENTER, fmt=PCT)
    w.cell(c_ctr, None, F_CAL, border=B_THIN, align=CENTER)
    w.cell(c_calc, "USD", F_CAL, border=B_THIN, align=CENTER)
    w.cell(c_min, rates.container_min_per_unit, F_CAL, border=B_THIN, align=CENTER, fmt=MONEY)
    w.cell(c_fin, "USD", F_CAL, border=B_THIN, align=CENTER)
    cr = w.row
    w.row += 1
    first = w.row

    def ctr_row(r):
        row = w.row
        w.cell(1, r.bl.bl_clean if r else None, F_HDR, border=B_THIN)
        p = r.part(G_CTR) if r else None
        col = None
        if p:
            col = (4 if r.brasil else 2) + (1 if p.currency == "EUR" else 0)
        for c, _, _ in amt:
            v = round(p.base, 2) if p and c == col else None
            w.cell(c, v, F_DATA, border=B_THIN, fmt=MONEY if v is not None else None)
        w.cell(c_ctr, p.min_ctr_units if p else None, F_DATA, border=B_THIN, align=HCENTER)
        calc = "+".join(f"{_L(c)}{row}*${_L(c)}${cr}" + (f"*{roe}" if n == "EURO" else "") for c, n, _ in amt)
        w.cell(c_calc, f"={calc}", F_DATA, border=B_THIN, fmt=MONEY)
        w.cell(c_min, f"={_L(c_ctr)}{row}*${_L(c_min)}${cr}", F_DATA, border=B_THIN, fmt=MONEY)
        w.cell(c_fin, f"=MAX({_L(c_calc)}{row},{_L(c_min)}{row})", F_DATA, border=B_THIN, fmt=MONEY)

    n_pend = 0
    if filas_modelo:
        modelo_rows(c_fin, ctr_row, (c_ctr,))
    for port, rs in grupos(lambda r: G_CTR in r.groups).items():
        w.port(port, c_fin, (c_ctr,))
        for r in rs:
            if not r.ok:
                n_pend += 1
                w.pending(r.bl.bl_clean, c_fin)
                continue
            ctr_row(r)
            w.row += 1
    last = max(w.row - 1, first)
    t = w.total_row("TOTAL CONT", {c: f"=SUM({_L(c)}{first}:{_L(c)}{last})" for c in range(2, c_fin + 1)},
                    c_fin, fmts={c: MONEY for c in range(2, c_fin + 1) if c != c_ctr},
                    aligns={c_ctr: HCENTER})
    tot["CTR"] = w.total_commission("TOTAL COMMISSION CONTAINERS", f"={_L(c_fin)}{t}", label_font=F_TOTAL,
                                    merge=False, value_border=B_MED, label_border=B_MED_LTB)
    ctr_total = f"{_L(c_ctr)}{t}"
    _pend_line(w, n_pend)
    w.skip(2, last_height=15.75)

    # ── 5. TOTALES ────────────────────────────────────────────────────
    w.cell(1, "TOTAL COMMISSION:", F_TOTCOMM, FILL_TOTAL, B_MED, align=HCENTER)
    w.cell(2, None, F_CAL, border=B_MED_RTB)
    w.cell(3, "=" + "+".join(tot.values()), F_BIG, FILL_TOTAL, B_MED, fmt=MONEY)
    w.merge(1, 2)
    w.height(21)
    tc = w.row
    w.row += 1
    w.cell(1, "CONTROL EQUIPMENT FEES PER CNTR:", F_HDR, border=B_MED, align=HCENTER)
    w.cell(2, None, F_CAL, border=B_MED_RTB)
    w.cell(3, config.CONTROL_EQUIPMENT_FEE_USD, F_CALB, border=B_MED, align=HCENTER, fmt=CEF_FMT)
    w.cell(4, f"={ctr_total}*C{w.row}", F_TOTAL, border=B_MED, fmt=MONEY)
    w.merge(1, 2)
    w.height(16.5)
    cef = w.row
    w.row += 1
    w.height(15.75)
    w.row += 1
    w.cell(1, "GRAND TOTAL:", F_BIG, FILL_TOTAL, B_MED, align=HCENTER)
    w.cell(2, None, F_CAL, border=B_MED_RTB)
    w.cell(3, f"=C{tc}+D{cef}", F_BIG, FILL_TOTAL, B_MED, fmt=MONEY)
    w.merge(1, 2)
    w.height(21)
    if pend:
        w.row += 1
        w.cell(1, f"Total NO cerrado: faltan {len(pend)} B/L en revisión.", F_PEND)
    return wb


def _pend_line(w, n):
    """Aviso de pendientes en la primera fila libre debajo de la sección."""
    if n:
        w.cell(1, f"PENDIENTES DE REVISIÓN: {n} B/L (no incluidos en el total)", F_PEND)


def _detail_sheet(wb, results):
    ws = wb.create_sheet("DETALLE B-L")
    headers = ["B/L", "PUERTO", "CARGA", "RODADOS", "ALL IN", "BRASIL", "MONEDA",
               "COMISIONABLE MANIFIESTO", "BASE COMISIÓN", "COMISIÓN USD", "CTRS CEF", "ESTADO", "NOTA / MOTIVO"]
    widths = [16, 20, 34, 12, 8, 8, 8, 16, 16, 14, 9, 11, 80]
    for i, (h, wd) in enumerate(zip(headers, widths), start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = F_HDR
        c.fill = FILL_PORT
        c.border = B_THIN
        c.alignment = CENTER
        ws.column_dimensions["ABCDEFGHIJKLM"[i - 1]].width = wd
    for row, r in enumerate(sorted(results, key=lambda r: (r.port, r.bl.bl_no)), start=2):
        base = sum(p.base for p in r.parts if p.group != G_ROD or p.mode == MODE_AGENCY)
        vals = [r.bl.bl_no, r.port, format_cargo(r.bl) or "", r.rodado_mode or "",
                "SI" if r.all_in else "NO", "SI" if r.brasil else "", r.currency,
                round(r.net, 2), round(base, 2), None if not r.ok else round(r.total_usd, 2),
                r.cef_units or None, "OK" if r.ok else "REVISAR",
                " | ".join(r.reasons + ([r.suggestion] if r.suggestion else []) + r.notes)]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = F_DATA
            c.border = B_THIN
            if col in (8, 9, 10):
                c.number_format = MONEY
            if col == 13:
                c.alignment = Alignment(wrap_text=True, vertical="top")
            if not r.ok:
                c.fill = FILL_PEND
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:M{len(results) + 1}"
