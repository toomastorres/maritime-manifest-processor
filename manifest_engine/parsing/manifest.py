"""
manifest.py - Lectura del CSV de manifiesto -> List[BLRecord].

El CSV es un reporte de ancho fijo envuelto en CSV: cada fila es
record_type,record_text,record_sort_id y el texto viene dividido en columnas
por pipes:

    0: prefijo   1: B/L No.   2: SHIPPER/CONSIGNEE/NOTIFY   3: Marks & Nos.
    4: Description Of Goods   5: Gross Weight   6: Measurement
    7: Charge Information     8: sufijo

Reglas de lectura que evitan los "datos basura":
  - Sólo son B/L válidos los que empiezan con S3. Los P3 (equipo vacío de la
    línea marítima) cierran el B/L anterior y se descartan con todas sus líneas.
  - Un bulto es sólo una línea 'N-Tipo' cuyo Tipo está en la hoja TIPOS DE
    BULTO del catálogo. El resto de las líneas numeradas son descripción.
    Un tipo con forma de bulto que no está en el catálogo se marca para revisión.
  - Los Mafis son equipo: llevan la carga Bbulk encima, no son carga.
  - El nombre de cada cargo incluye su encabezado ('Toll' + 'S/C(POD)'), así
    el S/C del Toll no se confunde con 'Over Hght S/C' o 'Terminal S/C(POL)'.
  - Los importes chicos vienen sin cero inicial ('.00P USD', '-.79P EUR').
"""

import re
from collections import OrderedDict
from typing import List, Optional

from .. import config
from ..catalog import Catalog, load_catalog
from ..domain.models import BLRecord, CargoItem, ChargeLine, TotalLine

# ── patrones ──────────────────────────────────────────────────────────
BL_PATTERN = re.compile(r"(?<![A-Z0-9])S3\d{8}(?![A-Z0-9])")
BL_PATTERN_P = re.compile(r"P3\d{8}")
BL_START_PATTERN = re.compile(r"^\[?S3\d{8}\]?(?:\[T\])?")
BL_GENERIC_PATTERN = re.compile(r"^[A-Z]\d{8}")

# Importe al final de una línea de cargo: '25,860.80P USD', '.00P USD', '-.79P EUR'.
AMOUNT = re.compile(r"(-?[\d,]*\.\d{2})([PC])\s+([A-Z]{3})\s*$")
CHARGE_DETAIL = re.compile(
    r"\s(?P<factor>-?[\d,]*\.\d+)\s+(?P<basis>[A-Z0-9/]+)\s+(?P<rate>-?[\d,]*\.\d+)\s+"
    r"-?[\d,]*\.\d{2}[PC]\s+[A-Z]{3}\s*$")
PCT_OF = re.compile(r"-?[\d.]+%\s*of\s*([\d.,]+)")
TOTAL_PATTERN = re.compile(
    r"total\[(.*?)\]\s+(-?[\d,]*\.\d{2})\s*([PC])\s*([A-Z]{3})", re.IGNORECASE)
BASIC_FRT_NEG = re.compile(
    r"basic\s+frt\.\s+.*?(-[\d,]*\.\d{2})\s*([PC])\s*([A-Z]{3})", re.IGNORECASE)

# Línea de bulto: '27-New Small Van(s)', '1-40 ft. Mafi 130 Tons 2.7 M'.
KIND_LINE = re.compile(r"^(\d+)\s*-\s*(.+?)\s*$")
# Si una línea numerada tiene alguno de estos rasgos y no está en el catálogo,
# es un tipo de bulto nuevo (no descripción): el B/L va a revisión.
KIND_LIKE = re.compile(r"\((?:S|s)\)|\bft\.|\bVan\b|\bCar\(|RoRo|Mafi|UNPACKED|SKID")

# Líneas de la columna de cargos que no son ni cargo ni encabezado de cargo.
_CHARGE_SKIP = ("Charge Information", "DESCRIPTION", "Booking", "Continue")
_NOT_HEADER = re.compile(r"R\.O\.E\.|^(Cross|TransShipment) Booking", re.IGNORECASE)

# VIN de 17 posiciones (el estándar excluye I, O y Q).
VIN_PATTERN = re.compile(r"(?<![A-Z0-9])[A-HJ-NPR-Z0-9]{17}(?![A-Z0-9])")

# Encabezado de página: |ATLANTIC STAR                 :ATL0426          |
SHIP_HEADER = re.compile(r"\|\s*([A-Z][A-Z0-9'. -]*[A-Z0-9])\s*:\s*([A-Z]{3})(\d{4})\s*\|")


def _f(s) -> float:
    return float(str(s).replace(",", ""))


def contar_pipes_y_posicion(texto: str, busqueda: str) -> int:
    for i, parte in enumerate(texto.split("|")):
        if busqueda.lower() in parte.lower():
            return i
    return -1


def extraer_puerto(linea: str, posicion: int) -> Optional[str]:
    partes = linea.split("|")
    if 0 <= posicion < len(partes):
        return partes[posicion].strip()
    return None


def extract_ship_name(filename: str, limit: int = 400) -> Optional[str]:
    """
    Nombre del buque según el encabezado del propio manifiesto.

    Evita depender del mapeo código -> nombre: cuando aparece un buque nuevo,
    el nombre sale igual bien.
    """
    with open(filename, "r", encoding="latin-1") as f:
        for i, line in enumerate(f):
            if i > limit:
                break
            m = SHIP_HEADER.search(line)
            if m:
                return m.group(1).strip()
    return None


# ── nombres de cargo ──────────────────────────────────────────────────

def line_desc(c: str) -> str:
    """Nombre del cargo en una línea con importe: sin importe, factor, base ni tarifa."""
    s = AMOUNT.sub("", c).rstrip()
    pct = re.search(r"\s-?[\d.]+%\s*of\s*[\d.,]+$", s)
    if pct:
        return s[:pct.start()].strip() + " (x% of y)"
    s = re.sub(r"\s+-?[\d,]*\.\d+\s+\S+\s+-?[\d,]*\.\d+$", "", s)
    return s.strip()


def concept_name(header: Optional[str], desc: str) -> str:
    """
    Nombre completo del cargo, tal como figura en la hoja CONCEPTOS DE CARGO.
    Algunos cargos ocupan dos líneas: el encabezado ('Toll', 'ETS+FUEL EU',
    'Over Hght') y la línea con el importe ('S/C(POD)', 'Surcharge', 'S/C').
    """
    if header and (re.search(r"\d+\.\d{3}", header) or header.startswith("Charge Information")):
        header = None
    name = f"{header} {desc}" if header else desc
    # El manifiesto a veces corta la palabra en dos líneas: 'Surcharge(P' + 'OL)'.
    if header:
        pegado = header + desc
        if re.search(r"\((POL|POD)\)", pegado) and not re.search(r"\((POL|POD)\)", name):
            name = pegado
    name = " ".join(name.split())
    if re.fullmatch(r"S/C\(PO[DL]\)", name):
        name = "Toll " + name
    return name


# ── parser ────────────────────────────────────────────────────────────

def parse_manifest(filename: str, op_type: str, catalog: Optional[Catalog] = None,
                   tariffs: Optional[config.LineCostTariffs] = None) -> List[BLRecord]:
    """Parser unificado IMPO/EXPO. Devuelve la lista de B/L válidos (S3...) del manifiesto."""
    catalog = catalog or load_catalog()
    tariffs = tariffs or config.LineCostTariffs()
    is_impo = (op_type.upper() == "IMPO")

    thc_label = "POD" if is_impo else "POL"
    entity_pattern = (re.compile(r"CN:[\-= ]?\s*([^|]+)") if is_impo
                      else re.compile(r"SH[:\-]?\s*([^|]+)"))
    thc_detail = re.compile(
        r"THC \(" + thc_label + r"\)\s+([\d,.]+)\s+(20|40|PU)\s+([\d,.]+)\s+"
        r"([\d,.]+)([PC])\s+([A-Z]{3})")

    all_bls: List[BLRecord] = []
    bl_dict = {} if not is_impo else None
    current: Optional[BLRecord] = None
    active_bl = None
    state = {}   # por B/L: bulto en curso, encabezado de cargo pendiente, último cargo

    port_of_loading = "Nulo"
    port_of_discharge = "Nulo"
    next_line_ports = False
    pos_carga = pos_descarga = -1
    search_entity = False
    montos_neg_frt = 0.0

    def new_bl(bl_number: str) -> BLRecord:
        return BLRecord(
            bl_no=bl_number,
            op_type=op_type.upper(),
            port_of_loading=port_of_loading,
            port_of_discharge=port_of_discharge,
        )

    def close_bl_impo():
        nonlocal current, montos_neg_frt
        if current:
            current.negative_basic_frt = montos_neg_frt
            all_bls.append(current)
            current = None
            montos_neg_frt = 0.0

    with open(filename, "r", encoding="latin-1") as f:
        for line in f:
            # ── puertos del encabezado de página ──────────────────────
            if next_line_ports:
                pc = extraer_puerto(line, pos_carga)
                pd_ = extraer_puerto(line, pos_descarga)
                if pc:
                    port_of_loading = pc
                if pd_:
                    port_of_discharge = pd_
                next_line_ports = False
                continue

            low = line.lower()
            if "port of loading" in low and "port of discharge" in low:
                pos_carga = contar_pipes_y_posicion(line, "port of loading")
                pos_descarga = contar_pipes_y_posicion(line, "port of discharge")
                next_line_ports = True
                continue

            parts = line.split("|")

            if is_impo:
                if search_entity and current:
                    m = entity_pattern.search(line)
                    if m:
                        current.entity = m.group(1).strip()
                        search_entity = False

                if BL_PATTERN_P.search(line):
                    close_bl_impo()
                    search_entity = False
                    continue

                bl_match = BL_PATTERN.search(line)
                if bl_match:
                    close_bl_impo()
                    current = new_bl(bl_match.group())
                    montos_neg_frt = 0.0
                    m = entity_pattern.search(line)
                    if m:
                        current.entity = m.group(1).strip()
                        search_entity = False
                    else:
                        search_entity = True

                entry = current
            else:
                is_new_bl = False
                if len(parts) > 1:
                    col1 = parts[1].strip()
                    if BL_START_PATTERN.match(col1):
                        raw_bl = re.sub(r"^\[?(S3\d{8}(?:\[T\])?)\]?$", r"\1", col1)
                        active_bl = raw_bl.replace("[T]", "")
                        is_new_bl = True
                        if active_bl not in bl_dict:
                            bl_dict[active_bl] = new_bl(raw_bl)
                        m = entity_pattern.search(line)
                        if m:
                            bl_dict[active_bl].entity = m.group(1).strip()
                            search_entity = False
                        else:
                            search_entity = True
                    elif BL_PATTERN_P.match(col1.strip("[]")) or BL_GENERIC_PATTERN.match(col1.strip("[]")):
                        active_bl = None
                        search_entity = False
                        continue

                if not active_bl:
                    continue

                entry = bl_dict[active_bl]

                if search_entity and not is_new_bl:
                    m = entity_pattern.search(line)
                    if m:
                        entry.entity = m.group(1).strip()
                    search_entity = False

            if entry is None:
                continue
            st = state.setdefault(id(entry), {"item": None, "hdr": None, "last": None})

            # ── shipper / consignee / notify (columna 2) ──────────────
            party = parts[2].strip() if len(parts) > 2 else ""
            if party:
                entry.party_lines.append(party)

            # ── descripción y bultos (columna 4) ──────────────────────
            desc_field = parts[4].strip() if len(parts) > 4 else ""
            if desc_field:
                entry.description_lines.append(desc_field)
                for vin in VIN_PATTERN.findall(desc_field):
                    if vin not in entry.vins:
                        entry.vins.append(vin)
                _read_kind(entry, desc_field, catalog, st)

            # ── medidas del bulto en curso (columnas 5 y 6) ───────────
            if len(parts) > 6:
                _read_measures(st["item"], parts[5].strip(), parts[6].strip())

            # ── peso bruto (columna 5), salteando líneas de TARA ──────
            if len(parts) > 5 and "tare" not in desc_field.lower():
                weight_raw = re.sub(r"[KkGgSs\s]+$", "", parts[5].strip()).strip()
                weight_raw = weight_raw.replace(",", "")
                try:
                    w = float(weight_raw)
                    if w > 0:
                        entry.weight += w
                except (ValueError, TypeError):
                    pass

            # ── THC discriminado por 20 y 40 pies ─────────────────────
            thc_m = thc_detail.search(line)
            if thc_m:
                try:
                    unit = thc_m.group(2)
                    rate = _f(thc_m.group(3))
                    amount = _f(thc_m.group(4))
                    target = None
                    if unit == "20":
                        target = entry.thc_20
                    elif unit == "40":
                        target = entry.thc_40
                    elif unit == "PU":
                        # Sólo contenedores facturados por unidad que coinciden con
                        # la tarifa THC 20'/40' cuando el B/L trae contenedores.
                        if entry.cargo_items and any(it.family == "CONTENEDOR" for it in entry.cargo_items):
                            if abs(rate - tariffs.thc_20) < 5:
                                target = entry.thc_20
                            elif abs(rate - tariffs.thc_40) < 5:
                                target = entry.thc_40
                    if target is not None:
                        target.monto += amount
                        target.letra = thc_m.group(5)
                        target.moneda = thc_m.group(6)
                except ValueError:
                    pass

            # ── totales por puerto ────────────────────────────────────
            total_m = TOTAL_PATTERN.search(line)
            if total_m:
                pais = total_m.group(1).strip().lower()
                monto = _f(total_m.group(2))
                cond = total_m.group(3).upper()
                moneda = total_m.group(4).upper()
                entry.total_lines.append(
                    TotalLine(" ".join(total_m.group(1).split()), monto, cond, moneda))
                # 'Buenos Aires - Partner' (booking de subagente local) también se paga
                # en Buenos Aires: es el mismo total a cobrar acá.
                if pais.startswith("buenos aires"):
                    pais = "buenos aires"
                if is_impo:
                    if pais == "buenos aires" and cond == "C":
                        if moneda == "USD":
                            entry.totals["Total Buenos Aires USD"] = monto
                        elif moneda == "EUR":
                            entry.totals["Total Buenos Aires EUR"] = monto
                else:
                    if cond == "P":
                        if pais == "buenos aires":
                            entry.totals["Total Buenos Aires Monto"] = monto
                            entry.totals["Total Buenos Aires Moneda"] = moneda
                        else:
                            entry.totals["Total Casa Matriz Monto"] = monto
                            entry.totals["Total Casa Matriz Moneda"] = moneda
                    elif cond == "C":
                        entry.collect.append(
                            {"Pais": pais, "Monto": monto, "Moneda": moneda})

            # ── Basic Frt. negativo ───────────────────────────────────
            if "basic frt." in low:
                neg_m = BASIC_FRT_NEG.search(line)
                if neg_m:
                    try:
                        val = _f(neg_m.group(1))
                        if is_impo:
                            montos_neg_frt += val
                        else:
                            entry.negative_basic_frt += val
                    except ValueError:
                        pass

            # ── columna de cargos ─────────────────────────────────────
            if len(parts) >= 8:
                _read_charge_column(entry, parts[-2].strip(), st, is_impo)

    if is_impo:
        close_bl_impo()
    records = consolidar(all_bls if is_impo else list(bl_dict.values()))
    for r in records:
        r.rebuild_cargo()
        r.party_id = extract_party_id(r.party_lines, "CN" if is_impo else "SH")
    return records


# Inicio de cada parte en la columna SHIPPER/CONSIGNEE/NOTIFY: 'SH:', 'CN:', 'NO[1]:'.
PARTY_START = re.compile(r"^(SH|CN|NO(?:\[\d+\])?)\s*[:\-=]")
CUIT = re.compile(r"C\.?U\.?I\.?T\.?\D{0,12}?(?<!\d)(\d{2})[-\s.]?(\d{8})[-\s.]?(\d)(?!\d)")
CUIT_WEIGHTS = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


def cuit_ok(digits: str) -> bool:
    """Dígito verificador de la CUIT (módulo 11): descarta números mal leídos."""
    if len(digits) != 11 or not digits.isdigit():
        return False
    resto = 11 - sum(int(d) * w for d, w in zip(digits, CUIT_WEIGHTS)) % 11
    dv = 0 if resto == 11 else resto
    return dv != 10 and dv == int(digits[10])


def extract_party_id(lines: List[str], role: str) -> str:
    """
    CUIT de la parte pedida ('CN' en IMPO, 'SH' en EXPO), si el manifiesto lo
    trae dentro de su bloque. El número puede venir partido en dos renglones
    ('... / CUIT' y abajo '30711111117'). Sólo se toma si el dígito verificador
    es correcto; si no figura o no verifica, queda vacío.
    """
    bloque, dentro = [], False
    for ln in lines:
        m = PARTY_START.match(ln)
        if m:
            if dentro:
                break
            dentro = m.group(1) == role
        if dentro:
            bloque.append(ln)
    for m in CUIT.finditer(" ".join(bloque).upper()):
        if cuit_ok(m.group(1) + m.group(2) + m.group(3)):
            return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    return ""


def consolidar(records: List[BLRecord]) -> List[BLRecord]:
    """Un B/L = un registro: junta los bloques repetidos del mismo número."""
    out = OrderedDict()
    for r in records:
        existente = out.get(r.bl_no)
        if existente is None:
            out[r.bl_no] = r
        else:
            existente.merge(r)
    return list(out.values())


def _read_kind(entry: BLRecord, desc_field: str, catalog: Catalog, st: dict) -> None:
    """Una línea 'N-Tipo' es un bulto sólo si el Tipo está en el catálogo."""
    m = KIND_LINE.match(desc_field)
    if not m:
        return
    kind = " ".join(m.group(2).split())
    rule = catalog.kind(kind)
    if rule is not None:
        item = CargoItem(qty=int(m.group(1)), kind=kind, family=rule.family,
                         equipment=rule.is_equipment)
        entry.items.append(item)
        st["item"] = item
    elif KIND_LIKE.search(kind):
        if kind not in entry.unknown_kinds:
            entry.unknown_kinds.append(kind)
        st["item"] = None


def _read_measures(item: Optional[CargoItem], weight: str, measure: str) -> None:
    if item is None:
        return
    wm = re.match(r"^([\d,]+\.\d+)", weight)
    if wm and item.tn is None:
        item.tn = _f(wm.group(1)) / 1000.0
    cm = re.match(r"^([\d,]+\.\d+)\s*CBM", measure)
    if cm and item.cbm is None:
        item.cbm = _f(cm.group(1))
    lm = re.match(r"^([\d,]+\.\d+)\s*LM", measure)
    if lm and item.lm is None:
        item.lm = _f(lm.group(1))


def _read_charge_column(entry: BLRecord, c: str, st: dict, is_impo: bool) -> None:
    """
    Recorre la columna 'Charge Information'. Las líneas sin importe son el
    encabezado del cargo siguiente; 'Tariff ID' identifica la cotización del
    cargo anterior; los separadores y los totales cortan el encabezado.
    """
    if not c or c.startswith(_CHARGE_SKIP):
        return
    if c.startswith("Tariff ID"):
        last = st["last"]
        if last is not None and last.tariff is None:
            last.tariff = c.split(":", 1)[1].strip()
        return
    if c.startswith("---") or c.startswith("Total["):
        st["hdr"] = None
        return
    if AMOUNT.search(c):
        st["last"] = _parse_charge(entry, c, st["hdr"], is_impo)
        st["hdr"] = None
        return
    st["hdr"] = None if _NOT_HEADER.search(c) else c


def _parse_charge(entry: BLRecord, c: str, header: Optional[str], is_impo: bool) -> ChargeLine:
    m = AMOUNT.search(c)
    amount = _f(m.group(1))
    letra, moneda = m.group(2), m.group(3)
    desc = line_desc(c)
    concept = concept_name(header, desc)
    desc_raw = desc.lower().replace(".", "")

    factor = basis = rate = None
    dm = CHARGE_DETAIL.search(c)
    if dm:
        try:
            factor = _f(dm.group("factor"))
            basis = dm.group("basis")
            rate = _f(dm.group("rate"))
        except ValueError:
            factor = basis = rate = None
    pct_of = None
    pm = PCT_OF.search(c)
    if pm:
        try:
            pct_of = _f(pm.group(1).rstrip("."))
        except ValueError:
            pct_of = None

    cl = ChargeLine(desc=desc_raw, total=amount, letra=letra, moneda=moneda,
                    factor=factor, basis=basis, rate=rate, concept=concept, pct_of=pct_of)
    entry.charge_lines.append(cl)

    # Columnas de gastos de línea de las biblias (TOLL, THC, SWP, BAF, ENS).
    toll_name = "Toll S/C(POD)" if is_impo else "Toll S/C(POL)"
    mapped_key = None
    if "basic frt" in desc_raw:
        mapped_key = "Basic FRT"
    elif not is_impo and desc_raw == "ens":
        mapped_key = "ENS"
    elif ("thc (pod)" if is_impo else "thc (pol)") in desc_raw:
        mapped_key = "THC"
    elif concept == toll_name:
        mapped_key = "Toll"
    elif "sweeping" in desc_raw:
        mapped_key = "Sweeping"
    elif is_impo and "baf" in desc_raw:
        mapped_key = "BAF"

    if mapped_key:
        monto = config.ENS_FIXED_AMOUNT if mapped_key == "ENS" else amount
        entry.set_charge(mapped_key, monto, letra, moneda)
        # Tarifa del S/C (TOLL) en carga suelta: se factura por tonelada (MT).
        if mapped_key == "Toll" and basis == "MT" and rate:
            entry.sc_loose_rates.append(rate)

    if amount < 0:
        for p in config.NEGATIVE_PREFIXES:
            if desc_raw == p or desc_raw.startswith(p + " "):
                entry.negative_adjustments += amount
                break
    return cl
