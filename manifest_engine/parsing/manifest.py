"""
manifest.py - Parser del CSV de manifiesto marítimo.

Migra parse_manifest() de manifest_processor_v3.py a una clase ManifestParser
que produce List[BLRecord]. Conserva la lectura por líneas (decisión 4 del
plan: reforzar sólo si los datos lo requieren) y la lógica IMPO/EXPO unificada.
"""

import re
from typing import Dict, List, Optional

from .. import config
from ..domain.models import BLRecord, ChargeLine
from ..domain.cargo import classify_description


# ═══════════════════════════════════════════════════════════════════════
# PATRONES (migrados del script original)
# ═══════════════════════════════════════════════════════════════════════

BL_PATTERN = re.compile(r"(?<![A-Z0-9])S3\d{8}(?![A-Z0-9])")
BL_PATTERN_P = re.compile(r"P3\d{8}")
BL_START_PATTERN = re.compile(r"^\[?S3\d{8}\]?(?:\[T\])?")
BL_GENERIC_PATTERN = re.compile(r"^[A-Z]\d{8}")
CHARGE_PATTERN = re.compile(r"^\s*(.*?)\s{2,}.*?(-?[\d,]+\.\d{2})([PC])\s+([A-Z]{3})$")
# Línea de cargo estructurada: DESCRIPTION FACTOR BASIS RATE TOTAL[P/C] CURR.
# FACTOR admite ".384" (sin cero inicial) y "109.060"; BASIS es 20/40/MT/PU/...
STRUCTURED_CHARGE = re.compile(
    r"^\s*(?P<desc>.+?)\s+(?P<factor>[\d,]*\.\d+)\s+(?P<basis>[A-Z0-9]{1,4})\s+"
    r"(?P<rate>-?[\d,]+\.\d{2})\s+(?P<total>-?[\d,]+\.\d{2})(?P<cond>[PC])\s+(?P<curr>[A-Z]{3})\s*$"
)
TOTAL_PATTERN = re.compile(
    r"total\[(.*?)\]\s+(-?[\d,]+\.\d{2})\s*([PC])\s*([A-Z]{3})", re.IGNORECASE
)
BASIC_FRT_NEG = re.compile(
    r"basic\s+frt\.\s+.*?(-[\d,]+\.\d{2})\s*([PC])\s*([A-Z]{3})", re.IGNORECASE
)
# VIN/chasis: token de 17 caracteres alfanuméricos contiguos (col Unit No. y
# Description). Se exige >=1 letra y >=1 dígito para no capturar números o
# códigos sueltos. Valida con alta precisión contra los manifiestos reales.
VIN_PATTERN = re.compile(r"(?<![A-Z0-9])([A-Z0-9]{17})(?![A-Z0-9])")


def _looks_like_vin(tok: str) -> bool:
    return any(c.isalpha() for c in tok) and any(c.isdigit() for c in tok)


def contar_pipes_y_posicion(texto: str, busqueda: str) -> int:
    for i, parte in enumerate(texto.split("|")):
        if busqueda.lower() in parte.lower():
            return i
    return -1


def extraer_puerto(linea: str, posicion: int) -> Optional[str]:
    partes = linea.split("|")
    if 0 <= posicion < len(partes):
        p = partes[posicion].strip()
        # Limpiar basura de truncado del origen (ej. 'CARTAGENA CO)' -> 'CARTAGENA CO')
        p = p.rstrip(") (").strip()
        return p
    return None


class ManifestParser:
    """Parser unificado IMPO/EXPO. Uso: ManifestParser(op_type).parse(path)."""

    def __init__(self, op_type: str):
        self.op_type = op_type.upper()
        self.is_impo = self.op_type == "IMPO"

        thc_label = "POD" if self.is_impo else "POL"
        self.entity_pattern = (
            re.compile(r"CN:[\-= ]?\s*([^|]+)") if self.is_impo
            else re.compile(r"SH[:\-]?\s*([^|]+)")
        )
        self.thc_detail = re.compile(
            rf"THC \({thc_label}\)\s+([\d,.]+)\s+(20|40|PU)\s+([\d,.]+)\s+([\d,.]+)([PC])\s+([A-Z]{{3}})"
        )
        # Estado para descripciones de cargo "envueltas" (ej. "Over Hght" en una
        # línea y " S/C  3.000 40 ... 3,000.00P EUR" en la siguiente).
        self._pending_extra = None

        if self.is_impo:
            self.charge_key_map = {
                "basic frt": "Basic FRT", "thc (pod)": "THC", "s/c(pod)": "Toll",
                "sweeping": "Sweeping", "baf": "BAF",
            }
        else:
            self.charge_key_map = {
                "basic frt": "Basic FRT", "ens": "ENS", "thc (pol)": "THC",
                "s/c(pol)": "Toll", "sweeping": "Sweeping",
            }

    # ------------------------------------------------------------------
    def parse(self, filename: str) -> List[BLRecord]:
        if self.is_impo:
            return self._parse_impo(filename)
        return self._parse_expo(filename)

    # ------------------------------------------------------------------
    def _new_bl(self, bl_number: str, pol: str, pod: str) -> BLRecord:
        return BLRecord(bl_no=bl_number, port_of_loading=pol, port_of_discharge=pod)

    # ------------------------------------------------------------------
    def _parse_impo(self, filename: str) -> List[BLRecord]:
        all_bls: List[BLRecord] = []
        current: Optional[BLRecord] = None
        pol = pod = "Nulo"
        next_line_ports = False
        pos_carga = pos_descarga = -1
        search_entity = False
        montos_neg_frt = 0.0

        def close_bl():
            nonlocal current, montos_neg_frt
            if current:
                current.montos_negativos_basic_frt = montos_neg_frt
                all_bls.append(current)
                current = None
                montos_neg_frt = 0.0

        with open(filename, "r", encoding="latin-1") as f:
            for line in f:
                if next_line_ports:
                    pc = extraer_puerto(line, pos_carga)
                    pd_ = extraer_puerto(line, pos_descarga)
                    if pc:
                        pol = pc
                    if pd_:
                        pod = pd_
                    next_line_ports = False
                    continue

                if "port of loading" in line.lower() and "port of discharge" in line.lower():
                    pos_carga = contar_pipes_y_posicion(line, "port of loading")
                    pos_descarga = contar_pipes_y_posicion(line, "port of discharge")
                    next_line_ports = True
                    continue

                # Entity en línea siguiente
                if search_entity and current:
                    m = self.entity_pattern.search(line)
                    if m:
                        current.entity = m.group(1).strip()
                        search_entity = False

                # P-BL cierra el BL actual
                if BL_PATTERN_P.search(line) and current:
                    close_bl()
                    search_entity = False
                    continue

                bl_match = BL_PATTERN.search(line)
                if bl_match:
                    close_bl()
                    current = self._new_bl(bl_match.group(), pol, pod)
                    montos_neg_frt = 0.0
                    self._pending_extra = None
                    m = self.entity_pattern.search(line)
                    if m:
                        current.entity = m.group(1).strip()
                        search_entity = False
                    else:
                        search_entity = True

                if current is None:
                    continue

                montos_neg_frt += self._consume_line(line, current)

        if current:
            current.montos_negativos_basic_frt = montos_neg_frt
            all_bls.append(current)
        return consolidate_bls(all_bls)

    # ------------------------------------------------------------------
    def _parse_expo(self, filename: str) -> List[BLRecord]:
        bl_dict: Dict[str, BLRecord] = {}
        active_bl: Optional[str] = None
        pol = pod = "Nulo"
        next_line_ports = False
        pos_carga = pos_descarga = -1
        search_entity = False

        with open(filename, "r", encoding="latin-1") as f:
            for line in f:
                if next_line_ports:
                    pc = extraer_puerto(line, pos_carga)
                    pd_ = extraer_puerto(line, pos_descarga)
                    if pc:
                        pol = pc
                    if pd_:
                        pod = pd_
                    next_line_ports = False

                if "port of loading" in line.lower() and "port of discharge" in line.lower():
                    pos_carga = contar_pipes_y_posicion(line, "port of loading")
                    pos_descarga = contar_pipes_y_posicion(line, "port of discharge")
                    next_line_ports = True
                    continue

                parts = line.split("|")
                is_new_bl = False
                if len(parts) > 1:
                    col1 = parts[1].strip()
                    if BL_START_PATTERN.match(col1):
                        raw_bl = re.sub(r"^\[?(S3\d{8}(?:\[T\])?)\]?$", r"\1", col1)
                        clean_bl = raw_bl.replace("[T]", "")
                        active_bl = clean_bl
                        is_new_bl = True
                        self._pending_extra = None
                        if active_bl not in bl_dict:
                            bl_dict[active_bl] = self._new_bl(raw_bl, pol, pod)
                        m = self.entity_pattern.search(line)
                        if m:
                            bl_dict[active_bl].entity = m.group(1).strip()
                            search_entity = False
                        else:
                            search_entity = True
                    elif BL_GENERIC_PATTERN.match(col1):
                        active_bl = None
                        search_entity = False
                        continue

                if not active_bl:
                    continue

                entry = bl_dict[active_bl]

                if search_entity and not is_new_bl:
                    m = self.entity_pattern.search(line)
                    if m:
                        entry.entity = m.group(1).strip()
                    search_entity = False

                neg = self._consume_line(line, entry)
                entry.montos_negativos_basic_frt += neg

        return list(bl_dict.values())

    # ------------------------------------------------------------------
    def _consume_line(self, line: str, entry: BLRecord) -> float:
        """
        Procesa una línea de datos para `entry` (carga, peso, THC, totales,
        cargos). Devuelve el monto negativo de Basic FRT detectado en la línea
        (que el caller acumula según IMPO/EXPO).
        """
        parts = line.split("|")
        neg_basic_frt = 0.0

        # --- Captura de líneas de descripción (col 4) para extraer marca ---
        if len(parts) > 4:
            d4 = parts[4].strip()
            if d4 and "TARE" not in d4.upper():
                entry.description_lines.append(d4)

        # --- Captura de VIN/chasis (col Unit No. y col Description) ---
        # La ubicación del VIN varía por marca (algunas en col 3, otras en col 4),
        # pero el token de 17 chars es estable; se escanean ambas columnas.
        for ci in (3, 4):
            if len(parts) > ci:
                for tok in VIN_PATTERN.findall(parts[ci]):
                    if _looks_like_vin(tok):
                        entry.add_vin(tok)

        # --- Clasificación de carga ---
        if len(parts) > 4:
            desc_field = parts[4].strip()
            if self.is_impo:
                cargo_m = re.match(r"^(\d+)\s*-\s*(.+)$", desc_field)
            else:
                cargo_m = re.match(r"^(\d+)[\s-]*(.+)$", desc_field)
            if cargo_m:
                try:
                    qty = int(cargo_m.group(1))
                    desc = cargo_m.group(2).strip()
                    classify_description(entry.cargo, qty, desc)
                except ValueError:
                    pass

        # --- Peso (col5), salteando TARE ---
        if len(parts) > 5:
            desc4 = parts[4].strip().lower() if len(parts) > 4 else ""
            if "tare" not in desc4:
                weight_raw = parts[5].strip()
                weight_raw = re.sub(r"[KkGgSs\s]+$", "", weight_raw).strip()
                weight_raw = weight_raw.replace(",", "")
                try:
                    w = float(weight_raw)
                    if w > 0:
                        entry.weight += w
                except (ValueError, TypeError):
                    pass

        # --- THC detalle (20 vs 40) ---
        thc_m = self.thc_detail.search(line)
        if thc_m:
            try:
                unit = thc_m.group(2)
                rate = float(thc_m.group(3).replace(",", ""))
                amount = float(thc_m.group(4).replace(",", ""))
                cond = thc_m.group(5)
                curr = thc_m.group(6)
                target = None
                if unit == "20":
                    target = entry.thc_20
                elif unit == "40":
                    target = entry.thc_40
                elif unit == "PU":
                    if abs(rate - self.tariff_thc_20) < 5:
                        target = entry.thc_20
                    elif abs(rate - self.tariff_thc_40) < 5:
                        target = entry.thc_40
                if target is not None:
                    target.add(amount, cond, curr)
            except ValueError:
                pass

        # --- Totales ---
        total_m = TOTAL_PATTERN.search(line)
        if total_m:
            pais = total_m.group(1).strip().lower()
            monto = float(total_m.group(2).replace(",", ""))
            cond = total_m.group(3).upper()
            moneda = total_m.group(4).upper()
            if self.is_impo:
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
                    elif pais == "matriz":
                        entry.totals["Total Matriz Monto"] = monto
                        entry.totals["Total Matriz Moneda"] = moneda
                elif cond == "C":
                    entry.collect.append({"Pais": pais, "Monto": monto, "Moneda": moneda})

        # --- Basic FRT negativo ---
        if "basic frt." in line.lower():
            neg_m = BASIC_FRT_NEG.search(line)
            if neg_m:
                try:
                    neg_basic_frt += float(neg_m.group(1).replace(",", ""))
                except ValueError:
                    pass

        # --- Descripción de cargo "envuelta": una línea sólo de texto cuyo
        # contenido es un comisionable extra (ej. "Over Hght"); el valor llega
        # en la línea siguiente. Se recuerda para fusionar la descripción.
        bare = parts[-2].strip() if len(parts) >= 2 else ""
        if bare and not STRUCTURED_CHARGE.match(bare):
            low_bare = bare.lower().rstrip(".")
            if any(low_bare.startswith(p) for p in config.COMMISSIONABLE_EXTRA):
                self._pending_extra = low_bare

        # --- Cargos ---
        if any(c in line for c in ("USD", "EUR", "BRL", "GBP")):
            charge_parts = line.split("|")
            if len(charge_parts) >= 2:
                charge_info = charge_parts[-2].strip()

                # Captura estructurada (FACTOR/BASIS/RATE) — aditiva.
                sc = STRUCTURED_CHARGE.match(charge_info)
                if sc:
                    try:
                        desc = sc.group("desc").strip().lower().rstrip(".")
                        # Si venía una descripción comisionable envuelta, la
                        # anteponemos (ej. "over hght" + "s/c" -> "over hght s/c").
                        if self._pending_extra:
                            desc = self._pending_extra
                            self._pending_extra = None
                        entry.charge_lines.append(ChargeLine(
                            desc=desc,
                            factor=float(sc.group("factor").replace(",", "")),
                            basis=sc.group("basis"),
                            rate=float(sc.group("rate").replace(",", "")),
                            total=float(sc.group("total").replace(",", "")),
                            letra=sc.group("cond"),
                            moneda=sc.group("curr"),
                        ))
                    except ValueError:
                        pass

                cm = CHARGE_PATTERN.search(charge_info)
                if cm:
                    desc_raw = cm.group(1).strip().lower().replace(".", "")
                    amount_str = cm.group(2).replace(",", "")
                    letra = cm.group(3)
                    moneda = cm.group(4)
                    try:
                        amount = float(amount_str)
                        mapped_key = None
                        for wl_key, int_key in self.charge_key_map.items():
                            if wl_key in desc_raw or desc_raw == wl_key:
                                mapped_key = int_key
                                break
                        if mapped_key:
                            if mapped_key == "ENS":
                                amount = config.ENS_FIXED_AMOUNT
                            entry.charge(mapped_key).add(amount, letra, moneda)
                        if amount < 0:
                            for p in config.NEGATIVE_PREFIXES:
                                if desc_raw == p or desc_raw.startswith(p + " "):
                                    entry.total_ajustes_negativos += amount
                                    break
                    except ValueError:
                        pass

        return neg_basic_frt

    # --- tarifas usadas en desambiguación THC PU ---
    @property
    def tariff_thc_20(self) -> float:
        return config.LineCostTariffs().thc_20

    @property
    def tariff_thc_40(self) -> float:
        return config.LineCostTariffs().thc_40


def _merge_charge(dst, src) -> None:
    dst.monto += src.monto
    if src.letra is not None:
        dst.letra = src.letra
    if src.moneda is not None:
        dst.moneda = src.moneda


def consolidate_bls(bls: List[BLRecord]) -> List[BLRecord]:
    """
    Consolida entradas con el mismo B/L (BL partido en varias páginas del
    manifiesto) en un único registro, conservando el orden de primera
    aparición. La verdad exacta muestra cada BL una sola vez.
    """
    merged: dict = {}
    order: List[str] = []
    for bl in bls:
        if bl.bl_no not in merged:
            merged[bl.bl_no] = bl
            order.append(bl.bl_no)
            continue
        dst = merged[bl.bl_no]
        for key, ch in bl.charges.items():
            _merge_charge(dst.charge(key), ch)
        _merge_charge(dst.thc_20, bl.thc_20)
        _merge_charge(dst.thc_40, bl.thc_40)
        dst.total_ajustes_negativos += bl.total_ajustes_negativos
        dst.montos_negativos_basic_frt += bl.montos_negativos_basic_frt
        for k, v in bl.cargo.containers.items():
            dst.cargo.add_container(k, v)
        for k, v in bl.cargo.vehicles.items():
            dst.cargo.add_vehicle(k, v)
        for k, v in bl.cargo.general.items():
            dst.cargo.add_general(k, v)
        dst.charge_lines.extend(bl.charge_lines)
        for vin in bl.vins:
            dst.add_vin(vin)
        dst.collect.extend(bl.collect)
        dst.totals.update(bl.totals)
        dst.weight += bl.weight
        if dst.entity in ("Nulo", "") and bl.entity not in ("Nulo", ""):
            dst.entity = bl.entity
    return [merged[k] for k in order]


def parse_manifest(filename: str, op_type: str) -> List[BLRecord]:
    """Helper compatible con la firma del script original."""
    return ManifestParser(op_type).parse(filename)
