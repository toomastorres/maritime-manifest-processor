"""
biblia.py - Reglas de las "Biblias" (prepaid/collect y formato de celdas).

Migra format_val, calcular_totales_impo y calcular_totales_expo del script
original, operando sobre BLRecord. Mantiene el comportamiento (verificado por
paridad) y centraliza las reglas de prepaid/collect IMPO/EXPO.
"""

from collections import defaultdict
from typing import Optional

from ..domain.models import BLRecord


def format_val(val, condition: Optional[str] = None, currency: Optional[str] = None,
               item_type: Optional[str] = None):
    """Formato de celda para importes (migrado de v3)."""
    if not val or val == 0:
        return ""
    if condition == "P" and item_type in ("THC", "SWP", "SC"):
        return "Prepaid"
    if currency and currency in ("EUR", "BRL", "GBP"):
        val_str = f"{val:,.2f}" if isinstance(val, (int, float)) else str(val)
        return f"{currency} {val_str}"
    if item_type in ("BAF", "COMM") and currency:
        val_str = f"{val:,.2f}" if isinstance(val, (int, float)) else str(val)
        return f"{val_str} {currency}"
    if isinstance(val, (int, float)):
        return float(val)
    return str(val)


def calcular_totales_impo(bl: BLRecord) -> dict:
    """Columnas DOLARES/EURO PREPAID/COLLECT + BAF + COMM (IMPO)."""
    basic = bl.charge("Basic FRT")
    basic_monto = basic.monto
    basic_letra = basic.letra
    basic_moneda = basic.moneda
    ajustes = bl.total_ajustes_negativos

    thc_monto = bl.charge("THC").monto
    thc_letra = bl.charge("THC").letra
    swp_monto = bl.charge("Sweeping").monto
    swp_letra = bl.charge("Sweeping").letra
    toll_monto = bl.charge("Toll").monto
    toll_letra = bl.charge("Toll").letra
    baf_monto = bl.charge("BAF").monto
    baf_letra = bl.charge("BAF").letra
    baf_moneda = bl.charge("BAF").moneda

    total_ba_usd = bl.totals.get("Total Buenos Aires USD", 0) or 0
    total_ba_eur = bl.totals.get("Total Buenos Aires EUR", 0) or 0

    dolares_prepaid = dolares_collect = euros_prepaid = euros_collect = baf_col = comm = ""
    basic_neto = basic_monto + ajustes

    if basic_letra == "P":
        res = max(basic_neto, 0)
        if basic_moneda == "USD":
            dolares_prepaid = format_val(res)
        elif basic_moneda == "EUR":
            euros_prepaid = format_val(res)
    elif basic_letra == "C":
        comm = format_val(basic_neto, item_type="COMM", currency=basic_moneda)
        if basic_moneda == "USD":
            ded = baf_monto
            if thc_letra == "C":
                ded += thc_monto
            if swp_letra == "C":
                ded += swp_monto
            if toll_letra == "C":
                ded += toll_monto
            dolares_collect = format_val(total_ba_usd - ded)
        elif basic_moneda == "EUR":
            euros_collect = format_val(total_ba_eur - baf_monto)

    if baf_letra == "C":
        baf_col = format_val(baf_monto, item_type="BAF", currency=baf_moneda)

    return {
        "dolares_prepaid": dolares_prepaid, "dolares_collect": dolares_collect,
        "euros_prepaid": euros_prepaid, "euros_collect": euros_collect,
        "baf": baf_col, "comm": comm,
    }


def calcular_totales_expo(bl: BLRecord) -> dict:
    """Columnas PREPAID/COLLECT/ABROAD + comisionable (EXPO)."""
    basic_monto = bl.charge("Basic FRT").monto
    basic_moneda = bl.charge("Basic FRT").moneda
    negativos = bl.total_ajustes_negativos
    comisionable_val = max(0, basic_monto + negativos)

    thc_monto = bl.charge("THC").monto
    thc_letra = bl.charge("THC").letra
    swp_monto = bl.charge("Sweeping").monto
    swp_letra = bl.charge("Sweeping").letra
    toll_monto = bl.charge("Toll").monto
    toll_letra = bl.charge("Toll").letra

    sum_p = 0
    if thc_letra == "P":
        sum_p += thc_monto
    if swp_letra == "P":
        sum_p += swp_monto
    if toll_letra == "P":
        sum_p += toll_monto

    total_ba = bl.totals.get("Total Buenos Aires Monto", 0) or 0
    is_matriz = sum_p > total_ba + 1.0

    thc_str = format_val(thc_monto, thc_letra, item_type="THC")
    swp_str = format_val(swp_monto, swp_letra, item_type="SWP")

    if thc_letra == "C":
        thc_str = "Collect"
    elif is_matriz and thc_letra == "P" and thc_monto > 0:
        thc_str = "Matriz"
    if swp_letra == "C":
        swp_str = "Collect"
    elif is_matriz and swp_letra == "P" and swp_monto > 0:
        swp_str = "Matriz"

    collect_by_curr = defaultdict(float)
    for item in bl.collect:
        collect_by_curr[item.get("Moneda", "USD")] += item.get("Monto", 0)
    collect_parts = [f"{m} {v:,.2f}" for m, v in sorted(collect_by_curr.items())]
    collect_str = " + ".join(collect_parts) if collect_parts else ""

    thc_p = thc_monto if thc_letra == "P" else 0
    swp_p = swp_monto if swp_letra == "P" else 0
    prepaid = total_ba - thc_p - swp_p
    if prepaid == 0 and total_ba == 0:
        prepaid = ""
    elif prepaid < 0:
        prepaid = total_ba
        if thc_letra == "P" and thc_monto > 0:
            thc_str = "Matriz"
        if swp_letra == "P" and swp_monto > 0:
            swp_str = "Matriz"

    abroad = bl.totals.get("Total Matriz Monto", "")
    abroad = float(abroad) if abroad else ""

    comisionable = format_val(comisionable_val, currency=basic_moneda)

    return {
        "comisionable": comisionable, "comisionable_val": comisionable_val,
        "comisionable_moneda": basic_moneda,
        "thc_str": thc_str, "swp_str": swp_str,
        "collect_total": collect_str, "abroad": abroad, "prepaid": prepaid,
    }
