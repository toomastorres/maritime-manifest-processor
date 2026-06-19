"""
commission.py - Cálculo del Monto Comisión y comisión por tipo de carga.

Regla (reglas_de_negocio.md):
- Monto Comisión = Basic Frt. + comisionables - descuentos; nunca negativo (->0).
- Rolling (cars/big vans): 1%.
- General Cargo (roros + proyecto): 2% IMPO / 4% EXPO.
- Contenedores: 2%/4% Y 25 USD por contenedor; el final es el MÁS ALTO de
  ambos, calculado POR BL (decisión 3 del plan).
- EUR -> USD por ROE (acá para el cómputo del máximo; en el reporte el ROE se
  aplica vía fórmula Excel, decisión 7).
"""

from typing import Tuple

from .. import config
from ..domain.models import BLRecord


def general_subcategory(bl: BLRecord) -> str:
    """
    Para la parte NO-contenedor del BL: 'cars' (ROLLING, 1%) o 'projects'
    (GENERAL CARGO).

    Regla de negocio (literal): ROLLING = "solo cars o big vans"; GENERAL CARGO =
    "roros y carga proyecto". Decodificado contra la verdad exacta:
    - Si hay RORO: GENERAL, salvo MOTORHOME/CAMPER -> ROLLING. Los motorhomes
      (RoRo autopropulsados) ruedan: la mayoría va a ROLLING en la verdad (2 de 3
      casos observados). Un RORO mezclado con autos sigue siendo GENERAL.
    - Si hay autos/vans (sin RORO): ROLLING, aunque convivan con carga general
      "units" incidental (la carga dominante son los autos).
    - Sólo carga general/proyecto (sin vehículos): GENERAL.

    Casos de borde NO determinables del manifiesto -> quedan en REVISIÓN (ver
    validation.py), no se sobre-ajustan acá:
    - MOTORHOME/CAMPER RORO: default ROLLING, pero algún caso aislado va a GENERAL
      sin patrón legible (mismo tipo, peso inverso).
    - Vehículos de uso especial (NCM 8705, ej. Sprinter ambulancia/bombero):
      la verdad los lleva a GENERAL aunque sean "van".
    """
    desc = " ".join(bl.description_lines).upper()
    has_roro = any("roro" in k for k in bl.cargo.vehicles)
    has_cars = any(("car" in k or "van" in k) for k in bl.cargo.vehicles)

    if has_roro:
        if any(kw in desc for kw in config.ROLLING_RORO_KEYWORDS):
            return "cars"
        return "projects"
    if has_cars:
        return "cars"
    return "projects"


def monto_comision(bl: BLRecord) -> Tuple[float, str]:
    """
    Monto Comisión del BL y su moneda original.
    Base = Basic Frt. + ajustes negativos (descuentos). Clamp a >= 0.

    TODO(A1/A2): sumar cargos comisionables extra (config.COMMISSIONABLE_EXTRA)
    una vez confirmada la lista de negocio.
    """
    basic = bl.charge("Basic FRT")
    base = basic.monto + bl.total_ajustes_negativos
    return max(0.0, base), (basic.moneda or "USD")


def _is_commissionable(desc: str) -> bool:
    """¿La descripción de la línea suma al Monto Comisión? (Basic Frt + extras)."""
    if desc == config.COMMISSIONABLE_BASE:
        return True
    return any(desc.startswith(p) for p in config.COMMISSIONABLE_EXTRA)


def _is_discount(cl) -> bool:
    return cl.total < 0 and any(
        cl.desc == p or cl.desc.startswith(p + " ") for p in config.NEGATIVE_PREFIXES
    )


def commissionable_split(bl: BLRecord) -> dict:
    """
    Separa el Monto Comisión del BL por tipo de carga usando el BASIS de cada
    línea de cargo (decisión confirmada por la verdad exacta):
      - basis 20/40  -> comisionable de CONTENEDORES (incluye Open Top S/C).
      - otro basis    -> comisionable de GENERAL/AUTOS.

    El BRUTO comisionable (positivos) se toma de charge_lines. Los DESCUENTOS se
    toman de bl.total_ajustes_negativos (autoritativo, con paridad), porque
    algunos llegan sin columnas FACTOR/BASIS (ej. "FAC -2% of 750"). Los
    descuentos estructurados (con basis) se restan en su grupo; el remanente no
    estructurado se imputa al grupo que tenga comisionable (general por defecto
    en BLs mixtos).
    """
    gross_cont = gross_gen = 0.0
    cont_cur = gen_cur = None
    struct_neg_cont = struct_neg_gen = 0.0

    for cl in bl.charge_lines:
        if _is_commissionable(cl.desc) and cl.total > 0:
            if cl.is_container_basis:
                gross_cont += cl.total
                cont_cur = cont_cur or cl.moneda
            else:
                gross_gen += cl.total
                gen_cur = gen_cur or cl.moneda
        elif _is_discount(cl):
            if cl.is_container_basis:
                struct_neg_cont += cl.total
            else:
                struct_neg_gen += cl.total

    unstruct_neg = bl.total_ajustes_negativos - (struct_neg_cont + struct_neg_gen)

    if gross_cont > 0 and gross_gen == 0:
        cont = gross_cont + struct_neg_cont + struct_neg_gen + unstruct_neg
        gen = 0.0
    elif gross_gen > 0 and gross_cont == 0:
        gen = gross_gen + struct_neg_gen + struct_neg_cont + unstruct_neg
        cont = 0.0
    else:
        cont = gross_cont + struct_neg_cont
        gen = gross_gen + struct_neg_gen + unstruct_neg

    return {
        "container": max(0.0, cont), "container_moneda": cont_cur or "USD",
        "general": max(0.0, gen), "general_moneda": gen_cur or "USD",
    }


def _to_usd(monto: float, moneda: str, roe: float) -> float:
    if moneda == "EUR":
        return monto * roe
    return monto


def container_commission_per_bl(bl: BLRecord, op_type: str, roe: float) -> dict:
    """
    Comisión de contenedores para un BL: max(rate * montoComisión_USD, 30 * ctrs).
    Devuelve componentes para auditar en la planilla.
    """
    rates = config.COMMISSION_RATES[op_type.upper()]
    monto, moneda = monto_comision(bl)
    monto_usd = _to_usd(monto, moneda, roe)
    n_ctr = bl.cargo.container_count

    pct = monto_usd * rates.containers
    minimo = n_ctr * rates.container_min_per_unit
    final = max(pct, minimo)
    return {
        "monto": monto, "moneda": moneda, "monto_usd": monto_usd,
        "n_ctr": n_ctr, "rate": rates.containers,
        "pct": pct, "minimo": minimo, "final": final,
        "min_wins": minimo >= pct,
    }
