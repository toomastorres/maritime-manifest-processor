"""
line_costs.py - Gastos de línea (THC, S/C alias TOLL, SWEEPING).

Modo híbrido: el motor NO recalcula los gastos, los LEE del manifiesto y los
compara contra la tarifa fija esperada (config.LineCostTariffs). Si no coinciden
es un error de carga de origen y se emite una alerta para la hoja REVISIÓN; el
importe que va a la planilla sigue siendo el del manifiesto.
"""

from collections import Counter
from typing import Iterable, List, Optional

from .. import config
from ..domain.models import BLRecord


def detect_loose_tariff(bls: Iterable[BLRecord], default: Optional[float] = None) -> float:
    """
    Tarifa del S/C suelto vigente en el viaje (ej. 18.0 histórico vs 20.0 actual,
    valores ilustrativos).

    Se lee del propio manifiesto: es el rate más frecuente entre las líneas de
    S/C facturadas por tonelada (basis MT). Si el manifiesto no trae ninguna,
    se usa el default.
    """
    if default is None:
        default = config.LineCostTariffs().sc_loose_per_ton
    rates = Counter()
    for bl in bls or []:
        for r in bl.sc_loose_rates:
            rates[round(r, 2)] += 1
    return rates.most_common(1)[0][0] if rates else default


def has_loose_toll(bl: BLRecord) -> bool:
    """True si el S/C del B/L se facturó por tonelada (carga suelta)."""
    return bool(bl.sc_loose_rates)


# Descripciones de línea que son S/C (TOLL) en el manifiesto.
SC_LINE_KEYS = ("s/c(pod)", "s/c(pol)")


def toll_split(bl: BLRecord):
    """
    Separa el TOLL del B/L por factor de facturación: (contenedor, suelto).

    El S/C por contenedor se factura por unidad (basis 20/40/PU) y el suelto
    por tonelada (basis MT). Cuando un B/L trae los dos, son cargas distintas
    conviviendo en el mismo B/L y el negocio pide discriminarlas.
    """
    suelto = sum(cl.total for cl in bl.charge_lines
                 if any(k in cl.desc for k in SC_LINE_KEYS) and cl.basis == "MT")
    total = bl.charge("Toll").monto
    return total - suelto, suelto


def validate_line_costs(bl: BLRecord, tariffs: config.LineCostTariffs) -> List[str]:
    """Devuelve las discrepancias entre lo facturado y la tarifa fija."""
    msgs: List[str] = []
    tol = tariffs.tolerance
    n_ctr = bl.cargo.n_containers

    # SWEEPING por contenedor (20' y 40' por igual).
    swp = bl.charge("Sweeping").monto
    if n_ctr and swp:
        esperado = n_ctr * tariffs.sweeping
        if abs(swp - esperado) > tol:
            msgs.append(
                f"SWEEPING {swp} != {esperado} ({n_ctr} ctr x {tariffs.sweeping})")

    # S/C (TOLL) por contenedor. Se compara sólo la parte facturada por unidad,
    # así el control también vale en los B/L que traen contenedor y carga suelta.
    toll_ctr, _ = toll_split(bl)
    if n_ctr and toll_ctr:
        esperado = n_ctr * tariffs.sc_container
        if abs(toll_ctr - esperado) > tol:
            msgs.append(
                f"TOLL {round(toll_ctr, 2)} != {esperado} "
                f"({n_ctr} ctr x {tariffs.sc_container})")

    # S/C (TOLL) suelto: el rate por tonelada tiene que ser el del viaje.
    for rate in sorted(set(round(r, 2) for r in bl.sc_loose_rates)):
        if abs(rate - tariffs.sc_loose_per_ton) > 0.01:
            msgs.append(
                f"TOLL suelto a {rate}/Tn (tarifa del viaje {tariffs.sc_loose_per_ton}/Tn)")

    # THC: el importe tiene que ser un múltiplo exacto de la tarifa por tamaño.
    for etiqueta, cargo, tarifa in (("20", bl.thc_20, tariffs.thc_20),
                                    ("40", bl.thc_40, tariffs.thc_40)):
        if not cargo.monto:
            continue
        unidades = cargo.monto / tarifa
        if abs(unidades - round(unidades)) * tarifa > tol:
            msgs.append(
                f"THC {etiqueta} {cargo.monto} no es multiplo de {tarifa}")

    return msgs
