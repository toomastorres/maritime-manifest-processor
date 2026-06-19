"""
line_costs.py - Gastos de línea (THC / TOLL / SWEEPING) en modo híbrido.

Decisión 2 del plan: se LEEN del manifiesto (lo hace el parser) y acá se VALIDAN
contra las tarifas esperadas de config.LineCostTariffs, registrando flags de
discrepancia para revisión humana (no se sobreescriben los importes).
"""

from collections import Counter
from typing import List

from .. import config
from ..domain.models import BLRecord


def detect_loose_tariff(bls: List[BLRecord],
                        known=(18, 20),
                        default: float = 20) -> float:
    """
    Detecta la era de la tarifa S/C suelta (TOLL por Tn) del propio manifiesto.

    La tarifa cambió de 18 a 20 (transición ~viajes 0226); NO se puede inferir
    del código de viaje. Pero el RATE de las líneas S/C de carga suelta (basis no
    contenedor) en el manifiesto ES la tarifa de la era. Se toma la dominante entre
    las tarifas conocidas; los rates distintos son errores de carga de origen
    (van a REVISIÓN, no cuentan para la detección). Si no hay datos, `default`.
    """
    c: Counter = Counter()
    for bl in bls:
        for cl in bl.charge_lines:
            if cl.desc.startswith("s/c") and not cl.is_container_basis and cl.rate > 0:
                r = round(cl.rate, 2)
                if r in known:
                    c[r] += 1
    return c.most_common(1)[0][0] if c else default


def validate_line_costs(bl: BLRecord, tariffs: config.LineCostTariffs) -> List[str]:
    """
    Compara los gastos de línea del BL contra las tarifas esperadas según la
    cantidad de contenedores/tonelaje. Devuelve mensajes de discrepancia.
    """
    flags: List[str] = []
    n_ctr = bl.cargo.container_count
    tol = tariffs.tolerance

    # THC por contenedor: 100 (20') / 120 (40'). Se valida la suma 20+40.
    thc_total = bl.thc_20.monto + bl.thc_40.monto
    if n_ctr > 0 and thc_total > 0:
        expected_20 = bl.thc_20.monto and tariffs.thc_20
        # nº de unidades implícito por monto/tarifa
        if bl.thc_20.monto > 0 and abs((bl.thc_20.monto % tariffs.thc_20)) > tol \
                and abs((bl.thc_20.monto % tariffs.thc_20) - tariffs.thc_20) > tol:
            flags.append(
                f"{bl.bl_no}: THC 20' {bl.thc_20.monto} no es múltiplo de {tariffs.thc_20}"
            )
        if bl.thc_40.monto > 0 and abs((bl.thc_40.monto % tariffs.thc_40)) > tol \
                and abs((bl.thc_40.monto % tariffs.thc_40) - tariffs.thc_40) > tol:
            flags.append(
                f"{bl.bl_no}: THC 40' {bl.thc_40.monto} no es múltiplo de {tariffs.thc_40}"
            )

    # SWEEPING: 20 por contenedor.
    swp = bl.charge("Sweeping").monto
    if n_ctr > 0 and swp > 0:
        expected = n_ctr * tariffs.sweeping
        if abs(swp - expected) > tol:
            flags.append(
                f"{bl.bl_no}: SWEEPING {swp} != {expected} ({n_ctr} ctr x {tariffs.sweeping})"
            )

    # S/C (TOLL): tarifa FIJA. 150/contenedor (basis 20/40) y 20/Tn en
    # carga suelta (basis MT/otro). Cualquier rate distinto = error de origen.
    for cl in bl.charge_lines:
        if not cl.desc.startswith("s/c"):
            continue
        if cl.is_container_basis:
            expected = tariffs.sc_container
        else:
            expected = tariffs.sc_loose_per_ton
        if abs(cl.rate - expected) > tol:
            flags.append(
                f"{bl.bl_no}: TOLL rate {cl.rate} (basis {cl.basis}) != tarifa fija "
                f"{expected} -> posible error de carga de origen"
            )

    return flags
