"""
validation.py - Capa REVISIÓN: SOLO alertas ACCIONABLES.

Filosofía (decisión del usuario): la hoja REVISIÓN debe mostrar únicamente lo que
un humano realmente tiene que mirar. Los casos RECURRENTES y EXPLICADOS (que no
son error del motor sino corrección humana o convención conocida) NO figuran acá:
se consideran "tratados como corregidos" y su contabilidad vive en el módulo de
auditoría (manifest_engine/audit.py), que los categoriza contra la verdad.

OMITIDOS a propósito (conocidos/recurrentes, no accionables):
  - BL sin Basic Frt → transbordo / Service B/L (el flete va en la parte service).
  - Cantidad descripción ≠ unidades comisionadas → sub-agrupación del manifiesto.
  - RORO motorhome/camper y vehículos NCM 8705 → clasificación rolling/general
    es criterio del operador (el motor usa un default razonable).
  - ENS forzado a 15 → es la regla, no una anomalía.
  - BL de autos sin VIN / VIN parcial → el manifiesto no siempre imprime los
    chasis; es normal y muy frecuente.

SE EMITEN (genuinas, accionables):
  - GASTO LINEA: TOLL/THC/SWEEPING fuera de tarifa fija = error de carga de origen.
  - COMISION: comisionable bruto > 0 pero neto clampeado a 0 (descuentos lo anulan).
  - MONEDA: comisionable en moneda no convertible (ni USD ni EUR).
  - CARGA: hay Basic Frt pero no se pudo clasificar la carga (posible parser miss).
"""

from dataclasses import dataclass
from typing import List

from . import config
from .domain.models import BLRecord
from .rules.line_costs import validate_line_costs
from .rules.commission import commissionable_split


@dataclass
class Alerta:
    bl: str
    tipo: str
    detalle: str


def revisar_bls(bls: List[BLRecord], op_type: str, cfg: config.EngineConfig) -> List[Alerta]:
    alertas: List[Alerta] = []
    _seen = set()

    def add(bl, tipo, detalle):
        key = (bl, tipo, detalle)
        if key in _seen:           # dedupe: misma alerta repetida (varias líneas)
            return
        _seen.add(key)
        alertas.append(Alerta(bl, tipo, detalle))

    for bl in bls:
        basic = bl.charge("Basic FRT")

        # 1) Gastos de línea fuera de tarifa (TOLL/THC/SWEEPING) = error de origen.
        for msg in validate_line_costs(bl, cfg.tariffs):
            add(bl.bl_no, "GASTO LINEA", msg.split(": ", 1)[-1])

        # 2) Comisionable bruto > 0 pero neto clampeado a 0 (descuentos lo anulan).
        sp = commissionable_split(bl)
        bruto = sum(cl.total for cl in bl.charge_lines
                    if cl.desc == config.COMMISSIONABLE_BASE and cl.total > 0)
        if bruto > 0 and (sp["container"] + sp["general"]) <= 0.005:
            add(bl.bl_no, "COMISION",
                f"Comisionable clampeado a 0 (bruto {bruto:.2f}, descuentos lo anulan)")

        # 3) Moneda no convertible en el comisionable.
        if basic.moneda and basic.moneda not in ("USD", "EUR"):
            add(bl.bl_no, "MONEDA", f"Basic Frt. en {basic.moneda} (no se convierte; revisar TC)")

        # 4) Hay Basic Frt pero no se clasificó la carga (posible parser miss).
        if basic.monto and bl.cargo.is_empty:
            add(bl.bl_no, "CARGA", "Hay Basic Frt. pero no se clasificó la carga")

    return alertas
