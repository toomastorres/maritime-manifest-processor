"""
validation.py - Capa REVISIÓN: SOLO alertas ACCIONABLES.

Filosofía: la hoja REVISIÓN debe mostrar únicamente lo que un analista humano
realmente tiene que mirar. Los casos recurrentes y explicados por la operación
normal NO figuran acá.

OMITIDOS a propósito (conocidos/recurrentes, no accionables):
  - B/L sin Basic Frt -> transbordo / Service B/L (el flete va en la parte service).
  - Cantidad descripción != unidades comisionadas -> sub-agrupación del manifiesto.
  - RORO motorhome/camper -> criterio operativo estándar.
  - ENS a monto fijo (config.ENS_FIXED_AMOUNT) -> es la regla, no una anomalía.
  - B/L de autos sin VIN / VIN parcial -> el manifiesto no siempre imprime todos
    los chasis; es normal y frecuente.

SE EMITEN (genuinas, accionables — junto con COMISIÓN - REVISAR, CARGA MIXTA y
LISTA COMERCIAL agregadas desde biblia/comisiones/__main__):
  - GASTO LINEA: TOLL/THC/SWEEPING fuera de tarifa fija = error de carga de origen.
  - MONEDA: comisionable mayor a 0 en moneda no convertible (ni USD ni EUR).
  - CARGA: hay Basic Frt pero no se pudo clasificar la carga (posible parser miss).
  - TOTAL MANIFIESTO: las líneas de cargo no suman lo mismo que los Total[...] del
    B/L (falta o sobra una línea, o el manifiesto imprimió mal un signo).
  - IMPORTES: el total a cobrar en Buenos Aires no entra en las columnas de la
    biblia (otra moneda, collect negativo, prepaid en una importación).
"""

from dataclasses import dataclass
from typing import List, Optional

from . import config
from .catalog import Catalog
from .domain.models import BLRecord
from .rules.billing import expo_money, impo_money, total_mismatches
from .rules.commission import _fmt, commissionable_split
from .rules.line_costs import validate_line_costs


@dataclass
class Alerta:
    bl: str
    tipo: str
    detalle: str


def revisar_bls(bls: List[BLRecord], op_type: str, cfg: config.EngineConfig,
                catalog: Optional[Catalog] = None) -> List[Alerta]:
    alertas: List[Alerta] = []
    _seen = set()

    def add(bl, tipo, detalle):
        key = (bl, tipo, detalle)
        if key in _seen:
            return
        _seen.add(key)
        alertas.append(Alerta(bl, tipo, detalle))

    for bl in bls:
        basic = bl.charge("Basic FRT")

        # 1) Gastos de línea fuera de tarifa (TOLL/THC/SWEEPING) = error de origen.
        for msg in validate_line_costs(bl, cfg.tariffs):
            add(bl.bl_no, "GASTO LINEA", msg.split(": ", 1)[-1])

        sp = commissionable_split(bl, catalog=catalog)

        # 2) Moneda no convertible en el comisionable.
        if basic.moneda and basic.moneda not in ("USD", "EUR") and sum(sp.values()) > 0.005:
            add(bl.bl_no, "MONEDA", f"Basic Frt. en {basic.moneda} (no se convierte; revisar TC)")

        # 3) Hay Basic Frt pero no se clasificó la carga (posible parser miss).
        if basic.monto and bl.cargo.is_empty:
            add(bl.bl_no, "CARGA", "Hay Basic Frt. pero no se clasificó la carga")

        # 4) Las líneas de cargo tienen que sumar lo mismo que los totales.
        for letra, moneda, lineas, totales in total_mismatches(bl):
            add(bl.bl_no, "TOTAL MANIFIESTO",
                f"Las líneas {letra} {moneda} suman {_fmt(lineas)} y los Total[...] del manifiesto "
                f"{_fmt(totales)} (diferencia {_fmt(lineas - totales)}): falta o sobra una línea, "
                f"o un signo mal impreso")

        # 5) Lo que se cobra en Buenos Aires tiene que entrar en las columnas.
        money = impo_money(bl) if op_type.upper() == "IMPO" else expo_money(bl)
        for w in money.warnings:
            add(bl.bl_no, "IMPORTES", w)

    return alertas
