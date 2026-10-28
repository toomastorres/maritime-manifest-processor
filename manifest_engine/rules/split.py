"""
split.py - Un B/L, una fila por cada carga que comisiona distinto.

Cada carga del B/L comisiona con su propia regla (contenedor, Bbulk, rodado).
Cuando un B/L trae más de una, la biblia abre una fila por carga y cada fila
lleva sólo lo que es de esa carga:

  - la carga y sus cantidades;
  - el TOLL facturado por esa carga;
  - el flete comisionable que el motor identificó como de esa carga (el mismo
    reparto que usa la planilla de comisiones) y su comisión.

Lo que es del B/L entero (el collect a cobrar, el BAF) va una sola vez, en la
primera fila, para que los totales no se dupliquen. El THC y el Sweeping son de
los contenedores: van en la fila del contenedor.

Una carga sin flete ni TOLL propio no abre fila: es lo que va dentro del
contenedor o los accesorios del rodado, y se muestra junto a la carga principal.

Si con los datos del manifiesto no se puede identificar el flete o el TOLL de
cada carga, el B/L NO se separa: queda en una sola fila 'MIXTO' y se avisa.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from ..catalog import CONTENEDOR
from ..domain.models import BLRecord, CargoItem, ChargeLine, cargo_data
from .commission import (
    G_BBULK, G_CTR, G_ROD, TOL, CommissionResult, Part, _group_of_family, _matches,
)

TIPO = {G_CTR: "CONTENEDOR", G_BBULK: "BBULK", G_ROD: "RODADO"}
MIXTO = "MIXTO"
SIN_CARGA = "SIN IDENTIFICAR"
ORDEN = (G_CTR, G_ROD, G_BBULK)   # el contenedor primero: lleva THC y Sweeping

TOLL_CONCEPT = {"IMPO": "Toll S/C(POD)", "EXPO": "Toll S/C(POL)"}


@dataclass
class Fila:
    """Una fila de la biblia: un B/L entero o una de sus cargas."""
    bl: BLRecord
    res: CommissionResult
    tipo: str                                   # CONTENEDOR | BBULK | RODADO | MIXTO | SIN IDENTIFICAR
    items: List[CargoItem] = field(default_factory=list)        # cuentan en cantidades
    extra_items: List[CargoItem] = field(default_factory=list)  # van en el texto, no en cantidades
    toll_lines: List[ChargeLine] = field(default_factory=list)  # TOLL que facturamos nosotros
    toll_prepaid: List[ChargeLine] = field(default_factory=list)  # IMPO: TOLL pagado en origen
    comm_lines: List[ChargeLine] = field(default_factory=list)
    parts: List[Part] = field(default_factory=list)
    first: bool = True
    container_costs: bool = True                # THC y Sweeping van en esta fila
    of: int = 1                                 # filas del B/L

    @property
    def toll(self) -> float:
        return round(sum(cl.total for cl in self.toll_lines), 2)

    def qty(self, family: str) -> int:
        return sum(it.qty for it in self.items if it.family == family)

    def qty_size(self, size: str) -> int:
        """Contenedores de 20' o de 40' (los de 45' cuentan como 40')."""
        return sum(it.qty for it in self.items
                   if it.family == CONTENEDOR and container_size(it.kind) == size)

    @property
    def cargo(self):
        return cargo_data(self.items + self.extra_items)

    @property
    def tons(self) -> Optional[float]:
        vals = [it.tn for it in self.items if it.tn is not None]
        return round(sum(vals), 3) if vals else None


def container_size(kind: str) -> Optional[str]:
    k = kind.lower().strip()
    if k.startswith("20"):
        return "20"
    if k.startswith(("40", "45")):
        return "40"
    return None


def _toll_groups(bl: BLRecord, res: CommissionResult, toll_lines: List[ChargeLine],
                 groups: List[str], partial: bool = False) -> Optional[Dict[int, str]]:
    """
    De qué carga es cada línea de TOLL. Orden:
      1. El TOLL por unidad (por contenedor) es del contenedor.
      2. El TOLL por tonelada es de la carga suelta cuya cotización (Tariff ID)
         ya está identificada en el flete, o cuyo peso coincide con el factor.
      3. Descarte: si todas las cargas sueltas menos una ya tienen su TOLL
         completo (toneladas facturadas = toneladas del manifiesto), lo que queda
         es de la otra.
    Devuelve None si alguna línea queda sin identificar (con partial=True,
    devuelve lo que se pudo identificar).
    """
    if len(groups) <= 1:
        g = groups[0] if groups else None
        return {id(cl): g for cl in toll_lines}

    sueltas = [g for g in groups if g != G_CTR]
    por_cotizacion: Dict[str, set] = {}
    for cl in res.lines:
        g = res.assignment.get(id(cl))
        if g and cl.tariff:
            por_cotizacion.setdefault(cl.tariff, set()).add(g)

    out: Dict[int, str] = {}
    for cl in toll_lines:
        if cl.basis != "MT":
            if G_CTR in groups:
                out[id(cl)] = G_CTR
            continue
        por_peso = set()
        if cl.factor is not None:
            for it in bl.cargo_items:
                g = _group_of_family(it.family)
                if g in sueltas and _matches(it, "MT", cl.factor, per_piece=True):
                    por_peso.add(g)
        por_cot = {g for g in por_cotizacion.get(cl.tariff, set()) if g in sueltas} if cl.tariff else set()
        if len(por_cot) == 1 and (not por_peso or por_cot <= por_peso):
            out[id(cl)] = next(iter(por_cot))
        elif len(por_peso) == 1 and not por_cot:
            out[id(cl)] = next(iter(por_peso))

    pendientes = [cl for cl in toll_lines if id(cl) not in out]
    if pendientes:
        incompletas = []
        for g in sueltas:
            tn = sum(it.tn or 0 for it in bl.cargo_items if _group_of_family(it.family) == g)
            facturado = sum(abs(cl.factor or 0) for cl in toll_lines if out.get(id(cl)) == g)
            if not (tn and abs(tn - facturado) <= TOL):
                incompletas.append(g)
        if len(incompletas) == 1:
            for cl in pendientes:
                out[id(cl)] = incompletas[0]

    if not partial and any(id(cl) not in out for cl in toll_lines):
        return None
    return out


def filas_del_bl(bl: BLRecord, res: CommissionResult, op: str) -> List[Fila]:
    """Filas de la biblia para un B/L (ver docstring del módulo)."""
    op = op.upper()
    tolls = [cl for cl in bl.charge_lines if cl.concept == TOLL_CONCEPT[op]]
    # En IMPO el TOLL prepaid se pagó en origen: no lo facturamos nosotros.
    toll_lines = [cl for cl in tolls if op == "EXPO" or cl.letra != "P"]
    toll_prepaid = [cl for cl in tolls if cl not in toll_lines]
    groups = [g for g in ORDEN if g in res.groups]

    def entero(tipo: str) -> List[Fila]:
        f = Fila(bl, res, tipo, items=list(bl.cargo_items), toll_lines=toll_lines,
                 toll_prepaid=toll_prepaid, comm_lines=list(res.lines), parts=list(res.parts))
        return [f]

    if not groups:
        return entero(SIN_CARGA)
    if len(groups) == 1:
        return entero(TIPO[groups[0]])

    toll_g = _toll_groups(bl, res, toll_lines, groups)
    if not res.split_ok or toll_g is None:
        return entero(MIXTO)

    def de(g):
        return [it for it in bl.cargo_items if _group_of_family(it.family) == g]

    propias = {g: [cl for cl in res.lines if res.assignment.get(id(cl)) == g] for g in groups}
    por_carga = {g: [cl for cl in toll_lines if toll_g.get(id(cl)) == g] for g in groups}
    pre_g = _toll_groups(bl, res, toll_prepaid, groups, partial=True)
    pre = {g: [cl for cl in toll_prepaid if pre_g.get(id(cl)) == g] for g in groups}

    # Una carga abre fila si tiene flete o TOLL propio, o si es contenedor o
    # rodado (el contenedor lleva THC y el rodado comisiona por unidad).
    con_fila = [g for g in groups if g in (G_CTR, G_ROD) or propias[g] or por_carga[g]]
    sin_fila = [g for g in groups if g not in con_fila]

    filas = []
    for g in con_fila:
        filas.append(Fila(bl, res, TIPO[g], items=de(g), toll_lines=por_carga[g], toll_prepaid=pre[g],
                          comm_lines=propias[g], parts=[p for p in res.parts if p.group == g]))
    # TOLL prepaid que no se pudo ubicar: con la primera fila (sólo se muestra el origen).
    filas[0].toll_prepaid += [cl for cl in toll_prepaid if id(cl) not in pre_g
                              or pre_g[id(cl)] not in con_fila]
    # Lo que no abre fila va con la primera: dentro del contenedor no cuenta
    # como bulto aparte; con un rodado, sí.
    for g in sin_fila:
        destino = filas[0]
        if destino.tipo == TIPO[G_CTR]:
            destino.extra_items += de(g)
        else:
            destino.items += de(g)
        destino.parts += [p for p in res.parts if p.group == g]
        destino.comm_lines += propias[g]

    tiene_ctr = any(f.tipo == TIPO[G_CTR] for f in filas)
    for i, f in enumerate(filas):
        f.first = (i == 0)
        f.of = len(filas)
        f.container_costs = (f.tipo == TIPO[G_CTR]) if tiene_ctr else f.first
    return filas


def split_alerts(filas: List[Fila]) -> List[str]:
    """Motivo por el que un B/L con varias cargas quedó en una sola fila."""
    out = []
    for f in filas:
        if f.tipo != MIXTO:
            continue
        cargas = " + ".join(TIPO[g] for g in ORDEN if g in f.res.groups)
        if not f.res.split_ok:
            out.append(f"Trae {cargas} y no se pudo identificar el flete de cada carga: "
                       f"queda en una fila MIXTO (ver el motivo en la comisión)")
        else:
            out.append(f"Trae {cargas} y no se pudo identificar el TOLL de cada carga: "
                       f"queda en una fila MIXTO; clasificar su TOLL a mano en la hoja TOLL")
    return out
