"""
billing.py - Importes del B/L según los totales del manifiesto.

Cada B/L cierra su bloque de cargos con líneas 'Total[Lugar] importe P/C
moneda'. Lo que se cobra localmente es lo que el manifiesto totaliza en Buenos
Aires (incluye 'Buenos Aires - Partner', bookings de subagente local). De ahí salen:

  IMPO  COLLECT por moneda = Total Buenos Aires collect - BAF collect - lo que
        va en sus propias columnas (TOLL, THC 20'/40', Sweeping collect).
        El THC por unidad de los rodados no tiene columna: queda en el collect.
  EXPO  PREPAID (incl. ENS) = Total Buenos Aires prepaid; THC y Sweeping prepaid
        se descuentan en la planilla con una fórmula. COLLECT por moneda = los
        totales collect; ABROAD = lo prepaid que se paga afuera.

Control: la suma de las líneas de cargo de cada condición y moneda tiene que dar
lo mismo que sus totales. Si no da, falta o sobra una línea (o el manifiesto
imprimió mal un signo) y el B/L se marca para revisar.
"""

from collections import defaultdict
from dataclasses import dataclass, field
from itertools import combinations
from typing import Dict, List, Optional, Tuple

from ..domain.models import BLRecord, ChargeLine, TotalLine

TOL_TOTAL = 0.10   # redondeo de las líneas 'x% of y' del propio manifiesto


def is_local(place: str) -> bool:
    """Total que se paga en Buenos Aires (incluye 'Buenos Aires - Partner')."""
    return place.lower().startswith("buenos aires")


def local_totals(bl: BLRecord, letra: Optional[str] = None) -> List[TotalLine]:
    return [t for t in bl.total_lines if is_local(t.place) and (letra is None or t.letra == letra)]


def by_currency(lines) -> Dict[str, float]:
    out: Dict[str, float] = defaultdict(float)
    for t in lines:
        out[t.moneda] += t.amount
    return dict(out)


def _is_baf(desc: str) -> bool:
    return "baf" in desc


@dataclass
class ImpoMoney:
    collect: Dict[str, float] = field(default_factory=dict)    # moneda -> importe
    baf: Dict[str, float] = field(default_factory=dict)        # BAF collect por moneda
    local_total: Dict[str, float] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)


def impo_money(bl: BLRecord, toll_concept: str = "Toll S/C(POD)") -> ImpoMoney:
    m = ImpoMoney()
    m.local_total = by_currency(local_totals(bl, "C"))
    for t in local_totals(bl):
        if t.letra == "C":
            continue
        lineas = local_lines(bl, t.letra, t.moneda)
        detalle = f" ({', '.join(dict.fromkeys(cl.concept for cl in lineas))})" if lineas else ""
        m.warnings.append(f"Importación con un total PREPAID en Buenos Aires: Total[{t.place}] "
                          f"{t.amount:,.2f}{t.letra} {t.moneda}{detalle}. No entra en las columnas "
                          f"collect; figura en el reporte de Facturación")

    baf = defaultdict(float)
    propias = defaultdict(float)      # lo que tiene columna propia en la biblia
    for cl in bl.charge_lines:
        if cl.letra != "C":
            continue
        if _is_baf(cl.desc):
            baf[cl.moneda] += cl.total
        elif cl.concept == toll_concept or "sweeping" in cl.desc:
            propias[cl.moneda] += cl.total
    for thc in (bl.thc_20, bl.thc_40):       # THC de contenedor: tiene columna
        if thc.monto and thc.letra == "C":
            propias[thc.moneda or "USD"] += thc.monto

    m.baf = {k: round(v, 2) for k, v in baf.items() if abs(v) > 0.005}
    for moneda, total in m.local_total.items():
        neto = round(total - baf.get(moneda, 0.0) - propias.get(moneda, 0.0), 2)
        if abs(neto) > 0.005:
            m.collect[moneda] = neto
        if neto < -0.005:
            m.warnings.append(f"El collect {moneda} da negativo ({neto:,.2f}): "
                              f"lo descontado supera el Total en Buenos Aires")
        if moneda not in ("USD", "EUR"):
            m.warnings.append(f"Total collect en Buenos Aires en {moneda}: la biblia sólo tiene "
                              f"columnas USD y EUR")
    return m


@dataclass
class ExpoMoney:
    local_prepaid: float = 0.0                                  # USD
    collect: Dict[str, float] = field(default_factory=dict)
    abroad: Dict[str, float] = field(default_factory=dict)
    warnings: List[str] = field(default_factory=list)


def expo_money(bl: BLRecord) -> ExpoMoney:
    m = ExpoMoney()
    pre = by_currency(local_totals(bl, "P"))
    m.local_prepaid = round(pre.pop("USD", 0.0), 2)
    for moneda, v in pre.items():
        m.warnings.append(f"Total prepaid en Buenos Aires en {moneda} ({v:,.2f}): "
                          f"la columna PREPAID es en USD")
    m.collect = {k: round(v, 2) for k, v in by_currency(
        [t for t in bl.total_lines if t.letra == "C"]).items()}
    m.abroad = {k: round(v, 2) for k, v in by_currency(
        [t for t in bl.total_lines if t.letra == "P" and not is_local(t.place)]).items()}
    return m


MAX_LINEAS_COMBINAR = 18


def local_lines(bl: BLRecord, letra: str, moneda: str) -> Optional[List[ChargeLine]]:
    """
    Líneas de cargo (condición y moneda dadas) que forman el total pagadero en
    Buenos Aires. Si en esa condición y moneda hay un solo total, son todas.
    Si hay otro total afuera (EXPO: Buenos Aires y Casa Matriz), el manifiesto no
    marca cada línea: se busca la combinación de líneas que suma exactamente el
    total de Buenos Aires, y sólo se acepta si es única. Si no se puede saber,
    devuelve None.
    """
    totales = [t for t in bl.total_lines if t.letra == letra and t.moneda == moneda]
    lineas = [cl for cl in bl.charge_lines if cl.letra == letra and cl.moneda == moneda]
    locales = [t for t in totales if is_local(t.place)]
    if len(locales) != 1:
        return None
    if len(totales) == 1:
        return lineas
    if len(lineas) > MAX_LINEAS_COMBINAR:
        return None
    objetivo = locales[0].amount
    firmas, elegida = set(), None
    for n in range(0, len(lineas) + 1):
        for combo in combinations(lineas, n):
            if abs(sum(cl.total for cl in combo) - objetivo) < 0.005:
                firmas.add(tuple(sorted((cl.concept, round(cl.total, 2)) for cl in combo)))
                if len(firmas) > 1:
                    return None
                elegida = elegida or list(combo)
    return elegida if len(firmas) == 1 else None


def total_mismatches(bl: BLRecord) -> List[Tuple[str, str, float, float]]:
    """(condición, moneda, suma de líneas, suma de totales) donde no cierran."""
    lineas: Dict[Tuple[str, str], float] = defaultdict(float)
    totales: Dict[Tuple[str, str], float] = defaultdict(float)
    for cl in bl.charge_lines:
        lineas[(cl.letra, cl.moneda)] += cl.total
    for t in bl.total_lines:
        totales[(t.letra, t.moneda)] += t.amount
    out = []
    for k in sorted(set(lineas) | set(totales), key=lambda k: (str(k[0]), str(k[1]))):
        if abs(lineas[k] - totales[k]) > TOL_TOTAL:
            out.append((k[0], k[1], round(lineas[k], 2), round(totales[k], 2)))
    return out
