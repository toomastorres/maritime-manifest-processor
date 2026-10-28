"""
commission.py - Comisión por B/L gobernada por catálogo.

Reglas (valores de tasas y mínimos en config.py son ilustrativos de portafolio;
las clasificaciones de conceptos, bultos y puertos vienen de la planilla Excel):

  Monto comisionable = conceptos que SUMAN (Basic Frt, Open Top, Over Hght…)
                       + conceptos que RESTAN (FAC, Rebate, Commercial Adj.).
                       Nunca negativo. Si hay importes en USD y en EUR, el EUR
                       se pasa a USD con el ROE.

  ALL IN (todo el B/L): sólo trae Basic Frt como flete. Cualquier RECARGO (BAF,
  ETS, BAF Decrease, THC del puerto del exterior…) u otro concepto que SUMA (Open
  Top, Over Hght) lo saca de ALL IN. Si es ALL IN se toma el 80% del Basic Frt.

  Contenedores  IMPO 2% | EXPO 4% (Brasil 3,5%). Mínimo 25 USD por contenedor:
                se toma el mayor entre el porcentaje (ya con 80%/Brasil) y el mínimo.
  Bbulk         IMPO 2% | EXPO 4% (Brasil 3,5%). Los Mafis llevan Bbulk.
  Rodados       Booking cerrado por la agencia AMA (lista de Comercial): 1% del flete.
                Si no: 2,50 USD por Car, 3,50 USD por Van, 12,00 USD por RoRo.
  Control Equipment Fees: 4,00 USD por contenedor (no Mafis).

  Conceptos de cargo nuevos (no están en la planilla): no influyen en la
  comisión; se registran automáticamente para sumarlos a la lista.

Si algo no se puede resolver con estas reglas el B/L queda EN REVISIÓN: su
comisión no se informa y el motivo se escribe. El motor nunca estima un valor.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .. import config
from ..catalog import (
    BBULK, CONTENEDOR, RECARGO, RESTA, RODADOS, SUMA, Catalog, load_catalog,
)
from ..domain.models import BLRecord, ChargeLine

# Grupos que consumen flete (el flete del B/L se reparte entre ellos).
G_CTR = "CONTENEDOR"
G_BBULK = "BBULK"
G_ROD = "RODADO"

MODE_AGENCY = "AMA"         # booking cerrado por la agencia (AMA): 1% del flete
MODE_UNIT = "UNIDAD"        # tarifa por unidad
MODE_ESTIMADO = "ESTIMADO"  # sin lista de Comercial: por unidad, a confirmar

TOL = 0.011  # tolerancia al comparar factor facturado vs medida del bulto


@dataclass
class Part:
    """Una porción del B/L que comisiona con una regla propia."""
    group: str                      # CONTENEDOR | BBULK | RODADO
    mode: Optional[str] = None      # sólo rodados: AMA | UNIDAD | ESTIMADO
    base: float = 0.0               # monto comisionable de la parte, ya con ALL IN
    base_manifest: float = 0.0      # antes de aplicar el 80%
    currency: str = "USD"
    rate: float = 0.0               # tanto por uno
    units: Dict[str, int] = field(default_factory=dict)
    min_ctr_units: int = 0
    calc_usd: float = 0.0           # porcentaje en USD
    minimum_usd: float = 0.0
    commission_usd: float = 0.0


@dataclass
class CommissionResult:
    bl: BLRecord
    op: str
    port: str                       # puerto por el que se agrupa en la planilla
    status: str = "OK"              # OK | REVISION
    reasons: List[str] = field(default_factory=list)
    suggestion: str = ""
    notes: List[str] = field(default_factory=list)
    all_in: bool = False
    all_in_breakers: List[str] = field(default_factory=list)
    brasil: bool = False
    currency: str = "USD"
    gross: float = 0.0
    net: float = 0.0
    parts: List[Part] = field(default_factory=list)
    cef_units: int = 0
    rodado_mode: Optional[str] = None
    groups: List[str] = field(default_factory=list)
    new_concepts: List[Tuple[str, str]] = field(default_factory=list)  # (concepto, ejemplo)
    lines: List[ChargeLine] = field(default_factory=list)              # líneas comisionables
    assignment: Dict[int, str] = field(default_factory=dict)           # id(línea) -> grupo de carga
    split_ok: bool = True           # el flete quedó identificado por carga

    @property
    def ok(self) -> bool:
        return self.status == "OK"

    @property
    def total_usd(self) -> Optional[float]:
        if not self.ok:
            return None
        return sum(p.commission_usd for p in self.parts)

    def part(self, group: str) -> Optional[Part]:
        return next((p for p in self.parts if p.group == group), None)


# ───────────────────────────── utilidades ─────────────────────────────

def _fmt(v: float) -> str:
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _is_basic(cl: ChargeLine) -> bool:
    return cl.concept.lower().startswith("basic frt")


def _group_of_family(family: str) -> Optional[str]:
    if family == CONTENEDOR:
        return G_CTR
    if family == BBULK:
        return G_BBULK
    if family in RODADOS:
        return G_ROD
    return None


def concept_rule(cl: ChargeLine, op: str, catalog: Catalog) -> Optional[Tuple[Optional[str], Optional[str]]]:
    """(categoría, efecto) del concepto para esta operación; None si es un concepto nuevo."""
    for prefijo, regla in config.OP_CONCEPT_OVERRIDES.get(op.upper(), {}).items():
        if cl.concept.startswith(prefijo):
            return regla
    rule = catalog.concept(cl.concept)
    if rule is None:
        return None
    return rule.category, rule.effect


def _commissionable_lines(bl: BLRecord, catalog: Catalog, op: str) -> List[ChargeLine]:
    out = []
    for cl in bl.charge_lines:
        r = concept_rule(cl, op, catalog)
        if r is not None and r[1] in (SUMA, RESTA):
            out.append(cl)
    return out


def _example(bl: BLRecord, cl: ChargeLine) -> str:
    detalle = f" {cl.factor:g} {cl.basis}" if cl.factor is not None else ""
    return f"{bl.bl_clean}: {cl.concept}{detalle} = {_fmt(cl.total)}{cl.letra or ''} {cl.moneda or ''}"


def commissionable_net(bl: BLRecord, catalog: Optional[Catalog] = None, op: Optional[str] = None) -> float:
    """Monto comisionable del manifiesto (sin ALL IN), nunca negativo, en la moneda del flete."""
    catalog = catalog or load_catalog()
    return max(0.0, sum(cl.total for cl in _commissionable_lines(bl, catalog, op or bl.op_type)))


def commissionable_currency(bl: BLRecord, catalog: Optional[Catalog] = None) -> str:
    catalog = catalog or load_catalog()
    curr = [cl.moneda for cl in _commissionable_lines(bl, catalog, bl.op_type) if cl.moneda]
    if curr:
        return curr[0]
    return bl.charge("Basic FRT").moneda or "USD"


def commissionable_split(bl: BLRecord, catalog: Optional[Catalog] = None) -> Dict[str, float]:
    """Compatibilidad con la hoja REVISIÓN: monto neto del B/L en una sola bolsa."""
    return {"total": commissionable_net(bl, catalog)}


# ───────────────────────── reparto del flete ──────────────────────────

def _matches(it, basis: str, factor: float, per_piece: bool = False) -> bool:
    m = it.metric(basis)
    if m is None:
        return False
    if abs(m - abs(factor)) <= TOL:
        return True
    return per_piece and it.qty > 1 and abs(m / it.qty - abs(factor)) <= TOL


def _metric_candidates(cl: ChargeLine, bl: BLRecord, groups: List[str]) -> set:
    """Grupos cuyas medidas coinciden con el factor facturado de la línea."""
    if cl.basis in ("20", "40"):
        return {G_CTR} if G_CTR in groups else set()
    if cl.factor is None or cl.basis is None:
        return set()
    out = set()
    for it in bl.cargo_items:
        g = _group_of_family(it.family)
        if g in groups and _matches(it, cl.basis, cl.factor):
            out.add(g)
    if cl.basis == "PU":
        for g in groups:
            tot = sum(it.qty for it in bl.cargo_items if _group_of_family(it.family) == g)
            if tot and abs(tot - abs(cl.factor)) <= TOL:
                out.add(g)
    return out


def _tariff_evidence(bl: BLRecord, groups: List[str]) -> Dict[str, set]:
    """
    Por cotización (Tariff ID): qué cargas tienen medidas que coinciden con el
    factor de ALGUNA línea de esa cotización (BAF, THC, Toll, Basic…), exacto o
    por pieza. Una cotización que apunta a una sola carga es de esa carga.
    """
    ev: Dict[str, set] = {}
    for cl in bl.charge_lines:
        if not cl.tariff or cl.factor is None or cl.basis in (None, "PU", "20", "40"):
            continue
        for it in bl.cargo_items:
            g = _group_of_family(it.family)
            if g in groups and _matches(it, cl.basis, cl.factor, per_piece=True):
                ev.setdefault(cl.tariff, set()).add(g)
    return ev


def _covered(bl: BLRecord, group: str, lines: List[ChargeLine], asign: Dict[int, str]) -> bool:
    """La carga ya tiene identificado su flete: cada bulto tiene una línea propia."""
    items = [it for it in bl.cargo_items if _group_of_family(it.family) == group]
    propias = [cl for cl in lines if asign.get(id(cl)) == group]
    if not items or not propias:
        return False
    if len(propias) >= sum(it.qty for it in items):
        return True
    return all(any(cl.factor is not None and cl.basis is not None and _matches(it, cl.basis, cl.factor)
                   for cl in propias) for it in items)


def _tariff_hint(bl: BLRecord, lines: List[ChargeLine]) -> str:
    """Indicios por cotización para que el analista decida rápido."""
    hints = []
    by_tariff: Dict[str, List[ChargeLine]] = {}
    for cl in bl.charge_lines:
        if cl.tariff:
            by_tariff.setdefault(cl.tariff, []).append(cl)
    vistos = set()
    for cl in lines:
        if not cl.tariff or cl.tariff in vistos:
            continue
        vistos.add(cl.tariff)
        evid = []
        for other in by_tariff.get(cl.tariff, []):
            if other.factor is None or other.basis in (None, "PU"):
                continue
            for it in bl.cargo_items:
                m = it.metric(other.basis)
                if m is None:
                    continue
                if abs(m - abs(other.factor)) <= TOL:
                    evid.append(f"{other.concept} {other.factor:g} {other.basis} = {it.qty} {it.kind}")
                elif it.qty > 1 and abs(m / it.qty - abs(other.factor)) <= TOL:
                    evid.append(f"{other.concept} {other.factor:g} {other.basis} = 1 de {it.qty} {it.kind}")
        if evid:
            hints.append(f"cotización {cl.tariff}: " + "; ".join(sorted(set(evid))))
    return " | ".join(hints)


def _assign(bl: BLRecord, lines: List[ChargeLine], groups: List[str],
            res: CommissionResult) -> Dict[int, str]:
    """
    Asigna cada línea comisionable a una carga. Orden de identificación:
      1. Medidas: el factor facturado coincide con el bulto (metros, volumen, peso,
         cantidad; 20'/40' es contenedor).
      2. Contenedor + otra carga: lo identificado como contenedor es del
         contenedor; todo lo demás, por descarte, de la otra carga.
      3. Cotización (Tariff ID): si las líneas de una cotización coinciden con las
         medidas de una sola carga, todo lo de esa cotización es de esa carga.
      4. Descarte: si una carga ya tiene identificado el flete de todos sus
         bultos, lo que queda es de la otra.
    Los descuentos 'x% of y' siguen a la línea cuyo total es su base.
    Lo que no se resuelve así queda sin asignar y el B/L va a revisión.
    """
    if len(groups) == 1:
        return {id(cl): groups[0] for cl in lines}

    basics = [cl for cl in lines if cl.pct_of is None]
    pcts = [cl for cl in lines if cl.pct_of is not None]
    cand = {id(cl): _metric_candidates(cl, bl, groups) for cl in basics}
    asign: Dict[int, str] = {}
    metodo = []

    if len(groups) == 2 and G_CTR in groups:
        otro = next(g for g in groups if g != G_CTR)
        for cl in basics:
            asign[id(cl)] = G_CTR if cand[id(cl)] == {G_CTR} else otro
        metodo.append(f"contenedor identificado por 20'/40', el resto es de {otro}")
    else:
        for cl in basics:
            if len(cand[id(cl)]) == 1:
                asign[id(cl)] = next(iter(cand[id(cl)]))
        if any(id(cl) in asign for cl in basics):
            metodo.append("medidas")
        if len(groups) == 2:
            ev = _tariff_evidence(bl, groups)
            usado = False
            for cl in basics:
                if id(cl) in asign or not cl.tariff:
                    continue
                g = ev.get(cl.tariff)
                if g and len(g) == 1:
                    gg = next(iter(g))
                    if not cand[id(cl)] or gg in cand[id(cl)]:
                        asign[id(cl)] = gg
                        usado = True
            if usado:
                metodo.append("cotización (Tariff ID)")
            pend = [cl for cl in basics if id(cl) not in asign]
            if pend:
                cubiertos = [g for g in groups if _covered(bl, g, basics, asign)]
                if len(cubiertos) == 1:
                    otro = next(g for g in groups if g != cubiertos[0])
                    for cl in pend:
                        asign[id(cl)] = otro
                    metodo.append(f"descarte ({cubiertos[0]} ya identificado, el resto es de {otro})")

    for cl in pcts:
        bases = {asign[id(b)] for b in basics
                 if id(b) in asign and abs(abs(b.total) - cl.pct_of) <= TOL}
        if len(bases) == 1:
            asign[id(cl)] = bases.pop()
        elif not bases and len(groups) == 2 and G_CTR in groups:
            asign[id(cl)] = next(g for g in groups if g != G_CTR)

    dudosas = [cl for cl in lines if id(cl) not in asign]
    if dudosas:
        det = "; ".join(
            f"{cl.concept} {cl.factor:g} {cl.basis} = {_fmt(cl.total)} {cl.moneda}" if cl.factor is not None
            else f"{cl.concept} = {_fmt(cl.total)} {cl.moneda}" for cl in dudosas)
        res.reasons.append(
            f"No se puede identificar qué parte del flete es de cada carga ({' + '.join(groups)}). "
            f"Líneas sin asignar: {det}")
        hint = _tariff_hint(bl, dudosas)
        if hint:
            res.suggestion = "Indicio: " + hint
    elif metodo:
        res.notes.append("Flete separado por " + " + ".join(metodo))
    return asign


# ───────────────────────────── evaluación ─────────────────────────────

def evaluate(bl: BLRecord, op: str, cfg: config.EngineConfig,
             catalog: Optional[Catalog] = None) -> CommissionResult:
    catalog = catalog or load_catalog()
    op = op.upper()
    port = bl.port_of_loading if op == "IMPO" else bl.port_of_discharge
    res = CommissionResult(bl=bl, op=op, port=port or "SIN PUERTO")
    rates = cfg.rates_for(op)
    roe = cfg.roe

    # 1. Conceptos: los nuevos se registran y no influyen; los indefinidos van a revisión.
    for cl in bl.charge_lines:
        r = concept_rule(cl, op, catalog)
        if r is None:
            if cl.concept not in [n for n, _ in res.new_concepts]:
                res.new_concepts.append((cl.concept, _example(bl, cl)))
        elif r[0] is None or r[1] is None:
            res.reasons.append(f"Concepto de cargo sin definir en CONCEPTOS DE CARGO: '{cl.concept}'")
    if res.new_concepts:
        res.notes.append("Concepto nuevo registrado, sin efecto en la comisión: "
                         + ", ".join(n for n, _ in res.new_concepts))
    for k in bl.unknown_kinds:
        res.reasons.append(f"Tipo de bulto nuevo, no está en TIPOS DE BULTO: '{k}'")
    for it in bl.items:
        rule = catalog.kind(it.kind)
        if rule is not None and not rule.defined:
            res.reasons.append(f"Tipo de bulto con definición incompleta en TIPOS DE BULTO: '{it.kind}'")

    # 2. Monto comisionable y moneda (USD + EUR: el EUR pasa a USD con el ROE).
    lines = _commissionable_lines(bl, catalog, op)
    for cl in lines:
        if concept_rule(cl, op, catalog)[1] == RESTA and cl.total > 0:
            res.reasons.append(f"Descuento con importe positivo: {cl.concept} {_fmt(cl.total)} {cl.moneda}")
    monedas = sorted({cl.moneda for cl in lines if cl.moneda})
    amt: Dict[int, float] = {id(cl): cl.total for cl in lines}
    if set(monedas) == {"USD", "EUR"}:
        for cl in lines:
            if cl.moneda == "EUR":
                amt[id(cl)] = cl.total * roe
        eur = ", ".join(f"{cl.concept} {_fmt(cl.total)} EUR" for cl in lines if cl.moneda == "EUR")
        res.notes.append(f"Montos en USD y EUR: {eur} pasado a USD con ROE {roe:g}")
        res.currency = "USD"
    else:
        res.currency = monedas[0] if monedas else (bl.charge("Basic FRT").moneda or "USD")
    res.gross = sum(amt.values())
    res.net = max(0.0, res.gross)
    if res.net > 0:
        if len(monedas) > 1 and set(monedas) != {"USD", "EUR"}:
            res.reasons.append(f"El monto comisionable mezcla monedas: {', '.join(monedas)}")
        elif res.currency not in config.COMMISSIONABLE_CURRENCIES:
            res.reasons.append(f"Comisionable en {res.currency}: no es USD ni EUR")

    # 3. ALL IN (todo el B/L): el 80% se aplica sobre el Basic Frt.
    has_basic = any(_is_basic(cl) for cl in bl.charge_lines)
    for cl in bl.charge_lines:
        r = concept_rule(cl, op, catalog)
        if r is None:
            continue
        if r[0] == RECARGO or (r[1] == SUMA and not _is_basic(cl)):
            if cl.concept not in res.all_in_breakers:
                res.all_in_breakers.append(cl.concept)
    res.all_in = has_basic and not res.all_in_breakers
    factor = config.ALL_IN_FACTOR if res.all_in else 1.0

    # 4. Carga del B/L.
    families = bl.families()
    if bl.equipment_items and BBULK not in families:
        res.reasons.append("Trae Mafi pero no se identificó el Bbulk que lleva encima")
    groups = []
    for fam in families:
        g = _group_of_family(fam)
        if g and g not in groups:
            groups.append(g)
    res.groups = groups
    if not groups:
        if res.net > 0:
            res.reasons.append("Hay flete comisionable pero no se identificó la carga del B/L")
        res.status = "REVISION" if res.reasons else "OK"
        return res

    # 5. Modo de los rodados.
    if G_ROD in groups:
        if cfg.agency_bls is None:
            res.rodado_mode = MODE_ESTIMADO
            res.notes.append("Estimado: no se ingresó la lista de Comercial (calculado por unidad)")
        elif bl.bl_clean in cfg.agency_bls:
            res.rodado_mode = MODE_AGENCY
        else:
            res.rodado_mode = MODE_UNIT

    # 6. Corroboración con lo facturado (la carga tiene que coincidir con el Basic Frt).
    basics = [cl for cl in bl.charge_lines if _is_basic(cl)]
    if groups == [G_ROD] and basics and all(cl.basis == "PU" for cl in basics):
        facturadas = sum(cl.factor or 0 for cl in basics)
        unidades = sum(bl.units(f) for f in RODADOS)
        if abs(facturadas - unidades) > TOL:
            res.reasons.append(f"Rodados: el manifiesto declara {unidades} unidades y el Basic Frt "
                               f"factura {facturadas:g} PU")
    if groups == [G_CTR] and basics and all(cl.basis in ("20", "40") for cl in basics):
        facturados = sum(cl.factor or 0 for cl in basics)
        ctrs = bl.units(CONTENEDOR)
        if abs(facturados - ctrs) > TOL:
            res.reasons.append(f"Contenedores: el manifiesto declara {ctrs} y el Basic Frt factura "
                               f"{facturados:g}")

    # 7. Brasil (sólo EXPO, sólo si hay porcentaje que depende del destino).
    if op == "EXPO" and (G_CTR in groups or G_BBULK in groups):
        prule = catalog.port(bl.port_of_discharge)
        if prule is None or prule.brasil is None:
            res.reasons.append(f"Puerto de descarga sin definir en PUERTOS EXPO: '{bl.port_of_discharge}'")
        else:
            res.brasil = prule.brasil

    # 8. Reparto del flete entre las cargas que lo consumen.
    asign = _assign(bl, lines, groups, res) if lines else {}
    res.lines = lines
    res.assignment = asign
    res.split_ok = all(id(cl) in asign for cl in lines)

    # 9. Comisión por parte.
    to_usd = (lambda v: v * roe) if res.currency == "EUR" else (lambda v: v)
    for g in groups:
        propias = [cl for cl in lines if asign.get(id(cl)) == g]
        p = Part(group=g, currency=res.currency)
        p.base_manifest = max(0.0, sum(amt[id(cl)] for cl in propias))
        p.base = max(0.0, sum(amt[id(cl)] * (factor if _is_basic(cl) else 1.0) for cl in propias))
        if g == G_CTR:
            p.units = {CONTENEDOR: bl.units(CONTENEDOR)}
            p.rate = rates.brasil_containers if res.brasil else rates.containers
            p.min_ctr_units = sum(
                it.qty for it in bl.cargo_items
                if it.family == CONTENEDOR and catalog.kind(it.kind) and catalog.kind(it.kind).min_ctr
            )
            p.calc_usd = to_usd(p.base) * p.rate
            p.minimum_usd = p.min_ctr_units * rates.container_min_per_unit
            p.commission_usd = max(p.calc_usd, p.minimum_usd)
        elif g == G_BBULK:
            p.units = {BBULK: bl.units(BBULK)}
            p.rate = rates.brasil_general_cargo if res.brasil else rates.general_cargo
            p.calc_usd = to_usd(p.base) * p.rate
            p.commission_usd = p.calc_usd
        else:
            p.mode = res.rodado_mode
            p.units = {f: bl.units(f) for f in RODADOS if bl.units(f)}
            if p.mode == MODE_AGENCY:
                p.rate = rates.rolling
                p.calc_usd = to_usd(p.base) * p.rate
            else:
                p.base = p.base_manifest = 0.0
                p.calc_usd = sum(n * config.PER_UNIT_TARIFF_USD[f] for f, n in p.units.items())
            p.commission_usd = p.calc_usd
        res.parts.append(p)

    res.cef_units = sum(
        it.qty for it in bl.cargo_items
        if it.family == CONTENEDOR and catalog.kind(it.kind) and catalog.kind(it.kind).cef
    )

    if res.all_in:
        res.notes.append("ALL IN: se comisiona el 80% del Basic Frt")
    if res.brasil:
        pct_br = f"{rates.brasil_containers * 100:g}%".replace(".", ",")
        res.notes.append(f"Destino Brasil: {pct_br}")
    res.status = "REVISION" if res.reasons else "OK"
    return res


def evaluate_all(bls: List[BLRecord], op: str, cfg: config.EngineConfig,
                 catalog: Optional[Catalog] = None) -> List[CommissionResult]:
    catalog = catalog or load_catalog()
    return [evaluate(bl, op, cfg, catalog) for bl in bls]


# ───────────────── compatibilidad con utilidades ─────────────────

def commission_family(bl: BLRecord) -> str:
    fams = bl.families()
    if CONTENEDOR in fams:
        return "container"
    if any(f in RODADOS for f in fams) and BBULK not in fams:
        return "rolling"
    return "general"


def commission_rate(bl: BLRecord, op_type: str, cfg: config.EngineConfig) -> float:
    rates = cfg.rates_for(op_type)
    fam = commission_family(bl)
    return {"container": rates.containers, "rolling": rates.rolling}.get(fam, rates.general_cargo)


def commissionable_gross(bl: BLRecord, catalog: Optional[Catalog] = None) -> float:
    catalog = catalog or load_catalog()
    return sum(cl.total for cl in _commissionable_lines(bl, catalog, bl.op_type))


def commission_amount(bl: BLRecord, op_type: str, cfg: config.EngineConfig,
                      catalog: Optional[Catalog] = None) -> Optional[float]:
    return evaluate(bl, op_type, cfg, catalog=catalog).total_usd
