"""
models.py - Modelo de datos tipado del motor.

Un BLRecord es la unidad de trabajo: todo lo que el manifiesto informa sobre un
B/L (bultos con medidas, líneas de cargo estructuradas, totales por lugar de
pago, CUIT validado y chasis/VINs) vive en dataclases con accesores explícitos.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class Charge:
    """Un cargo agregado del B/L (Basic FRT, THC, Toll, Sweeping, BAF, ENS)."""
    monto: float = 0.0
    letra: Optional[str] = None     # 'P' prepaid / 'C' collect
    moneda: Optional[str] = None    # 'USD', 'EUR', ...

    @property
    def is_prepaid(self) -> bool:
        return self.letra == "P"

    @property
    def is_collect(self) -> bool:
        return self.letra == "C"

    def __bool__(self) -> bool:
        return bool(self.monto)


@dataclass
class ChargeLine:
    """Una línea individual del bloque 'Charge Information' del manifiesto."""
    desc: str                       # normalizada: minúsculas y sin puntos
    total: float
    letra: Optional[str] = None
    moneda: Optional[str] = None
    factor: Optional[float] = None  # cantidad facturada (Tns, contenedores, ...)
    basis: Optional[str] = None     # 'MT', 'WMM', 'CM', 'PU', '20', '40', ...
    rate: Optional[float] = None    # tarifa unitaria
    concept: str = ""               # nombre completo con encabezado ('Toll S/C(POD)')
    tariff: Optional[str] = None    # 'Tariff ID' de la cotización a la que pertenece
    pct_of: Optional[float] = None  # en cargos 'x% of y': el importe base 'y'


@dataclass
class TotalLine:
    """Una línea 'Total[Buenos Aires] 11,035.09C USD' del manifiesto."""
    place: str                      # 'Buenos Aires', 'Buenos Aires - Partner', 'Casa Matriz', ...
    amount: float
    letra: str                      # 'P' | 'C'
    moneda: str


@dataclass
class CargoItem:
    """Un bulto declarado en el manifiesto ('27-New Small Van(s)')."""
    qty: int
    kind: str                       # tal cual el manifiesto
    family: Optional[str] = None    # CONTENEDOR | BBULK | CAR | VAN | RORO (según catálogo)
    equipment: bool = False         # Mafi: lleva carga encima, no es carga
    tn: Optional[float] = None
    cbm: Optional[float] = None
    lm: Optional[float] = None

    def metric(self, basis: str) -> Optional[float]:
        """Medida del bulto comparable con el factor de una línea de flete."""
        if basis == "WMM":
            vals = [v for v in (self.tn, self.cbm) if v is not None]
            return max(vals) if vals else None
        if basis == "CM":
            return self.cbm
        if basis == "MT":
            return self.tn
        if basis in ("ME", "LM"):
            return self.lm
        if basis == "PU":
            return float(self.qty)
        return None


@dataclass
class CargoData:
    """Qué viaja en el B/L, desagregado en las tres familias clásicas."""
    containers: Dict[str, int] = field(default_factory=dict)
    vehicles: Dict[str, int] = field(default_factory=dict)
    general: Dict[str, int] = field(default_factory=dict)

    def add(self, bucket: str, key: str, qty: int) -> None:
        d = getattr(self, bucket)
        d[key] = d.get(key, 0) + qty

    @property
    def n_containers(self) -> int:
        return sum(self.containers.values())

    @property
    def n_vehicles(self) -> int:
        return sum(self.vehicles.values())

    @property
    def n_general(self) -> int:
        return sum(self.general.values())

    @property
    def n_roro(self) -> int:
        return sum(v for k, v in self.vehicles.items() if "roro" in k.lower())

    @property
    def n_cars(self) -> int:
        return sum(v for k, v in self.vehicles.items()
                   if "car" in k.lower() or "van" in k.lower())

    @property
    def is_empty(self) -> bool:
        return not (self.containers or self.vehicles or self.general)


def cargo_data(items: List[CargoItem]) -> CargoData:
    """Resumen por familia de un conjunto de bultos (todo el B/L o una de sus cargas)."""
    c = CargoData()
    for it in items:
        usado = it.kind.lower().startswith("used")
        if it.family == "CONTENEDOR":
            c.add("containers", it.kind.lower(), it.qty)
        elif it.family == "CAR":
            c.add("vehicles", "used car" if usado else "new car", it.qty)
        elif it.family == "VAN":
            c.add("vehicles", "used van" if usado else "new van", it.qty)
        elif it.family == "RORO":
            c.add("vehicles", "used roro" if usado else "new roro", it.qty)
        elif it.family == "BBULK":
            c.add("general", it.kind.lower().replace("(s)", "s"), it.qty)
    return c


@dataclass
class BLRecord:
    """Un B/L del manifiesto, ya parseado y consolidado."""
    bl_no: str
    op_type: str = "IMPO"
    port_of_loading: str = "Nulo"
    port_of_discharge: str = "Nulo"
    entity: str = "Nulo"

    charges: Dict[str, Charge] = field(default_factory=dict)
    thc_20: Charge = field(default_factory=Charge)
    thc_40: Charge = field(default_factory=Charge)
    charge_lines: List[ChargeLine] = field(default_factory=list)

    negative_adjustments: float = 0.0   # rebate/adj/fac/baf decrease (<= 0)
    negative_basic_frt: float = 0.0     # líneas Basic Frt. negativas

    cargo: CargoData = field(default_factory=CargoData)
    totals: Dict[str, object] = field(default_factory=dict)
    collect: List[Dict[str, object]] = field(default_factory=list)
    total_lines: List[TotalLine] = field(default_factory=list)
    party_lines: List[str] = field(default_factory=list)  # columna SHIPPER/CONSIGNEE/NOTIFY
    party_id: str = ""                  # CUIT del cliente (CN en IMPO, SH en EXPO), si figura

    weight: float = 0.0                 # kg brutos acumulados
    description_lines: List[str] = field(default_factory=list)
    vins: List[str] = field(default_factory=list)
    sc_loose_rates: List[float] = field(default_factory=list)

    items: List[CargoItem] = field(default_factory=list)
    unknown_kinds: List[str] = field(default_factory=list)   # bultos que no están en el catálogo

    # ── accesores ────────────────────────────────────────────────────
    @property
    def bl_clean(self) -> str:
        """Número de B/L sin el sufijo [T] de transbordo."""
        return self.bl_no.replace("[T]", "")

    def charge(self, key: str) -> Charge:
        """Cargo agregado por nombre (case-insensitive); vacío si no existe."""
        return self.charges.get(key.lower(), Charge())

    def set_charge(self, key: str, monto: float, letra=None, moneda=None) -> None:
        c = self.charges.setdefault(key.lower(), Charge())
        c.monto += monto
        if letra:
            c.letra = letra
        if moneda:
            c.moneda = moneda

    @property
    def weight_tons(self) -> float:
        return self.weight / 1000.0

    @property
    def description(self) -> str:
        return " ".join(self.description_lines).upper()

    @property
    def cargo_items(self) -> List[CargoItem]:
        """Bultos que son carga (excluye el equipo: Mafis)."""
        return [it for it in self.items if not it.equipment]

    @property
    def equipment_items(self) -> List[CargoItem]:
        return [it for it in self.items if it.equipment]

    def units(self, family: str) -> int:
        return sum(it.qty for it in self.cargo_items if it.family == family)

    def families(self) -> List[str]:
        out = []
        for it in self.cargo_items:
            if it.family and it.family not in out:
                out.append(it.family)
        return out

    def rebuild_cargo(self) -> None:
        """
        Deriva el resumen por familia (containers / vehicles / general) que usan
        las biblias a partir de los bultos clasificados con el catálogo.
        """
        self.cargo = cargo_data(self.cargo_items)

    def merge(self, other: "BLRecord") -> None:
        """
        Absorbe otro registro del MISMO B/L.

        El manifiesto repite el número de B/L al continuar en otra página, por lo
        que un mismo B/L puede aparecer partido en varios bloques: uno con la
        carga y otro con los cargos. Se consolidan en un único registro.
        """
        if self.entity in ("", "Nulo") and other.entity not in ("", "Nulo"):
            self.entity = other.entity
        for attr in ("port_of_loading", "port_of_discharge"):
            if getattr(self, attr) in ("", "Nulo"):
                setattr(self, attr, getattr(other, attr))

        for key, ch in other.charges.items():
            self.set_charge(key, ch.monto, ch.letra, ch.moneda)
        for mio, suyo in ((self.thc_20, other.thc_20), (self.thc_40, other.thc_40)):
            mio.monto += suyo.monto
            mio.letra = mio.letra or suyo.letra
            mio.moneda = mio.moneda or suyo.moneda

        for bucket in ("containers", "vehicles", "general"):
            for k, v in getattr(other.cargo, bucket).items():
                self.cargo.add(bucket, k, v)

        self.charge_lines.extend(other.charge_lines)
        self.items.extend(other.items)
        self.unknown_kinds.extend(k for k in other.unknown_kinds if k not in self.unknown_kinds)
        self.negative_adjustments += other.negative_adjustments
        self.negative_basic_frt += other.negative_basic_frt
        self.weight += other.weight
        self.collect.extend(other.collect)
        self.total_lines.extend(other.total_lines)
        self.party_lines.extend(other.party_lines)
        if not self.party_id:
            self.party_id = other.party_id
        self.sc_loose_rates.extend(other.sc_loose_rates)
        self.description_lines.extend(other.description_lines)
        for vin in other.vins:
            if vin not in self.vins:
                self.vins.append(vin)

        for k, v in other.totals.items():
            actual = self.totals.get(k)
            if actual is None:
                self.totals[k] = v
            elif isinstance(v, (int, float)) and isinstance(actual, (int, float)):
                self.totals[k] = actual + v
