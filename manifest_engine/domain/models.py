"""
models.py - Modelo de datos tipado de un Bill of Lading (BL).

Reemplaza los dict/defaultdict anidados del parser original por dataclasses
explícitas. Los nombres de campo conservan la semántica del código v3 para que
la migración del parser sea directa, pero el acceso es tipado y con helpers.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass
class Charge:
    """Un cargo del BL (Basic Frt., THC, Toll, Sweeping, BAF, ENS...)."""
    monto: float = 0.0
    letra: Optional[str] = None   # 'P' (prepaid) / 'C' (collect)
    moneda: Optional[str] = None  # 'USD' / 'EUR' / 'BRL' / 'GBP'

    def add(self, monto: float, letra: Optional[str], moneda: Optional[str]) -> None:
        self.monto += monto
        if letra is not None:
            self.letra = letra
        if moneda is not None:
            self.moneda = moneda


@dataclass
class TotalLine:
    """Una línea Total[Pais] del manifiesto."""
    pais: str
    monto: float
    letra: str       # 'P' / 'C'
    moneda: str


@dataclass
class ChargeLine:
    """
    Una línea estructurada de la sección de cargos, con sus columnas
    DESCRIPTION / FACTOR / BASIS / RATE / TOTAL / cond / CURR.

    Permite separar el comisionable por tipo de carga (basis 20/40 =
    contenedor; otro basis = general/autos) y validar cantidades (B1/B2),
    y derivar el factor del TOLL para la columna PESO (C).
    """
    desc: str          # normalizada en minúsculas, sin punto final
    factor: float      # FACTOR (nº de unidades / tonelaje según basis)
    basis: str         # '20' / '40' / 'MT' / 'PU' / 'ME' / 'WMM' / 'CM' / ...
    rate: float        # RATE por unidad
    total: float       # TOTAL de la línea
    letra: str         # 'P' / 'C'
    moneda: str        # 'USD' / 'EUR' / ...

    @property
    def is_container_basis(self) -> bool:
        return self.basis in ("20", "40")


@dataclass
class CargoData:
    """Conteo de carga del BL, separado por tipo."""
    containers: Dict[str, int] = field(default_factory=dict)
    vehicles: Dict[str, int] = field(default_factory=dict)
    general: Dict[str, int] = field(default_factory=dict)

    def add_container(self, kind: str, qty: int) -> None:
        self.containers[kind] = self.containers.get(kind, 0) + qty

    def add_vehicle(self, kind: str, qty: int) -> None:
        self.vehicles[kind] = self.vehicles.get(kind, 0) + qty

    def add_general(self, kind: str, qty: int) -> None:
        self.general[kind] = self.general.get(kind, 0) + qty

    @property
    def container_count(self) -> int:
        return sum(self.containers.values())

    @property
    def vehicle_count(self) -> int:
        return sum(self.vehicles.values())

    @property
    def roro_count(self) -> int:
        return sum(v for k, v in self.vehicles.items() if "roro" in k.lower())

    @property
    def is_empty(self) -> bool:
        return not (self.containers or self.vehicles or self.general)


@dataclass
class BLRecord:
    """Un Bill of Lading completo extraído del manifiesto."""
    bl_no: str
    port_of_loading: str = "Nulo"
    port_of_discharge: str = "Nulo"
    entity: str = "Nulo"

    charges: Dict[str, Charge] = field(default_factory=dict)
    thc_20: Charge = field(default_factory=Charge)
    thc_40: Charge = field(default_factory=Charge)

    total_ajustes_negativos: float = 0.0
    montos_negativos_basic_frt: float = 0.0

    cargo: CargoData = field(default_factory=CargoData)
    totals: Dict[str, object] = field(default_factory=dict)
    collect: List[Dict[str, object]] = field(default_factory=list)
    weight: float = 0.0

    # Líneas de cargo estructuradas (FACTOR/BASIS/RATE), para comisionable
    # partido por basis y derivación del factor del TOLL.
    charge_lines: List[ChargeLine] = field(default_factory=list)

    # Líneas de la descripción de mercadería (col 4); se usan para extraer la
    # marca/tipo del vehículo (DELTAR/MISTRAL/LYRA...) y descriptores.
    description_lines: List[str] = field(default_factory=list)

    # Chasis/VIN del BL (autos y RoRo). Tokens de 17 chars hallados en las
    # columnas Unit No. (col 3) y Description (col 4). Orden de aparición, sin
    # duplicados. Vacío si el manifiesto no imprime los VIN de ese BL.
    vins: List[str] = field(default_factory=list)

    def add_vin(self, vin: str) -> None:
        if vin not in self._vin_seen:
            self._vin_seen.add(vin)
            self.vins.append(vin)

    _vin_seen: set = field(default_factory=set, repr=False, compare=False)

    def charge(self, key: str) -> Charge:
        """Devuelve (creando si hace falta) el Charge para `key`."""
        if key not in self.charges:
            self.charges[key] = Charge()
        return self.charges[key]
