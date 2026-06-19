"""
config.py - Única fuente de verdad de los parámetros de negocio "modificables".

Todo lo que el negocio puede cambiar (tarifas de gastos de línea, rates de
comisión, mapeos de marcas, abreviaturas de puerto, etc.) vive acá. El resto
del motor sólo lee de este módulo; ninguna regla hardcodea estos valores.

Los valores por defecto replican el comportamiento del script original
(manifest_processor_v3.py) salvo donde el plan indica una corrección.
"""

from dataclasses import dataclass, field
from typing import Dict, List, Tuple


# ═══════════════════════════════════════════════════════════════════════
# GASTOS DE LÍNEA (modificables) — modo híbrido: se leen del manifiesto y se
# validan contra estas tarifas esperadas. Ver rules/line_costs.py.
# ═══════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class LineCostTariffs:
    """
    Tarifas FIJAS de los gastos de línea (USD). Todo importe del manifiesto que
    no coincida con la tarifa esperada es un ERROR DE CARGA DE ORIGEN y se marca
    para revisión (no se recalcula sobre el dato erróneo).

    NOTA: estos valores son **ilustrativos/ficticios** (versión de portafolio).
    Las tarifas reales se inyectan por configuración; no se publican.
    """
    thc_20: float = 100.0          # THC POD/POL por contenedor de 20 Tns (ilustrativo)
    thc_40: float = 120.0          # THC POD/POL por contenedor de 40 Tns (ilustrativo)
    sc_container: float = 150.0    # S/C (TOLL) por contenedor (20 y 40) (ilustrativo)
    sc_loose_per_ton: float = 20.0  # S/C (TOLL) por Tn en carga suelta (ilustrativo)
    sweeping: float = 10.0         # SWEEPING por contenedor (ilustrativo)

    # Tolerancia (USD) al comparar importe del manifiesto vs tarifa esperada.
    tolerance: float = 0.5


# ═══════════════════════════════════════════════════════════════════════
# COMISIONES — rates por operación y tipo de carga.
# ═══════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class CommissionRates:
    rolling: float           # cars / big vans
    general_cargo: float     # roros + carga proyecto
    containers: float        # contenedores
    container_min_per_unit: float = 25.0  # USD por contenedor (mínimo, ilustrativo)


# Rates de comisión por operación y tipo de carga. Valores **ilustrativos**
# (versión de portafolio); los reales se inyectan por configuración.
COMMISSION_RATES: Dict[str, CommissionRates] = {
    "IMPO": CommissionRates(rolling=0.01, general_cargo=0.02, containers=0.02),
    "EXPO": CommissionRates(rolling=0.01, general_cargo=0.04, containers=0.04),
}


# ═══════════════════════════════════════════════════════════════════════
# CONVERSIÓN DE MONEDA
# ═══════════════════════════════════════════════════════════════════════

DEFAULT_ROE: float = 1.22  # EU 1,00 = X USD (TC pactado). Aplicado en Excel.

# Monedas distintas de USD que pueden aparecer como comisionables.
FOREIGN_CURRENCIES: Tuple[str, ...] = ("EUR", "BRL", "GBP")


# ═══════════════════════════════════════════════════════════════════════
# MONTO COMISIÓN — qué cargos suman/restan a la base comisionable.
# (Pendiente de confirmación de negocio A1/A2; valores base = código original.)
# ═══════════════════════════════════════════════════════════════════════

# Descripción base siempre comisionable.
COMMISSIONABLE_BASE: str = "basic frt"

# Cargos adicionales comisionables (además de Basic Frt.). Se identifican por
# prefijo de la descripción de la línea de cargo. En los datos reales aparecen
# como "Open Top S/C"; "Over Height"/"Over Weight" se incluyen por robustez aún
# sin casos observados (variantes de ortografía contempladas).
COMMISSIONABLE_EXTRA: Tuple[str, ...] = (
    "open top", "over height", "over high", "over hght", "over weight", "over wight",
)

# Prefijos de descripción que representan descuentos (restan al comisionable).
NEGATIVE_PREFIXES: Tuple[str, ...] = ("rebate", "adj", "fac", "baf decrease")


# ═══════════════════════════════════════════════════════════════════════
# CLASIFICACIÓN DE CARGA
# ═══════════════════════════════════════════════════════════════════════

CONTAINER_KEYWORDS: Tuple[str, ...] = (
    "Tank Container", "Dry Cargo", "High Cube", "Open Top",
)

GENERAL_KEYWORDS: Tuple[str, ...] = (
    "PALLET(S)", "UNIT(S)", "PACKAGE(S)", "BOX(S)", "CASE(S)",
    "BUNDLE(S)", "CRATE(S)", "BAG(S)", "PIECE(S)", "UNPACKED(S)",
    "STEEL(S)", "MAFI(S)",
)


# Mapeo nombre de shipper/consignee -> marca. El orden importa: se evalúa de
# arriba hacia abajo (primer match gana).
# NOTA: las marcas y razones sociales son **ficticias** (versión de portafolio).
BRAND_MAPPING: Dict[str, str] = {
    "AUTOMOTRIZ AVALON": "AVALON",
    "GRUPO BIONDA": "BIONDA",
    "CORVEX": "CORVEX",
    # El importador del grupo trae mayormente DELTAR -> marca por defecto DELTAR;
    # las excepciones (MISTRAL) se marcan para REVISIÓN. Las variantes específicas
    # (DELCAR/PREMIUMCARS) se evalúan con sus propias claves más abajo.
    "IMPORTADORA DELTAR": "DELTAR",
    "PREMIUM AUTO": "EVORA",
    "DELCAR": "DELTAR - DELCAR",
    "NORTHWAGEN S.A.": "FENIX",
    "GRIFON": "GRIFON",
    "TO THE ORDER OF PREMIUMCARS GROUP": "DELTAR - PREMIUMCARS",
    "HELIOS ARGENTINA SA": "HELIOS",
    "AGENCIA MARITIMA OCEANLINK": "OCEANLINK",
    "MILANO MOTORS S.A.": "IONIC",
    "EXIMPORT INTL SA": "JUNO",
    "PEGASO": "PEGASO",
    "NOVAX INTERNATIONAL DEVELOPMENT": "NOVAX",
    "ORBIS": "ORBIS",
}


# Marca/tipo del vehículo: se extrae de la DESCRIPCIÓN de la mercadería.
# Lista ordenada (primer token encontrado gana); valor = nombre canónico/abrev.
# Las multi-palabra y prefijos VIN deben ir antes que los genéricos. Marcas ficticias.
MAKE_KEYWORDS: Tuple[Tuple[str, str], ...] = (
    ("EVORA-PREMIUM", "EVORA"), ("EVORA", "EVORA"),
    ("IONIC", "IONIC"),
    ("KRONOS", "KRONOS"), ("JUNO", "JUNO"),
    ("DELTAR", "DELTAR"), ("ZZDLTR", "DELTAR"),   # ZZDLTR = prefijo VIN ficticio
    ("FENIX", "FENIX"),
    ("MINIC", "CORVEX"), ("CORVEX", "CORVEX"),
    ("BIONDA", "BIONDA"),
    ("AVALON", "AVALON"),
    ("LYRA", "LYRA"),
    ("MISTRAL", "MISTRAL"),
    ("NOVAX", "NOVAX"), ("ORBIS", "ORBIS"),
    ("GRIFON", "GRIFON"), ("PEGASO", "PEGASO"), ("QUASAR", "QUASAR"),
)

# Sufijo de agencia/forwarder según el consignee (se antepone " - <SUFIJO>").
AGENCY_SUFFIXES: Tuple[Tuple[str, str], ...] = (
    ("OCEANLINK", "OCEANLINK"),
    ("DELCAR", "DELCAR"),
    ("PREMIUMCARS", "PREMIUMCARS"),
    ("EXIMPORT", "EXIMPORT"),
)


# ═══════════════════════════════════════════════════════════════════════
# PUERTOS Y BUQUES
# ═══════════════════════════════════════════════════════════════════════

PORT_ABBREVIATIONS: Dict[str, str] = {
    "ZARATE": "ZTE", "BUENOS AIRES": "BUE", "ROSARIO": "ROS",
    "MONTEVIDEO": "MVD", "SANTOS": "STS", "RIO": "RIO",
    "PARANAGUA": "PNG", "ITAJAI": "ITJ",
}

# Flota ficticia (los nombres reales de la línea fueron anonimizados).
SHIP_NAMES: Dict[str, str] = {
    "ATL": "ATLANTIC STAR",
    "ABA": "ATLANTIC BAHIA",
    "ANI": "ATLANTIC NEBULA",
    "ASP": "ATLANTIC SAN PEDRO",
    "AFR": "ATLANTIC FRONTIER",
    "ATX": "ATLANTIC TEXAS",
    "AHO": "ATLANTIC HORIZON",
}


# ═══════════════════════════════════════════════════════════════════════
# OTROS
# ═══════════════════════════════════════════════════════════════════════

ENS_FIXED_AMOUNT: float = 10.0  # ENS EXPO: monto fijo por BL (USD, ilustrativo).

# RORO que, pese a ser RORO, cuenta como ROLLING (los motorhomes/campers ruedan;
# el resto de los RORO van a GENERAL CARGO). Se busca en la descripción.
ROLLING_RORO_KEYWORDS: Tuple[str, ...] = ("MOTORHOME", "MOTOR HOME", "CAMPER")


@dataclass
class EngineConfig:
    """Agrupa la configuración de una corrida; permite override puntual."""
    tariffs: LineCostTariffs = field(default_factory=LineCostTariffs)
    roe: float = DEFAULT_ROE
    commission_rates: Dict[str, CommissionRates] = field(
        default_factory=lambda: dict(COMMISSION_RATES)
    )

    @property
    def tonnage_factor(self) -> float:
        """Divisor del PESO en la planilla (=TOLL/factor): la tarifa fija suelta."""
        return self.tariffs.sc_loose_per_ton

    def rates_for(self, op_type: str) -> CommissionRates:
        return self.commission_rates[op_type.upper()]
