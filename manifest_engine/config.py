"""
config.py - Única fuente de verdad de los parámetros de negocio "modificables".

Todo lo que el negocio puede cambiar (tarifas de gastos de línea, tasas de
comisión, mapeos de marcas, abreviaturas de puerto, nombres de buques) vive acá.
El resto del motor sólo lee de este módulo y del catálogo Excel
(`Definiciones_reglas_comisiones.xlsx`); ninguna regla hardcodea estos valores.

NOTA DE PORTAFOLIO: todas las tarifas, tasas, razones sociales y buques en este
módulo son valores **ilustrativos y ficticios** para demostración técnica
(estándar Atlantic Maritime Agency — AMA).
"""

from dataclasses import dataclass, field
from typing import Dict, FrozenSet, Optional, Tuple


# ═══════════════════════════════════════════════════════════════════════
# GASTOS DE LÍNEA (modificables) — modo híbrido: se leen del manifiesto y se
# validan contra estas tarifas esperadas. Ver rules/line_costs.py.
# ═══════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class LineCostTariffs:
    """
    Tarifas FIJAS de los gastos de línea (USD, valores ilustrativos de portafolio).
    El TOLL es fijo: 20.0/Tn en carga suelta (project/roro/autos) y 150.0/contenedor.
    Cualquier importe del manifiesto que no coincida con la tarifa esperada es un
    ERROR DE CARGA DE ORIGEN y se marca para revisión (no se recalcula sobre el
    dato erróneo).

    Nota histórica ilustrativa: el S/C suelto se cobraba 18.0/Tn en viajes
    anteriores (ej. ATL 0125); el motor auto-detecta la tarifa vigente del viaje
    a partir de las líneas MT del manifiesto.
    """
    thc_20: float = 100.0          # THC POD/POL por contenedor de 20' (ilustrativo)
    thc_40: float = 120.0          # THC POD/POL por contenedor de 40' (ilustrativo)
    sc_container: float = 150.0    # S/C (TOLL) por contenedor (20' y 40') (ilustrativo)
    sc_loose_per_ton: float = 20.0  # S/C (TOLL) por Tn en carga suelta (ilustrativo; hist. 18.0)
    sweeping: float = 10.0         # SWEEPING por contenedor (20' y 40') (ilustrativo)

    # Tolerancia (USD) al comparar importe del manifiesto vs tarifa esperada.
    tolerance: float = 0.5


# ═══════════════════════════════════════════════════════════════════════
# COMISIONES — tasas por operación y tipo de carga (valores ilustrativos).
# ═══════════════════════════════════════════════════════════════════════

@dataclass(frozen=True)
class CommissionRates:
    rolling: float                        # rodados (car / van / roro) cuyo booking cerró AMA
    general_cargo: float                  # Bbulk (carga suelta)
    containers: float                     # contenedores
    container_min_per_unit: float = 25.0  # USD por contenedor (mínimo ilustrativo)
    brasil_general_cargo: float = 0.0     # Bbulk con destino Brasil (sólo EXPO)
    brasil_containers: float = 0.0        # contenedores con destino Brasil (sólo EXPO)


# Tasas ilustrativas de portafolio:
#   IMPO: contenedores 2,0%, Bbulk 2,0%.
#   EXPO: contenedores 4,0% (Brasil 3,5%), Bbulk 4,0% (Brasil 3,5%).
#   Rodados cerrados por la agencia (AMA): 1,0% en ambos.
COMMISSION_RATES: Dict[str, CommissionRates] = {
    "IMPO": CommissionRates(rolling=0.01, general_cargo=0.02, containers=0.02),
    "EXPO": CommissionRates(
        rolling=0.01,
        general_cargo=0.04,
        containers=0.04,
        brasil_general_cargo=0.035,
        brasil_containers=0.035,
    ),
}

# Rodados NO cerrados por la agencia: tarifa fija ilustrativa en USD por unidad
# (IMPO y EXPO), nuevos o usados, con o sin Basic Frt (siempre que el B/L sea S3...).
PER_UNIT_TARIFF_USD: Dict[str, float] = {"CAR": 2.5, "VAN": 3.5, "RORO": 12.0}

# ALL IN: el B/L sólo trae Basic Frt como flete -> se comisiona el 80% del Basic Frt.
ALL_IN_FACTOR: float = 0.80

# Clasificación que depende de la operación. Un gasto en el puerto del EXTERIOR
# es parte del flete (rompe ALL IN); el del puerto local no. El catálogo Excel
# está expresado desde EXPO (THC POL = local, THC POD = exterior); en IMPO es al
# revés: el THC (POL) es en origen (afuera) y rompe ALL IN, mientras que THC (POD)
# es local y no lo afecta (ej. caso sintético S30000105 en ATL 0426).
OP_CONCEPT_OVERRIDES: Dict[str, Dict[str, Tuple[str, str]]] = {
    "IMPO": {
        "THC (POL)": ("RECARGO", "NO AFECTA"),   # origen, afuera: parte del flete
        "THC (POD)": ("GASTO", "NO AFECTA"),     # Zárate, local: no afecta ALL IN
    },
}

# Control Equipment Fees: USD por contenedor (IMPO y EXPO, ilustrativo). Los Mafis no cuentan.
CONTROL_EQUIPMENT_FEE_USD: float = 4.0

# Tipos de bulto que son EQUIPO (llevan la carga encima), no la carga en sí.
EQUIPMENT_KIND_TOKENS: Tuple[str, ...] = ("Mafi",)

# Monedas en las que se puede comisionar; cualquier otra va siempre a revisión.
COMMISSIONABLE_CURRENCIES: Tuple[str, ...] = ("USD", "EUR")


# ═══════════════════════════════════════════════════════════════════════
# CONVERSIÓN DE MONEDA
# ═══════════════════════════════════════════════════════════════════════

DEFAULT_ROE: float = 1.22  # EU 1,00 = X USD (TC pactado ilustrativo).

# Monedas distintas de USD que pueden aparecer en cargos del manifiesto.
FOREIGN_CURRENCIES: Tuple[str, ...] = ("EUR", "BRL", "GBP")


# ═══════════════════════════════════════════════════════════════════════
# PREFIJOS DE DESCUENTO Y PALABRAS CLAVE DE RESPALDO
# ═══════════════════════════════════════════════════════════════════════

COMMISSIONABLE_BASE: str = "basic frt"

COMMISSIONABLE_EXTRA: Tuple[str, ...] = (
    "open top", "over height", "over high", "over hght", "over weight", "over wight",
)

NEGATIVE_PREFIXES: Tuple[str, ...] = ("rebate", "adj", "fac", "baf decrease")


# ═══════════════════════════════════════════════════════════════════════
# CLASIFICACIÓN DE MARCA Y CONSIGNATARIOS (100% ficticios — portafolio AMA)
# ═══════════════════════════════════════════════════════════════════════

CONTAINER_KEYWORDS: Tuple[str, ...] = (
    "Tank Container", "Dry Cargo", "High Cube", "Open Top",
)

GENERAL_KEYWORDS: Tuple[str, ...] = (
    "PALLET(S)", "UNIT(S)", "PACKAGE(S)", "BOX(S)", "CASE(S)",
    "BUNDLE(S)", "CRATE(S)", "BAG(S)", "PIECE(S)", "UNPACKED(S)",
    "STEEL(S)", "MAFI(S)",
)

# Mapeo nombre de shipper/consignee -> marca. El orden importa: primer match gana.
# NOTA: marcas y razones sociales 100% ficticias (versión de portafolio).
BRAND_MAPPING: Dict[str, str] = {
    "AUTOMOTRIZ AVALON": "AVALON",
    "GRUPO BIONDA": "BIONDA",
    "CORVEX": "CORVEX",
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
# Lista ordenada (primer token encontrado gana); las multi-palabra y prefijos VIN
# van antes que los genéricos. Marcas y prefijos VIN ficticios.
MAKE_KEYWORDS: Tuple[Tuple[str, str], ...] = (
    ("EVORA-PREMIUM", "EVORA"), ("EVORA", "EVORA"),
    ("IONIC", "IONIC"),
    ("KRONOS", "KRONOS"), ("JUNO", "JUNO"),
    ("DELTAR", "DELTAR"), ("ZZDLTR", "DELTAR"),
    ("FENIX", "FENIX"),
    ("MINIC", "CORVEX"), ("CORVEX", "CORVEX"),
    ("BIONDA", "BIONDA"),
    ("AVALON", "AVALON"), ("ZZAVLN", "AVALON"),
    ("LYRA", "LYRA"),
    ("MISTRAL", "MISTRAL"),
    ("NOVAX", "NOVAX"), ("ORBIS", "ORBIS"),
    ("GRIFON", "GRIFON"), ("PEGASO", "PEGASO"), ("QUASAR", "QUASAR"),
)

# Sufijo de agencia/forwarder según el consignee (se pospone " - <SUFIJO>").
AGENCY_SUFFIXES: Tuple[Tuple[str, str], ...] = (
    ("OCEANLINK", "OCEANLINK"),
    ("DELCAR", "DELCAR"),
    ("PREMIUMCARS", "PREMIUMCARS"),
    ("EXIMPORT", "EXIMPORT"),
)


# ═══════════════════════════════════════════════════════════════════════
# PUERTOS Y BUQUES (flota ficticia Atlantic Maritime Agency)
# ═══════════════════════════════════════════════════════════════════════

PORT_ABBREVIATIONS: Dict[str, str] = {
    "ZARATE": "ZTE", "BUENOS AIRES": "BUE", "ROSARIO": "ROS",
    "MONTEVIDEO": "MVD", "SANTOS": "STS", "RIO": "RIO",
    "RIO DE JANEIRO": "RIO", "PARANAGUA": "PNG", "ITAJAI": "ITJ",
}

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

ENS_FIXED_AMOUNT: float = 10.0  # ENS EXPO: monto fijo por B/L (USD, ilustrativo).
ENS_ISSUER_ADDRESS: str = "Av. Puerto Madero 100 - Buenos Aires (Argentina)"
ENS_PRINCIPAL_NAME: str = "Atlantic Maritime Lines S.A."

ROLLING_RORO_KEYWORDS: Tuple[str, ...] = ("MOTORHOME", "MOTOR HOME", "CAMPER")


@dataclass
class EngineConfig:
    """Agrupa la configuración de una corrida; permite override puntual."""
    tariffs: LineCostTariffs = field(default_factory=LineCostTariffs)
    roe: float = DEFAULT_ROE
    commission_rates: Dict[str, CommissionRates] = field(
        default_factory=lambda: dict(COMMISSION_RATES)
    )
    # B/L de rodados cuyo booking cerró la agencia AMA (los informa Comercial).
    # None = no se recibió la lista: se calcula como estimación (todos por unidad)
    # y se avisa. Conjunto vacío = Comercial confirmó que no hay ninguno.
    agency_bls: Optional[FrozenSet[str]] = None

    @property
    def tonnage_factor(self) -> float:
        """Divisor del PESO en la planilla (=TOLL/factor): la tarifa fija suelta."""
        return self.tariffs.sc_loose_per_ton

    def rates_for(self, op_type: str) -> CommissionRates:
        return self.commission_rates[op_type.upper()]
