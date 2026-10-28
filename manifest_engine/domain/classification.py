"""
classification.py - Clasificación de carga y resolución de marca.

Tres preguntas que el resto del motor le hace a un B/L:
  - ¿qué familia de carga es?        -> cargo_category()
  - ¿de qué marca son los autos?     -> resolve_brand()
  - ¿cómo se escribe en la planilla? -> format_cargo()
"""

import re
from typing import Optional

from .. import config
from .models import BLRecord, CargoData


# Orden de evaluación: la marca REAL sale de la descripción de la mercadería;
# el mapeo por shipper/consignee es el respaldo.
_MAKE_PATTERNS = tuple(
    (re.compile(r"\b" + re.escape(token) + r"\b"), name)
    for token, name in config.MAKE_KEYWORDS
)

_SINGULARIZE = (
    ("PACKAGES", "PACKAGE"), ("PALLETS", "PALLET"), ("CASES", "CASE"),
    ("PIECES", "PIECE"), ("UNITS", "UNIT"), ("BUNDLES", "BUNDLE"),
    ("CRATES", "CRATE"), ("BAGS", "BAG"), ("UNPACKEDS", "UNPACKED"),
    ("BOXS", "BOX"), ("STEELS", "STEEL"), ("MAFIS", "MAFI"),
)


def abreviar_puerto(puerto: str) -> str:
    """ZARATE -> ZTE, BUENOS AIRES -> BUE; desconocido -> primeras 3 letras."""
    if not puerto or puerto == "Nulo":
        return ""
    p = puerto.upper().strip()
    return config.PORT_ABBREVIATIONS.get(p, p[:3])


def detect_brand(entity: str) -> Optional[str]:
    """Marca según el shipper/consignee (BRAND_MAPPING)."""
    e = str(entity or "").upper()
    for key, val in config.BRAND_MAPPING.items():
        if key in e:
            return val
    return None


def detect_make(description: str) -> Optional[str]:
    """Marca según la DESCRIPCIÓN de la mercadería (MAKE_KEYWORDS)."""
    d = str(description or "").upper()
    for rx, name in _MAKE_PATTERNS:
        if rx.search(d):
            return name
    return None


def agency_suffix(entity: str) -> Optional[str]:
    e = str(entity or "").upper()
    for key, suf in config.AGENCY_SUFFIXES:
        if key in e:
            return suf
    return None


def resolve_brand(bl: BLRecord) -> Optional[str]:
    """
    Marca final del B/L: la de la descripción manda; si no hay, la del
    shipper/consignee. Si la marca vino de la descripción y el consignee es un
    forwarder conocido, se le agrega el sufijo de agencia (ej. DELTAR - DELCAR).
    """
    make = detect_make(bl.description)
    if not make:
        return detect_brand(bl.entity)
    suf = agency_suffix(bl.entity)
    if suf and suf not in make:
        return f"{make} - {suf}"
    return make


def cargo_category(cargo: CargoData) -> Optional[str]:
    """'containers' | 'cars' | 'projects' | None."""
    if cargo.containers:
        return "containers"
    has_cars = cargo.n_cars > 0
    has_roro = cargo.n_roro > 0
    if has_cars and not has_roro and not cargo.general:
        return "cars"
    if has_roro or cargo.general:
        return "projects"
    if has_cars:
        return "cars"
    return None


def is_rolling(bl: BLRecord) -> bool:
    """Rolling = sólo autos/vans. Los motorhome/camper ruedan aunque sean RORO."""
    if cargo_category(bl.cargo) == "cars":
        return True
    if bl.cargo.n_roro and not bl.cargo.containers and not bl.cargo.general:
        desc = bl.description
        return any(k in desc for k in config.ROLLING_RORO_KEYWORDS)
    return False


def _singular(word: str) -> str:
    w = word.upper()
    for plural, singular in _SINGULARIZE:
        w = w.replace(plural, singular)
    return w


def format_cargo(bl: BLRecord, cargo: Optional[CargoData] = None) -> str:
    """
    Texto de la columna CARGA: '3 USED RORO + 30 PACKAGE', '18 NEW CAR - AVALON'.

    `cargo` permite renderizar sólo una parte de la carga del B/L (la planilla
    abre un renglón por carga cuando conviven contenedor, rodado y suelta).
    """
    brand = resolve_brand(bl)
    cargo = bl.cargo if cargo is None else cargo
    parts = []

    if cargo.containers:
        parts.append(str(cargo.n_containers))

    # El tipo (CAR / VAN / RORO) siempre se muestra: cada uno comisiona distinto.
    for v_type, count in cargo.vehicles.items():
        v = v_type.upper()
        base = f"{v} - {brand}" if brand else v
        parts.append(f"{count} {base}")

    if not cargo.containers:
        for g_type, count in cargo.general.items():
            desc = f"{count} {_singular(g_type)}"
            if brand:
                desc += f" - {brand}"
            parts.append(desc)

    return " + ".join(parts)
