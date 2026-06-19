"""
cargo.py - Clasificación de carga y detección de marca.

Migra desde manifest_processor_v3.py:
- patrones NEW/USED CAR/RORO
- classify_description (bloque de clasificación del parser)
- cargo_category, detect_brand, format_cargo_output
"""

import re
from typing import Optional

from .. import config
from .models import CargoData


# Patrones de vehículos (migrados del script original)
NEW_CAR_PATTERN = re.compile(
    r"New Car(?:\(s\))?|New Small Van(?:\(s\))?|New Big Van(?:\(s\))?", re.IGNORECASE
)
USED_CAR_PATTERN = re.compile(
    r"Used Car(?:\(s\))?|Used Small Van(?:\(s\))?|Used Big Van(?:\(s\))?", re.IGNORECASE
)
NEW_RORO_PATTERN = re.compile(r"New LM RoRo", re.IGNORECASE)
USED_RORO_PATTERN = re.compile(r"Used LM RoRo", re.IGNORECASE)


def classify_description(cargo: CargoData, qty: int, desc: str) -> Optional[str]:
    """
    Clasifica una línea de descripción de carga y la acumula en `cargo`.
    Devuelve el tipo agrupado ('containers'/'vehicles'/'general') o None.
    """
    low = desc.lower()

    for kw in config.CONTAINER_KEYWORDS:
        if kw.lower() in low:
            cargo.add_container(kw.lower(), qty)
            return "containers"

    if NEW_CAR_PATTERN.search(desc):
        cargo.add_vehicle("new car", qty)
        return "vehicles"
    if USED_CAR_PATTERN.search(desc):
        cargo.add_vehicle("used car", qty)
        return "vehicles"
    if NEW_RORO_PATTERN.search(desc):
        cargo.add_vehicle("new roro", qty)
        return "vehicles"
    if USED_RORO_PATTERN.search(desc):
        cargo.add_vehicle("used roro", qty)
        return "vehicles"

    for kw in config.GENERAL_KEYWORDS:
        base = kw.replace("(S)", "").replace("(s)", "")
        if base.lower() in low:
            clean = kw.lower().replace("(s)", "s")
            cargo.add_general(clean, qty)
            return "general"

    return None


def cargo_category(cargo: CargoData) -> Optional[str]:
    """Retorna 'containers', 'cars', 'projects' o None."""
    has_containers = bool(cargo.containers)
    has_general = bool(cargo.general)
    has_cars = any(("car" in v.lower() or "van" in v.lower()) for v in cargo.vehicles)
    has_roro = any("roro" in v.lower() for v in cargo.vehicles)

    if has_containers:
        return "containers"
    if has_cars and not has_roro and not has_general:
        return "cars"
    if has_roro or has_general:
        return "projects"
    if has_cars:
        return "cars"
    return None


def detect_brand(shipper: str = "", consignee: str = "") -> Optional[str]:
    entity = (str(shipper) + " " + str(consignee)).upper()
    for key, val in config.BRAND_MAPPING.items():
        if key in entity:
            return val
    return None


def extract_make(description_lines) -> Optional[str]:
    """Marca/tipo real del vehículo leído de la descripción de la mercadería."""
    if not description_lines:
        return None
    text = " ".join(description_lines).upper()
    for kw, canon in config.MAKE_KEYWORDS:
        if kw in text:
            return canon
    return None


def agency_suffix(entity: str) -> Optional[str]:
    """Sufijo de agencia/forwarder según el consignee (OCEANLINK/DELCAR/...)."""
    e = (entity or "").upper()
    for kw, suf in config.AGENCY_SUFFIXES:
        if kw in e:
            return suf
    return None


_PLURALS = [
    ("PACKAGES", "PACKAGE"), ("PALLETS", "PALLET"), ("CASES", "CASE"),
    ("PIECES", "PIECE"), ("UNITS", "UNIT"), ("BUNDLES", "BUNDLE"),
    ("CRATES", "CRATE"), ("BAGS", "BAG"), ("UNPACKEDS", "UNPACKED"),
]


def format_cargo_output(cargo: CargoData, entity: str = "", description_lines=None) -> str:
    """
    Texto de la columna CARGA. La marca real del vehículo sale de la
    DESCRIPCIÓN (extract_make); el sufijo de agencia, del consignee.
    """
    make = extract_make(description_lines)
    suffix = agency_suffix(entity)
    desc_text = " ".join(description_lines or []).upper()
    brand_fallback = detect_brand(entity, entity)
    parts = []

    if cargo.containers:
        parts.append(str(cargo.container_count))

    if cargo.vehicles:
        for v_type, count in cargo.vehicles.items():
            v_upper = v_type.upper()
            newused = "NEW" if "NEW" in v_upper else "USED"
            if "RORO" in v_upper:
                base = f"{newused} RORO"
                special = make if make == "GRIFON" else brand_fallback if brand_fallback in ("GRIFON", "HELIOS") else None
                if special:
                    base += f" - {special}"
                elif "MOTORHOME" in desc_text:
                    base += " - MOTORHOME"
                elif suffix:
                    base += f" - {suffix}"
            else:  # car / van
                mk = make or brand_fallback
                base = f"{newused} {mk}" if mk else v_upper
                if suffix and (not mk or suffix not in mk):
                    base += f" - {suffix}"
            parts.append(f"{count} {base}")

    # La carga general se muestra también cuando convive con vehículos
    # (ej. "1 USED RORO + 24 PACKAGE"); en BLs de contenedor no se desglosa.
    if not cargo.containers and cargo.general:
        for g_type, count in cargo.general.items():
            g = g_type.upper()
            for plural, singular in _PLURALS:
                g = g.replace(plural, singular)
            parts.append(f"{count} {g}")

    return " + ".join(parts)
