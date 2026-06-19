"""
Pruebas del motor de manifiestos sobre los manifiestos de ejemplo sintéticos.
Cubren parsing IMPO/EXPO, clasificación de carga, extracción de VIN/marca,
sufijo de agencia y generación de los reportes Excel end-to-end.
"""

from pathlib import Path

import pytest

from manifest_engine.parsing import parse_manifest
from manifest_engine.domain.cargo import extract_make, agency_suffix
from manifest_engine.__main__ import process_directory

SAMPLES = Path(__file__).resolve().parents[1] / "samples"
IMPO = str(SAMPLES / "ATL0426_IMPO_1.CSV")
EXPO = str(SAMPLES / "ATL0426_EXPO_1.CSV")


@pytest.fixture(scope="module")
def impo():
    return parse_manifest(IMPO, "IMPO")


@pytest.fixture(scope="module")
def expo():
    return parse_manifest(EXPO, "EXPO")


def test_impo_cantidad_de_bls(impo):
    assert len(impo) == 3
    assert [b.bl_no for b in impo] == ["S329600001", "S329600002", "S329600003"]


def test_impo_puerto_de_descarga(impo):
    assert all(b.port_of_discharge == "ZARATE" for b in impo)


def test_clasificacion_contenedor(impo):
    bl = impo[0]
    assert bl.cargo.container_count == 3
    assert bl.cargo.vehicle_count == 0
    # THC de contenedor de 40 ft a tarifa esperada
    assert bl.thc_40.monto == pytest.approx(120.0)
    assert {"Basic FRT", "THC", "Toll", "Sweeping", "BAF"} <= set(bl.charges)


def test_clasificacion_vehiculos(impo):
    bl = impo[1]
    assert bl.cargo.vehicle_count == 5
    assert bl.cargo.container_count == 0


def test_extraccion_de_vin_y_marca(impo):
    bl = impo[1]
    assert bl.vins == ["ZZDLTR8K9AA123456"]
    # El prefijo de VIN ficticio ZZDLTR resuelve a la marca DELTAR
    assert extract_make(bl.description_lines + bl.vins) == "DELTAR"


def test_sufijo_de_agencia(impo):
    # El consignee "AGENCIA MARITIMA OCEANLINK" aporta el sufijo OCEANLINK
    assert agency_suffix(impo[1].entity) == "OCEANLINK"


def test_expo_basico(expo):
    assert len(expo) == 2
    assert all(b.port_of_loading == "ZARATE" for b in expo)
    assert expo[0].cargo.container_count == 2
    assert expo[1].cargo.vehicle_count == 6


def test_generacion_de_reportes(tmp_path):
    process_directory(str(SAMPLES), str(tmp_path))
    reportes = sorted(p.name for p in tmp_path.glob("*.xlsx"))
    # 8 reportes esperados para un viaje con IMPO y EXPO
    assert len(reportes) == 8
    for kind in ("Impo", "Expo", "ENS", "PorBL", "Chasis", "PorMarca",
                 "MontocomZte", "MontoExpoZte"):
        assert any(kind in r for r in reportes), f"falta reporte {kind}"
