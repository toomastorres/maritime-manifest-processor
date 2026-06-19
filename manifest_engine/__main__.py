"""
CLI / batch del motor. Uso:

    python -m manifest_engine [carpeta_entrada] [carpeta_salida]

Busca CSV de manifiestos (GBA0625_IMPO_1.CSV, ...), los agrupa por buque-viaje y
genera, por cada uno, los 5 reportes: planilla IMPO/EXPO, comisiones IMPO/EXPO y
ENS. Pide el ROE (EUR->USD) y el factor TOLL suelto al iniciar.
"""

import glob
import os
import re
import sys
import time

from . import config
from .parsing import parse_manifest
from .rules.line_costs import detect_loose_tariff
from .reports import (
    generar_planilla_impo, generar_planilla_expo, generar_comisiones, generar_ens,
    generar_por_bl, generar_chassis, generar_por_marca,
)


def extract_ship_info(filename):
    """GBA0126_IMPO_1.CSV -> ('GBA', '0126', 'IMPO')."""
    base = os.path.basename(filename).upper()
    m = re.match(r"([A-Z]{3})(\d{4})_(IMPO|EXPO)_", base)
    return (m.group(1), m.group(2), m.group(3)) if m else (None, None, None)


def process_directory(input_dir, output_dir=None, cfg=None):
    cfg = cfg or config.EngineConfig()
    if output_dir is None:
        output_dir = os.path.join(input_dir, "output")
    os.makedirs(output_dir, exist_ok=True)

    csv_files = glob.glob(os.path.join(input_dir, "*.CSV")) + \
        glob.glob(os.path.join(input_dir, "*.csv"))

    ships = {}
    for f in csv_files:
        code, voyage, op = extract_ship_info(f)
        if code:
            ships.setdefault((code, voyage), {})[op] = f

    if not ships:
        print(f"No se encontraron manifiestos CSV en: {input_dir}")
        return

    print("=" * 64)
    print("  MOTOR DE PROCESAMIENTO DE MANIFIESTOS")
    print("=" * 64)
    print(f"  ROE: EU 1,00 = {cfg.roe} USD   |   TOLL suelto: autodetectado por viaje")
    print(f"  Entrada: {input_dir}")
    print(f"  Salida:  {output_dir}")
    print(f"  Buques:  {len(ships)}")
    print("=" * 64)

    t0 = time.time()
    generados = 0
    for (code, voyage), files in sorted(ships.items()):
        ship = config.SHIP_NAMES.get(code, code)
        prefix = f"{code}_{voyage}"
        print(f"\n[{prefix}] {ship}")

        impo_bls = parse_manifest(files["IMPO"], "IMPO") if "IMPO" in files else None
        expo_bls = parse_manifest(files["EXPO"], "EXPO") if "EXPO" in files else None

        # Era de la tarifa TOLL suelta: se autodetecta del manifiesto (18 vs
        # 20). No depende del código de viaje. El ROE viene del cfg global.
        detect_src = (impo_bls or []) + (expo_bls or [])
        toll_loose = detect_loose_tariff(detect_src, default=cfg.tariffs.sc_loose_per_ton)
        vcfg = config.EngineConfig(
            tariffs=config.LineCostTariffs(sc_loose_per_ton=toll_loose),
            roe=cfg.roe,
        )
        print(f"   TOLL suelto detectado: {toll_loose}/Tn")

        if impo_bls is not None:
            generar_planilla_impo(impo_bls, ship, voyage, _p(output_dir, prefix, "Impo"), cfg=vcfg)
            generar_comisiones(impo_bls, "IMPO", ship, voyage, _p(output_dir, prefix, "MontocomZte"), cfg=vcfg)
            generados += 2
            print(f"   IMPO: {len(impo_bls):3d} BLs -> Impo + MontocomZte")

        if expo_bls is not None:
            generar_planilla_expo(expo_bls, ship, voyage, _p(output_dir, prefix, "Expo"), cfg=vcfg)
            generar_comisiones(expo_bls, "EXPO", ship, voyage, _p(output_dir, prefix, "MontoExpoZte"), cfg=vcfg)
            generar_ens(expo_bls, ship, voyage, _p(output_dir, prefix, "ENS"), cfg=vcfg)
            generados += 3
            print(f"   EXPO: {len(expo_bls):3d} BLs -> Expo + MontoExpoZte + ENS")

        if impo_bls or expo_bls:
            generar_por_bl(impo_bls, expo_bls, ship, voyage, _p(output_dir, prefix, "PorBL"), cfg=vcfg)
            generar_chassis(impo_bls, expo_bls, ship, voyage, _p(output_dir, prefix, "Chasis"), cfg=vcfg)
            generar_por_marca(impo_bls, expo_bls, ship, voyage, _p(output_dir, prefix, "PorMarca"), cfg=vcfg)
            generados += 3
            print(f"   PorBL + Chasis + PorMarca")

    print("\n" + "=" * 64)
    print(f"  Listo: {generados} reportes en {time.time()-t0:.1f}s")
    print(f"  Revisá la hoja 'REVISIÓN' de cada planilla para los controles.")
    print("=" * 64)


def _p(output_dir, prefix, kind):
    return os.path.join(output_dir, f"{prefix}-{kind}.xlsx")


def _ask_float(prompt, default):
    raw = input(f"{prompt} [{default}]: ").strip().replace(",", ".")
    try:
        return float(raw) if raw else default
    except ValueError:
        print(f"  valor inválido, uso {default}")
        return default


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else "."
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None

    print("Configuración de la corrida (Enter = valor por defecto):")
    roe = _ask_float("  ROE (EU 1,00 = X USD)", config.DEFAULT_ROE)
    # El TOLL suelto (18 vs 20) se autodetecta por viaje desde el manifiesto.

    cfg = config.EngineConfig(roe=roe)
    process_directory(input_dir, output_dir, cfg)


if __name__ == "__main__":
    main()
