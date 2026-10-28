"""
Ejecutable del motor (v2.0). Uso por bandeja:

    python -m manifest_engine

  1. Toma los CSV de la carpeta 'Manifiestos a procesar' (ATL0426_IMPO_1.CSV, ...)
     y los agrupa por buque-viaje.
  2. Por cada viaje pide el ROE (EU 1,00 = X USD) y los B/L de rodados cuyo
     booking cerró la agencia AMA (lista de Comercial).
  3. Genera los reportes en 'Salida Motor/<COD> <VIAJE>/' y guarda el manifiesto
     procesado en 'Salida Motor/<COD> <VIAJE>/Manifiestos/' (sale de la bandeja).

Otros usos:

    python -m manifest_engine <carpeta_entrada> [carpeta_salida]
        procesa otra carpeta (por ejemplo, samples/ o la carpeta Manifiestos de
        un viaje para reprocesarlo). Los CSV de esa carpeta se copian, no se mueven.

    python -m manifest_engine --modelos [carpeta]
        genera las planillas de comisiones vacías (modelos).

Las reglas de conceptos de cargo, tipos de bulto y puertos se leen de
'Definiciones_reglas_comisiones.xlsx' (raíz del proyecto).
"""

import glob
import os
import re
import shutil
import sys
import time
from datetime import datetime

from . import config
from .catalog import CatalogError, load_catalog, register_new_concepts
from .parsing import extract_ship_name, parse_manifest
from .reports import (
    generar_biblia, generar_chassis, generar_comisiones, generar_ens,
    generar_facturacion, generar_modelo_comisiones, generar_por_bl,
    generar_por_marca,
)
from .rules.line_costs import detect_loose_tariff

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INBOX = os.path.join(PROJECT_DIR, "Manifiestos a procesar")
OUTPUT = os.path.join(PROJECT_DIR, "Salida Motor")
MODELOS = os.path.join(PROJECT_DIR, "Modelos")

COMM_NAME = {"IMPO": "Comm. Inward", "EXPO": "Comm. Outward"}

ROE_MIN, ROE_MAX = 0.5, 3.0     # control de tipeo: un EUR vale entre 0,5 y 3 USD
_ASK = object()


def extract_ship_info(filename):
    """ATL0426_IMPO_1.CSV -> ('ATL', '0426', 'IMPO')."""
    base = os.path.basename(filename).upper()
    m = re.match(r"([A-Z]{3})(\d{4})_(IMPO|EXPO)_", base)
    return (m.group(1), m.group(2), m.group(3)) if m else (None, None, None)


def _csv_files(folder):
    """CSV de la carpeta, sin repetir (en Windows '*.CSV' y '*.csv' son lo mismo)."""
    vistos = {}
    for f in glob.glob(os.path.join(folder, "*.csv")) + glob.glob(os.path.join(folder, "*.CSV")):
        vistos.setdefault(os.path.normcase(os.path.abspath(f)), f)
    return sorted(vistos.values())


def process_directory(input_dir, output_dir=None, roe=_ASK, agency_override=_ASK,
                      catalog=None, mover=None, register_concepts=True):
    """
    input_dir:       carpeta con los CSV.
    output_dir:      carpeta donde se crea una subcarpeta por viaje (default: 'Salida Motor').
    roe:             float (ej. 1.22); None usa config.DEFAULT_ROE; si no se pasa, pregunta por consola.
    agency_override: frozenset de B/L AMA, frozenset() si Comercial confirmó ninguno,
                     None para modo ESTIMACIÓN sin preguntar; si no se pasa, pregunta por consola.
    mover:           True = mueve los CSV procesados; False = los copia.
                     Default: mueve sólo si la entrada es la bandeja 'Manifiestos a procesar'.
    register_concepts: True = registra conceptos nuevos en el catálogo Excel.
    Devuelve la lista de carpetas de viaje generadas.
    """
    catalog = catalog or load_catalog()
    output_dir = output_dir or OUTPUT
    if mover is None:
        mover = os.path.normcase(os.path.abspath(input_dir)) == os.path.normcase(os.path.abspath(INBOX))
    os.makedirs(output_dir, exist_ok=True)

    ships, ignorados = {}, []
    for f in _csv_files(input_dir):
        code, voyage, op = extract_ship_info(f)
        if not code:
            ignorados.append(os.path.basename(f))
            continue
        ships.setdefault((code, voyage), {}).setdefault(op, []).append(f)

    print("=" * 66)
    print("  PROCESADOR DE MANIFIESTOS MARÍTIMOS — AMA (v2.0)")
    print("=" * 66)
    print(f"  Entrada: {input_dir}")
    print(f"  Salida:  {output_dir}")
    print(f"  Reglas:  {catalog.path.name}")
    print(f"  Viajes:  {len(ships)}")
    print("=" * 66)
    for nombre in ignorados:
        print(f"  ATENCIÓN: '{nombre}' no tiene el nombre de un manifiesto "
              f"(ej. ATL0426_IMPO_1.CSV): no se procesa.")
    if not ships:
        print(f"\nNo hay manifiestos para procesar en: {input_dir}")
        return []

    t0 = time.time()
    carpetas = []
    for (code, voyage), por_op in sorted(ships.items()):
        viaje = f"{code} {voyage}"
        print(f"\n[{viaje}]")
        dobles = {op: fs for op, fs in por_op.items() if len(fs) > 1}
        if dobles:
            for op, fs in dobles.items():
                print(f"   ERROR: hay {len(fs)} manifiestos {op} de este viaje "
                      f"({', '.join(os.path.basename(f) for f in fs)}). Dejá uno solo y volvé a correr.")
            print("   El viaje NO se procesó.")
            continue
        files = {op: fs[0] for op, fs in por_op.items()}
        try:
            carpeta = _procesar_viaje(
                code, voyage, files, output_dir, roe, agency_override, catalog,
                register_concepts=register_concepts,
            )
        except PermissionError as e:
            print(f"   ERROR: no se pudo guardar '{os.path.basename(e.filename or '')}'. "
                  f"Si está abierto en Excel, cerralo y volvé a correr.")
            print("   El viaje NO se terminó de procesar; el manifiesto queda donde estaba.")
            continue
        _guardar_manifiestos(files, carpeta, mover)
        carpetas.append(carpeta)

    print("\n" + "=" * 66)
    print(f"  Listo en {time.time() - t0:.1f}s. Reportes en:")
    for c in carpetas:
        print(f"    {c}")
    print("  Revisá la hoja 'REVISIÓN' de cada planilla y las filas en rojo.")
    print("=" * 66)
    return carpetas


def _procesar_viaje(code, voyage, files, output_root, roe, agency_override, catalog,
                    register_concepts=True):
    out = os.path.join(output_root, f"{code} {voyage}")
    os.makedirs(out, exist_ok=True)

    # El nombre sale del encabezado del manifiesto; SHIP_NAMES es respaldo.
    ship = (extract_ship_name(next(iter(files.values())))
            or config.SHIP_NAMES.get(code) or code)
    print(f"   Buque: {ship}   ({', '.join(sorted(files))})")

    impo_bls = parse_manifest(files["IMPO"], "IMPO", catalog) if "IMPO" in files else None
    expo_bls = parse_manifest(files["EXPO"], "EXPO", catalog) if "EXPO" in files else None

    # Tarifa del TOLL suelto (ej. 18.0 histórico / 20.0 actual): se lee del manifiesto.
    toll_loose = detect_loose_tariff((impo_bls or []) + (expo_bls or []),
                                     default=config.LineCostTariffs().sc_loose_per_ton)

    if roe is _ASK:
        roe_viaje = _ask_roe(f"{code} {voyage}")
    elif roe is None:
        roe_viaje = config.DEFAULT_ROE
    else:
        roe_viaje = float(roe)

    if agency_override is _ASK:
        agency = _ask_agency(f"{code} {voyage}")
    else:
        agency = agency_override

    alertas = _agency_alerts(agency, impo_bls or [], expo_bls or [])
    print(f"   ROE: EU 1,00 = {roe_viaje:g} USD   |   TOLL suelto: {toll_loose:g}/Tn")
    if agency is None:
        print("   ATENCIÓN: sin lista de Comercial -> rodados estimados por unidad.")
    else:
        print(f"   B/L de la agencia (AMA): {', '.join(sorted(agency)) or 'ninguno'}")
    for a in alertas:
        print(f"   ATENCIÓN {a.bl}: {a.detalle}")

    cfg = config.EngineConfig(
        tariffs=config.LineCostTariffs(sc_loose_per_ton=toll_loose),
        roe=roe_viaje,
        agency_bls=agency,
    )

    def name(n):
        return os.path.join(out, f"{code} {voyage} - {n}.xlsx")

    nuevos = {}
    if impo_bls is not None:
        generar_biblia(impo_bls, "IMPO", ship, voyage, name("IMPO"),
                       cfg=cfg, catalog=catalog, extra_alerts=_for(alertas, impo_bls))
        res = generar_comisiones(impo_bls, "IMPO", ship, voyage, name(COMM_NAME["IMPO"]),
                                 cfg=cfg, catalog=catalog, extra_alerts=_for(alertas, impo_bls))
        _collect_new(nuevos, res, "IMPO")
        print(f"   IMPO: {len(impo_bls):3d} B/L -> IMPO + Comm. Inward   {_resumen(res)}")

    if expo_bls is not None:
        generar_biblia(expo_bls, "EXPO", ship, voyage, name("EXPO"),
                       cfg=cfg, catalog=catalog, extra_alerts=_for(alertas, expo_bls))
        res = generar_comisiones(expo_bls, "EXPO", ship, voyage, name(COMM_NAME["EXPO"]),
                                 cfg=cfg, catalog=catalog, extra_alerts=_for(alertas, expo_bls))
        _collect_new(nuevos, res, "EXPO")
        generar_ens(expo_bls, ship, voyage, name("ENS"), cfg=cfg)
        print(f"   EXPO: {len(expo_bls):3d} B/L -> EXPO + Comm. Outward + ENS   {_resumen(res)}")

    generar_facturacion(impo_bls, expo_bls, ship, voyage, name("Facturación"))
    generar_por_bl(impo_bls, expo_bls, ship, voyage, name("PorBL"), cfg=cfg)
    generar_chassis(impo_bls, expo_bls, ship, voyage, name("Chasis"), cfg=cfg)
    generar_por_marca(impo_bls, expo_bls, ship, voyage, name("PorMarca"), cfg=cfg)
    print("   Facturación (para Caja) + PorBL + Chasis + PorMarca")

    if nuevos and register_concepts:
        agregados, error = register_new_concepts(nuevos, f"{code} {voyage}", catalog.path)
        if agregados:
            print(f"   Conceptos nuevos registrados en {catalog.path.name}: {', '.join(agregados)}")
        if error:
            print(f"   ATENCIÓN: {error}")
    return out


def _guardar_manifiestos(files, carpeta_viaje, mover):
    """Guarda los CSV procesados en '<viaje>/Manifiestos/' (si ya había uno, lo renombra)."""
    destino = os.path.join(carpeta_viaje, "Manifiestos")
    os.makedirs(destino, exist_ok=True)
    guardados = 0
    for f in files.values():
        dst = os.path.join(destino, os.path.basename(f))
        if os.path.normcase(os.path.abspath(f)) == os.path.normcase(os.path.abspath(dst)):
            continue
        if os.path.exists(dst):
            base, ext = os.path.splitext(dst)
            os.replace(dst, f"{base} (anterior {datetime.now():%Y-%m-%d %H%M}){ext}")
        if mover:
            shutil.move(f, dst)
        else:
            shutil.copy2(f, dst)
        guardados += 1
    if guardados:
        print(f"   Manifiesto {'movido' if mover else 'copiado'} a {destino}")


BL_TOKEN = re.compile(r"S3\d{8}")


def parse_agency(raw):
    """
    Interpreta lo que se ingresa como lista de Comercial.
      ''        -> None (no hay lista: se estima)
      'NINGUNO' -> conjunto vacío (Comercial confirmó que no hay)
      'S30000101, S30000102' -> esos B/L
    Devuelve (lista, tokens_invalidos).
    """
    raw = (raw or "").strip().upper()
    if not raw:
        return None, []
    if raw == "NINGUNO":
        return frozenset(), []
    tokens = [t for t in re.split(r"[\s,;]+", raw.replace("[T]", "")) if t]
    validos = [t for t in tokens if BL_TOKEN.fullmatch(t)]
    invalidos = [t for t in tokens if not BL_TOKEN.fullmatch(t)]
    return frozenset(validos), invalidos


def parse_roe(raw):
    """'1,22' o '1.22' -> 1.22; None si no es un ROE razonable."""
    try:
        v = float((raw or "").strip().replace(",", "."))
    except ValueError:
        return None
    return v if ROE_MIN <= v <= ROE_MAX else None


def _ask_roe(viaje):
    print(f"   ROE del viaje {viaje} (cuántos USD vale 1 EUR, ej. 1,22; Enter = {config.DEFAULT_ROE:g}):")
    while True:
        try:
            raw = input("     > ")
        except EOFError:
            return config.DEFAULT_ROE
        if not raw.strip():
            return config.DEFAULT_ROE
        v = parse_roe(raw)
        if v is not None:
            return v
        print(f"     '{raw}' no es un ROE válido (tiene que estar entre {ROE_MIN:g} y {ROE_MAX:g}).")


def _ask_agency(viaje):
    print(f"   B/L de rodados cerrados por la agencia AMA (lista de Comercial) para {viaje}:")
    print("     separados por coma o espacio | NINGUNO = Comercial confirmó que no hay")
    print("     Enter vacío = no tengo la lista (se calcula como ESTIMACIÓN, todo por unidad)")
    while True:
        try:
            raw = input("     > ")
        except EOFError:
            return None
        lista, invalidos = parse_agency(raw)
        if not invalidos:
            return lista
        print(f"     No reconozco como B/L: {', '.join(invalidos)}. Volvé a ingresar la lista.")


def _agency_alerts(agency, impo_bls, expo_bls):
    """Controles de la lista de Comercial contra los manifiestos del viaje."""
    from .validation import Alerta
    if not agency:
        return []
    todos = {b.bl_clean: b for b in list(impo_bls) + list(expo_bls)}
    out = []
    for bl in sorted(agency):
        b = todos.get(bl)
        if b is None:
            out.append(Alerta(bl, "LISTA COMERCIAL",
                              "Informado por Comercial pero no está en los manifiestos de este viaje"))
        elif not any(b.units(f) for f in ("CAR", "VAN", "RORO")):
            out.append(Alerta(bl, "LISTA COMERCIAL",
                              "Informado como booking de la agencia pero el B/L no tiene rodados: "
                              "no se aplica el 1%"))
    return out


def _for(alertas, bls):
    """Alertas que corresponden a un manifiesto (las de B/L inexistentes van a todos)."""
    propios = {b.bl_clean for b in bls}
    return [a for a in alertas if a.bl in propios or "no está en los manifiestos" in a.detalle]


def _collect_new(nuevos, results, op):
    for r in results:
        for nombre, ejemplo in r.new_concepts:
            d = nuevos.setdefault(nombre, {"IMPO": 0, "EXPO": 0, "ejemplo": ejemplo})
            d[op] += 1


def _resumen(results):
    pend = sum(1 for r in results if not r.ok)
    total = sum(r.total_usd or 0 for r in results)
    txt = f"comisión calculada USD {total:,.2f}"
    return txt + (f" | {pend} B/L A REVISAR" if pend else " | sin pendientes")


def main():
    args = sys.argv[1:]
    if args and args[0] == "--modelos":
        destino = args[1] if len(args) > 1 else MODELOS
        os.makedirs(destino, exist_ok=True)
        for op in ("IMPO", "EXPO"):
            out = os.path.join(destino, f"MODELO - {COMM_NAME[op]}.xlsx")
            generar_modelo_comisiones(op, out)
            print(f"Modelo generado: {out}")
        return

    input_dir = args[0] if args else INBOX
    output_dir = args[1] if len(args) > 1 else OUTPUT
    if input_dir == INBOX:
        os.makedirs(INBOX, exist_ok=True)
    if not os.path.isdir(input_dir):
        print(f"ERROR: no existe la carpeta {input_dir}")
        sys.exit(1)

    try:
        catalog = load_catalog()
    except CatalogError as e:
        print(f"ERROR: {e}")
        print("El motor necesita la planilla de definiciones de reglas para calcular comisiones.")
        sys.exit(1)

    process_directory(input_dir, output_dir, catalog=catalog)


if __name__ == "__main__":
    main()
