"""
audit.py - Auditoría de paridad a NIVEL BL contra la verdad exacta.

El diff celda-por-celda es engañoso (penaliza layout/espaciado). Este módulo
compara por número de B/L, campo por campo, EVALUANDO las fórmulas del motor
(no comparando strings de fórmula vs valor cacheado de la verdad), y CATEGORIZA
cada discrepancia en:

  - genuina (el motor debería poder resolverla)  -> a investigar/arreglar
  - conocida/explicada (corrección humana, no es error del motor) -> "corregido"

Objetivo: llegar a "100% o explicado". Toda celda que no coincide debe caer en
una categoría conocida; lo que no, es trabajo real de paridad.

Uso:
    python -m manifest_engine.audit [carpeta_datos]   # default: cwd
"""

import os
import re
import sys
import glob
import warnings
from collections import Counter, defaultdict

warnings.filterwarnings("ignore")
import openpyxl
from openpyxl.utils import column_index_from_string

from .parsing import parse_manifest
from . import config
from .rules.line_costs import detect_loose_tariff
from .reports import generar_planilla_impo, generar_planilla_expo


# ─────────────────────────── evaluador de fórmulas ───────────────────────────

_CELL = re.compile(r"\$?([A-Z]{1,2})\$?(\d+)")
_TEXT_NONNUM = ("", "Prepaid", "MATRIZ", "Matriz", "Matriz", "Collect", "***")


def to_num(v):
    """Convierte una celda (número, o texto con moneda US/español) a float o None."""
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip()
    if s in _TEXT_NONNUM:
        return None
    s = re.sub(r"[A-Za-z$]", "", s).strip()
    if not s or not re.search(r"\d", s):
        return None
    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):       # español: 2.790,00
            s = s.replace(".", "").replace(",", ".")
        else:                                  # US: 2,790.00
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if re.search(r",\d{1,2}$", s) else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _grid(ws):
    return {(c.column, c.row): c.value for row in ws.iter_rows() for c in row
            if c.value is not None}


def _eval(g, col, row, memo, depth=0):
    key = (col, row)
    if key in memo:
        return memo[key]
    if depth > 25:
        return None
    v = g.get(key)
    if v is None:
        memo[key] = None
        return None
    if not (isinstance(v, str) and v.startswith("=")):
        memo[key] = to_num(v)
        return memo[key]
    memo[key] = None  # corta ciclos
    expr = v[1:]

    def sub_sum(m):
        c1, c2 = column_index_from_string(m.group(1)), column_index_from_string(m.group(3))
        r1, r2 = int(m.group(2)), int(m.group(4))
        tot = 0.0
        for ci in range(c1, c2 + 1):
            for rr in range(r1, r2 + 1):
                x = _eval(g, ci, rr, memo, depth + 1)
                if x:
                    tot += x
        return str(tot)

    expr = re.sub(r"SUM\(\$?([A-Z]{1,2})\$?(\d+):\$?([A-Z]{1,2})\$?(\d+)\)", sub_sum, expr)

    def sub_max(m):
        a = _eval(g, column_index_from_string(m.group(1)), int(m.group(2)), memo, depth + 1) or 0
        b = _eval(g, column_index_from_string(m.group(3)), int(m.group(4)), memo, depth + 1) or 0
        return str(max(a, b))

    expr = re.sub(r"MAX\(\$?([A-Z]{1,2})\$?(\d+),\$?([A-Z]{1,2})\$?(\d+)\)", sub_max, expr)

    def sub_ref(m):
        x = _eval(g, column_index_from_string(m.group(1)), int(m.group(2)), memo, depth + 1)
        return str(x) if x is not None else "0"

    expr = _CELL.sub(sub_ref, expr)
    try:
        res = eval(expr, {"__builtins__": {}}, {})
    except Exception:
        res = None
    memo[key] = res
    return res


# ─────────────────────────── extracción por BL ───────────────────────────

_BLP = re.compile(r"^S3\d{8}(\[T\])?$")
IMPO_FIELDS = {"TOLL": 6, "THC20": 7, "THC40": 8, "SWP": 9, "PRE$": 10,
               "COL$": 11, "PRE€": 12, "COL€": 13, "BAF": 14, "COMM": 15}
# El motor diverge del layout de la verdad en EXPO (COLLECT en 4 columnas
# USD/EUR/GBP/BRL = 9..12, y todo lo de la derecha corrido +3). La verdad
# (hecha a mano) usa el layout viejo: COLLECT en 1 columna (9), ABROAD 10,
# COMIS 11, TOLL 19. Cada lado se extrae con SU mapa.
EXPO_FIELDS_GEN = {"THC20": 4, "THC40": 5, "SWP": 6, "ENS": 7, "PREPAID": 8,
                   "COLLECT": (9, 12), "ABROAD": 13, "COMIS": 14, "TOLL": 22}
EXPO_FIELDS_TRUTH = {"THC20": 4, "THC40": 5, "SWP": 6, "ENS": 7, "PREPAID": 8,
                     "COLLECT": 9, "ABROAD": 10, "COMIS": 11, "TOLL": 19}


def _eval_field(g, fc, row, memo):
    """fc puede ser una columna (int) o un rango (a, b) que se suma."""
    if isinstance(fc, tuple):
        a, b = fc
        tot = None
        for c in range(a, b + 1):
            x = _eval(g, c, row, memo)
            if x is not None:
                tot = (tot or 0) + x
        return tot
    return _eval(g, fc, row, memo)


def _extract(path, bl_col, fields):
    wb = openpyxl.load_workbook(path, data_only=False)
    out = {}
    for ws in wb.worksheets:
        if ws.title.strip() in ("TOLL", "REVISIÓN", "REVISI�N"):
            continue
        g = _grid(ws)
        memo = {}
        for (col, row), v in list(g.items()):
            if col == bl_col and isinstance(v, str) and _BLP.match(v.strip()):
                bl = v.strip().replace("[T]", "")
                out[bl] = {fn: _eval_field(g, fc, row, memo) for fn, fc in fields.items()}
    return out


# ─────────────────────────── localizar la verdad ───────────────────────────

def truth_dir(code, voy):
    for d in glob.glob(f"{code} V*{voy}"):
        if os.path.isdir(d):
            return d
    return None


def truth_file(d, kind):
    for f in os.listdir(d):
        if f.startswith("~$"):
            continue
        if f.lower().startswith(kind.lower() + " "):
            return os.path.join(d, f)
    return None


# ─────────────────────────── categorización ───────────────────────────

def _zero(x):
    return x is None or abs(x) < 0.5


def _match(gv, tv):
    if _zero(gv) and _zero(tv):
        return True
    if gv is None or tv is None:
        return False
    return abs(gv - tv) <= 0.5


def categorize(bl, gv, tv, rec, in_gen):
    """Clasifica una discrepancia de campo. `rec` = BLRecord (o None)."""
    if not in_gen:
        return "CONEXION"  # BL en la verdad que no está en el manifiesto (Ex S3-)
    desc = " ".join(rec.description_lines).upper() if rec else ""
    basic = rec.charge("Basic FRT").monto if rec else 0
    has_roro = rec and any("roro" in k for k in rec.cargo.vehicles)
    if rec and basic == 0 and (rec.cargo.vehicles or rec.cargo.general or rec.cargo.containers):
        return "TRANSBORDO/SERVICE-BL"  # sin Basic Frt pese a tener carga
    if has_roro and any(k in desc for k in config.ROLLING_RORO_KEYWORDS):
        return "MOTORHOME (ambiguo)"
    if "8705" in desc:
        return "VEHICULO ESPECIAL (NCM 8705)"
    if _zero(gv) and not _zero(tv):
        return "GEN_VACIO (motor sin dato)"
    if _zero(tv) and not _zero(gv):
        return "VERDAD_VACIA (celda en blanco a mano)"
    return "VALOR (numérico distinto)"


# ─────────────────────────── auditoría de un viaje ───────────────────────────

def audit_voyage(code, voy, csv_impo, csv_expo, tmpdir):
    impo = parse_manifest(csv_impo, "IMPO") if csv_impo else None
    expo = parse_manifest(csv_expo, "EXPO") if csv_expo else None
    toll = detect_loose_tariff((impo or []) + (expo or []))
    cfg = config.EngineConfig(tariffs=config.LineCostTariffs(sc_loose_per_ton=toll))

    results = {}
    d = truth_dir(code, voy)
    if not d:
        return results, toll

    for op, bls, gen_fn, gen_fields, truth_fields, bl_col, kind, write in (
        ("IMPO", impo, "i.xlsx", IMPO_FIELDS, IMPO_FIELDS, 1, "IMPO", generar_planilla_impo),
        ("EXPO", expo, "e.xlsx", EXPO_FIELDS_GEN, EXPO_FIELDS_TRUTH, 2, "Expo", generar_planilla_expo),
    ):
        if bls is None:
            continue
        tp = truth_file(d, kind)
        if not tp:
            continue
        gen_path = os.path.join(tmpdir, gen_fn)
        write(bls, code, voy, gen_path, cfg=cfg)
        g = _extract(gen_path, bl_col, gen_fields)
        t = _extract(tp, bl_col, truth_fields)
        blmap = {b.bl_no.replace("[T]", ""): b for b in bls}

        fld = {fn: [0, 0] for fn in gen_fields}     # comparables, matches
        cats = Counter()
        for bl in set(g) | set(t):
            in_gen = bl in g
            grec = g.get(bl, {})
            trec = t.get(bl, {})
            for fn in gen_fields:
                gv, tv = grec.get(fn), trec.get(fn)
                if _zero(gv) and _zero(tv):
                    continue
                fld[fn][0] += 1
                if _match(gv, tv):
                    fld[fn][1] += 1
                else:
                    cats[categorize(bl, gv, tv, blmap.get(bl), in_gen)] += 1
        results[op] = {"fields": fld, "cats": cats}
    return results, toll


def main():
    base = sys.argv[1] if len(sys.argv) > 1 else "."
    os.chdir(base)
    ships = {}
    for f in glob.glob("*_IMPO_1.CSV") + glob.glob("*_EXPO_1.CSV"):
        m = re.match(r"([A-Z]{3})(\d{4})_(IMPO|EXPO)_", os.path.basename(f).upper())
        if m:
            ships.setdefault((m.group(1), m.group(2)), {})[m.group(3)] = f

    tmpdir = "_audit_tmp"
    os.makedirs(tmpdir, exist_ok=True)

    agg = {"IMPO": {"fields": defaultdict(lambda: [0, 0]), "cats": Counter()},
           "EXPO": {"fields": defaultdict(lambda: [0, 0]), "cats": Counter()}}
    per_voy = {}

    for (code, voy), files in sorted(ships.items()):
        res, toll = audit_voyage(code, voy, files.get("IMPO"), files.get("EXPO"), tmpdir)
        for op, r in res.items():
            comp = mat = 0
            for fn, (c, m) in r["fields"].items():
                agg[op]["fields"][fn][0] += c
                agg[op]["fields"][fn][1] += m
                comp += c
                mat += m
            agg[op]["cats"] += r["cats"]
            per_voy[f"{code}{voy}-{op}"] = (mat, comp, toll)

    for op in ("IMPO", "EXPO"):
        print(f"\n{'='*60}\n  {op} — paridad a nivel BL\n{'='*60}")
        T = [0, 0]
        for fn, (c, m) in agg[op]["fields"].items():
            T[0] += c
            T[1] += m
            if c:
                print(f"  {fn:8} {m:5}/{c:<5} {100*m/c:5.1f}%")
        if T[0]:
            print(f"  {'GLOBAL':8} {T[1]:5}/{T[0]:<5} {100*T[1]/T[0]:5.1f}%")
        print("  Discrepancias por categoría:")
        for cat, n in agg[op]["cats"].most_common():
            print(f"     {n:4}  {cat}")

    print(f"\n{'='*60}\n  Por viaje\n{'='*60}")
    for v, (m, c, toll) in sorted(per_voy.items()):
        if c:
            print(f"  {v:16} {100*m/c:5.1f}%   (toll {toll})")

    import shutil
    shutil.rmtree(tmpdir, ignore_errors=True)


if __name__ == "__main__":
    main()
