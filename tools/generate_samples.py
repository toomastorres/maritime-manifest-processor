"""
generate_samples.py - Genera manifiestos de ejemplo SINTÉTICOS.

Los datos (B/L, consignatarios, VIN, montos) son totalmente inventados; no
provienen de ningún manifiesto real. Reproducen el formato del export que
consume `manifest_engine` (CSV que envuelve una tabla pipe-delimited de ancho
fijo) para que la demo y las pruebas funcionen sin datos del cliente.

Uso:
    python tools/generate_samples.py            # escribe en ./samples
    python tools/generate_samples.py <carpeta>
"""

import sys
from pathlib import Path

SHIP = "ATLANTIC STAR"
VOY = "ATL0426"


def row(c1="", c2="", c3="", c4="", c5="", c6="", c7=""):
    """Construye una línea de datos del manifiesto (7 columnas pipe-delimited).

    El parser hace line.split("|"): parts[0]='2,  ', parts[1..7]=columnas,
    parts[-1]=',0;;'. La columna de cargos (c7) queda como parts[-2].
    """
    cols = [
        f"{c1:<15}", f"{c2:<36}", f"{c3:<19}", f"{c4:<30}",
        f"{c5:<12}", f"{c6:<12}", f"{c7:<55}",
    ]
    return "2,  |" + "|".join(cols) + "|,0;;"


def sep():
    return "2,  |" + "-" * 140 + "|,0;;"


def port_header(pol, pod):
    """Header de puertos + línea de valores (mismas posiciones de pipe).

    El parser localiza "port of loading"/"port of discharge" por índice de
    sección entre pipes y lee el mismo índice en la línea siguiente: aquí
    "Port Of Loading" cae en idx 4 y "Port Of Discharge" en idx 5.
    """
    label = "2,  |Nationality|Name Of Master|Place Of Receipt|Port Of Loading|Port Of Discharge|Place|,0;;"
    values = f"2,  |Italy|Capt. Demo|  |{pol}|{pod}|  |,0;;"
    return [label, values]


def header_block(pol, pod):
    out = [sep()]
    out.append("2,  |FREIGHT MANIFEST                         Atlantic Maritime Agency                         |,0;;")
    out.append(sep())
    out.append(f"2,  |Name Of Ship And Voyage No.|Move Type|Origin Port|,0;;")
    out.append(f'2,  |{SHIP}                :{VOY}|H : H|Demo Origin|,0;;')
    out.append(sep())
    out.extend(port_header(pol, pod))
    out.append(sep())
    out.append(row("B/L No.", "SHIPPER(SH), CONSIGNEE(CN), NOTIFY(NO",
                   "Unit No.", "Description Of Goods", "Weight(Kgs)",
                   "Measurement", "Charge Information"))
    out.append(sep())
    return out


def charge(desc, factor, basis, rate, total, cond, curr):
    """Línea de cargo estructurada (DESCRIPTION FACTOR BASIS RATE TOTAL+COND CURR)."""
    txt = f"{desc:<14} {factor:>7} {basis:<3} {rate:>10} {total:>12}{cond} {curr}"
    return row(c7=txt)


def total_line(label, amount, cond, curr):
    return row(c7=f"Total[{label}]              {amount:>12}{cond} {curr}")


# ─────────────────────────────────────────────────────────────────────────
# B/Ls de ejemplo (sintéticos)
# ─────────────────────────────────────────────────────────────────────────

def bl_container(bl, shipper, consignee, qty_ctr, container_no):
    """B/L de contenedor (40 ft High Cube)."""
    L = []
    # Valores ILUSTRATIVOS (no son tarifas reales).
    L.append(row(f"[{bl}]", f"SH:{shipper}", "", f"{qty_ctr}-40 ft. High Cube",
                 "", "", "Basic Frt.      1.000   40      500.00       500.00P USD"))
    L.append(row("", "ADDRESS LINE 1", "", "CONTAINER SAID TO CONTAIN", "", "",
                 "THC (POD)       1.000   40      120.00       120.00C USD"))
    L.append(row("", f"CN:{consignee}", "", f"{qty_ctr} PALLET(S)", "21217.50", "",
                 "S/C(POD)        1.000   40      150.00       150.00C USD"))
    L.append(row("", "CUIT-NO. 30-71111111-7", "", "", "", "",
                 "Sweeping        1.000   40       10.00        10.00C USD"))
    L.append(row("", "", f"CN : {container_no}", "", "", "",
                 "BAF             1.000   40      100.00       100.00P USD"))
    L.append(row(c7=""))
    L.append(total_line("Buenos Aires", "280.00", "C", "USD"))
    L.append(sep())
    return L


def bl_vehicles(bl, shipper, consignee, qty, make_vin):
    """B/L de autos nuevos (rolling)."""
    L = []
    # Valores ILUSTRATIVOS (no son tarifas reales).
    L.append(row(f"[{bl}]", f"SH:{shipper}", "", f"{qty}-New Car", "", "",
                 "Basic Frt.      1.000   PU      300.00     1,500.00P USD"))
    L.append(row("", "ADDRESS LINE 1", "", f"{make_vin['make']}", "", "",
                 "THC (POD)       1.000   PU      100.00       500.00C USD"))
    L.append(row("", f"CN:{consignee}", make_vin["vin"], "USED CAR", "1500.00", "",
                 "S/C(POD)        1.000   PU       20.00       100.00C USD"))
    L.append(row("", "CUIT-NO. 30-72222222-2", "", "", "", "",
                 "Sweeping        1.000   PU       10.00        50.00C USD"))
    L.append(row(c7=""))
    L.append(total_line("Buenos Aires", "650.00", "C", "USD"))
    L.append(sep())
    return L


def build_impo():
    lines = []
    lines.append("record_type,record_text,record_sort_id;;")
    lines.extend(header_block("ANTWERP", "ZARATE"))
    lines.extend(bl_container("S329600001", "COMPO DEMO GMBH",
                              "DEMO IMPORT SRL", 3, "DEMU1234567"))
    lines.extend(bl_vehicles("S329600002", "IMPORTADORA DELTAR",
                             "AGENCIA MARITIMA OCEANLINK", 5,
                             {"make": "DELTAR X3", "vin": "ZZDLTR8K9AA123456"}))
    lines.extend(bl_vehicles("S329600003", "CORVEX MOTORS",
                             "DEMO MOTORS SA", 4,
                             {"make": "CORVEX X5", "vin": "ZZCRVX9C50BC98765"}))
    return "\n".join(lines) + "\n"


def build_expo():
    lines = []
    lines.append("record_type,record_text,record_sort_id;;")
    lines.extend(header_block("ZARATE", "HAMBURG"))
    lines.extend(bl_container("S329700001", "DEMO EXPORT SA",
                              "EURO RECEIVER GMBH", 2, "DEMU7654321"))
    lines.extend(bl_vehicles("S329700002", "DEMO CARS SA",
                             "EXIMPORT INTL SA", 6,
                             {"make": "AVALON 500", "vin": "ZZAVL120000065432"}))
    return "\n".join(lines) + "\n"


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("samples")
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{VOY}_IMPO_1.CSV").write_text(build_impo(), encoding="latin-1")
    (out_dir / f"{VOY}_EXPO_1.CSV").write_text(build_expo(), encoding="latin-1")
    print(f"Manifiestos de ejemplo escritos en: {out_dir.resolve()}")


if __name__ == "__main__":
    main()
