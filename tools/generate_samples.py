"""
generate_samples.py - Genera manifiestos de ejemplo 100 % SINTÉTICOS (v2.0).

Todos los datos (buque, viaje, números de B/L, embarcadores, consignatarios,
CUIT con dígito verificador módulo 11, VIN de 17 caracteres y tarifas) son
totalmente ficticios con fines demostrativos de portafolio.

Reproduce fielmente el formato pipe-delimited envuelto en CSV que consume
`manifest_engine` v2.0 para ejercitar todas sus capacidades:
  - Filtrado de reposicionamientos vacíos (`P3...`) vs B/L comerciales (`S3...`).
  - Condición ALL IN (80 % comisionable) vs B/L con recargos de flete (BAF/ETS).
  - Mínimo por contenedor, tasa diferencial a Brasil y Control Equipment Fees.
  - Rodados (Car, Van, RoRo) con VINs ficticios y sufijo de agencia forwarder.
  - B/L de carga mixta (RoRo + Break-Bulk) con partición automática exacta.
  - Liquidación de Caja (`Facturación.xlsx`): CUIT módulo 11, multimoneda
    (USD/EUR), subagente `Buenos Aires - Partner` y búsqueda combinatoria de
    conceptos locales prepagos en exportación.

Uso:
    python tools/generate_samples.py            # escribe en ./samples
    python tools/generate_samples.py <carpeta>
"""

import sys
from pathlib import Path

SHIP = "ATLANTIC STAR"
VOY = "ATL0426"
CUIT_WEIGHTS = (5, 4, 3, 2, 7, 6, 5, 4, 3, 2)


def make_cuit(prefix: str, body8: str) -> str:
    """Construye una CUIT sintética válida según módulo 11."""
    assert len(prefix) == 2 and len(body8) == 8
    base = prefix + body8
    resto = 11 - sum(int(d) * w for d, w in zip(base, CUIT_WEIGHTS)) % 11
    if resto == 10:
        base = "33" + body8
        resto = 11 - sum(int(d) * w for d, w in zip(base, CUIT_WEIGHTS)) % 11
    dv = 0 if resto == 11 else resto
    return f"{base[:2]}-{base[2:10]}-{dv}"


def row(c1="", c2="", c3="", c4="", c5="", c6="", c7=""):
    """Construye una línea de datos del manifiesto (7 columnas pipe-delimited)."""
    cols = [
        f"{c1:<15}",
        f"{c2:<36}",
        f"{c3:<19}",
        f"{c4:<30}",
        f"{c5:<12}",
        f"{c6:<14}",
        f"{c7:<55}",
    ]
    return "2,  |" + "|".join(cols) + "|,0;;"


def sep():
    return "2,  |" + "-" * 145 + "|,0;;"


def port_header(pol: str, pod: str):
    """Encabezado de puertos + valores en las mismas posiciones de pipe."""
    label = "2,  |Nationality|Name Of Master|Place Of Receipt|Port Of Loading|Port Of Discharge|Place|,0;;"
    values = f"2,  |Panama|Capt. Demo|  |{pol}|{pod}|  |,0;;"
    return [label, values]


def header_block(pol: str, pod: str):
    out = [sep()]
    out.append("2,  |FREIGHT MANIFEST                         Atlantic Maritime Agency                         |,0;;")
    out.append(sep())
    out.append("2,  |Name Of Ship And Voyage No.|Move Type|Origin Port|,0;;")
    out.append(f"2,  |{SHIP:<28}:{VOY}|H : H|Demo Port|,0;;")
    out.append(sep())
    out.extend(port_header(pol, pod))
    out.append(sep())
    out.append(
        row(
            "B/L No.",
            "SHIPPER(SH), CONSIGNEE(CN), NOTIFY(NO",
            "Unit No.",
            "Description Of Goods",
            "Weight(Kgs)",
            "Measurement",
            "Charge Information",
        )
    )
    out.append(sep())
    return out


def charge_str(desc: str, factor: str, basis: str, rate: str, total: str, cond: str, curr: str) -> str:
    return f"{desc:<15} {factor:>7} {basis:<3} {rate:>10} {total:>12}{cond} {curr}"


def total_line(label: str, amount: str, cond: str, curr: str) -> str:
    return row(c7=f"Total[{label}]          {amount:>12}{cond} {curr}")


# ─────────────────────────────────────────────────────────────────────────
# Bloques sintéticos de IMPO
# ─────────────────────────────────────────────────────────────────────────

def bl_empty_equipment(bl_p: str, pol_or_pod: str) -> list[str]:
    """B/L P3... de reposicionamiento vacío (debe ser descartado por el parser)."""
    return [
        row(f"[{bl_p}]", "SH:ATLANTIC MARITIME LINES", "", "2-40 ft. High Cube", "4,400.00", "", ""),
        row("", "CN:ATLANTIC MARITIME AGENCY", "ATLU0000099", "EMPTY REPOSITIONING", "", "", ""),
        sep(),
    ]


def build_impo() -> str:
    cuit_1 = make_cuit("30", "71000101")
    cuit_2 = make_cuit("30", "71000102")
    cuit_3 = make_cuit("30", "71000103")
    cuit_4 = make_cuit("30", "71000104")
    cuit_5 = make_cuit("30", "71000105")

    lines = ["record_type,record_text,record_sort_id;;"]
    lines.extend(header_block("ANTWERP", "ZARATE"))

    # 0) Equipo vacío P3 (descartado automáticamente)
    lines.extend(bl_empty_equipment("P300000901", "ZARATE"))

    # 1) S300000101 — Contenedores con recargos (NO ALL IN), aplica mínimo 3x25=75 USD,
    #    multimoneda (USD + EUR) y discrepancia ilustrativa de Sweeping para auditoría.
    lines.extend([
        row("[S300000101]", "SH:COMPO NORDIC GMBH", "", "3-40 ft. High Cube", "21,217.50", "180.000 CBM",
            charge_str("Basic Frt.", "3.000", "40", "600.00", "1,800.00", "P", "USD")),
        row("", "HARBORSTRASSE 12", "ATLU1040011", "INDUSTRIAL COMPONENTS", "", "",
            charge_str("BAF", "3.000", "40", "100.00", "300.00", "P", "USD")),
        row("", "CN:INDUSTRIAL PAMPA SRL", "ATLU1040012", "30 PALLETS", "", "",
            charge_str("ISPS(POL)", "3.000", "40", "12.00", "36.00", "P", "EUR")),
        row("", f"CUIT {cuit_1}", "ATLU1040013", "", "", "",
            charge_str("THC (POD)", "3.000", "40", "120.00", "360.00", "C", "USD")),
        row("", "NO:INDUSTRIAL PAMPA SRL", "", "", "", "", "Toll"),
        row("", "", "", "", "", "",
            charge_str("S/C(POD)", "3.000", "40", "150.00", "450.00", "C", "USD")),
        row("", "", "", "", "", "",
            charge_str("Sweeping", "3.000", "40", "12.00", "36.00", "C", "USD")),
        total_line("Antwerp", "2,100.00", "P", "USD"),
        total_line("Antwerp", "36.00", "P", "EUR"),
        total_line("Buenos Aires", "846.00", "C", "USD"),
        sep(),
    ])

    # 2) S300000102 — Contenedor ALL IN (sólo Basic Frt + gastos locales POD -> 80% comisionable)
    #    y cobro vía subagente local 'Buenos Aires - Partner'.
    lines.extend([
        row("[S300000102]", "SH:EUROPARTS BV", "", "1-20 ft. Dry Cargo", "14,500.00", "28.000 CBM",
            charge_str("Basic Frt.", "1.000", "20", "2,000.00", "2,000.00", "P", "USD")),
        row("", "ROTTERDAMSEWEG 44", "ATLU2020022", "AUTO SPARE PARTS", "", "",
            charge_str("THC (POD)", "1.000", "20", "100.00", "100.00", "C", "USD")),
        row("", "CN:LOGISTICA DEL PLATA SA", "", "", "", "",
            charge_str("S/C(POD)", "1.000", "20", "150.00", "150.00", "C", "USD")),
        row("", f"CUIT-NO. {cuit_2}", "", "", "", "", ""),
        total_line("Antwerp", "2,000.00", "P", "USD"),
        total_line("Buenos Aires - Partner", "250.00", "C", "USD"),
        sep(),
    ])

    # 3) S300000103 — Rodados (5 New Car(s)), flete en EUR, consignatario forwarder OCEANLINK,
    #    5 VINs sintéticos de 17 caracteres en columna de descripción.
    lines.extend([
        row("[S300000103]", "SH:DELTAR AUTOMOTIVE", "", "5-New Car(s)", "7,500.00", "62.500 CBM",
            charge_str("Basic Frt.", "5.000", "PU", "400.00", "2,000.00", "P", "EUR")),
        row("", "ANTWERP TERMINAL", "", "DELTAR X3 SEDAN", "", "",
            charge_str("BAF", "5.000", "PU", "40.00", "200.00", "P", "EUR")),
        row("", "CN:AGENCIA MARITIMA OCEANLINK", "", "ZZDLTR8K9AA100001", "", "",
            charge_str("THC (POD)", "5.000", "PU", "90.00", "450.00", "C", "USD")),
        row("", f"CUIT {cuit_3}", "", "ZZDLTR8K9AA100002", "", "",
            charge_str("S/C(POD)", "7.500", "MT", "20.00", "150.00", "C", "USD")),
        row("", "", "", "ZZDLTR8K9AA100003", "", "", ""),
        row("", "", "", "ZZDLTR8K9AA100004", "", "", ""),
        row("", "", "", "ZZDLTR8K9AA100005", "", "", ""),
        total_line("Antwerp", "2,200.00", "P", "EUR"),
        total_line("Buenos Aires", "600.00", "C", "USD"),
        sep(),
    ])

    # 4) S300000104 — Utilitarios livianos (4 New Small Van(s)) ALL IN en USD.
    lines.extend([
        row("[S300000104]", "SH:CORVEX COMMERCIAL", "", "4-New Small Van(s)", "8,800.00", "76.000 CBM",
            charge_str("Basic Frt.", "4.000", "PU", "450.00", "1,800.00", "P", "USD")),
        row("", "INDUSTRIEWEG 8", "", "CORVEX CARGO VAN", "", "",
            charge_str("THC (POD)", "4.000", "PU", "95.00", "380.00", "C", "USD")),
        row("", "CN:AUTOMOTORES AUSTRAL SA", "", "ZZCRVX9C50BC20001", "", "",
            charge_str("S/C(POD)", "8.800", "MT", "20.00", "176.00", "C", "USD")),
        row("", f"C.U.I.T. {cuit_4}", "", "ZZCRVX9C50BC20002", "", "", ""),
        row("", "", "", "ZZCRVX9C50BC20003", "", "", ""),
        row("", "", "", "ZZCRVX9C50BC20004", "", "", ""),
        total_line("Antwerp", "1,800.00", "P", "USD"),
        total_line("Buenos Aires", "556.00", "C", "USD"),
        sep(),
    ])

    # 5) S300000105 — Carga mixta (1 New LM RoRo + 4 CASE(S)) con partición automática exacta
    #    por coincidencia de medidas (6.500 LM y 32.000 CBM en flete; 18.000 Tn y 5.000 Tn en TOLL).
    lines.extend([
        row("[S300000105]", "SH:HELIOS MACHINERY GMBH", "", "1-New LM RoRo", "18,000.00", "6.500 LM",
            charge_str("Basic Frt.", "6.500", "ME", "400.00", "2,600.00", "P", "EUR")),
        row("", "WERKSTRASSE 90", "", "HELIOS EXCAVATOR HX180", "", "",
            charge_str("Basic Frt.", "32.000", "WMM", "80.00", "2,560.00", "P", "EUR")),
        row("", "CN:VIALIDAD DEL SUR SA", "", "ZZHLSM4X700030001", "", "", "ETS+FUEL EU"),
        row("", f"CUIT {cuit_5}", "", "4-CASE(S)", "5,000.00", "32.000 CBM",
            charge_str("Surcharge", "1.000", "AA", "140.00", "140.00", "P", "EUR")),
        row("", "", "", "ACCESSORY CRATES", "", "",
            charge_str("S/C(POD)", "18.000", "MT", "20.00", "360.00", "C", "USD")),
        row("", "", "", "", "", "",
            charge_str("S/C(POD)", "5.000", "MT", "20.00", "100.00", "C", "USD")),
        total_line("Antwerp", "5,300.00", "P", "EUR"),
        total_line("Buenos Aires", "460.00", "C", "USD"),
        sep(),
    ])

    return "\n".join(lines) + "\n"


# ─────────────────────────────────────────────────────────────────────────
# Bloques sintéticos de EXPO
# ─────────────────────────────────────────────────────────────────────────

def build_expo() -> str:
    cuit_e1 = make_cuit("30", "72000201")
    cuit_e2 = make_cuit("30", "72000202")
    cuit_e3 = make_cuit("30", "72000203")
    cuit_e4 = make_cuit("30", "72000204")

    lines = ["record_type,record_text,record_sort_id;;"]

    # ── Destino Europa (HAMBURG — tasa general EXPO 4,0% + ENS) ──────────
    lines.extend(header_block("ZARATE", "HAMBURG"))
    lines.extend(bl_empty_equipment("P300000902", "HAMBURG"))

    # 1) S300000201 — Contenedores a Hamburgo con ENS y búsqueda combinatoria Prepaid:
    #    6 cargos P USD en total (3.135 USD), de los cuales 4 son locales en Buenos Aires (585 USD)
    #    y 2 se pagan en Casa Matriz / Exterior (2.550 USD).
    lines.extend([
        row("[S300000201]", "SH:AGROEXPORT PAMPA SA", "", "2-40 ft. High Cube", "42,000.00", "132.000 CBM",
            charge_str("Basic Frt.", "2.000", "40", "1,100.00", "2,200.00", "P", "USD")),
        row("", f"AV. CORDOBA 1500 / CUIT {cuit_e1}", "ATLU4080011", "ORGANIC GRAINS", "", "",
            charge_str("BAF", "2.000", "40", "175.00", "350.00", "P", "USD")),
        row("", "CN:NORDIC FOODS GMBH", "ATLU4080012", "40 BAGS", "", "",
            charge_str("THC (POL)", "2.000", "40", "120.00", "240.00", "P", "USD")),
        row("", "HAMBURG PORT", "", "", "", "",
            charge_str("S/C(POL)", "2.000", "40", "150.00", "300.00", "P", "USD")),
        row("", "", "", "", "", "",
            charge_str("Sweeping", "2.000", "40", "10.00", "20.00", "P", "USD")),
        row("", "", "", "", "", "",
            charge_str("ENS", "1.000", "PS", "25.00", "25.00", "P", "USD")),
        row("", "", "", "", "", "",
            charge_str("ISPS(POD)", "2.000", "40", "12.00", "24.00", "C", "EUR")),
        total_line("Buenos Aires", "585.00", "P", "USD"),
        total_line("Head Office", "2,550.00", "P", "USD"),
        total_line("Hamburg", "24.00", "C", "EUR"),
        sep(),
    ])

    # 2) S300000202 — Rodados de exportación (6 New Car(s) marca AVALON) con ENS.
    lines.extend([
        row("[S300000202]", "SH:TERMINAL AUTOMOTRIZ DEL SUR SA", "", "6-New Car(s)", "9,000.00", "75.000 CBM",
            charge_str("Basic Frt.", "6.000", "PU", "350.00", "2,100.00", "C", "EUR")),
        row("", f"CUIT {cuit_e2}", "", "AVALON 500 PICKUP", "", "",
            charge_str("THC (POL)", "6.000", "PU", "85.00", "510.00", "P", "USD")),
        row("", "CN:EURO AUTO LOGISTICS NV", "", "ZZAVLN12000060001", "", "",
            charge_str("S/C(POL)", "9.000", "MT", "20.00", "180.00", "P", "USD")),
        row("", "", "", "ZZAVLN12000060002", "", "",
            charge_str("ENS", "1.000", "PS", "25.00", "25.00", "P", "USD")),
        row("", "", "", "ZZAVLN12000060003", "", "", ""),
        row("", "", "", "ZZAVLN12000060004", "", "", ""),
        row("", "", "", "ZZAVLN12000060005", "", "", ""),
        row("", "", "", "ZZAVLN12000060006", "", "", ""),
        total_line("Buenos Aires", "715.00", "P", "USD"),
        total_line("Hamburg", "2,100.00", "C", "EUR"),
        sep(),
    ])

    # ── Destino Brasil (SANTOS — tasa regional EXPO 3,5%) ────────────────
    lines.extend(header_block("ZARATE", "SANTOS"))

    # 3) S300000203 — Contenedores ALL IN a Santos (2x20', 80% x 3,5% = 67,20 USD > mín 50 USD).
    lines.extend([
        row("[S300000203]", "SH:QUIMICA RIOPLATENSE SA", "", "2-20 ft. Dry Cargo", "32,000.00", "56.000 CBM",
            charge_str("Basic Frt.", "2.000", "20", "1,200.00", "2,400.00", "P", "USD")),
        row("", f"CUIT {cuit_e3}", "ATLU2050091", "CHEMICAL PRODUCTS", "", "",
            charge_str("THC (POL)", "2.000", "20", "100.00", "200.00", "P", "USD")),
        row("", "CN:SUL QUIMICA LTDA", "ATLU2050092", "", "", "",
            charge_str("S/C(POL)", "2.000", "20", "150.00", "300.00", "P", "USD")),
        total_line("Buenos Aires", "2,900.00", "P", "USD"),
        sep(),
    ])

    # 4) S300000204 — Break-Bulk sobre equipo Mafi 40 a Santos (tasa Brasil 3,5%, sin CEF ni mín ctr).
    lines.extend([
        row("[S300000204]", "SH:METALURGICA AUSTRAL SA", "", "1-Mafi 40", "24,000.00", "48.000 CBM",
            charge_str("Basic Frt.", "48.000", "WMM", "50.00", "2,400.00", "P", "USD")),
        row("", f"CUIT {cuit_e4}", "MAFI4001", "2-CASE(S)", "", "",
            charge_str("BAF", "48.000", "WMM", "8.00", "384.00", "P", "USD")),
        row("", "CN:ACOS PAULISTAS SA", "", "STEEL COILS ON MAFI", "", "",
            charge_str("S/C(POL)", "24.000", "MT", "20.00", "480.00", "P", "USD")),
        total_line("Buenos Aires", "3,264.00", "P", "USD"),
        sep(),
    ])

    return "\n".join(lines) + "\n"


def main():
    out_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("samples")
    out_dir.mkdir(parents=True, exist_ok=True)
    (out_dir / f"{VOY}_IMPO_1.CSV").write_text(build_impo(), encoding="latin-1")
    (out_dir / f"{VOY}_EXPO_1.CSV").write_text(build_expo(), encoding="latin-1")
    print(f"Manifiestos sintéticos v2.0 escritos en: {out_dir.resolve()}")


if __name__ == "__main__":
    main()
