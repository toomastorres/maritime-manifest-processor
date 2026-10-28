"""
streamlit_app.py - Demo web interactiva del motor de manifiestos marítimos (v2.0).

Permite procesar los manifiestos sintéticos de ejemplo o subir archivos CSV
propios (IMPO/EXPO), inspeccionar Bills of Lading, liquidación de comisiones
gobernada por catálogo Excel, alertas accionables de auditoría (REVISIÓN) y
descargar las 9 planillas Excel del viaje (incluyendo Facturación para Caja).

Ejecución local:
    streamlit run streamlit_app.py
"""

import io
import re
import tempfile
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

from manifest_engine import config
from manifest_engine.__main__ import parse_agency, process_directory
from manifest_engine.catalog import CONTENEDOR, RODADOS, load_catalog
from manifest_engine.domain import format_cargo
from manifest_engine.parsing import parse_manifest
from manifest_engine.rules.commission import evaluate_all
from manifest_engine.rules.line_costs import detect_loose_tariff
from manifest_engine.rules.split import filas_del_bl, split_alerts
from manifest_engine.validation import Alerta, revisar_bls

st.set_page_config(
    page_title="Procesador de Manifiestos — AMA v2.0",
    page_icon="🚢",
    layout="wide",
)

SAMPLES = Path(__file__).parent / "samples"
NAME_RE = re.compile(r"([A-Z]{3})(\d{4})_(IMPO|EXPO)_", re.IGNORECASE)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Cabecera ──────────────────────────────────────────────────────────────
st.title("🚢 Procesador de Manifiestos Marítimos — v2.0")
st.caption(
    "Motor determinístico gobernado por catálogo Excel (`Definiciones_reglas_comisiones.xlsx`): "
    "liquida Biblias IMPO/EXPO, Comisiones Inward/Outward, Facturación para Caja (CUIT Módulo 11), "
    "ENS, Chasis y resúmenes operativos."
)
st.info(
    "Datos de demostración **100 % sintéticos** (*Atlantic Maritime Agency — AMA*): "
    "buques, clientes, CUIT, VIN y tarifas son ficticios e ilustrativos de portafolio.",
    icon="ℹ️",
)

# ── Controles laterales ───────────────────────────────────────────────────
if "demo_mode" not in st.session_state:
    st.session_state["demo_mode"] = False

with st.sidebar:
    st.header("⚙️ Parámetros del Viaje")
    roe = st.number_input("ROE (EUR → USD)", 0.50, 5.00, config.DEFAULT_ROE, 0.01)
    agency_raw = st.text_input(
        "B/L de rodados cerrados por AMA",
        value="S300000103",
        help=(
            "Números S3... cerrados comercialmente por AMA (comisionan 1 % del flete). "
            "Escribí NINGUNO si no hubo ninguno, o dejalo vacío para modo ESTIMACIÓN."
        ),
    )
    st.divider()
    st.markdown("**Entrada de manifiestos**")
    if st.button("▶️ Usar manifiestos de ejemplo", use_container_width=True, type="primary"):
        st.session_state["demo_mode"] = True
    uploaded = st.file_uploader(
        "…o subí tus CSV",
        type=["csv"],
        accept_multiple_files=True,
        help="Formato esperado: ATL0426_IMPO_1.CSV / ATL0426_EXPO_1.CSV",
    )
    if uploaded:
        st.session_state["demo_mode"] = False


def _familia_label(bl) -> str:
    fams = bl.families()
    if len(fams) > 1:
        return "Carga mixta (" + " + ".join(fams) + ")"
    if CONTENEDOR in fams:
        return "Contenedor"
    if any(f in RODADOS for f in fams):
        return "Rodados"
    if fams:
        return "Break-Bulk"
    return "Sin clasificar"


def bls_to_df(bls, op: str) -> pd.DataFrame:
    rows = []
    for b in bls:
        frt = b.charge("Basic FRT")
        toll = b.charge("Toll")
        thc = b.charge("THC")
        row_data = {
            "B/L": b.bl_no,
            "Operación": op,
            "Tipo": _familia_label(b),
            "Puerto": b.port_of_loading if op == "IMPO" else b.port_of_discharge,
            "Cliente (CN/SH)": b.entity or "—",
            "CUIT (Mód. 11)": b.party_id or "—",
            "Carga": format_cargo(b) or "—",
            "Peso (kg)": round(b.weight, 1),
            "Basic FRT": f"{frt.monto:,.2f} {frt.moneda}".strip() if frt.monto else "—",
            "THC": f"{thc.monto:,.2f} {thc.moneda}".strip() if thc.monto else "—",
            "TOLL": f"{toll.monto:,.2f} {toll.moneda}".strip() if toll.monto else "—",
            "VINs": f"{len(b.vins)} ({', '.join(b.vins[:2])}{'…' if len(b.vins) > 2 else ''})" if b.vins else "—",
        }
        if op == "IMPO":
            row_data["Collect Bs.As. USD"] = round(b.totals.get("Total Buenos Aires USD", 0.0), 2)
        else:
            row_data["ENS (Biblia)"] = round(b.charge("ENS").monto, 2)
            row_data["Prepaid Bs.As."] = (
                f"{b.totals['Total Buenos Aires Monto']:,.2f} {b.totals['Total Buenos Aires Moneda']}"
                if b.totals.get("Total Buenos Aires Monto")
                else "—"
            )
        rows.append(row_data)
    return pd.DataFrame(rows)


def commissions_to_df(comm_results) -> pd.DataFrame:
    rows = []
    for r in comm_results:
        cef_usd = r.cef_units * config.CONTROL_EQUIPMENT_FEE_USD
        comm_parts = r.total_usd or 0.0
        rows.append({
            "B/L": r.bl.bl_no,
            "Op.": r.op,
            "Puerto": r.port,
            "Grupos": " + ".join(r.groups) if r.groups else "—",
            "ALL IN (80%)": "SÍ" if r.all_in else "NO",
            "Brasil": "SÍ" if r.brasil else "NO",
            "Modo Rodados": r.rodado_mode or "—",
            "CEF (ctr)": r.cef_units,
            "Comisión Flete (USD)": round(comm_parts, 2) if r.ok else None,
            "CEF (USD)": round(cef_usd, 2) if r.ok else None,
            "Total Agencia (USD)": round(comm_parts + cef_usd, 2) if r.ok else None,
            "Estado": r.status,
            "Notas / Motivo": "; ".join(r.reasons if not r.ok else r.notes) or "—",
        })
    return pd.DataFrame(rows)


def collect_alerts(bls, op: str, cfg: config.EngineConfig, catalog, comm_results) -> list[dict]:
    alerts = list(revisar_bls(bls, op, cfg, catalog=catalog))
    for r in comm_results:
        if not r.ok:
            alerts.append(Alerta(r.bl.bl_no, "COMISIÓN - REVISAR", "; ".join(r.reasons)))
        filas = filas_del_bl(r.bl, r, op)
        for aviso in split_alerts(filas):
            alerts.append(Alerta(r.bl.bl_no, "CARGA MIXTA", aviso))
    if cfg.agency_bls is None and any(r.rodado_mode for r in comm_results):
        alerts.insert(
            0,
            Alerta(
                "—",
                "LISTA COMERCIAL",
                "Sin lista de B/L de Comercial: rodados calculados en modo ESTIMACIÓN por unidad.",
            ),
        )
    return [{"Operación": op, "B/L": a.bl, "Categoría": a.tipo, "Detalle": a.detalle} for a in alerts]


def zip_reports(files: list[Path]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in files:
            zf.write(f, f.name)
    return buf.getvalue()


def process(files_payload: list[tuple[str, bytes]], agency_bls):
    valid = [(n, b) for (n, b) in files_payload if NAME_RE.match(n)]
    if not valid:
        st.error("Ningún archivo coincide con el patrón `<COD><VIAJE>_IMPO_1.CSV` / `_EXPO_1.CSV`.")
        return

    catalog = load_catalog()

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        in_dir = tmp_dir / "in"
        out_dir = tmp_dir / "out"
        in_dir.mkdir()
        out_dir.mkdir()

        for name, data in valid:
            (in_dir / name).write_bytes(data)

        t0 = time.perf_counter()
        process_directory(
            str(in_dir),
            str(out_dir),
            roe=roe,
            agency_override=agency_bls,
            register_concepts=False,
        )
        elapsed = time.perf_counter() - t0

        reports = sorted(p for p in out_dir.rglob("*.xlsx"))

        impo_bls, expo_bls = [], []
        for name, _ in valid:
            op = "IMPO" if "IMPO" in name.upper() else "EXPO"
            parsed = parse_manifest(str(in_dir / name), op, catalog=catalog)
            if op == "IMPO":
                impo_bls.extend(parsed)
            else:
                expo_bls.extend(parsed)

        loose = detect_loose_tariff(impo_bls + expo_bls)
        cfg = config.EngineConfig(
            roe=roe,
            line_tariffs=config.LineCostTariffs(sc_loose_per_ton=loose),
            agency_bls=agency_bls,
        )

        comm_impo = evaluate_all(impo_bls, "IMPO", cfg, catalog=catalog) if impo_bls else []
        comm_expo = evaluate_all(expo_bls, "EXPO", cfg, catalog=catalog) if expo_bls else []
        all_comm = comm_impo + comm_expo

        alerts_rows = []
        if impo_bls:
            alerts_rows.extend(collect_alerts(impo_bls, "IMPO", cfg, catalog, comm_impo))
        if expo_bls:
            alerts_rows.extend(collect_alerts(expo_bls, "EXPO", cfg, catalog, comm_expo))

        impo_df = bls_to_df(impo_bls, "IMPO") if impo_bls else pd.DataFrame()
        expo_df = bls_to_df(expo_bls, "EXPO") if expo_bls else pd.DataFrame()
        comm_df = commissions_to_df(all_comm)
        alerts_df = pd.DataFrame(alerts_rows)

        total_comm_usd = sum(
            (r.total_usd or 0.0) + r.cef_units * config.CONTROL_EQUIPMENT_FEE_USD
            for r in all_comm
            if r.ok
        )

        # ── KPIs principales ──
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Bills of Lading (S3)", len(impo_bls) + len(expo_bls))
        k2.metric("Comisión Agencia (USD)", f"${total_comm_usd:,.2f}")
        k3.metric("Alertas REVISIÓN", len(alerts_rows))
        k4.metric("Reportes Excel", len(reports))
        k5.metric("Tiempo de motor", f"{elapsed * 1000:.0f} ms")

        tab_resumen, tab_impo, tab_expo, tab_rev, tab_dl = st.tabs(
            [
                "📊 Resumen y Comisiones",
                "📥 Biblia IMPO",
                "📤 Biblia EXPO",
                f"🔍 Alertas REVISIÓN ({len(alerts_rows)})",
                "⬇️ Descargas Excel",
            ]
        )

        with tab_resumen:
            all_bls_df = pd.concat(
                [d[["B/L", "Operación", "Tipo"]] for d in (impo_df, expo_df) if not d.empty],
                ignore_index=True,
            )
            if not all_bls_df.empty:
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown("**Bills of Lading por familia de carga**")
                    st.bar_chart(all_bls_df.groupby("Tipo")["B/L"].count().rename("B/Ls"), color="#1f4e79")
                with col_b:
                    st.markdown("**Comisión total (USD) por operación**")
                    if not comm_df.empty and comm_df["Total Agencia (USD)"].notna().any():
                        st.bar_chart(
                            comm_df.groupby("Op.")["Total Agencia (USD)"].sum(),
                            color="#3d6fa5",
                        )
            st.markdown("#### Liquidación gobernada por catálogo (`DETALLE B-L`)")
            st.dataframe(comm_df, use_container_width=True, hide_index=True)
            st.caption(
                f"ROE aplicado: EUR 1,00 = {roe:.2f} USD · "
                f"TOLL carga suelta detectado: {loose:g} USD/Tn · "
                f"Catálogo activo: {len(catalog.concepts)} conceptos, {len(catalog.kinds)} bultos, {len(catalog.ports)} puertos EXPO."
            )

        with tab_impo:
            if not impo_df.empty:
                st.dataframe(impo_df, use_container_width=True, hide_index=True)
            else:
                st.info("Sin manifiesto IMPO en la selección actual.")

        with tab_expo:
            if not expo_df.empty:
                st.dataframe(expo_df, use_container_width=True, hide_index=True)
            else:
                st.info("Sin manifiesto EXPO en la selección actual.")

        with tab_rev:
            if not alerts_df.empty:
                st.warning(
                    "El motor nunca inventa ni pisa importes: las anomalías se exponen en la hoja **REVISIÓN** de cada planilla.",
                    icon="⚠️",
                )
                st.dataframe(alerts_df, use_container_width=True, hide_index=True)
            else:
                st.success("Sin alertas de revisión para este viaje.")

        with tab_dl:
            st.success(f"✅ {len(reports)} planillas Excel generadas (incluye Facturación para Caja y Comisiones v2.0).")
            cols = st.columns(3)
            for i, r in enumerate(reports):
                cols[i % 3].download_button(
                    f"⬇️ {r.name}",
                    r.read_bytes(),
                    file_name=r.name,
                    mime=XLSX_MIME,
                    use_container_width=True,
                )
            st.divider()
            st.download_button(
                "📦 Descargar paquete completo (.zip)",
                zip_reports(reports),
                file_name="reportes_viaje_v2.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )


agency_parsed, invalid_agency_tokens = parse_agency(agency_raw)
if invalid_agency_tokens:
    st.sidebar.warning(f"Tokens ignorados en lista AMA: {', '.join(invalid_agency_tokens)}")

if uploaded:
    process([(f.name, f.getvalue()) for f in uploaded], agency_parsed)
elif st.session_state["demo_mode"]:
    process([(f.name, f.read_bytes()) for f in sorted(SAMPLES.glob("*.CSV"))], agency_parsed)
else:
    st.markdown(
        "👈 Hacé clic en **Usar manifiestos de ejemplo** en la barra lateral o subí tus propios archivos CSV.\n\n"
        "**Novedades de la arquitectura v2.0:**\n"
        "- **Catálogo gobernado en Excel (`Definiciones_reglas_comisiones.xlsx`):** 84 conceptos de cargo, "
        "38 tipos de bulto y 20 puertos de exportación con auto-descubrimiento de conceptos inéditos.\n"
        "- **Reglas `ALL IN` (80 %) y partición de cargas mixtas (`split.py`):** separación exacta de flete y TOLL "
        "por coincidencia de medidas, `Tariff ID` o descarte determinístico.\n"
        "- **Facturación para Caja (`Facturación.xlsx`):** extracción y validación Módulo 11 de CUIT, "
        "subagente local `Buenos Aires - Partner` y búsqueda combinatoria de líneas prepagas locales.\n"
        "- **Principio *Never Invent Values*:** 8 categorías tipadas de auditoría en la hoja `REVISIÓN`."
    )
