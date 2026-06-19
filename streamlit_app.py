"""
streamlit_app.py - Demo web interactiva del motor de procesamiento de manifiestos.

Sube manifiestos CSV (IMPO/EXPO) o usa los de ejemplo, parsea los Bills of Lading,
muestra métricas, tablas y un gráfico de carga, y deja descargar los reportes Excel.
Reutiliza `manifest_engine` sin modificarlo.

Local:   streamlit run streamlit_app.py
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
from manifest_engine.parsing import parse_manifest
from manifest_engine.domain.cargo import format_cargo_output
from manifest_engine.__main__ import process_directory

st.set_page_config(page_title="Procesador de Manifiestos", page_icon="🚢", layout="wide")

SAMPLES = Path(__file__).parent / "samples"
NAME_RE = re.compile(r"([A-Z]{3})(\d{4})_(IMPO|EXPO)_", re.IGNORECASE)
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Cabecera ──────────────────────────────────────────────────────────────
st.title("🚢 Procesador de Manifiestos Marítimos")
st.caption(
    "Lee manifiestos de carga (IMPO/EXPO) y genera en segundos las planillas de "
    "liquidación: gastos de línea, comisiones, ENS, chasis y resúmenes por B/L y marca."
)
st.info(
    "Datos de ejemplo **100% sintéticos**: el motor y las reglas de negocio son reales, "
    "pero los manifiestos de muestra no contienen información de ningún cliente.",
    icon="ℹ️",
)

# ── Controles ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuración")
    roe = st.number_input("ROE (EUR → USD)", 0.5, 5.0, config.DEFAULT_ROE, 0.01)
    st.divider()
    st.markdown("**Entrada**")
    use_samples = st.button("▶️ Usar manifiestos de ejemplo", use_container_width=True, type="primary")
    uploaded = st.file_uploader(
        "…o subí tus CSV", type=["csv"], accept_multiple_files=True,
        help="Nombre esperado: ATLNNNN_IMPO_1.CSV / ATLNNNN_EXPO_1.CSV",
    )


def _tipo(cargo):
    if cargo.containers:
        return "Contenedor"
    if cargo.vehicles:
        return "Vehículos"
    if cargo.general:
        return "Carga general"
    return "Otro"


def bls_to_df(bls, op):
    rows = []
    for b in bls:
        rows.append({
            "B/L": b.bl_no,
            "Operación": op,
            "Tipo": _tipo(b.cargo),
            "Puerto": b.port_of_discharge if op == "IMPO" else b.port_of_loading,
            "Cliente / Consignee": b.entity,
            "Carga": format_cargo_output(b.cargo, b.entity, b.description_lines) or "—",
            "Peso (kg)": round(b.weight, 1),
            "Basic FRT": round(b.charges.get("Basic FRT").monto, 2) if "Basic FRT" in b.charges else 0.0,
            "THC": round(b.thc_20.monto + b.thc_40.monto, 2),
            "VINs": ", ".join(b.vins) if b.vins else "—",
        })
    return pd.DataFrame(rows)


def run_engine(input_dir: Path):
    out_dir = input_dir / "out"
    out_dir.mkdir(exist_ok=True)
    t0 = time.perf_counter()
    process_directory(str(input_dir), str(out_dir), config.EngineConfig(roe=roe))
    return out_dir, time.perf_counter() - t0


def zip_reports(out_dir: Path) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for f in sorted(out_dir.glob("*.xlsx")):
            zf.write(f, f.name)
    return buf.getvalue()


def process(files_payload):
    valid = [(n, b) for (n, b) in files_payload if NAME_RE.match(n)]
    if not valid:
        st.error("Ningún archivo coincide con `ATLNNNN_IMPO_1.CSV` / `_EXPO_1.CSV`.")
        return

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        for name, data in valid:
            (tmp_dir / name).write_bytes(data)

        # Parseo para visualización
        dfs = []
        for name, _ in valid:
            op = "IMPO" if "IMPO" in name.upper() else "EXPO"
            dfs.append((op, bls_to_df(parse_manifest(str(tmp_dir / name), op), op)))

        out_dir, elapsed = run_engine(tmp_dir)
        reports = sorted(out_dir.glob("*.xlsx"))

        impo_df = pd.concat([d for o, d in dfs if o == "IMPO"], ignore_index=True) if any(o == "IMPO" for o, _ in dfs) else pd.DataFrame()
        expo_df = pd.concat([d for o, d in dfs if o == "EXPO"], ignore_index=True) if any(o == "EXPO" for o, _ in dfs) else pd.DataFrame()
        total_bls = len(impo_df) + len(expo_df)

        # ── Métricas ──
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Bills of Lading", total_bls)
        c2.metric("Reportes generados", len(reports))
        c3.metric("Manifiestos", len(valid))
        c4.metric("Tiempo", f"{elapsed*1000:.0f} ms")

        tab_resumen, tab_impo, tab_expo, tab_dl = st.tabs(
            ["📊 Resumen", "📥 IMPO", "📤 EXPO", "⬇️ Descargas"]
        )

        with tab_resumen:
            all_df = pd.concat([impo_df, expo_df], ignore_index=True)
            if not all_df.empty:
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown("**B/Ls por tipo de carga**")
                    st.bar_chart(all_df.groupby("Tipo")["B/L"].count().rename("B/Ls"), color="#1e3a5f")
                with col_b:
                    st.markdown("**B/Ls por operación**")
                    st.bar_chart(all_df.groupby("Operación")["B/L"].count().rename("B/Ls"), color="#3d6fa5")
                st.caption(f"ROE aplicado: EUR 1,00 = {roe} USD")

        with tab_impo:
            st.dataframe(impo_df, use_container_width=True, hide_index=True) if not impo_df.empty else st.info("Sin manifiesto IMPO.")
        with tab_expo:
            st.dataframe(expo_df, use_container_width=True, hide_index=True) if not expo_df.empty else st.info("Sin manifiesto EXPO.")

        with tab_dl:
            st.success(f"✅ {len(reports)} reportes listos.")
            cols = st.columns(2)
            for i, r in enumerate(reports):
                cols[i % 2].download_button(f"⬇️ {r.name}", r.read_bytes(), file_name=r.name,
                                            mime=XLSX_MIME, use_container_width=True)
            st.download_button("⬇️ Descargar todo (.zip)", zip_reports(out_dir),
                               file_name="reportes.zip", mime="application/zip",
                               type="primary", use_container_width=True)


if use_samples:
    process([(f.name, f.read_bytes()) for f in SAMPLES.glob("*.CSV")])
elif uploaded:
    process([(f.name, f.getvalue()) for f in uploaded])
else:
    st.markdown(
        "👈 Usá **manifiestos de ejemplo** desde la barra lateral, o subí tus propios CSV.\n\n"
        "**Qué genera por viaje:** planilla IMPO/EXPO · comisiones IMPO/EXPO · ENS · "
        "Por B/L · Por Marca · Chasis — cada una con una hoja **REVISIÓN** de controles."
    )
