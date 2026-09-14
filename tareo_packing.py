import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="tareito_de_packingcito", page_icon="⚙️")
st.title("⚙️ TAREO PACKING")
st.caption("Sube tu Excel, aplica las transformaciones y descarga el archivo limpio.")

archivo = st.file_uploader("📂 Sube tu archivo Excel", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo)

    st.subheader("Vista previa — original")
    st.dataframe(df.head(10), use_container_width=True)

    # ── Transformaciones ─────────────────────────────────────────────────────

    # 1. Convertir INICIO y FIN a datetime
    df["INICIO"] = pd.to_datetime(df["INICIO"], dayfirst=True, errors="coerce")
    df["FIN"]    = pd.to_datetime(df["FIN"],    dayfirst=True, errors="coerce")

    # 2. Columna HORA = (FIN - INICIO) * 24  →  horas decimales
    df["HORA"] = (df["FIN"] - df["INICIO"]).dt.total_seconds() / 3600
    df["HORA"] = df["HORA"].round(4)

    # 3. Reemplazar TURNO: DIURNO → DIA, NOCTURNO → NOCHE
    df["TURNO"] = df["TURNO"].str.strip().str.upper().replace({
        "DIURNO":   "DIA",
        "NOCTURNO": "NOCHE"
    })

    # ── Vista previa resultado ───────────────────────────────────────────────
    st.subheader("Vista previa — transformado")
    st.dataframe(df.head(10), use_container_width=True)

    # ── Estadísticas rápidas ─────────────────────────────────────────────────
    col1, col2, col3 = st.columns(3)
    col1.metric("Total filas", len(df))
    col2.metric("Filas DIA",   len(df[df["TURNO"] == "DIA"]))
    col3.metric("Filas NOCHE", len(df[df["TURNO"] == "NOCHE"]))

    # ── Descarga ─────────────────────────────────────────────────────────────
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Transformado")
    buffer.seek(0)

    st.download_button(
        label="⬇️ Descargar Excel limpio",
        data=buffer,
        file_name="turnos_transformado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )