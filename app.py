import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="centered"
)

@st.cache_data
def cargar_datos():
    archivo = "MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"

    df = pd.read_excel(
        archivo,
        sheet_name="8"
    )

    columnas_codigo = [
        "EAN PADRE ",
        "Código"
    ]

    for col in columnas_codigo:
        if col in df.columns:
            df[col] = df[col].astype(str)

    return df

try:
    df = cargar_datos()
except Exception as e:
    st.error(f"Error cargando archivo: {e}")
    st.stop()

st.title("🛍️ Consulta de Descuentos")
st.write("Escanea o escribe el código de barras")

codigo = st.text_input(
    "Código de barras",
    placeholder="Escanea aquí"
)

if codigo:

    codigo = codigo.strip()

    resultado = df[
        (df["EAN PADRE "] == codigo)
        |
        (df["Código"] == codigo)
    ]

    if not resultado.empty:

        fila = resultado.iloc[0]

        st.success("Producto encontrado")

        col1, col2 = st.columns(2)
        st.error("Código no encontrado")