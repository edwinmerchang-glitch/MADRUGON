import streamlit as st
import pandas as pd

# CONFIGURACIÓN
ARCHIVO_EXCEL = "MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"
HOJA = "8"

# CARGAR EXCEL
@st.cache_data
def cargar_datos():
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA)

    # Limpiar columnas
    df.columns = df.columns.str.strip()

    # Convertir código a texto
    df['Código'] = df['Código'].astype(str)

    return df

# Cargar información
try:
    df = cargar_datos()
except Exception as e:
    st.error(f"Error cargando el archivo: {e}")
    st.stop()

# INTERFAZ
st.set_page_config(page_title="Scanner de Descuentos", layout="centered")

st.title("🛒 Scanner de Descuentos")
st.write("Escanea o escribe el código de barras")

# INPUT CÓDIGO
codigo = st.text_input("Código de barras")

# BÚSQUEDA
if codigo:

    resultado = df[df['Código'].str.strip() == codigo.strip()]

    if not resultado.empty:

        producto = resultado.iloc[0]

        marca = producto.get('Marca', 'No disponible')

        descuento = producto.get('% Descuento', producto.get('Porcentaje Descuento', 0))

        precio_final = producto.get('Precio de venta con descuento', 0)

        precio_original = producto.get('Precio de venta', producto.get('Precio de venta ', 0))

        st.success("Producto encontrado")

        col1, col2 = st.columns(2)

        with col1:
            st.metric("Marca", marca)
            st.metric("Descuento", f"{descuento}%")
        st.error("Código no encontrado")