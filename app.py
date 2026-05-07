import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="centered"
)

@st.cache_data
def cargar_datos(archivo):
    try:
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
    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")
        return None

st.title("🛍️ Consulta de Descuentos")

# Cargar archivo Excel
archivo_subido = st.file_uploader(
    "Sube el archivo Excel de descuentos",
    type=["xlsx", "xls"],
    help="Selecciona el archivo 'MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx'"
)

if archivo_subido is not None:
    df = cargar_datos(archivo_subido)
    
    if df is not None and not df.empty:
        st.success("Archivo cargado correctamente")
        
        st.write("Escanea o escribe el código de barras")
        
        codigo = st.text_input(
            "Código de barras",
            placeholder="Escanea aquí",
            key="codigo_input"  # Asegura que el input se actualice correctamente
        )
        
        if codigo:
            codigo = codigo.strip()
            
            # Verificar si las columnas existen
            columna_ean = "EAN PADRE " if "EAN PADRE " in df.columns else None
            columna_codigo = "Código" if "Código" in df.columns else None
            
            if columna_ean and columna_codigo:
                resultado = df[
                    (df[columna_ean] == codigo) |
                    (df[columna_codigo] == codigo)
                ]
            elif columna_ean:
                resultado = df[df[columna_ean] == codigo]
            elif columna_codigo:
                resultado = df[df[columna_codigo] == codigo]
            else:
                st.error("No se encontraron las columnas necesarias")
                st.stop()
            
            if not resultado.empty:
                fila = resultado.iloc[0]
                st.success("Producto encontrado")
                
                # Mostrar información del producto
                st.write("### Información del Producto")
                for columna in resultado.columns:
                    valor = fila[columna]
                    st.write(f"**{columna}:** {valor}")
            else:
                st.error("Código no encontrado")
    else:
        st.warning("No se pudieron cargar los datos del archivo")
else:
    st.info("Por favor, sube el archivo Excel para comenzar")