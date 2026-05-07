import streamlit as st
import pandas as pd

st.set_page_config(page_title="Scanner de Descuentos", layout="centered")

st.title("🛒 Scanner de Descuentos")
st.write("Escanea o escribe el código de barras")

# Subir archivo Excel
archivo_subido = st.file_uploader(
    "📂 Cargar archivo Excel", 
    type=['xlsx', 'xls'],
    help="Sube el archivo MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"
)

if archivo_subido is not None:
    try:
        HOJA = "8"
        
        # Cargar el archivo subido
        df = pd.read_excel(archivo_subido, sheet_name=HOJA)
        
        # Limpiar columnas
        df.columns = df.columns.str.strip()
        
        # Convertir código a texto
        df['Código'] = df['Código'].astype(str)
        
        st.success("✅ Archivo cargado correctamente")
        
        # INPUT CÓDIGO
        codigo = st.text_input("Código de barras", placeholder="Ej: 1234567890")
        
        # BÚSQUEDA
        if codigo:
            resultado = df[df['Código'].str.strip() == codigo.strip()]
            
            if not resultado.empty:
                producto = resultado.iloc[0]
                
                marca = producto.get('Marca', 'No disponible')
                descuento = producto.get('% Descuento', producto.get('Porcentaje Descuento', 0))
                precio_final = producto.get('Precio de venta con descuento', 0)
                precio_original = producto.get('Precio de venta', producto.get('Precio de venta ', 0))
                
                st.success("✅ Producto encontrado")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.metric("🏷️ Marca", str(marca))
                    st.metric("🔻 Descuento", f"{descuento}%")
                
                with col2:
                    if pd.notna(precio_original) and precio_original != 0:
                        st.metric("💰 Precio Original", f"${precio_original:,.0f}")
                    else:
                        st.metric("💰 Precio Original", "No disponible")
                    
                    if pd.notna(precio_final) and precio_final != 0:
                        st.metric("💵 Precio Final", f"${precio_final:,.0f}")
                    else:
                        st.metric("💵 Precio Final", "No disponible")
            else:
                st.error("❌ Código no encontrado")
                
    except Exception as e:
        st.error(f"Error cargando el archivo: {e}")
        st.info("Asegúrate de que la hoja se llame '8' y tenga las columnas correctas")
else:
    st.info("👈 Por favor, carga el archivo Excel en el botón de arriba")