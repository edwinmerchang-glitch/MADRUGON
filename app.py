import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="centered"
)

# CSS personalizado para resaltar el descuento
st.markdown("""
<style>
    .descuento-container {
        background: linear-gradient(135deg, #FF416C, #FF4B2B);
        color: white;
        padding: 30px;
        border-radius: 20px;
        text-align: center;
        margin: 20px 0;
        box-shadow: 0 10px 30px rgba(255, 65, 108, 0.3);
    }
    .porcentaje-descuento {
        font-size: 80px;
        font-weight: 900;
        line-height: 1;
        margin: 10px 0;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .texto-descuento {
        font-size: 24px;
        font-weight: 600;
    }
    .precio-original {
        font-size: 20px;
        text-decoration: line-through;
        opacity: 0.8;
    }
    .precio-descuento {
        font-size: 36px;
        font-weight: 700;
    }
</style>
""", unsafe_allow_html=True)

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

def calcular_descuento(fila):
    """
    Intenta encontrar y calcular el porcentaje de descuento
    Prioriza columnas comunes de descuento en diferentes órdenes
    """
    # Lista de posibles nombres de columnas para precio original y descuento
    columnas_precio_original = [
        'Precio Original', 'Precio Normal', 'Precio Lista', 'Precio Base',
        'PRECIO ORIGINAL', 'PRECIO NORMAL', 'PRECIO LISTA', 'PRECIO BASE',
        'precio_original', 'precio_normal', 'precio_lista', 'precio_base'
    ]
    
    columnas_precio_descuento = [
        'Precio Descuento', 'Precio Oferta', 'Precio Especial', 'Precio Madrugón',
        'PRECIO DESCUENTO', 'PRECIO OFERTA', 'PRECIO ESPECIAL', 'PRECIO MADRUGON',
        'precio_descuento', 'precio_oferta', 'precio_especial', 'precio_madrugon'
    ]
    
    columnas_porcentaje = [
        'Descuento', '% Descuento', 'Porcentaje', 'Descuento %',
        'DESCUENTO', '% DESCUENTO', 'PORCENTAJE', 'DESCUENTO %',
        'descuento', 'porcentaje'
    ]
    
    # 1. Buscar columna de porcentaje de descuento directo
    for col in columnas_porcentaje:
        if col in fila.index and pd.notna(fila[col]):
            try:
                valor = float(fila[col])
                if 0 < valor <= 100:
                    return valor, None, None
                elif valor > 1:  # Asume que si es mayor a 1 y menor a 100, es porcentaje
                    return valor, None, None
            except:
                continue
    
    # 2. Calcular desde precio original y precio descuento
    precio_original = None
    precio_descuento = None
    
    for col in columnas_precio_original:
        if col in fila.index and pd.notna(fila[col]):
            try:
                precio_original = float(fila[col])
                break
            except:
                continue
    
    for col in columnas_precio_descuento:
        if col in fila.index and pd.notna(fila[col]):
            try:
                precio_descuento = float(fila[col])
                break
            except:
                continue
    
    if precio_original and precio_descuento and precio_original > 0:
        descuento = ((precio_original - precio_descuento) / precio_original) * 100
        return round(descuento, 1), precio_original, precio_descuento
    
    return None, None, None

def mostrar_producto(fila):
    """Muestra la información del producto con el descuento destacado"""
    
    # Calcular descuento
    descuento, precio_original, precio_descuento = calcular_descuento(fila)
    
    # Mostrar descuento de manera destacada
    if descuento is not None:
        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">¡AHORRA!</div>
            <div class="porcentaje-descuento">{descuento}%</div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Mostrar precios si están disponibles
        if precio_original and precio_descuento:
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f'<p class="precio-original">Antes: ${precio_original:,.0f}</p>', 
                          unsafe_allow_html=True)
            with col2:
                st.markdown(f'<p class="precio-descuento">Ahora: ${precio_descuento:,.0f}</p>', 
                          unsafe_allow_html=True)
    
    # Mostrar resto de información del producto
    st.write("### 📋 Detalles del Producto")
    
    # Columnas a mostrar (excluyendo las de código)
    columnas_mostrar = [col for col in fila.index 
                       if col not in ['EAN PADRE ', 'Código'] 
                       and pd.notna(fila[col])]
    
    # Crear dos columnas para mostrar la información
    mitad = len(columnas_mostrar) // 2
    col1, col2 = st.columns(2)
    
    for i, columna in enumerate(columnas_mostrar):
        if i < mitad:
            with col1:
                st.metric(label=columna.strip(), value=fila[columna])
        else:
            with col2:
                st.metric(label=columna.strip(), value=fila[columna])

# Interfaz principal
st.title("🛍️ Consulta de Descuentos - Madrugón Mayo 2026")

# Seleccionar método de carga
metodo_carga = st.radio(
    "Elige cómo cargar el archivo:",
    ["Subir archivo", "Usar archivo local"],
    horizontal=True
)

df = None

if metodo_carga == "Subir archivo":
    archivo_subido = st.file_uploader(
        "Sube el archivo Excel",
        type=["xlsx", "xls"],
        help="Selecciona el archivo 'MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx'"
    )
    
    if archivo_subido is not None:
        df = cargar_datos(archivo_subido)
else:
    # Intentar cargar archivo local
    archivo_local = "MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"
    if os.path.exists(archivo_local):
        df = cargar_datos(archivo_local)
    else:
        st.error(f"No se encuentra el archivo '{archivo_local}' en el directorio")

if df is not None and not df.empty:
    st.success("✅ Datos cargados correctamente")
    
    # Input de código de barras
    st.write("### 🔍 Buscar Producto")
    codigo = st.text_input(
        "Código de barras",
        placeholder="Escanea o escribe el código aquí",
        key="codigo_input"
    )
    
    if codigo:
        codigo = codigo.strip()
        
        # Buscar en diferentes columnas
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
            st.error("No se encontraron columnas de código en el archivo")
            st.stop()
        
        if not resultado.empty:
            fila = resultado.iloc[0]
            st.success("✨ Producto encontrado")
            mostrar_producto(fila)
        else:
            st.error("❌ Código no encontrado")
            st.info("Verifica que el código sea correcto o intenta con otro código")
else:
    st.info("ℹ️ Carga o selecciona un archivo Excel para comenzar")