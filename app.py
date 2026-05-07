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
        color: #999;
    }
    .precio-descuento {
        font-size: 36px;
        font-weight: 700;
        color: #FF416C;
    }
    .precios-container {
        background: #f8f9fa;
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
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
        
        # Convertir columnas de código a string
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
    Detecta y calcula el porcentaje de descuento correctamente
    """
    # Lista de posibles nombres de columnas
    columnas_precio_original = [
        'Precio Original', 'Precio Normal', 'Precio Lista', 'Precio Base',
        'PRECIO ORIGINAL', 'PRECIO NORMAL', 'PRECIO LISTA', 'PRECIO BASE',
        'precio_original', 'precio_normal', 'precio_lista', 'precio_base',
        'Precio', 'PRECIO', 'precio'
    ]
    
    columnas_precio_descuento = [
        'Precio Descuento', 'Precio Oferta', 'Precio Especial', 'Precio Madrugón',
        'PRECIO DESCUENTO', 'PRECIO OFERTA', 'PRECIO ESPECIAL', 'PRECIO MADRUGON',
        'precio_descuento', 'precio_oferta', 'precio_especial', 'precio_madrugon',
        'Precio Final', 'PRECIO FINAL', 'precio_final'
    ]
    
    columnas_porcentaje = [
        'Descuento', '% Descuento', 'Porcentaje', 'Descuento %',
        'DESCUENTO', '% DESCUENTO', 'PORCENTAJE', 'DESCUENTO %',
        'descuento', 'porcentaje', 'Dscto', 'DSCTO', 'dscto'
    ]
    
    # 1. Buscar columna de porcentaje de descuento directo
    for col in columnas_porcentaje:
        if col in fila.index and pd.notna(fila[col]):
            try:
                valor = float(fila[col])
                if valor == 0:
                    continue
                
                # DETECCIÓN INTELIGENTE DEL FORMATO:
                # Si el valor es menor o igual a 1 (como 0.3), es formato decimal
                if valor <= 1:
                    porcentaje = valor * 100
                # Si el valor es mayor a 1 (como 30), ya está en porcentaje
                else:
                    porcentaje = valor
                
                return round(porcentaje, 1), None, None
            except:
                continue
    
    # 2. Calcular desde precio original y precio descuento
    precio_original = None
    precio_descuento = None
    
    for col in columnas_precio_original:
        if col in fila.index and pd.notna(fila[col]):
            try:
                precio_original = float(fila[col])
                if precio_original > 0:
                    break
            except:
                continue
    
    for col in columnas_precio_descuento:
        if col in fila.index and pd.notna(fila[col]):
            try:
                precio_descuento = float(fila[col])
                if precio_descuento > 0:
                    break
            except:
                continue
    
    if precio_original and precio_descuento and precio_original > 0:
        if precio_descuento < precio_original:
            descuento = ((precio_original - precio_descuento) / precio_original) * 100
            return round(descuento, 1), precio_original, precio_descuento
    
    return None, None, None

def mostrar_producto(fila):
    """Muestra la información del producto con el descuento destacado"""
    
    # Calcular descuento
    descuento, precio_original, precio_descuento = calcular_descuento(fila)
    
    # Mostrar descuento de manera destacada
    if descuento is not None and descuento > 0:
        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">🔥 ¡AHORRA!</div>
            <div class="porcentaje-descuento">{descuento:.0f}%</div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)
        
        # Mostrar precios si están disponibles
        if precio_original and precio_descuento:
            st.markdown('<div class="precios-container">', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown(f'<p class="precio-original">Antes: ${precio_original:,.0f}</p>', 
                          unsafe_allow_html=True)
            with col2:
                st.markdown(f'<p class="precio-descuento">Ahora: ${precio_descuento:,.0f}</p>', 
                          unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
    
    # Mostrar resto de información del producto
    st.write("### 📋 Detalles del Producto")
    
    # Filtrar columnas para mostrar
    columnas_excluir = ['EAN PADRE ', 'Código', 'Descuento', '% Descuento', 'Porcentaje', 
                       'DESCUENTO', '% DESCUENTO', 'PORCENTAJE']
    
    columnas_mostrar = [col for col in fila.index 
                       if col not in columnas_excluir 
                       and pd.notna(fila[col])
                       and str(fila[col]).strip() != '']
    
    if columnas_mostrar:
        # Mostrar en formato de tabla
        datos_mostrar = {}
        for columna in columnas_mostrar:
            valor = fila[columna]
            # Formatear números
            try:
                valor_float = float(valor)
                if valor_float == int(valor_float):
                    datos_mostrar[columna.strip()] = int(valor_float)
                else:
                    datos_mostrar[columna.strip()] = f"${valor_float:,.2f}"
            except:
                datos_mostrar[columna.strip()] = valor
        
        # Mostrar como métricas en columnas
        cols = st.columns(min(3, len(datos_mostrar)))
        for i, (key, value) in enumerate(datos_mostrar.items()):
            with cols[i % 3]:
                st.metric(label=key, value=value)

# Interfaz principal
st.title("🛍️ Consulta de Descuentos")
st.caption("Madrugón Mayo 2026")

# Seleccionar método de carga
metodo_carga = st.radio(
    "Elige cómo cargar el archivo:",
    ["📤 Subir archivo", "📁 Usar archivo local"],
    horizontal=True
)

df = None

if metodo_carga == "📤 Subir archivo":
    archivo_subido = st.file_uploader(
        "Sube el archivo Excel",
        type=["xlsx", "xls"],
        help="Selecciona el archivo 'MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx'"
    )
    
    if archivo_subido is not None:
        with st.spinner("Cargando datos..."):
            df = cargar_datos(archivo_subido)
else:
    # Intentar cargar archivo local
    archivo_local = "MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"
    if os.path.exists(archivo_local):
        with st.spinner("Cargando datos..."):
            df = cargar_datos(archivo_local)
    else:
        st.error(f"❌ No se encuentra el archivo '{archivo_local}'")

if df is not None and not df.empty:
    st.success(f"✅ {len(df)} productos cargados correctamente")
    
    # Input de código de barras
    st.write("### 🔍 Buscar Producto")
    
    # Campo de búsqueda con auto-focus
    codigo = st.text_input(
        "Código de barras",
        placeholder="Escanea o escribe el código aquí",
        key="codigo_input",
        label_visibility="visible"
    )
    
    if codigo:
        codigo = codigo.strip()
        
        # Buscar en diferentes columnas
        mascara = pd.Series(False, index=df.index)
        
        if "EAN PADRE " in df.columns:
            mascara |= (df["EAN PADRE "] == codigo)
        
        if "Código" in df.columns:
            mascara |= (df["Código"] == codigo)
        
        # También buscar en otras posibles columnas de código
        for col in df.columns:
            if 'ean' in col.lower() or 'codigo' in col.lower() or 'código' in col.lower():
                if col not in ["EAN PADRE ", "Código"]:
                    try:
                        mascara |= (df[col].astype(str) == codigo)
                    except:
                        pass
        
        resultado = df[mascara]
        
        if not resultado.empty:
            fila = resultado.iloc[0]
            st.success("✨ Producto encontrado")
            mostrar_producto(fila)
        else:
            st.error("❌ Código no encontrado")
            st.info("💡 Verifica que el código sea correcto o intenta con otro código")
            
            # Sugerencia: mostrar algunos códigos de ejemplo
            with st.expander("Ver códigos de ejemplo"):
                if "EAN PADRE " in df.columns:
                    ejemplos = df["EAN PADRE "].dropna().head(5).tolist()
                elif "Código" in df.columns:
                    ejemplos = df["Código"].dropna().head(5).tolist()
                else:
                    ejemplos = []
                
                if ejemplos:
                    st.write("Algunos códigos del archivo:")
                    for ej in ejemplos:
                        st.code(ej)
else:
    st.info("ℹ️ Carga o selecciona un archivo Excel para comenzar")
    
    # Mostrar instrucciones
    with st.expander("📖 Instrucciones"):
        st.write("""
        1. Sube el archivo Excel con los datos del Madrugón
        2. Escanea o escribe un código de barras
        3. Visualiza el descuento y detalles del producto
        
        **Columnas esperadas:**
        - EAN PADRE o Código (para buscar)
        - Descuento o % Descuento (porcentaje de descuento)
        - O Precio Original y Precio Descuento (para calcular)
        """)