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
    .producto-card {
        background: white;
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin: 20px 0;
    }
    .producto-nombre {
        font-size: 28px;
        font-weight: 700;
        color: #1a1a1a;
        margin-bottom: 20px;
        text-align: center;
    }
    .producto-marca {
        font-size: 16px;
        color: #666;
        text-align: center;
        margin-bottom: 20px;
        background: #f0f0f0;
        padding: 5px 15px;
        border-radius: 20px;
        display: inline-block;
    }
    .precio-container {
        display: flex;
        justify-content: center;
        gap: 20px;
        margin-top: 20px;
    }
    .precio-original {
        text-decoration: line-through;
        color: #999;
        font-size: 20px;
        text-align: center;
    }
    .precio-descuento {
        color: #FF416C;
        font-size: 32px;
        font-weight: 700;
        text-align: center;
    }
    .precio-label {
        font-size: 14px;
        color: #666;
        margin-bottom: 5px;
    }
    .stats-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 15px;
        margin: 20px 0;
    }
    .upload-area {
        border: 3px dashed #FF416C;
        border-radius: 20px;
        padding: 50px;
        text-align: center;
        background: #fafafa;
    }
</style>
""", unsafe_allow_html=True)

# Inicializar estado de la sesión
if 'datos_cargados' not in st.session_state:
    st.session_state.datos_cargados = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'nombre_archivo' not in st.session_state:
    st.session_state.nombre_archivo = ""

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

def obtener_info_producto(fila):
    """
    Extrae la información específica del producto:
    - Nombre
    - Precio de venta
    - Precio de venta con descuento
    - Marca
    - Porcentaje de descuento
    """
    info = {
        'nombre': None,
        'precio_venta': None,
        'precio_descuento': None,
        'marca': None,
        'porcentaje_descuento': None
    }
    
    # Mapeo de posibles nombres de columnas (del español al estándar)
    mapeo_columnas = {
        'nombre': ['name', 'nombre', 'producto', 'product', 'descripción', 'descripcion',
                   'Name', 'Nombre', 'Producto', 'Product', 'Descripción', 'Descripcion',
                   'NAME', 'NOMBRE', 'PRODUCTO', 'DESCRIPCIÓN', 'DESCRIPCION'],
        'marca': ['marca', 'brand', 'Marca', 'Brand', 'MARCA', 'BRAND'],
        'precio_venta': ['Precio de venta', 'precio de venta', 'PRECIO DE VENTA',
                        'Precio', 'precio', 'PRECIO', 'Precio Original', 'precio original',
                        'Precio Normal', 'precio normal', 'Precio Lista', 'precio lista'],
        'precio_descuento': ['Precio de venta con descuento', 'precio de venta con descuento',
                           'PRECIO DE VENTA CON DESCUENTO', 'Precio Descuento', 'precio descuento',
                           'PRECIO DESCUENTO', 'Precio Oferta', 'precio oferta',
                           'Precio Final', 'precio final', 'PRECIO FINAL'],
        'porcentaje': ['PORCENTAJE DESCUENTO', 'porcentaje descuento', 'Porcentaje Descuento',
                      'Descuento', 'descuento', 'DESCUENTO', '% Descuento', '% descuento',
                      'Dscto', 'dscto', 'DSCTO']
    }
    
    # Buscar nombre
    for col in mapeo_columnas['nombre']:
        if col in fila.index and pd.notna(fila[col]):
            info['nombre'] = str(fila[col]).strip()
            break
    
    # Buscar marca
    for col in mapeo_columnas['marca']:
        if col in fila.index and pd.notna(fila[col]):
            info['marca'] = str(fila[col]).strip()
            break
    
    # Buscar precio de venta original
    for col in mapeo_columnas['precio_venta']:
        if col in fila.index and pd.notna(fila[col]):
            try:
                info['precio_venta'] = float(fila[col])
                break
            except:
                continue
    
    # Buscar precio con descuento
    for col in mapeo_columnas['precio_descuento']:
        if col in fila.index and pd.notna(fila[col]):
            try:
                info['precio_descuento'] = float(fila[col])
                break
            except:
                continue
    
    # Buscar porcentaje de descuento
    for col in mapeo_columnas['porcentaje']:
        if col in fila.index and pd.notna(fila[col]):
            try:
                valor = float(fila[col])
                if valor > 0:
                    # Detectar formato decimal vs porcentaje
                    if valor <= 1:
                        info['porcentaje_descuento'] = valor * 100
                    else:
                        info['porcentaje_descuento'] = valor
                    break
            except:
                continue
    
    # Si no hay porcentaje directo, calcularlo de los precios
    if info['porcentaje_descuento'] is None and info['precio_venta'] and info['precio_descuento']:
        if info['precio_venta'] > 0 and info['precio_descuento'] < info['precio_venta']:
            info['porcentaje_descuento'] = ((info['precio_venta'] - info['precio_descuento']) / 
                                           info['precio_venta']) * 100
    
    return info

def mostrar_producto(info):
    """Muestra la información del producto de forma atractiva"""
    
    # Contenedor principal del descuento
    if info['porcentaje_descuento'] is not None and info['porcentaje_descuento'] > 0:
        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">🔥 ¡AHORRA!</div>
            <div class="porcentaje-descuento">{info['porcentaje_descuento']:.0f}%</div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Tarjeta del producto
    st.markdown('<div class="producto-card">', unsafe_allow_html=True)
    
    # Nombre del producto
    if info['nombre']:
        st.markdown(f'<div class="producto-nombre">{info["nombre"]}</div>', unsafe_allow_html=True)
    
    # Marca
    if info['marca']:
        st.markdown(f'<div style="text-align: center;"><span class="producto-marca">🏷️ {info["marca"]}</span></div>', 
                   unsafe_allow_html=True)
    
    # Precios
    if info['precio_venta'] is not None or info['precio_descuento'] is not None:
        st.markdown('<div class="precio-container">', unsafe_allow_html=True)
        
        # Crear columnas para los precios
        col1, col2 = st.columns(2)
        
        with col1:
            if info['precio_venta'] is not None:
                st.markdown(f"""
                <div style="text-align: center;">
                    <div class="precio-label">Precio Original</div>
                    <div class="precio-original">${info['precio_venta']:,.0f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        with col2:
            if info['precio_descuento'] is not None:
                st.markdown(f"""
                <div style="text-align: center;">
                    <div class="precio-label">Precio con Descuento</div>
                    <div class="precio-descuento">${info['precio_descuento']:,.0f}</div>
                </div>
                """, unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    # Ahorro total
    if info['precio_venta'] and info['precio_descuento']:
        ahorro = info['precio_venta'] - info['precio_descuento']
        if ahorro > 0:
            st.markdown(f"""
            <div style="text-align: center; margin-top: 15px; padding: 10px; background: #e8f5e9; border-radius: 10px;">
                <span style="color: #2e7d32; font-weight: 600;">💰 Te ahorras: ${ahorro:,.0f}</span>
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# ============ INTERFAZ PRINCIPAL CON PESTAÑAS ============

tab1, tab2 = st.tabs(["📤 Cargar Archivo", "🔍 Consultar Productos"])

# Pestaña 1: Cargar Archivo
with tab1:
    st.title("📤 Cargar Archivo de Datos")
    st.markdown("---")
    
    # Método de carga
    metodo_carga = st.radio(
        "Selecciona el método de carga:",
        ["📁 Subir archivo desde mi computadora", "💻 Usar archivo del servidor"],
        horizontal=False
    )
    
    if metodo_carga == "📁 Subir archivo desde mi computadora":
        st.markdown('<div class="upload-area">', unsafe_allow_html=True)
        
        archivo_subido = st.file_uploader(
            "Arrastra o selecciona el archivo Excel",
            type=["xlsx", "xls"],
            help="Formatos aceptados: .xlsx, .xls"
        )
        
        st.markdown('</div>', unsafe_allow_html=True)
        
        if archivo_subido is not None:
            with st.spinner("🔄 Procesando archivo..."):
                df = cargar_datos(archivo_subido)
                
                if df is not None and not df.empty:
                    st.session_state.df = df
                    st.session_state.datos_cargados = True
                    st.session_state.nombre_archivo = archivo_subido.name
                    
                    # Mostrar estadísticas
                    st.markdown(f"""
                    <div class="stats-container">
                        <h2>✅ ¡Archivo cargado exitosamente!</h2>
                        <h3>📊 {len(df):,} productos encontrados</h3>
                        <p>📄 Archivo: {archivo_subido.name}</p>
                        <p>📏 Tamaño: {archivo_subido.size/1024/1024:.1f} MB</p>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.success("👉 Ve a la pestaña **'Consultar Productos'** para comenzar a buscar")
    
    else:  # Archivo del servidor
        st.info("💡 Esta opción busca el archivo en el servidor donde está alojada la app")
        
        nombre_archivo = st.text_input(
            "Nombre del archivo en el servidor:",
            value="MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx"
        )
        
        if st.button("🔍 Buscar y cargar archivo", type="primary"):
            if os.path.exists(nombre_archivo):
                with st.spinner("🔄 Cargando archivo del servidor..."):
                    df = cargar_datos(nombre_archivo)
                    
                    if df is not None and not df.empty:
                        st.session_state.df = df
                        st.session_state.datos_cargados = True
                        st.session_state.nombre_archivo = nombre_archivo
                        
                        st.markdown(f"""
                        <div class="stats-container">
                            <h2>✅ ¡Archivo cargado exitosamente!</h2>
                            <h3>📊 {len(df):,} productos encontrados</h3>
                            <p>📄 Archivo: {nombre_archivo}</p>
                        </div>
                        """, unsafe_allow_html=True)
                        
                        st.success("👉 Ve a la pestaña **'Consultar Productos'** para comenzar a buscar")
            else:
                st.error(f"❌ No se encontró el archivo '{nombre_archivo}' en el servidor")

# Pestaña 2: Consultar Productos
with tab2:
    st.title("🔍 Consultar Productos")
    st.markdown("---")
    
    if not st.session_state.datos_cargados:
        st.warning("⚠️ Primero debes cargar un archivo en la pestaña 'Cargar Archivo'")
        st.info("""
        📋 **Instrucciones:**
        1. Ve a la pestaña **'Cargar Archivo'**
        2. Sube tu archivo Excel con los datos del Madrugón
        3. Vuelve a esta pestaña para buscar productos
        """)
    else:
        # Mostrar info del archivo cargado
        st.markdown(f"""
        <div style="background: #e8f5e9; padding: 15px; border-radius: 10px; margin-bottom: 20px;">
            <p style="margin: 0; color: #2e7d32;">
                ✅ <strong>{st.session_state.nombre_archivo}</strong> - 
                {len(st.session_state.df):,} productos disponibles
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Campo de búsqueda
        st.write("### 🎯 Buscar Producto")
        
        codigo = st.text_input(
            "Código de barras",
            placeholder="Escanea o escribe el código aquí",
            key="codigo_input"
        )
        
        if codigo:
            codigo = codigo.strip()
            df = st.session_state.df
            
            # Búsqueda en múltiples columnas
            mascara = pd.Series(False, index=df.index)
            
            if "EAN PADRE " in df.columns:
                mascara |= (df["EAN PADRE "] == codigo)
            
            if "Código" in df.columns:
                mascara |= (df["Código"] == codigo)
            
            # Buscar en otras columnas que contengan 'ean' o 'codigo'
            for col in df.columns:
                col_lower = col.lower()
                if ('ean' in col_lower or 'codigo' in col_lower or 'código' in col_lower):
                    if col not in ["EAN PADRE ", "Código"]:
                        try:
                            mascara |= (df[col].astype(str).str.strip() == codigo)
                        except:
                            pass
            
            resultado = df[mascara]
            
            if not resultado.empty:
                fila = resultado.iloc[0]
                st.success("✨ ¡Producto encontrado!")
                
                # Extraer y mostrar solo la información necesaria
                info_producto = obtener_info_producto(fila)
                mostrar_producto(info_producto)
                
            else:
                st.error("❌ Código no encontrado")
                
                # Sugerencias de códigos
                with st.expander("💡 ¿Necesitas ayuda?"):
                    st.write("**Códigos de ejemplo del archivo:**")
                    if "EAN PADRE " in df.columns:
                        ejemplos = df["EAN PADRE "].dropna().sample(min(5, len(df))).tolist()
                    elif "Código" in df.columns:
                        ejemplos = df["Código"].dropna().sample(min(5, len(df))).tolist()
                    else:
                        ejemplos = []
                    
                    for ej in ejemplos:
                        st.code(ej)
        
        # Botón para cambiar de archivo
        if st.button("🔄 Cargar otro archivo", use_container_width=True):
            st.session_state.datos_cargados = False
            st.session_state.df = None
            st.session_state.nombre_archivo = ""
            st.rerun()