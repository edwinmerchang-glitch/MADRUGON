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
    .debug-box {
        background: #fff3cd;
        border: 1px solid #ffc107;
        padding: 10px;
        border-radius: 5px;
        margin: 10px 0;
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
if 'mostrar_columnas' not in st.session_state:
    st.session_state.mostrar_columnas = False

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

def mostrar_columnas_excel(df):
    """Muestra todas las columnas del Excel para diagnóstico"""
    st.markdown("### 📋 Columnas encontradas en el archivo:")
    
    # Crear lista de columnas con ejemplos
    datos_columnas = []
    for col in df.columns:
        col_limpio = col.strip()
        ejemplo = df[col].dropna().iloc[0] if not df[col].dropna().empty else "N/A"
        datos_columnas.append({
            "Nombre Original": col,
            "Nombre Limpio": col_limpio,
            "Ejemplo": str(ejemplo)[:50] + "..." if len(str(ejemplo)) > 50 else str(ejemplo)
        })
    
    df_columnas = pd.DataFrame(datos_columnas)
    st.dataframe(df_columnas, use_container_width=True, hide_index=True)
    
    # Descargar lista de columnas
    columnas_texto = "\n".join([f"'{col}'" for col in df.columns])
    st.download_button(
        "📋 Copiar nombres de columnas",
        columnas_texto,
        "columnas_encontradas.txt",
        "text/plain"
    )

def obtener_info_producto(fila, df=None):
    """
    Extrae la información específica del producto
    Ahora busca en TODAS las columnas que contengan palabras clave
    """
    info = {
        'nombre': None,
        'precio_venta': None,
        'precio_descuento': None,
        'marca': None,
        'porcentaje_descuento': None
    }
    
    # Palabras clave para buscar en nombres de columnas
    keywords_nombre = ['name', 'nombre', 'product', 'producto', 'descrip']
    keywords_marca = ['marca', 'brand']
    keywords_precio_original = ['precio', 'venta', 'original', 'normal', 'lista', 'base']
    keywords_precio_descuento = ['descuento', 'oferta', 'especial', 'final', 'madrugón', 'madrugon']
    keywords_porcentaje = ['descuento', 'dscto', 'porcentaje']
    
    # Buscar en cada columna del DataFrame
    for col in fila.index:
        col_lower = col.strip().lower()
        valor = fila[col]
        
        if pd.isna(valor) or str(valor).strip() == '':
            continue
        
        # Buscar nombre del producto
        if info['nombre'] is None:
            if any(keyword in col_lower for keyword in keywords_nombre):
                # Verificar que no sea una columna de precios
                if not any(keyword in col_lower for keyword in ['precio', 'descuento', 'ean', 'código', 'codigo', 'supplier', 'day']):
                    try:
                        # Si se puede convertir a número, probablemente no es un nombre
                        float(str(valor).replace('$', '').replace(',', ''))
                    except:
                        info['nombre'] = str(valor).strip()
        
        # Buscar marca
        if info['marca'] is None:
            if any(keyword in col_lower for keyword in keywords_marca):
                info['marca'] = str(valor).strip()
        
        # Buscar precio original
        if info['precio_venta'] is None:
            if any(keyword in col_lower for keyword in keywords_precio_original):
                if not any(keyword in col_lower for keyword in ['descuento', 'oferta', 'especial']):
                    try:
                        valor_limpio = str(valor).replace('$', '').replace(',', '').strip()
                        precio = float(valor_limpio)
                        if precio > 0:
                            info['precio_venta'] = precio
                    except:
                        continue
        
        # Buscar precio con descuento
        if info['precio_descuento'] is None:
            if 'descuento' in col_lower or 'oferta' in col_lower or 'final' in col_lower or 'especial' in col_lower:
                try:
                    valor_limpio = str(valor).replace('$', '').replace(',', '').strip()
                    precio = float(valor_limpio)
                    if precio > 0:
                        info['precio_descuento'] = precio
                except:
                    continue
        
        # Buscar porcentaje de descuento
        if info['porcentaje_descuento'] is None:
            if any(keyword in col_lower for keyword in keywords_porcentaje):
                try:
                    valor_float = float(str(valor).replace('%', '').strip())
                    if valor_float > 0:
                        if valor_float <= 1:
                            info['porcentaje_descuento'] = valor_float * 100
                        else:
                            info['porcentaje_descuento'] = valor_float
                except:
                    continue
    
    # Si no se encontró precio con descuento específico, intentar con el que sea diferente al original
    if info['precio_venta'] is not None and info['precio_descuento'] is None:
        for col in fila.index:
            col_lower = col.strip().lower()
            valor = fila[col]
            
            if pd.isna(valor):
                continue
                
            if 'precio' in col_lower and 'descuento' not in col_lower and 'original' not in col_lower:
                try:
                    valor_limpio = str(valor).replace('$', '').replace(',', '').strip()
                    precio_temp = float(valor_limpio)
                    if precio_temp > 0 and precio_temp != info['precio_venta']:
                        info['precio_descuento'] = precio_temp
                        break
                except:
                    continue
    
    # Calcular porcentaje si tenemos precios pero no porcentaje
    if info['porcentaje_descuento'] is None and info['precio_venta'] and info['precio_descuento']:
        if info['precio_venta'] > 0 and info['precio_descuento'] < info['precio_venta']:
            info['porcentaje_descuento'] = ((info['precio_venta'] - info['precio_descuento']) / 
                                           info['precio_venta']) * 100
    
    return info

def mostrar_producto(info, debug=False):
    """Muestra la información del producto de forma atractiva"""
    
    if debug:
        st.markdown("---")
        st.markdown("### 🔍 Debug: Información detectada")
        st.json(info)
        st.markdown("---")
    
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
        
        col1, col2 = st.columns(2)
        
        with col1:
            if info['precio_venta'] is not None:
                st.markdown(f"""
                <div style="text-align: center;">
                    <div class="precio-label">Precio Original</div>
                    <div class="precio-original">${info['precio_venta']:,.0f}</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                <div style="text-align: center;">
                    <div class="precio-label">Precio Original</div>
                    <div style="color: #999;">No disponible</div>
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
            else:
                st.markdown("""
                <div style="text-align: center;">
                    <div class="precio-label">Precio con Descuento</div>
                    <div style="color: #999;">No disponible</div>
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
                    
                    # Mostrar columnas encontradas
                    with st.expander("🔍 Ver columnas del archivo"):
                        mostrar_columnas_excel(df)
                    
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
                        
                        with st.expander("🔍 Ver columnas del archivo"):
                            mostrar_columnas_excel(df)
                        
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
        
        # Opciones de debug
        with st.expander("⚙️ Opciones avanzadas"):
            modo_debug = st.checkbox("Modo debug (mostrar datos detectados)", value=False)
            if st.button("📋 Mostrar todas las columnas"):
                mostrar_columnas_excel(st.session_state.df)
        
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
            
            # Buscar en TODAS las columnas que puedan contener códigos
            for col in df.columns:
                col_lower = col.strip().lower()
                if any(keyword in col_lower for keyword in ['ean', 'código', 'codigo', 'code', 'barras']):
                    try:
                        mascara |= (df[col].astype(str).str.strip() == codigo)
                    except:
                        pass
            
            resultado = df[mascara]
            
            if not resultado.empty:
                fila = resultado.iloc[0]
                st.success("✨ ¡Producto encontrado!")
                
                # Extraer y mostrar solo la información necesaria
                info_producto = obtener_info_producto(fila, df)
                mostrar_producto(info_producto, debug=modo_debug)
                
                # Si hay modo debug, mostrar la fila completa
                if modo_debug:
                    with st.expander("🔍 Ver fila completa del Excel"):
                        st.dataframe(pd.DataFrame([fila]), use_container_width=True)
                
            else:
                st.error("❌ Código no encontrado")
                
                # Sugerencias
                with st.expander("💡 ¿Necesitas ayuda?"):
                    st.write("**Primeros 5 registros del archivo:**")
                    st.dataframe(df.head(5), use_container_width=True)
        
        # Botón para cambiar de archivo
        if st.button("🔄 Cargar otro archivo", use_container_width=True):
            st.session_state.datos_cargados = False
            st.session_state.df = None
            st.session_state.nombre_archivo = ""
            st.rerun()