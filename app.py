import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="centered"
)

# CSS personalizado
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
    .upload-area {
        border: 3px dashed #FF416C;
        border-radius: 20px;
        padding: 50px;
        text-align: center;
        background: #fafafa;
    }
    .stats-container {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 20px;
        border-radius: 15px;
        margin: 20px 0;
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
    columnas_porcentaje = [
        'Descuento', '% Descuento', 'Porcentaje', 'Descuento %',
        'DESCUENTO', '% DESCUENTO', 'PORCENTAJE', 'DESCUENTO %',
        'descuento', 'porcentaje', 'Dscto', 'DSCTO', 'dscto'
    ]
    
    columnas_precio_original = [
        'Precio Original', 'Precio Normal', 'Precio Lista', 'Precio Base',
        'PRECIO ORIGINAL', 'PRECIO NORMAL', 'PRECIO LISTA', 'PRECIO BASE',
        'Precio', 'PRECIO'
    ]
    
    columnas_precio_descuento = [
        'Precio Descuento', 'Precio Oferta', 'Precio Especial', 
        'PRECIO DESCUENTO', 'PRECIO OFERTA', 'PRECIO ESPECIAL',
        'Precio Final', 'PRECIO FINAL'
    ]
    
    # Buscar porcentaje directo
    for col in columnas_porcentaje:
        if col in fila.index and pd.notna(fila[col]):
            try:
                valor = float(fila[col])
                if valor == 0:
                    continue
                if valor <= 1:
                    porcentaje = valor * 100
                else:
                    porcentaje = valor
                return round(porcentaje, 1), None, None
            except:
                continue
    
    # Calcular desde precios
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
    descuento, precio_original, precio_descuento = calcular_descuento(fila)
    
    if descuento is not None and descuento > 0:
        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">🔥 ¡AHORRA!</div>
            <div class="porcentaje-descuento">{descuento:.0f}%</div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)
        
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
    
    # Mostrar detalles del producto
    st.write("### 📋 Detalles del Producto")
    
    columnas_excluir = ['EAN PADRE ', 'Código', 'Descuento', '% Descuento', 'Porcentaje', 
                       'DESCUENTO', '% DESCUENTO', 'PORCENTAJE']
    
    columnas_mostrar = [col for col in fila.index 
                       if col not in columnas_excluir 
                       and pd.notna(fila[col])
                       and str(fila[col]).strip() != '']
    
    if columnas_mostrar:
        datos_mostrar = {}
        for columna in columnas_mostrar:
            valor = fila[columna]
            try:
                valor_float = float(valor)
                if valor_float == int(valor_float):
                    datos_mostrar[columna.strip()] = int(valor_float)
                else:
                    datos_mostrar[columna.strip()] = f"${valor_float:,.2f}"
            except:
                datos_mostrar[columna.strip()] = valor
        
        # Mostrar en grid de 3 columnas
        for i in range(0, len(datos_mostrar), 3):
            cols = st.columns(3)
            items = list(datos_mostrar.items())[i:i+3]
            for j, (key, value) in enumerate(items):
                with cols[j]:
                    st.metric(label=key, value=value)

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
                    
                    # Vista previa de los datos
                    with st.expander("🔍 Vista previa de los datos"):
                        st.dataframe(
                            df.head(10),
                            use_container_width=True,
                            hide_index=True
                        )
                        st.caption(f"Mostrando 10 de {len(df):,} productos")
                    
                    st.success("👉 Ve a la pestaña **'Consultar Productos'** para comenzar a buscar")
    
    else:  # Archivo del servidor
        st.info("💡 Esta opción busca el archivo en el servidor donde está alojada la app")
        
        nombre_archivo = st.text_input(
            "Nombre del archivo en el servidor:",
            value="MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx",
            help="Nombre exacto del archivo incluyendo la extensión"
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
                        
                        with st.expander("🔍 Vista previa de los datos"):
                            st.dataframe(df.head(10), use_container_width=True, hide_index=True)
                        
                        st.success("👉 Ve a la pestaña **'Consultar Productos'** para comenzar a buscar")
            else:
                st.error(f"❌ No se encontró el archivo '{nombre_archivo}' en el servidor")
                
                # Mostrar archivos disponibles
                st.write("📁 Archivos en el directorio actual:")
                archivos_encontrados = [f for f in os.listdir() if f.endswith(('.xlsx', '.xls'))]
                if archivos_encontrados:
                    for archivo in archivos_encontrados:
                        st.code(archivo)
                else:
                    st.warning("No se encontraron archivos Excel en el directorio")

# Pestaña 2: Consultar Productos
with tab2:
    st.title("🔍 Consultar Productos")
    st.markdown("---")
    
    if not st.session_state.datos_cargados:
        st.warning("⚠️ Primero debes cargar un archivo en la pestaña 'Cargar Archivo'")
        
        # Mostrar estado actual
        st.info("""
        📋 **Instrucciones:**
        1. Ve a la pestaña **'Cargar Archivo'**
        2. Sube tu archivo Excel con los datos del Madrugón
        3. Vuelve a esta pestaña para buscar productos
        
        **Formatos aceptados:** .xlsx, .xls
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
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            codigo = st.text_input(
                "Código de barras",
                placeholder="Escanea o escribe el código aquí",
                key="codigo_input",
                label_visibility="collapsed"
            )
        
        with col2:
            buscar_btn = st.button("🔍 Buscar", type="primary", use_container_width=True)
        
        if codigo or buscar_btn:
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
                    st.success(f"✨ ¡Producto encontrado! ({len(resultado)} coincidencia(s))")
                    mostrar_producto(fila)
                    
                    # Si hay múltiples coincidencias, mostrarlas
                    if len(resultado) > 1:
                        with st.expander(f"📋 Ver todas las coincidencias ({len(resultado)})"):
                            st.dataframe(resultado, use_container_width=True, hide_index=True)
                else:
                    st.error("❌ Código no encontrado")
                    
                    # Sugerencias
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
            else:
                st.warning("⚠️ Por favor ingresa un código para buscar")
        
        # Separador visual
        st.markdown("---")
        
        # Botón para cambiar de archivo
        if st.button("🔄 Cargar otro archivo", use_container_width=True):
            st.session_state.datos_cargados = False
            st.session_state.df = None
            st.session_state.nombre_archivo = ""
            st.rerun()