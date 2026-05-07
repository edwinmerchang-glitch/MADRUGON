import streamlit as st
import pandas as pd
import os

st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============ CSS PERSONALIZADO ============
st.markdown("""
<style>
    .main {
        padding: 0rem 1rem;
    }
    
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
        animation: pulse 2s infinite;
    }
    .texto-descuento {
        font-size: 24px;
        font-weight: 600;
    }
    
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
    }
    
    .producto-card {
        background: white;
        border-radius: 15px;
        padding: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin: 20px 0;
    }
    .producto-nombre {
        font-size: 24px;
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
        background: #f8f9fa;
        border-radius: 15px;
        padding: 20px;
        margin-top: 20px;
    }
    .precio-original {
        text-decoration: line-through;
        color: #999;
        font-size: 20px;
    }
    .precio-descuento {
        color: #FF416C;
        font-size: 32px;
        font-weight: 700;
    }
    .precio-label {
        font-size: 14px;
        color: #666;
        margin-bottom: 5px;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    
    .ahorro-badge {
        display: inline-block;
        background: #4CAF50;
        color: white;
        padding: 8px 20px;
        border-radius: 25px;
        font-weight: 600;
        margin-top: 10px;
    }
    
    .status-card {
        background: white;
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
    }
    .status-title {
        font-size: 12px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #999;
    }
    .status-value {
        font-size: 20px;
        font-weight: 700;
        color: #1a1a1a;
    }
    
    .search-container {
        background: white;
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.05);
        margin: 20px 0;
    }
    
    .stTextInput > label {
        display: none;
    }
    
    .stTextInput input {
        font-size: 18px !important;
        padding: 12px 15px !important;
        border: 2px solid #e0e0e0 !important;
        border-radius: 10px !important;
    }
    
    .stTextInput input:focus {
        border-color: #FF416C !important;
        box-shadow: 0 0 0 2px rgba(255, 65, 108, 0.2) !important;
    }
    
    @media (max-width: 768px) {
        .porcentaje-descuento {
            font-size: 60px;
        }
        .producto-nombre {
            font-size: 20px;
        }
    }
</style>
""", unsafe_allow_html=True)

# ============ INICIALIZAR ESTADO ============
if 'datos_cargados' not in st.session_state:
    st.session_state.datos_cargados = False
if 'df' not in st.session_state:
    st.session_state.df = None
if 'nombre_archivo' not in st.session_state:
    st.session_state.nombre_archivo = ""
if 'total_productos' not in st.session_state:
    st.session_state.total_productos = 0
if 'modo_debug' not in st.session_state:
    st.session_state.modo_debug = False
if 'resultado_actual' not in st.session_state:
    st.session_state.resultado_actual = None
if 'codigo_actual' not in st.session_state:
    st.session_state.codigo_actual = ""
if 'mensaje_error' not in st.session_state:
    st.session_state.mensaje_error = ""
# Key dinámica para limpiar el input
if 'input_key' not in st.session_state:
    st.session_state.input_key = 0

# ============ FUNCIONES ============

@st.cache_data
def cargar_datos(archivo):
    """Carga y procesa el archivo Excel"""
    try:
        df = pd.read_excel(archivo, sheet_name="8")
        
        for col in df.columns:
            col_lower = col.strip().lower()
            if any(keyword in col_lower for keyword in ['ean', 'código', 'codigo', 'code']):
                df[col] = df[col].astype(str).str.strip()
        
        return df
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        return None

def buscar_producto(codigo, df):
    """Busca un producto por código"""
    if not codigo or df is None:
        return pd.DataFrame()
    
    mascara = pd.Series(False, index=df.index)
    for col in df.columns:
        col_lower = col.strip().lower()
        if any(kw in col_lower for kw in ['ean', 'código', 'codigo', 'code', 'barras']):
            try:
                mascara |= (df[col].astype(str).str.strip() == codigo)
            except:
                pass
    return df[mascara]

def obtener_info_producto(fila):
    """Extrae información del producto"""
    info = {
        'nombre': None,
        'precio_venta': None,
        'precio_descuento': None,
        'marca': None,
        'porcentaje_descuento': None
    }
    
    # Identificar porcentaje de descuento
    for col in fila.index:
        col_lower = col.strip().lower()
        valor = fila[col]
        
        if pd.isna(valor) or str(valor).strip() == '':
            continue
        
        if any(kw in col_lower for kw in ['porcentaje', '%', 'dscto']) and \
           not any(kw in col_lower for kw in ['precio', 'venta', 'valor']):
            try:
                valor_float = float(str(valor).replace('%', '').replace(',', '.').strip())
                if 0 < valor_float <= 100:
                    if valor_float <= 1:
                        info['porcentaje_descuento'] = valor_float * 100
                    else:
                        info['porcentaje_descuento'] = valor_float
                    continue
            except:
                pass
    
    # Buscar precios
    for col in fila.index:
        col_lower = col.strip().lower()
        valor = fila[col]
        
        if pd.isna(valor) or str(valor).strip() == '':
            continue
        
        try:
            valor_str = str(valor).replace('$', '').replace(',', '').strip()
            valor_float = float(valor_str)
            
            if valor_float <= 100 and any(kw in col_lower for kw in ['%', 'porcentaje', 'descuento']):
                continue
            
            if valor_float > 100000000:
                continue
            
            if info['precio_venta'] is None:
                if any(kw in col_lower for kw in ['precio', 'venta']) and \
                   not any(kw in col_lower for kw in ['descuento', 'final', 'dscto', 'oferta', 'especial']):
                    if valor_float > 100:
                        info['precio_venta'] = valor_float
                        continue
            
            if info['precio_descuento'] is None:
                if any(kw in col_lower for kw in ['precio']) and \
                   any(kw in col_lower for kw in ['descuento', 'final', 'dscto', 'oferta', 'especial']):
                    if valor_float > 100:
                        info['precio_descuento'] = valor_float
                        continue
            
        except (ValueError, TypeError):
            pass
    
    # Buscar nombre
    for col in fila.index:
        col_lower = col.strip().lower()
        valor = fila[col]
        
        if pd.isna(valor) or str(valor).strip() == '':
            continue
        
        if not any(kw in col_lower for kw in ['precio', 'descuento', 'ean', 'codigo', 'supplier', 
                                                'day', 'linea', 'estado', 'number', 'número']):
            try:
                float(str(valor).replace('$', '').replace(',', ''))
            except:
                if info['nombre'] is None:
                    valor_str = str(valor).strip()
                    if len(valor_str) > 3 and any(c.isalpha() for c in valor_str):
                        info['nombre'] = valor_str
        
        if info['marca'] is None:
            if any(kw in col_lower for kw in ['marca', 'brand']):
                info['marca'] = str(valor).strip()
    
    # Calcular precio con descuento si no existe
    if info['precio_descuento'] is None and info['precio_venta'] and info['porcentaje_descuento']:
        descuento_decimal = info['porcentaje_descuento'] / 100
        info['precio_descuento'] = info['precio_venta'] * (1 - descuento_decimal)
    
    # Calcular porcentaje si no existe
    if info['porcentaje_descuento'] is None and info['precio_venta'] and info['precio_descuento']:
        if info['precio_venta'] > info['precio_descuento']:
            info['porcentaje_descuento'] = ((info['precio_venta'] - info['precio_descuento']) / 
                                           info['precio_venta']) * 100
    
    return info

def mostrar_producto(info):
    """Muestra la información del producto formateada"""
    
    if info['porcentaje_descuento'] and info['porcentaje_descuento'] > 0:
        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">🔥 ¡AHORRA!</div>
            <div class="porcentaje-descuento">{info['porcentaje_descuento']:.0f}%</div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('<div class="producto-card">', unsafe_allow_html=True)
    
    if info['nombre']:
        st.markdown(f'<div class="producto-nombre">{info["nombre"]}</div>', unsafe_allow_html=True)
    
    if info['marca']:
        st.markdown(f'<div style="text-align: center;"><span class="producto-marca">🏷️ {info["marca"]}</span></div>', 
                   unsafe_allow_html=True)
    
    if info['precio_venta'] or info['precio_descuento']:
        st.markdown('<div class="precio-container">', unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown('<div class="precio-label">Precio Original</div>', unsafe_allow_html=True)
            if info['precio_venta']:
                st.markdown(f'<div class="precio-original">${info["precio_venta"]:,.0f}</div>', 
                          unsafe_allow_html=True)
            else:
                st.markdown('<div style="color: #999;">No disponible</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="precio-label">Precio con Descuento</div>', unsafe_allow_html=True)
            if info['precio_descuento']:
                st.markdown(f'<div class="precio-descuento">${info["precio_descuento"]:,.0f}</div>', 
                          unsafe_allow_html=True)
            else:
                st.markdown('<div style="color: #999;">No disponible</div>', unsafe_allow_html=True)
        
        st.markdown('</div>', unsafe_allow_html=True)
    
    if info['precio_venta'] and info['precio_descuento']:
        ahorro = info['precio_venta'] - info['precio_descuento']
        if ahorro > 0:
            st.markdown(f'''
            <div style="text-align: center; margin-top: 15px;">
                <span class="ahorro-badge">💰 Te ahorras: ${ahorro:,.0f}</span>
            </div>
            ''', unsafe_allow_html=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# ============ SIDEBAR ============
with st.sidebar:
    st.title("🛒 Descuentos App")
    st.markdown("---")
    
    st.header("📂 Datos")
    
    archivo_subido = st.file_uploader(
        "Cargar archivo Excel",
        type=["xlsx", "xls"],
        help="Selecciona el archivo con los datos de productos"
    )
    
    if archivo_subido is not None:
        with st.spinner("Procesando..."):
            df = cargar_datos(archivo_subido)
            if df is not None and not df.empty:
                st.session_state.df = df
                st.session_state.datos_cargados = True
                st.session_state.nombre_archivo = archivo_subido.name
                st.session_state.total_productos = len(df)
                st.session_state.resultado_actual = None
                st.session_state.codigo_actual = ""
                st.success("✅ Archivo cargado!")
    
    st.markdown("---")
    
    st.header("📊 Estado")
    
    if st.session_state.datos_cargados:
        st.markdown(f"""
        <div class="status-card">
            <div class="status-title">Archivo</div>
            <div class="status-value" style="font-size: 14px;">{st.session_state.nombre_archivo}</div>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        <div class="status-card">
            <div class="status-title">Productos</div>
            <div class="status-value">{st.session_state.total_productos:,}</div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.codigo_actual:
            st.markdown(f"""
            <div class="status-card">
                <div class="status-title">Última búsqueda</div>
                <div class="status-value" style="font-size: 14px;">{st.session_state.codigo_actual}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Sin archivo cargado")
    
    st.markdown("---")
    
    st.header("⚙️ Opciones")
    
    st.session_state.modo_debug = st.checkbox("Modo diagnóstico", value=False)
    
    if st.session_state.datos_cargados:
        if st.button("🗑️ Limpiar datos", use_container_width=True):
            st.session_state.datos_cargados = False
            st.session_state.df = None
            st.session_state.nombre_archivo = ""
            st.session_state.total_productos = 0
            st.session_state.resultado_actual = None
            st.session_state.codigo_actual = ""
            st.rerun()

# ============ CONTENIDO PRINCIPAL ============

st.title("🔍 Consulta de Productos")
st.caption("Escanea o ingresa un código de barras para ver descuentos")

if not st.session_state.datos_cargados:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 50px 20px;">
            <div style="font-size: 80px; margin-bottom: 20px;">📤</div>
            <h2>Sin archivo de datos</h2>
            <p style="color: #666; font-size: 16px;">
                Para comenzar, carga un archivo Excel usando el panel lateral izquierdo.
            </p>
            <div style="background: #f0f0f0; padding: 15px; border-radius: 10px; margin-top: 20px;">
                <p style="margin: 0; color: #999; font-size: 14px;">
                    ⬅️ Usa el menú lateral para cargar tus datos
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
else:
    # Buscador
    st.markdown('<div class="search-container">', unsafe_allow_html=True)
    
    col1, col2 = st.columns([4, 1])
    
    with col1:
        # Usar key dinámica para poder limpiar el input
        codigo_input = st.text_input(
            "Código de barras",
            placeholder="🔍 Escanea o escribe el código y presiona Enter...",
            label_visibility="collapsed",
            key=f"input_busqueda_{st.session_state.input_key}"
        )
    
    with col2:
        buscar_btn = st.button("🔍 Buscar", type="primary", use_container_width=True)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # PROCESAR BÚSQUEDA
    if codigo_input:
        codigo = codigo_input.strip()
        
        if codigo and codigo != st.session_state.codigo_actual:
            # Guardar código
            st.session_state.codigo_actual = codigo
            
            # Buscar
            resultado = buscar_producto(codigo, st.session_state.df)
            
            if not resultado.empty:
                st.session_state.resultado_actual = resultado.iloc[0]
                st.session_state.mensaje_error = ""
            else:
                st.session_state.resultado_actual = None
                st.session_state.mensaje_error = f"❌ Producto no encontrado: {codigo}"
            
            # Incrementar key para limpiar el input
            st.session_state.input_key += 1
            st.rerun()
    
    # También buscar con el botón
    if buscar_btn:
        # Obtener el valor actual del input (antes de limpiar)
        input_actual = st.session_state.get(f"input_busqueda_{st.session_state.input_key}", "").strip()
        
        if input_actual and input_actual != st.session_state.codigo_actual:
            st.session_state.codigo_actual = input_actual
            
            resultado = buscar_producto(input_actual, st.session_state.df)
            
            if not resultado.empty:
                st.session_state.resultado_actual = resultado.iloc[0]
                st.session_state.mensaje_error = ""
            else:
                st.session_state.resultado_actual = None
                st.session_state.mensaje_error = f"❌ Producto no encontrado: {input_actual}"
            
            st.session_state.input_key += 1
            st.rerun()
        elif not input_actual:
            st.warning("⚠️ Ingresa un código para buscar")
    
    # MOSTRAR RESULTADOS
    if st.session_state.resultado_actual is not None:
        st.success(f"✅ Producto encontrado: {st.session_state.codigo_actual}")
        
        info_producto = obtener_info_producto(st.session_state.resultado_actual)
        mostrar_producto(info_producto)
        
        if st.session_state.modo_debug:
            with st.expander("🔧 Diagnóstico"):
                st.write("**Código buscado:**", st.session_state.codigo_actual)
                st.write("**Información detectada:**")
                st.json(info_producto)
                st.write("**Datos completos:**")
                st.dataframe(pd.DataFrame([st.session_state.resultado_actual]), use_container_width=True)
    
    elif st.session_state.mensaje_error:
        st.error(st.session_state.mensaje_error)
        
        with st.expander("💡 Ayuda"):
            st.write("**Primeros 5 productos en la base:**")
            st.dataframe(st.session_state.df.head(5), use_container_width=True)
    
    elif not st.session_state.codigo_actual:
        st.info("👆 Ingresa un código de barras y presiona Enter")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #999; font-size: 12px;'>"
    "App de Consulta de Descuentos v2.5 | Desarrollado con Streamlit</p>",
    unsafe_allow_html=True)