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

# ============ ESTADOS ============
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

if 'buscar_auto' not in st.session_state:
    st.session_state.buscar_auto = False

if 'ultimo_codigo' not in st.session_state:
    st.session_state.ultimo_codigo = ""

# ============ FUNCIONES ============

@st.cache_data
def cargar_datos(archivo):
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

def buscar_producto():
    st.session_state.buscar_auto = True
    st.session_state.ultimo_codigo = st.session_state.codigo_input

def obtener_info_producto(fila):

    info = {
        'nombre': None,
        'precio_venta': None,
        'precio_descuento': None,
        'marca': None,
        'porcentaje_descuento': None
    }

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

            except:
                pass

    for col in fila.index:

        col_lower = col.strip().lower()
        valor = fila[col]

        if pd.isna(valor) or str(valor).strip() == '':
            continue

        try:
            valor_str = str(valor).replace('$', '').replace(',', '').strip()
            valor_float = float(valor_str)

            if valor_float > 100000000:
                continue

            if info['precio_venta'] is None:

                if any(kw in col_lower for kw in ['precio', 'venta']) and \
                   not any(kw in col_lower for kw in ['descuento', 'final', 'dscto']):

                    if valor_float > 100:
                        info['precio_venta'] = valor_float

            if info['precio_descuento'] is None:

                if any(kw in col_lower for kw in ['precio']) and \
                   any(kw in col_lower for kw in ['descuento', 'final', 'dscto']):

                    if valor_float > 100:
                        info['precio_descuento'] = valor_float

        except:
            pass

    for col in fila.index:

        col_lower = col.strip().lower()
        valor = fila[col]

        if pd.isna(valor) or str(valor).strip() == '':
            continue

        if not any(kw in col_lower for kw in [
            'precio', 'descuento', 'ean', 'codigo',
            'supplier', 'day', 'linea', 'estado'
        ]):

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

    if info['precio_descuento'] is None and \
       info['precio_venta'] and \
       info['porcentaje_descuento']:

        descuento_decimal = info['porcentaje_descuento'] / 100

        info['precio_descuento'] = (
            info['precio_venta'] * (1 - descuento_decimal)
        )

    if info['porcentaje_descuento'] is None and \
       info['precio_venta'] and \
       info['precio_descuento']:

        if info['precio_venta'] > info['precio_descuento']:

            info['porcentaje_descuento'] = (
                (info['precio_venta'] - info['precio_descuento']) /
                info['precio_venta']
            ) * 100

    return info

def mostrar_producto(info):

    if info['porcentaje_descuento'] and info['porcentaje_descuento'] > 0:

        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">🔥 ¡AHORRA!</div>
            <div class="porcentaje-descuento">
                {info['porcentaje_descuento']:.0f}%
            </div>
            <div class="texto-descuento">DE DESCUENTO</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="producto-card">', unsafe_allow_html=True)

    if info['nombre']:
        st.markdown(
            f'<div class="producto-nombre">{info["nombre"]}</div>',
            unsafe_allow_html=True
        )

    if info['marca']:
        st.markdown(
            f'''
            <div style="text-align:center;">
                <span class="producto-marca">
                    🏷️ {info["marca"]}
                </span>
            </div>
            ''',
            unsafe_allow_html=True
        )

    if info['precio_venta'] or info['precio_descuento']:

        st.markdown('<div class="precio-container">', unsafe_allow_html=True)

        col1, col2 = st.columns(2)

        with col1:

            st.markdown(
                '<div class="precio-label">Precio Original</div>',
                unsafe_allow_html=True
            )

            if info['precio_venta']:
                st.markdown(
                    f'<div class="precio-original">${info["precio_venta"]:,.0f}</div>',
                    unsafe_allow_html=True
                )

        with col2:

            st.markdown(
                '<div class="precio-label">Precio con Descuento</div>',
                unsafe_allow_html=True
            )

            if info['precio_descuento']:
                st.markdown(
                    f'<div class="precio-descuento">${info["precio_descuento"]:,.0f}</div>',
                    unsafe_allow_html=True
                )

        st.markdown('</div>', unsafe_allow_html=True)

    if info['precio_venta'] and info['precio_descuento']:

        ahorro = info['precio_venta'] - info['precio_descuento']

        if ahorro > 0:

            st.markdown(f'''
            <div style="text-align:center; margin-top:15px;">
                <span class="ahorro-badge">
                    💰 Te ahorras: ${ahorro:,.0f}
                </span>
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
        type=["xlsx", "xls"]
    )

    if archivo_subido is not None:

        with st.spinner("Procesando..."):

            df = cargar_datos(archivo_subido)

            if df is not None and not df.empty:

                st.session_state.df = df
                st.session_state.datos_cargados = True
                st.session_state.nombre_archivo = archivo_subido.name
                st.session_state.total_productos = len(df)

                st.success("✅ Archivo cargado")

    st.markdown("---")

    st.session_state.modo_debug = st.checkbox(
        "Modo diagnóstico",
        value=False
    )

# ============ PRINCIPAL ============
st.title("🔍 Consulta de Productos")
st.caption("Escanea o escribe el código de barras")

if not st.session_state.datos_cargados:

    st.info("⬅️ Carga un archivo Excel para comenzar")

else:

    st.markdown(
        '<div class="search-container">',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns([4, 1])

    with col1:

        codigo = st.text_input(
            "",
            placeholder="🔍 Escanea o escribe el código...",
            key="codigo_input",
            label_visibility="collapsed",
            on_change=buscar_producto
        )

    with col2:

        buscar = st.button(
            "🔍 Buscar",
            type="primary",
            use_container_width=True
        )

    st.markdown('</div>', unsafe_allow_html=True)

    # ========= BUSQUEDA =========
    if buscar or st.session_state.buscar_auto:

        codigo = st.session_state.ultimo_codigo.strip()

        if codigo:

            df = st.session_state.df

            mascara = pd.Series(False, index=df.index)

            for col in df.columns:

                col_lower = col.strip().lower()

                if any(kw in col_lower for kw in [
                    'ean', 'código', 'codigo',
                    'code', 'barras'
                ]):

                    try:
                        mascara |= (
                            df[col]
                            .astype(str)
                            .str.strip() == codigo
                        )

                    except:
                        pass

            resultado = df[mascara]

            if not resultado.empty:

                fila = resultado.iloc[0]

                info_producto = obtener_info_producto(fila)

                mostrar_producto(info_producto)

                if st.session_state.modo_debug:

                    with st.expander("🔧 Diagnóstico"):

                        st.json(info_producto)

                        st.dataframe(
                            pd.DataFrame([fila]),
                            use_container_width=True
                        )

            else:

                st.error("❌ Producto no encontrado")

        # ========= RESET =========
        st.session_state.buscar_auto = False
        st.session_state.codigo_input = ""

        # ========= AUTOFOCUS =========
        st.components.v1.html(
            """
            <script>
                const inputs =
                    window.parent.document.querySelectorAll('input');

                for (const input of inputs) {
                    if (input.type === 'text') {
                        input.focus();
                        break;
                    }
                }
            </script>
            """,
            height=0
        )

# Footer
st.markdown("---")

st.markdown(
    """
    <p style='text-align:center;
              color:#999;
              font-size:12px;'>
        App de Consulta de Descuentos v2.0
    </p>
    """,
    unsafe_allow_html=True
)