import streamlit as st
import pandas as pd

# =========================================
# CONFIGURACIÓN
# =========================================
st.set_page_config(
    page_title="Consulta de Descuentos",
    page_icon="🛒",
    layout="wide"
)

# =========================================
# CSS
# =========================================
st.markdown("""
<style>

.main {
    padding-top: 1rem;
}

.descuento-container {
    background: linear-gradient(135deg, #FF416C, #FF4B2B);
    color: white;
    padding: 30px;
    border-radius: 20px;
    text-align: center;
    margin-top: 20px;
    margin-bottom: 20px;
}

.porcentaje-descuento {
    font-size: 80px;
    font-weight: 900;
    line-height: 1;
}

.texto-descuento {
    font-size: 24px;
    font-weight: 600;
}

.producto-card {
    background: white;
    border-radius: 20px;
    padding: 25px;
    box-shadow: 0px 4px 15px rgba(0,0,0,0.08);
    margin-top: 20px;
}

.producto-nombre {
    font-size: 28px;
    font-weight: bold;
    text-align: center;
    margin-bottom: 15px;
}

.producto-marca {
    text-align: center;
    font-size: 18px;
    color: gray;
    margin-bottom: 20px;
}

.precio-original {
    text-decoration: line-through;
    color: #888;
    font-size: 22px;
}

.precio-descuento {
    color: #FF416C;
    font-size: 36px;
    font-weight: bold;
}

.ahorro {
    margin-top: 20px;
    background: #4CAF50;
    color: white;
    padding: 10px 20px;
    border-radius: 30px;
    display: inline-block;
    font-weight: bold;
}

.search-box {
    background: white;
    padding: 20px;
    border-radius: 20px;
    box-shadow: 0px 2px 10px rgba(0,0,0,0.05);
    margin-top: 20px;
    margin-bottom: 20px;
}

</style>
""", unsafe_allow_html=True)

# =========================================
# SESSION STATE
# =========================================
if "df" not in st.session_state:
    st.session_state.df = None

if "buscar_auto" not in st.session_state:
    st.session_state.buscar_auto = False

if "ultimo_codigo" not in st.session_state:
    st.session_state.ultimo_codigo = ""

if "reset_input" not in st.session_state:
    st.session_state.reset_input = False

# =========================================
# FUNCIONES
# =========================================

@st.cache_data
def cargar_excel(archivo):
    try:
        df = pd.read_excel(archivo, sheet_name="8")

        for col in df.columns:
            if any(x in col.lower() for x in ["ean", "codigo", "código", "barras"]):
                df[col] = df[col].astype(str).str.strip()

        return df

    except Exception as e:
        st.error(f"Error leyendo Excel: {e}")
        return None


def buscar_producto():
    st.session_state.buscar_auto = True
    st.session_state.ultimo_codigo = st.session_state.codigo_input


def obtener_info_producto(fila):

    info = {
        "nombre": None,
        "marca": None,
        "precio_original": None,
        "precio_descuento": None,
        "porcentaje_descuento": None
    }

    for col in fila.index:

        valor = fila[col]

        if pd.isna(valor):
            continue

        col_lower = col.lower()

        # MARCA
        if info["marca"] is None:
            if "marca" in col_lower or "brand" in col_lower:
                info["marca"] = str(valor)

        # NOMBRE
        if info["nombre"] is None:
            if not any(x in col_lower for x in [
                "codigo", "ean", "precio",
                "descuento", "marca"
            ]):

                texto = str(valor)

                if len(texto) > 3:
                    info["nombre"] = texto

        # PORCENTAJE
        try:

            valor_num = float(
                str(valor)
                .replace("%", "")
                .replace(",", ".")
            )

            if (
                0 < valor_num <= 100 and
                any(x in col_lower for x in [
                    "%", "descuento", "dscto"
                ])
            ):
                info["porcentaje_descuento"] = valor_num

        except:
            pass

        # PRECIOS
        try:

            valor_num = float(
                str(valor)
                .replace("$", "")
                .replace(",", "")
            )

            if valor_num > 100:

                if (
                    info["precio_original"] is None and
                    "precio" in col_lower and
                    "descuento" not in col_lower
                ):
                    info["precio_original"] = valor_num

                if (
                    info["precio_descuento"] is None and
                    (
                        "descuento" in col_lower or
                        "final" in col_lower or
                        "oferta" in col_lower
                    )
                ):
                    info["precio_descuento"] = valor_num

        except:
            pass

    # CALCULAR PRECIO DESCUENTO
    if (
        info["precio_descuento"] is None and
        info["precio_original"] and
        info["porcentaje_descuento"]
    ):

        info["precio_descuento"] = (
            info["precio_original"] *
            (1 - info["porcentaje_descuento"] / 100)
        )

    # CALCULAR %
    if (
        info["porcentaje_descuento"] is None and
        info["precio_original"] and
        info["precio_descuento"]
    ):

        info["porcentaje_descuento"] = (
            (
                info["precio_original"] -
                info["precio_descuento"]
            ) /
            info["precio_original"]
        ) * 100

    return info


def mostrar_producto(info):

    if info["porcentaje_descuento"]:

        st.markdown(f"""
        <div class="descuento-container">
            <div class="texto-descuento">
                🔥 ¡AHORRA!
            </div>

            <div class="porcentaje-descuento">
                {info["porcentaje_descuento"]:.0f}%
            </div>

            <div class="texto-descuento">
                DE DESCUENTO
            </div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="producto-card">', unsafe_allow_html=True)

    if info["nombre"]:
        st.markdown(
            f'<div class="producto-nombre">{info["nombre"]}</div>',
            unsafe_allow_html=True
        )

    if info["marca"]:
        st.markdown(
            f'<div class="producto-marca">🏷️ {info["marca"]}</div>',
            unsafe_allow_html=True
        )

    col1, col2 = st.columns(2)

    with col1:

        st.caption("Precio Original")

        if info["precio_original"]:
            st.markdown(
                f'<div class="precio-original">${info["precio_original"]:,.0f}</div>',
                unsafe_allow_html=True
            )

    with col2:

        st.caption("Precio con Descuento")

        if info["precio_descuento"]:
            st.markdown(
                f'<div class="precio-descuento">${info["precio_descuento"]:,.0f}</div>',
                unsafe_allow_html=True
            )

    if (
        info["precio_original"] and
        info["precio_descuento"]
    ):

        ahorro = (
            info["precio_original"] -
            info["precio_descuento"]
        )

        st.markdown(
            f"""
            <div style="text-align:center;">
                <div class="ahorro">
                    💰 Ahorras ${ahorro:,.0f}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("</div>", unsafe_allow_html=True)

# =========================================
# SIDEBAR
# =========================================
with st.sidebar:

    st.title("🛒 Descuentos")

    archivo = st.file_uploader(
        "Cargar Excel",
        type=["xlsx", "xls"]
    )

    if archivo:

        df = cargar_excel(archivo)

        if df is not None:
            st.session_state.df = df
            st.success("Archivo cargado")

            st.write(f"Productos: {len(df):,}")

# =========================================
# MAIN
# =========================================
st.title("🔍 Consulta de Productos")
st.caption("Escanea o escribe el código de barras")

if st.session_state.df is None:

    st.info("⬅️ Carga un archivo Excel")

else:

    st.markdown(
        '<div class="search-box">',
        unsafe_allow_html=True
    )

    codigo = st.text_input(
        "",
        value="",
        placeholder="🔍 Escanea el código...",
        key="codigo_input",
        label_visibility="collapsed",
        on_change=buscar_producto
    )

    st.markdown(
        "</div>",
        unsafe_allow_html=True
    )

    # =========================================
    # BUSCAR
    # =========================================
    if st.session_state.buscar_auto:

        codigo = st.session_state.ultimo_codigo.strip()

        df = st.session_state.df

        mascara = pd.Series(False, index=df.index)

        for col in df.columns:

            col_lower = col.lower()

            if any(x in col_lower for x in [
                "ean",
                "codigo",
                "código",
                "barras",
                "code"
            ]):

                mascara |= (
                    df[col]
                    .astype(str)
                    .str.strip() == codigo
                )

        resultado = df[mascara]

        if not resultado.empty:

            fila = resultado.iloc[0]

            info = obtener_info_producto(fila)

            mostrar_producto(info)

        else:

            st.error("❌ Producto no encontrado")

        # RESET
        st.session_state.buscar_auto = False

        # AUTOFOCUS
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

# =========================================
# FOOTER
# =========================================
st.markdown("---")

st.caption("App de Consulta de Descuentos v2.0")