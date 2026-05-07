import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime
import io

# Configurar la pagina
st.set_page_config(
    page_title="Scanner de Descuentos", 
    layout="centered",
    page_icon="🛒"
)

# Inicializar base de datos SQLite
def init_database():
    conn = sqlite3.connect('productos.db')
    c = conn.cursor()
    
    # Crear tabla si no existe
    c.execute('''CREATE TABLE IF NOT EXISTS productos (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        codigo TEXT UNIQUE,
        marca TEXT,
        descuento REAL,
        precio_final REAL,
        precio_original REAL,
        fecha_actualizacion TIMESTAMP
    )''')
    
    conn.commit()
    conn.close()

# Cargar datos desde archivo (Excel o CSV)
def cargar_datos_a_sqlite(archivo, tipo='excel', hoja="8"):
    try:
        if tipo == 'excel':
            # Intentar leer Excel sin openpyxl (usar xlrd para .xls)
            try:
                df = pd.read_excel(archivo, sheet_name=hoja)
            except:
                # Si falla, intentar con diferentes motores
                try:
                    df = pd.read_excel(archivo, sheet_name=hoja, engine='xlrd')
                except:
                    st.error("No se pudo leer el Excel. Por favor, guarda el archivo como CSV y usalo en su lugar.")
                    return False, "Error leyendo Excel"
        else:  # CSV
            df = pd.read_csv(archivo)
        
        # Limpiar columnas
        df.columns = df.columns.str.strip()
        
        # Identificar columnas (buscar nombres similares)
        col_codigo = None
        col_marca = None
        col_descuento = None
        col_precio_final = None
        col_precio_original = None
        
        for col in df.columns:
            col_lower = col.lower().strip()
            if 'codigo' in col_lower or 'código' in col_lower or 'cod' in col_lower:
                col_codigo = col
            elif 'marca' in col_lower:
                col_marca = col
            elif 'descuento' in col_lower or '%' in col:
                col_descuento = col
            elif 'final' in col_lower or 'con descuento' in col_lower:
                col_precio_final = col
            elif 'original' in col_lower or 'venta' in col_lower:
                col_precio_original = col
        
        # Usar las columnas encontradas o las originales
        df.rename(columns={
            col_codigo: 'codigo' if col_codigo else 'Código',
            col_marca: 'marca' if col_marca else 'Marca',
            col_descuento: 'descuento' if col_descuento else '% Descuento',
            col_precio_final: 'precio_final' if col_precio_final else 'Precio de venta con descuento',
            col_precio_original: 'precio_original' if col_precio_original else 'Precio de venta'
        }, inplace=True)
        
        # Tomar solo las columnas necesarias
        columnas_necesarias = ['codigo', 'marca', 'descuento', 'precio_final', 'precio_original']
        columnas_existentes = [col for col in columnas_necesarias if col in df.columns]
        
        if not columnas_existentes:
            return False, "No se encontraron las columnas necesarias"
        
        df = df[columnas_existentes]
        
        # Limpiar datos
        df['codigo'] = df['codigo'].astype(str).str.strip()
        df['marca'] = df['marca'].astype(str).str.strip()
        df['descuento'] = pd.to_numeric(df['descuento'], errors='coerce').fillna(0)
        df['precio_final'] = pd.to_numeric(df['precio_final'], errors='coerce').fillna(0)
        df['precio_original'] = pd.to_numeric(df['precio_original'], errors='coerce').fillna(0)
        
        # Conectar a SQLite
        conn = sqlite3.connect('productos.db')
        
        # Limpiar tabla existente
        conn.execute("DELETE FROM productos")
        
        # Insertar datos
        for _, row in df.iterrows():
            conn.execute("""
                INSERT OR REPLACE INTO productos 
                (codigo, marca, descuento, precio_final, precio_original, fecha_actualizacion)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                row['codigo'],
                row['marca'],
                float(row['descuento']),
                float(row['precio_final']),
                float(row['precio_original']),
                datetime.now()
            ))
        
        conn.commit()
        conn.close()
        
        return True, len(df)
    except Exception as e:
        return False, str(e)

# Buscar producto en SQLite
def buscar_producto(codigo):
    conn = sqlite3.connect('productos.db')
    c = conn.cursor()
    
    c.execute("SELECT * FROM productos WHERE codigo = ?", (codigo.strip(),))
    resultado = c.fetchone()
    
    conn.close()
    
    if resultado:
        return {
            'codigo': resultado[1],
            'marca': resultado[2],
            'descuento': resultado[3],
            'precio_final': resultado[4],
            'precio_original': resultado[5]
        }
    return None

# Obtener estadisticas
def get_estadisticas():
    conn = sqlite3.connect('productos.db')
    c = conn.cursor()
    
    c.execute("SELECT COUNT(*) FROM productos")
    total = c.fetchone()[0]
    
    c.execute("SELECT AVG(descuento) FROM productos WHERE descuento > 0")
    promedio_desc = c.fetchone()[0]
    
    conn.close()
    
    return total, promedio_desc

# Inicializar base de datos
init_database()

# Interfaz principal
st.title("🛒 Scanner de Descuentos")
st.write("Escanea o escribe el codigo de barras para ver el descuento")

# Sidebar
with st.sidebar:
    st.header("📋 Configuracion")
    
    # Opcion para elegir tipo de archivo
    st.subheader("1. Cargar datos")
    tipo_archivo = st.radio(
        "Tipo de archivo:",
        ["Excel (.xls)", "CSV (Recomendado)"],
        help="CSV es mas compatible y no necesita librerias adicionales"
    )
    
    # Subir archivo
    if tipo_archivo == "CSV (Recomendado)":
        archivo = st.file_uploader(
            "Cargar archivo CSV",
            type=['csv'],
            help="Guarda tu Excel como CSV (Archivo > Guardar como > CSV)"
        )
    else:
        archivo = st.file_uploader(
            "Cargar archivo Excel",
            type=['xls', 'xlsx'],
            help="Para archivos .xlsx necesitas instalar openpyxl"
        )
    
    if archivo:
        with st.spinner("Cargando datos a la base de datos..."):
            if tipo_archivo == "CSV (Recomendado)":
                success, result = cargar_datos_a_sqlite(archivo, tipo='csv')
            else:
                success, result = cargar_datos_a_sqlite(archivo, tipo='excel')
            
            if success:
                st.success(f"✅ Datos cargados! {result} productos en la base de datos")
                st.balloons()
            else:
                st.error(f"❌ Error: {result}")
                if "openpyxl" in str(result):
                    st.info("💡 Consejo: Guarda tu archivo Excel como CSV y sube ese archivo en su lugar.")
    
    # Mostrar estadisticas
    st.divider()
    st.subheader("📊 Estadisticas")
    
    total_productos, promedio_desc = get_estadisticas()
    st.metric("Total de productos", total_productos)
    if promedio_desc and promedio_desc > 0:
        st.metric("Descuento promedio", f"{promedio_desc:.1f}%")
    
    # Opciones
    st.divider()
    if st.button("🗑️ Limpiar base de datos"):
        conn = sqlite3.connect('productos.db')
        conn.execute("DELETE FROM productos")
        conn.commit()
        conn.close()
        st.success("Base de datos limpiada!")
        st.rerun()

# Area principal de busqueda
st.markdown("---")

# Busqueda por codigo
codigo = st.text_input(
    "🔍 Codigo de barras",
    placeholder="Ej: 1234567890",
    help="Escanea o escribe el codigo de barras del producto",
    key="codigo_input"
)

# Boton de busqueda
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    buscar_click = st.button("🔎 Buscar producto", use_container_width=True)

# Realizar busqueda
if codigo and buscar_click:
    with st.spinner("Buscando..."):
        producto = buscar_producto(codigo)
    
    if producto:
        st.success("✅ Producto encontrado!")
        
        # Mostrar informacion del producto
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric("🏷️ Marca", producto['marca'])
            if producto['descuento'] > 0:
                st.metric("🔻 Descuento", f"{producto['descuento']:.0f}%", 
                         delta=f"-{producto['descuento']:.0f}%", 
                         delta_color="inverse")
            else:
                st.metric("🔻 Descuento", "0%")
        
        with col2:
            if producto['precio_original'] > 0:
                st.metric("💰 Precio Original", f"${producto['precio_original']:,.0f}")
            else:
                st.metric("💰 Precio Original", "-")
            
            if producto['precio_final'] > 0:
                ahorro = producto['precio_original'] - producto['precio_final']
                if ahorro > 0:
                    st.metric("💵 Precio Final", f"${producto['precio_final']:,.0f}", 
                             delta=f"Ahorro: ${ahorro:,.0f}")
                else:
                    st.metric("💵 Precio Final", f"${producto['precio_final']:,.0f}")
            else:
                st.metric("💵 Precio Final", "-")
    else:
        st.error("❌ Producto no encontrado")

# Mostrar productos cargados
st.markdown("---")
with st.expander("📋 Ver productos en la base de datos"):
    conn = sqlite3.connect('productos.db')
    df_preview = pd.read_sql_query(
        "SELECT codigo, marca, descuento, precio_final FROM productos LIMIT 20", 
        conn
    )
    conn.close()
    
    if not df_preview.empty:
        st.dataframe(df_preview, use_container_width=True)
    else:
        st.info("No hay productos cargados. Por favor, carga un archivo.")

# Instrucciones
with st.expander("📖 Instrucciones para usar CSV (Recomendado)"):
    st.markdown("""
    ### Como convertir Excel a CSV:
    
    1. **Abre tu archivo Excel**
    2. **Ve a Archivo > Guardar como**
    3. **Elige "CSV UTF-8 (Comma delimited)"** 
    4. **Guarda el archivo**
    5. **Sube el archivo CSV a esta app**
    
    ### Ventajas del CSV:
    - ✅ No necesita librerias adicionales
    - ✅ Funciona perfectamente en Streamlit Cloud
    - ✅ Es mas rapido de procesar
    - ✅ Compatible con todos los sistemas
    """)

# Footer
st.markdown("---")
st.caption("Scanner de Descuentos v1.0 - Usa CSV para mejor compatibilidad")