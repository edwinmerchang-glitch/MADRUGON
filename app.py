import streamlit as st
import pandas as pd
import sqlite3
import os
from datetime import datetime

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

# Cargar datos desde Excel a SQLite
def cargar_excel_a_sqlite(archivo_excel, hoja="8"):
    try:
        # Leer Excel
        df = pd.read_excel(archivo_excel, sheet_name=hoja, engine='openpyxl')
        
        # Limpiar columnas
        df.columns = df.columns.str.strip()
        
        # Renombrar columnas para estandarizar
        df.rename(columns={
            'Código': 'codigo',
            'Marca': 'marca',
            '% Descuento': 'descuento',
            'Precio de venta con descuento': 'precio_final',
            'Precio de venta': 'precio_original'
        }, inplace=True)
        
        # Limpiar datos
        df['codigo'] = df['codigo'].astype(str).str.strip()
        df['marca'] = df['marca'].astype(str).str.strip()
        
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
                float(row['descuento']) if pd.notna(row['descuento']) else 0,
                float(row['precio_final']) if pd.notna(row['precio_final']) else 0,
                float(row['precio_original']) if pd.notna(row['precio_original']) else 0,
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
    
    # Opcion para cargar Excel
    st.subheader("1. Cargar datos")
    archivo_excel = st.file_uploader(
        "Cargar archivo Excel",
        type=['xlsx', 'xls'],
        help="Sube el archivo Excel con los productos"
    )
    
    if archivo_excel:
        with st.spinner("Cargando datos a la base de datos..."):
            success, result = cargar_excel_a_sqlite(archivo_excel)
            if success:
                st.success(f"✅ Datos cargados! {result} productos en la base de datos")
                st.balloons()
            else:
                st.error(f"❌ Error: {result}")
    
    # Mostrar estadisticas
    st.divider()
    st.subheader("📊 Estadisticas")
    
    total_productos, promedio_desc = get_estadisticas()
    st.metric("Total de productos", total_productos)
    if promedio_desc:
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
if codigo and (buscar_click or codigo != st.session_state.get('last_codigo', '')):
    st.session_state['last_codigo'] = codigo
    
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
                st.metric("💰 Precio Original", "No disponible")
            
            if producto['precio_final'] > 0:
                ahorro = producto['precio_original'] - producto['precio_final']
                if ahorro > 0:
                    st.metric("💵 Precio Final", f"${producto['precio_final']:,.0f}", 
                             delta=f"Ahorro: ${ahorro:,.0f}")
                else:
                    st.metric("💵 Precio Final", f"${producto['precio_final']:,.0f}")
            else:
                st.metric("💵 Precio Final", "No disponible")
        
        # Mostrar codigo
        st.caption(f"📌 Codigo: `{producto['codigo']}`")
        
    else:
        st.error("❌ Producto no encontrado")
        st.info("""
        **Sugerencias:**
        - Verifica que el codigo sea correcto
        - Asegurate de haber cargado los datos del Excel
        - Prueba con otro codigo
        """)

# Mostrar ultimos productos (opcional)
st.markdown("---")
with st.expander("📋 Ver ultimos productos cargados"):
    conn = sqlite3.connect('productos.db')
    df_preview = pd.read_sql_query(
        "SELECT codigo, marca, descuento, precio_final FROM productos LIMIT 20", 
        conn
    )
    conn.close()
    
    if not df_preview.empty:
        st.dataframe(df_preview, use_container_width=True)
    else:
        st.info("No hay productos cargados. Por favor, carga un archivo Excel.")

# Footer
st.markdown("---")
st.caption("Scanner de Descuentos v1.0 - Powered by SQLite | Escanea y ahorra")