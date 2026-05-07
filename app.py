import streamlit as st
import pandas as pd
import sqlite3
import os
from io import BytesIO

# Configuración de la página
st.set_page_config(
    page_title="Sistema de Consulta de Productos",
    page_icon="🛒",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilos personalizados
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 1rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .product-card {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
        padding: 1.5rem;
        border-radius: 15px;
        margin: 1rem 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .price {
        font-size: 2rem;
        font-weight: bold;
        color: #28a745;
    }
    .original-price {
        text-decoration: line-through;
        color: #dc3545;
        font-size: 1.2rem;
    }
    .discount {
        background-color: #ffc107;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-weight: bold;
        display: inline-block;
    }
    .success {
        background-color: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 10px;
    }
    .error {
        background-color: #f8d7da;
        color: #721c24;
        padding: 1rem;
        border-radius: 10px;
    }
    .stButton button {
        background-color: #667eea;
        color: white;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)


class ProductosApp:
    def __init__(self):
        self.db_path = "productos.db"
        self.init_database()
    
    def init_database(self):
        """Inicializar la base de datos SQLite"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS productos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                codigo TEXT NOT NULL UNIQUE,
                nombre TEXT NOT NULL,
                marca TEXT,
                porcentaje_descuento REAL,
                precio_original REAL,
                precio_final REAL,
                estado TEXT,
                proveedor TEXT,
                categoria TEXT
            )
        ''')
        
        conn.commit()
        conn.close()
    
    def cargar_desde_excel(self, archivo_bytes):
        """Cargar datos desde archivo Excel"""
        try:
            df = pd.read_excel(BytesIO(archivo_bytes), sheet_name=8)
            
            conn = sqlite3.connect(self.db_path)
            
            # Limpiar y preparar datos
            df_clean = df.copy()
            
            # Mapear columnas
            df_clean['codigo'] = df_clean['EAN PADRE'].astype(str).str.strip()
            df_clean['nombre'] = df_clean['name'].fillna('')
            df_clean['marca'] = df_clean['Marca'].fillna('')
            df_clean['porcentaje_descuento'] = pd.to_numeric(df_clean['Porcentaje Descuento'], errors='coerce').fillna(0) * 100
            df_clean['precio_original'] = pd.to_numeric(df_clean['Precio de venta'], errors='coerce').fillna(0)
            df_clean['precio_final'] = pd.to_numeric(df_clean['Precio de venta con descuento'], errors='coerce').fillna(0)
            df_clean['estado'] = df_clean['ESTADO'].fillna('')
            df_clean['proveedor'] = df_clean['supplier'].fillna('')
            df_clean['categoria'] = df_clean['NUMERO LINEA'].fillna('').astype(str)
            
            # Insertar datos
            for _, row in df_clean.iterrows():
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT OR REPLACE INTO productos 
                    (codigo, nombre, marca, porcentaje_descuento, precio_original, 
                     precio_final, estado, proveedor, categoria)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    str(row['codigo']),
                    str(row['nombre'])[:500],
                    str(row['marca']),
                    float(row['porcentaje_descuento']),
                    float(row['precio_original']),
                    float(row['precio_final']),
                    str(row['estado']),
                    str(row['proveedor']),
                    str(row['categoria'])
                ))
                conn.commit()
            
            conn.close()
            return len(df_clean)
        except Exception as e:
            st.error(f"Error al cargar datos: {e}")
            return 0
    
    def buscar_producto(self, codigo):
        """Buscar producto por código de barras"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT codigo, nombre, marca, porcentaje_descuento, 
                   precio_original, precio_final, estado, proveedor, categoria
            FROM productos 
            WHERE codigo = ? OR codigo LIKE ? OR codigo LIKE ?
        ''', (codigo, f'%{codigo}%', f'{codigo}%'))
        
        resultado = cursor.fetchone()
        conn.close()
        return resultado
    
    def buscar_por_nombre(self, termino):
        """Buscar productos por nombre"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT codigo, nombre, marca, porcentaje_descuento, 
                   precio_original, precio_final, estado
            FROM productos 
            WHERE LOWER(nombre) LIKE ? AND estado = 'ACTIVO'
            LIMIT 50
        ''', (f'%{termino.lower()}%',))
        
        resultados = cursor.fetchall()
        conn.close()
        return resultados
    
    def get_marcas_top(self, limit=10):
        """Obtener marcas con más productos"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT marca, COUNT(*) as total 
            FROM productos 
            WHERE marca != '' AND estado = 'ACTIVO'
            GROUP BY marca 
            ORDER BY total DESC 
            LIMIT ?
        ''', (limit,))
        
        resultados = cursor.fetchall()
        conn.close()
        return resultados
    
    def get_productos_mas_descuento(self, limit=20):
        """Obtener productos con mayor descuento"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT nombre, marca, porcentaje_descuento, precio_original, precio_final
            FROM productos 
            WHERE estado = 'ACTIVO' AND porcentaje_descuento > 0
            ORDER BY porcentaje_descuento DESC 
            LIMIT ?
        ''', (limit,))
        
        resultados = cursor.fetchall()
        conn.close()
        return resultados
    
    def get_stats(self):
        """Obtener estadísticas de la base de datos"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute("SELECT COUNT(*) FROM productos")
        total = cursor.fetchone()[0]
        
        cursor.execute("SELECT COUNT(*) FROM productos WHERE estado = 'ACTIVO'")
        activos = cursor.fetchone()[0]
        
        cursor.execute("SELECT AVG(porcentaje_descuento) FROM productos WHERE estado = 'ACTIVO'")
        promedio_descuento = cursor.fetchone()[0] or 0
        
        conn.close()
        return total, activos, promedio_descuento


def main():
    st.markdown('<div class="main-header"><h1>🛒 Sistema de Consulta de Productos</h1><p>Escanea o ingresa el código de barras</p></div>', unsafe_allow_html=True)
    
    # Inicializar app
    if 'app' not in st.session_state:
        st.session_state.app = ProductosApp()
    
    app = st.session_state.app
    
    # Sidebar para carga de datos
    with st.sidebar:
        st.header("📁 Carga de Datos")
        
        uploaded_file = st.file_uploader(
            "Cargar archivo Excel", 
            type=['xlsx', 'xls'],
            help="Selecciona el archivo Excel con los datos de productos"
        )
        
        if uploaded_file is not None:
            if st.button("📥 Cargar a Base de Datos"):
                with st.spinner("Cargando productos..."):
                    total = app.cargar_desde_excel(uploaded_file.getvalue())
                    st.success(f"✅ {total} productos cargados exitosamente!")
        
        st.divider()
        
        # Estadísticas
        st.header("📊 Estadísticas")
        total, activos, promedio_desc = app.get_stats()
        st.metric("Total Productos", f"{total:,}")
        st.metric("Productos Activos", f"{activos:,}")
        st.metric("Descuento Promedio", f"{promedio_desc:.1f}%")
        
        st.divider()
        
        # Marcas destacadas
        st.header("🏷️ Marcas Destacadas")
        marcas = app.get_marcas_top(8)
        for marca, total in marcas:
            st.write(f"• {marca}: {total} productos")
    
    # Pestañas principales
    tab1, tab2, tab3, tab4 = st.tabs(["🔍 Escáner", "📋 Búsqueda", "🔥 Ofertas", "ℹ️ Ayuda"])
    
    # Tab 1: Escáner de códigos
    with tab1:
        st.subheader("🔍 Escanear Código de Barras")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            codigo = st.text_input(
                "Ingresa o escanea el código de barras:",
                placeholder="Ej: 7709990384987",
                key="codigo_input",
                help="Puedes usar un lector USB o ingresar manualmente"
            )
        
        with col2:
            buscar = st.button("🔍 Buscar", type="primary", use_container_width=True)
        
        if buscar and codigo:
            with st.spinner("Buscando producto..."):
                producto = app.buscar_producto(codigo)
                
                if producto:
                    codigo, nombre, marca, descuento, precio_orig, precio_final, estado, proveedor, categoria = producto
                    
                    # Mostrar resultado
                    st.markdown(f"""
                    <div class="product-card">
                        <h2>📦 {nombre}</h2>
                        <p><strong>🏷️ Marca:</strong> {marca}</p>
                        <p><strong>🔢 Código:</strong> {codigo}</p>
                        <p><strong>📌 Categoría:</strong> {categoria if categoria else 'No especificada'}</p>
                        <p><strong>🏢 Proveedor:</strong> {proveedor if proveedor else 'No especificado'}</p>
                        <p><strong>📊 Estado:</strong> {'✅ ACTIVO' if estado == 'ACTIVO' else '❌ INACTIVO'}</p>
                        <br>
                        <div style="text-align: center;">
                            <span class="discount">🎯 {descuento:.0f}% DE DESCUENTO</span>
                            <br><br>
                            <span class="original-price">Precio original: ${precio_orig:,.0f}</span>
                            <br>
                            <span class="price">💰 Precio final: ${precio_final:,.0f}</span>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Mostrar ahorro
                    ahorro = precio_orig - precio_final
                    if ahorro > 0:
                        st.success(f"💸 ¡Ahorras ${ahorro:,.0f} en este producto!")
                else:
                    st.markdown(f"""
                    <div class="error">
                        ❌ No se encontró ningún producto con el código: <strong>{codigo}</strong><br>
                        💡 Verifica que el código sea correcto o carga los datos en el sidebar.
                    </div>
                    """, unsafe_allow_html=True)
        
        # Sugerencia de códigos comunes
        with st.expander("📋 Ver códigos de ejemplo"):
            st.code("""
            Códigos de ejemplo:
            - 7709990384987 (SHOT-B MULTIVITAMINICO)
            - 7707066043844 (PANTORRILL.NVX SPORT)
            - 7707066043646 (PANTORRILL.NVX SPORT AGUAMA)
            - 7509546684765 (DESOD.LADY SPEED S.MINIMIZER)
            """)
    
    # Tab 2: Búsqueda por nombre
    with tab2:
        st.subheader("📋 Búsqueda por Nombre")
        
        termino = st.text_input(
            "Ingresa el nombre del producto:",
            placeholder="Ej: SHOT-B, VOGUE, MAYBELLINE",
            key="search_input"
        )
        
        if termino:
            resultados = app.buscar_por_nombre(termino)
            
            if resultados:
                st.write(f"🔍 Se encontraron **{len(resultados)}** productos:")
                
                for prod in resultados:
                    codigo, nombre, marca, descuento, precio_orig, precio_final, estado = prod
                    
                    with st.container():
                        col1, col2, col3 = st.columns([3, 1, 1])
                        with col1:
                            st.write(f"**{nombre}**")
                            st.write(f"📌 {marca} | Código: {codigo}")
                        with col2:
                            if descuento > 0:
                                st.write(f"🎯 {descuento:.0f}% OFF")
                            st.write(f"💰 ${precio_final:,.0f}")
                        with col3:
                            if precio_orig > precio_final:
                                st.write(f"~~${precio_orig:,.0f}~~")
                        st.divider()
            else:
                st.warning(f"No se encontraron productos con: '{termino}'")
    
    # Tab 3: Ofertas destacadas
    with tab3:
        st.subheader("🔥 Productos con Mejores Descuentos")
        
        productos_oferta = app.get_productos_mas_descuento(20)
        
        if productos_oferta:
            cols = st.columns(2)
            for i, prod in enumerate(productos_oferta):
                nombre, marca, descuento, precio_orig, precio_final = prod
                
                with cols[i % 2]:
                    st.markdown(f"""
                    <div style="background: linear-gradient(135deg, #fff5e6 0%, #ffe6cc 100%); 
                                padding: 1rem; border-radius: 10px; margin: 0.5rem 0;">
                        <h4>🔥 {nombre[:50]}</h4>
                        <p><strong>{marca}</strong></p>
                        <p><span class="discount">{descuento:.0f}% DESCUENTO</span></p>
                        <p><span class="original-price">${precio_orig:,.0f}</span> → 
                        <span style="font-size: 1.5rem; font-weight: bold; color: #28a745;">${precio_final:,.0f}</span></p>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("Carga los productos para ver las ofertas destacadas")
    
    # Tab 4: Ayuda
    with tab4:
        st.subheader("ℹ️ Instrucciones de uso")
        
        st.markdown("""
        ### 📱 Cómo usar esta aplicación:
        
        1. **Carga de datos**:
           - En el panel izquierdo, carga tu archivo Excel
           - La aplicación procesará automáticamente los datos
        
        2. **Escaneo de códigos**:
           - Usa un lector de códigos de barras USB
           - O ingresa el código manualmente
           - Presiona Enter o el botón Buscar
        
        3. **Búsqueda por nombre**:
           - Escribe parte del nombre del producto
           - Ejemplo: "VOGUE", "MAYBELLINE", "SHOT-B"
        
        4. **Ofertas**:
           - Revisa los productos con mayores descuentos
        
        ### 🔧 Requisitos técnicos:
        - Navegador web actualizado
        - Conexión a internet (para Streamlit Cloud)
        - El archivo Excel debe estar en el formato correcto
        
        ### 💡 Tips:
        - Los códigos de barras pueden tener 13 dígitos
        - Puedes configurar un lector para que envíe Enter automáticamente
        - Los datos se guardan localmente en tu sesión
        """)
        
        st.info("📌 La base de datos se mantiene durante tu sesión. Para usar datos nuevos, recarga la página y vuelve a cargar el archivo.")


if __name__ == "__main__":
    main()