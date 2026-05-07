""")
st.stop()

# Verificar openpyxl primero
verificar_instalar_openpyxl()

# Configurar la página
st.set_page_config(
page_title="Scanner de Descuentos", 
layout="centered",
page_icon="🛒"
)

# Título de la app
st.title("🛒 Scanner de Descuentos")
st.write("Escanea o escribe el código de barras para ver el descuento")

# Sidebar con información
with st.sidebar:
st.header("📋 Instrucciones")
st.markdown("""
1. Carga el archivo Excel
2. Escanea o escribe el código
3. Mira el descuento aplicado

**Formato requerido del Excel:**
- Hoja llamada: **8**
- Columnas necesarias:
- `Código`
- `Marca`
- `% Descuento`
- `Precio de venta con descuento`
- `Precio de venta`
""")

st.divider()

# Opción para usar archivo local (alternativa)
usar_archivo_local = st.checkbox("Usar archivo local desde mi PC")

if usar_archivo_local:
ruta_local = st.text_input(
"Ruta del archivo",
placeholder="Ej: C:/Users/edwin/Downloads/archivo.xlsx"
)

# Subir archivo Excel (principal)
if not usar_archivo_local:
archivo_subido = st.file_uploader(
"📂 Cargar archivo Excel", 
type=['xlsx', 'xls'],
help="Sube el archivo MADRUGON MAYO 2026 PUNTO DE VENTA.xlsx",
key="file_uploader"
)

if archivo_subido is not None:
try:
HOJA = "8"

# Mostrar spinner mientras carga
with st.spinner("Cargando archivo..."):
    # Cargar el archivo subido
    df = pd.read_excel(archivo_subido, sheet_name=HOJA, engine='openpyxl')

# Limpiar columnas
df.columns = df.columns.str.strip()

# Convertir código a texto
df['Código'] = df['Código'].astype(str).str.strip()

# Mostrar información del archivo
st.success(f"✅ Archivo cargado correctamente")
st.info(f"📊 Total de productos: {len(df)}")

# INPUT CÓDIGO
st.markdown("---")
codigo = st.text_input(
    "🔍 Código de barras", 
    placeholder="Ej: 1234567890",
    help="Escribe o escanea el código de barras del producto"
)

# Botón para buscar (opcional)
buscar = st.button("🔎 Buscar producto", use_container_width=True)

# BÚSQUEDA (con Enter o con botón)
if codigo and (buscar or True):
    resultado = df[df['Código'] == codigo.strip()]
    
    if not resultado.empty:
        producto = resultado.iloc[0]
        
        # Obtener valores con manejo de errores
        marca = producto.get('Marca', 'No disponible')
        descuento = producto.get('% Descuento', producto.get('Porcentaje Descuento', 0))
        precio_final = producto.get('Precio de venta con descuento', 0)
        precio_original = producto.get('Precio de venta', producto.get('Precio de venta ', 0))
        
        st.balloons()
        st.success("✅ ¡Producto encontrado!")
        
        # Mostrar tarjeta del producto
        with st.container():
            st.markdown("### 📦 Información del producto")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.metric("🏷️ Marca", str(marca))
                if descuento and float(descuento) > 0:
                    st.metric("🔻 Descuento", f"{descuento}%", delta=f"-{descuento}%", delta_color="inverse")
                else:
                    st.metric("🔻 Descuento", "0%")
            
            with col2:
                if pd.notna(precio_original) and precio_original != 0:
                    st.metric("💰 Precio Original", f"${precio_original:,.0f}")
                else:
                    st.metric("💰 Precio Original", "No disponible")
                
                if pd.notna(precio_final) and precio_final != 0:
                    ahorro = precio_original - precio_final if pd.notna(precio_original) else 0
                    st.metric("💵 Precio Final", f"${precio_final:,.0f}", delta=f"Ahorro: ${ahorro:,.0f}" if ahorro > 0 else None)
                else:
                    st.metric("💵 Precio Final", "No disponible")
        
        # Mostrar código encontrado
        st.caption(f"✅ Código escaneado: `{codigo}`")
        
    else:
        st.error("❌ Código no encontrado")
        
        # Sugerir códigos similares
        if len(codigo) >= 3:
            codigos_similares = df[df['Código'].str.contains(codigo[:3], na=False)].head(5)
            if not codigos_similares.empty:
                st.info("💡 **¿Buscabas alguno de estos códigos?**")
                for idx, row in codigos_similares.iterrows():
                    st.code(f"📌 {row['Código']} - {row.get('Marca', 'Sin marca')}")

# Mostrar vista previa (opcional)
with st.expander("📋 Ver vista previa de productos"):
    st.dataframe(
        df[['Código', 'Marca', '% Descuento', 'Precio de venta con descuento']].head(10),
        use_container_width=True
    )
    
except Exception as e:
st.error(f"❌ Error: {e}")
st.info("""
**Posibles soluciones:**
1. Verifica que el archivo no esté corrupto
2. Asegúrate de que la hoja se llame exactamente **'8'**
3. Confirma que el archivo tenga las columnas necesarias
4. Revisa que los datos estén en el formato correcto
""")
else:
st.info("👈 **Por favor, carga un archivo Excel para comenzar**")
st.markdown("""
### 🚀 Cómo usar la app:
1. Haz clic en **'Browse files'**
2. Selecciona tu archivo Excel (debe tener una hoja llamada '8')
3. Espera a que se cargue
4. Escanea o escribe el código de barras
5. ¡Mira el descuento aplicado!
""")

# Opción de archivo local
else:
if ruta_local and os.path.exists(ruta_local):
try:
HOJA = "8"

with st.spinner("Cargando archivo local..."):
    df = pd.read_excel(ruta_local, sheet_name=HOJA, engine='openpyxl')

df.columns = df.columns.str.strip()
df['Código'] = df['Código'].astype(str).str.strip()

st.success(f"✅ Archivo local cargado: {os.path.basename(ruta_local)}")
st.info(f"📊 Total de productos: {len(df)}")

# INPUT CÓDIGO
st.markdown("---")
codigo = st.text_input("🔍 Código de barras", placeholder="Ej: 1234567890")

if codigo:
    resultado = df[df['Código'] == codigo.strip()]
    
    if not resultado.empty:
        producto = resultado.iloc[0]
        
        marca = producto.get('Marca', 'No disponible')
        descuento = producto.get('% Descuento', producto.get('Porcentaje Descuento', 0))
        precio_final = producto.get('Precio de venta con descuento', 0)
        precio_original = producto.get('Precio de venta', producto.get('Precio de venta ', 0))
        
        st.balloons()
        st.success("✅ ¡Producto encontrado!")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric("🏷️ Marca", str(marca))
            st.metric("🔻 Descuento", f"{descuento}%")
        
        with col2:
            if pd.notna(precio_original) and precio_original != 0:
                st.metric("💰 Precio Original", f"${precio_original:,.0f}")
            else:
                st.metric("💰 Precio Original", "No disponible")
            
            if pd.notna(precio_final) and precio_final != 0:
                st.metric("💵 Precio Final", f"${precio_final:,.0f}")
            else:
                st.metric("💵 Precio Final", "No disponible")
    else:
        st.error("❌ Código no encontrado")
        
except Exception as e:
st.error(f"Error: {e}")
elif usar_archivo_local:
st.warning("⚠️ Por favor, ingresa una ruta válida al archivo Excel")

# Footer
st.markdown("---")
st.caption("🔍 Scanner de Descuentos v1.0 - Escanea y ahorra")