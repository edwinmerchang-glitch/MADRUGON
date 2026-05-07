import streamlit as st
        st.divider()

        col1, col2 = st.columns(2)

        # ==========================
        # COLUMNA 1
        # ==========================

        with col1:

            st.metric(
                "Marca",
                str(fila.get('Marca', 'Sin dato'))
            )

            st.metric(
                "Descuento",
                f"{fila.get('Porcentaje Descuento', 0)}%"
            )

        # ==========================
        # COLUMNA 2
        # ==========================

        with col2:

            try:
                precio_original = int(fila.get('Precio de venta ', 0))
            except:
                precio_original = 0

            try:
                precio_final = int(fila.get('Precio de venta con descuento ', 0))
            except:
                precio_final = 0

            st.metric(
                "Precio Original",
                f"$ {precio_original:,}"
            )

            st.metric(
                "Precio Final",
                f"$ {precio_final:,}"
            )

        st.divider()

        st.subheader("Producto")
        st.info(str(fila.get('name', 'Sin nombre')))

        st.subheader("Estado")
        st.write(str(fila.get('ESTADO', 'Sin estado')))

    else:
        st.error("Código no encontrado")