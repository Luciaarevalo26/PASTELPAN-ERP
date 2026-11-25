st.warning("Carga el Excel maestro.")

    else:

        ventas = leer_hoja("Atención al cliente", "A:J")
        clientes = leer_hoja("Clientes")

        if not ventas.empty:

            ventas["Cantidad (unid.)"] = pd.to_numeric(
                ventas["Cantidad (unid.)"],
                errors="coerce"
            )

            ventas["Fecha de venta"] = pd.to_datetime(
                ventas["Fecha de venta"],
                errors="coerce"
            )

            ventas["Calificación NPS (0-5)"] = pd.to_numeric(
                ventas["Calificación NPS (0-5)"],
                errors="coerce"
            )

            total_ventas = len(ventas)

            total_unidades = ventas["Cantidad (unid.)"].sum()

            promedio = ventas["Calificación NPS (0-5)"].mean()

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "Registros de venta",
                f"{total_ventas:,}"
            )

            c2.metric(
                "Unidades vendidas",
                f"{total_unidades:,.0f}"
            )

            c3.metric(
                "Calificación promedio",
                f"{promedio:.2f}/5" if pd.notna(promedio) else "Sin datos"
            )

            st.divider()

            col1, col2 = st.columns(2)

            with col1:

                st.subheader("Ventas por producto")

                grafico_productos = (
                    ventas.groupby("Producto")["Cantidad (unid.)"]
                    .sum()
                )

                st.bar_chart(grafico_productos)

            with col2:

                st.subheader("Evolución de ventas")

                ventas_mes = ventas.dropna(
                    subset=["Fecha de venta"]
                ).copy()

                ventas_mes["Mes"] = (
                    ventas_mes["Fecha de venta"]
                    .dt.to_period("M")
                    .astype(str)
                )

                grafico_meses = (
                    ventas_mes.groupby("Mes")["Cantidad (unid.)"]
                    .sum()
                    .sort_index()
                )

                if not grafico_meses.empty:

                    st.line_chart(grafico_meses)

                else:

                    st.info("No hay fechas disponibles para mostrar la evolución.")

            st.divider()

            if not clientes.empty:

                st.subheader("Estado de los clientes")

                st.bar_chart(
                    clientes["Estado"].value_counts()
                )

        else:

            st.warning(
                "No se encontraron registros de ventas para generar el Dashboard."
            )


# ==========================================
# PIE DE PÁGINA
# ==========================================

st.divider()

st.caption(
    "PASTELPAN ERP | Una tentación irresistible"
)
