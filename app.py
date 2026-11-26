filtro_orden = filtro_opciones(ordenes, "Estado", "Estado de órdenes") if not ordenes.empty else None
        with f4:
            filtro_proveedor = filtro_opciones(proveedores, "Calificación", "Calificación de proveedor") if not proveedores.empty else None

        clientes = aplicar_filtro(clientes, "Estado", filtro_lead)
        atencion = aplicar_filtro(atencion, "Calificación NPS (0-5)", filtro_atencion)
        ordenes = aplicar_filtro(ordenes, "Estado", filtro_orden)
        proveedores = aplicar_filtro(proveedores, "Calificación", filtro_proveedor)

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            if not clientes.empty and "Estado" in clientes.columns:
                conteo = clientes["Estado"].fillna("Sin estado").value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                grafico_barras(conteo, "Estado", "Cantidad", "CRM | Leads por estado")
            if not proveedores.empty and "Calificación" in proveedores.columns:
                proveedores["Calificación"] = pd.to_numeric(proveedores["Calificación"], errors="coerce")
                prov_plot = proveedores.dropna(subset=["Calificación"])
                if not prov_plot.empty:
                    grafico_barras(prov_plot, "Nombre del proveedor", "Calificación", "Compras | Calificación de proveedores")
        with col_g2:
            if not atencion.empty and "Calificación NPS (0-5)" in atencion.columns:
                atencion["Calificación NPS (0-5)"] = pd.to_numeric(atencion["Calificación NPS (0-5)"], errors="coerce")
                distrib = atencion["Calificación NPS (0-5)"].value_counts().sort_index().rename_axis("Calificación").reset_index(name="Cantidad")
                if not distrib.empty:
                    fig_donut = px.pie(distrib, names="Calificación", values="Cantidad", hole=0.58,
                                       title="CRM | Distribución de calificaciones", color_discrete_sequence=PALETA_PASTEL)
                    fig_donut.update_layout(title_x=0.02, paper_bgcolor="white", margin=dict(l=15,r=15,t=55,b=15), legend_title_text="")
                    st.plotly_chart(fig_donut, use_container_width=True)
            if not ordenes.empty and "Estado" in ordenes.columns:
                conteo_ord = ordenes["Estado"].fillna("Sin estado").value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                if not conteo_ord.empty:
                    fig_pie = px.pie(conteo_ord, names="Estado", values="Cantidad",
                                     title="Abastecimiento | Órdenes por estado", color_discrete_sequence=PALETA_PASTEL)
                    fig_pie.update_layout(title_x=0.02, paper_bgcolor="white", margin=dict(l=15,r=15,t=55,b=15), legend_title_text="")
                    st.plotly_chart(fig_pie, use_container_width=True)

        if not atencion.empty and "Fecha de venta" in atencion.columns:
            col_cantidad = next((c for c in ["Cantidad (unid.)", "Cantidad", "Unidades"] if c in atencion.columns), None)
            if col_cantidad:
                ventas_tiempo = atencion[["Fecha de venta", col_cantidad]].copy()
                ventas_tiempo["Fecha de venta"] = pd.to_datetime(ventas_tiempo["Fecha de venta"], errors="coerce")
                ventas_tiempo[col_cantidad] = pd.to_numeric(ventas_tiempo[col_cantidad], errors="coerce")
                ventas_tiempo = ventas_tiempo.dropna()
                if not ventas_tiempo.empty:
                    ventas_tiempo["Mes"] = ventas_tiempo["Fecha de venta"].dt.to_period("M").astype(str)
                    ventas_mes = ventas_tiempo.groupby("Mes", as_index=False)[col_cantidad].sum()
                    grafico_linea(ventas_mes, "Mes", col_cantidad, "Ventas | Evolución mensual de unidades")

        df_comp_raw = rango_df("INDICADORES ERP", "A12:G48", True)
        if not df_comp_raw.empty:
            st.subheader("Indicadores complementarios")
            df_comp = df_comp_raw.copy()
            if len(df_comp.columns) >= 4:
                df_comp = df_comp.iloc[:, :4].copy()
                df_comp.columns = ["Indicador", "Módulo", "Resultado", "Unidad"]
            else:
                st.warning("La estructura de indicadores complementarios no contiene las columnas esperadas.")
                df_comp = pd.DataFrame()
            if not df_comp.empty:
                df_comp = df_comp.dropna(how="all")
                fc1, fc2 = st.columns([1, 2])
                with fc1:
                    modulos_comp = ["Todos"] + sorted(df_comp["Módulo"].dropna().astype(str).unique().tolist())
                    modulo_comp = st.selectbox("Filtrar por módulo", modulos_comp, key="dashboard_filtro_modulo")
                with fc2:
                    buscar_indicador = st.text_input("Buscar indicador", placeholder="Nombre del indicador", key="dashboard_buscar_indicador")
                if modulo_comp != "Todos":
                    df_comp = df_comp[df_comp["Módulo"].astype(str) == modulo_comp]
                if buscar_indicador:
                    df_comp = df_comp[df_comp["Indicador"].astype(str).str.contains(buscar_indicador, case=False, na=False)]
                if not df_comp.empty:
                    vista_comp = df_comp[["Indicador", "Resultado", "Módulo"]].copy()
                    def formato_resultado_fila(fila):
                        val = pd.to_numeric(pd.Series([fila["Resultado"]]), errors="coerce").iloc[0]
                        unidad = str(fila.get("Unidad", "")).lower()
                        if pd.isna(val):
                            return str(fila["Resultado"])
                        if "porcentaje" in unidad or "%" in unidad:
                            return formato_porcentaje(val)
                        if "puntos" in unidad:
                            return f"{formato_numero(val)}/5"
                        return formato_numero(val)
                    vista_comp["Resultado"] = df_comp.apply(formato_resultado_fila, axis=1)
                    mostrar_tabla(vista_comp, 300)
                else:
                    st.info("No hay indicadores que coincidan con los filtros seleccionados.")

        st.caption("Los gráficos responden a los filtros seleccionados y se actualizan con los registros del Excel maestro.")


# =========================================================
# PIE DE PÁGINA Y DESCARGA GENERAL
# =========================================================
st.divider()
col_a, col_b = st.columns([3, 1])
with col_a:
    st.caption("PASTELPAN ERP – Una tentación irresistible")
with col_b:
    if bytes_actuales():
        st.download_button(
            "Descargar Excel maestro",
            data=bytes_actuales(),
            file_name="PASTELPAN_ERP.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="descarga_general"
        )
