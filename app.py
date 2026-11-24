
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="PASTELPAN ERP",
    page_icon="🎂",
    layout="wide"
)

# PALETA DE MARCA
ROJO = "#C83E3E"
VERDE = "#A9D18E"
CORAL = "#E88F91"
BEIGE = "#F3E8D5"

st.markdown("""
<style>
.stApp {background-color: #FFFDF9;}
[data-testid="stSidebar"] {background-color: #F3E8D5;}
h1, h2, h3 {color: #C83E3E;}
div.stButton > button {
    background-color: #C83E3E;
    color: white;
    border-radius: 9px;
}
</style>
""", unsafe_allow_html=True)


# ==========================================
# CONEXION CON EXCEL
# ==========================================

def obtener_excel():
    contenido = st.session_state.get("excel_bytes")
    if contenido is None:
        return None
    return pd.ExcelFile(BytesIO(contenido), engine="openpyxl")


def leer_hoja(nombre, columnas=None):
    excel = obtener_excel()
    if excel is None or nombre not in excel.sheet_names:
        return pd.DataFrame()

    return pd.read_excel(
        excel,
        sheet_name=nombre,
        header=2,
        usecols=columnas
    )


# ==========================================
# CARGA DEL ARCHIVO
# ==========================================

st.sidebar.markdown("# 🎂 PASTELPAN")
st.sidebar.caption("Una tentación irresistible")
st.sidebar.divider()

archivo = st.sidebar.file_uploader(
    "Cargar Excel maestro",
    type=["xlsx"]
)

if archivo is not None:
    st.session_state["excel_bytes"] = archivo.getvalue()
    st.session_state["excel_nombre"] = archivo.name

excel = obtener_excel()

MODULOS = [
    "Inicio",
    "Información",
    "CRM",
    "Producción",
    "Inventarios",
    "Compras y Proveedores",
    "Dashboard"
]

seccion = st.sidebar.radio("MENÚ PRINCIPAL", MODULOS)

st.sidebar.divider()
st.sidebar.caption("Sistema de gestión empresarial")


# ==========================================
# INICIO
# ==========================================

if seccion == "Inicio":

    st.title("🎂 PASTELPAN ERP")
    st.subheader("Una tentación irresistible")

    st.write("Sistema de gestión integral de PASTELPAN.")

    if excel is None:
        st.warning("Carga el Excel maestro desde el menú lateral.")
    else:
        st.success("Archivo maestro conectado correctamente.")
        st.metric("Hojas disponibles", len(excel.sheet_names))

    st.divider()
    st.subheader("Módulos del sistema")

    columnas = st.columns(3)

    modulos = [
        ("🏢 Información", "Datos generales de la empresa."),
        ("👥 CRM", "Gestión de clientes y atención."),
        ("🏭 Producción", "Planificación y control productivo."),
        ("📦 Inventarios", "Ingredientes y materiales."),
        ("🛒 Compras", "Proveedores y órdenes."),
        ("📊 Dashboard", "Indicadores de gestión.")
    ]

    for i, (nombre, descripcion) in enumerate(modulos):
        with columnas[i % 3]:
            with st.container(border=True):
                st.markdown(f"### {nombre}")
                st.write(descripcion)


# ==========================================
# INFORMACION
# ==========================================

elif seccion == "Información":

    st.title("🏢 Información de PASTELPAN")

    if excel is not None:
        datos = leer_hoja("Presentación de la empresa ")
        st.dataframe(datos, use_container_width=True)
    else:
        st.warning("Carga el Excel maestro.")


# ==========================================
# CRM
# ==========================================

elif seccion == "CRM":

    st.title("👥 Gestión de clientes (CRM)")
    st.caption("Administración comercial y atención al cliente")

    if excel is None:
        st.warning("Primero carga el Excel maestro.")
    else:

        clientes = leer_hoja("Clientes")
        ventas = leer_hoja("Atención al cliente", "A:J")

        if clientes.empty or ventas.empty:
            st.error("No se pudieron leer las hojas del CRM.")
        else:

            # Normalización de fechas
            clientes["Fecha de registro"] = pd.to_datetime(
                clientes["Fecha de registro"], errors="coerce"
            )

            ventas["Fecha de venta"] = pd.to_datetime(
                ventas["Fecha de venta"], errors="coerce"
            )

            ventas["Calificación NPS (0-5)"] = pd.to_numeric(
                ventas["Calificación NPS (0-5)"],
                errors="coerce"
            )

            # PESTAÑAS DEL CRM
            tab1, tab2 = st.tabs([
                "📋 Directorio de clientes",
                "🧾 Atención al cliente y ventas"
            ])

            # ----------------------------------
            # DIRECTORIO DE CLIENTES
            # ----------------------------------

            with tab1:

                st.subheader("Directorio de clientes")

                c1, c2, c3 = st.columns(3)

                c1.metric("Clientes registrados", len(clientes))
                c2.metric(
                    "Clientes con venta",
                    (clientes["Estado"] == "Venta").sum()
                )
                c3.metric(
                    "Clientes pendientes",
                    (clientes["Estado"] == "Pendiente").sum()
                )

                st.divider()

                st.markdown("#### Filtros de búsqueda")

                f1, f2, f3 = st.columns(3)

                with f1:
                    buscar = st.text_input(
                        "Buscar por nombre, apellido o ID"
                    )

                with f2:
                    estados = sorted(
                        clientes["Estado"].dropna().astype(str).unique()
                    )
                    filtro_estado = st.multiselect(
                        "Estado",
                        estados,
                        default=estados
                    )

                with f3:
                    asesoras = sorted(
                        clientes["Asesor asignado"].dropna().astype(str).unique()
                    )
                    filtro_asesora = st.multiselect(
                        "Asesora asignada",
                        asesoras,
                        default=asesoras
                    )

                f4, f5 = st.columns(2)

                with f4:
                    productos = sorted(
                        clientes["Producto de interés"].dropna().astype(str).unique()
                    )
                    filtro_producto = st.multiselect(
                        "Producto de interés",
                        productos,
                        default=productos
                    )

                with f5:
                    canales = sorted(
                        clientes["Canal de contacto"].dropna().astype(str).unique()
                    )
                    filtro_canal = st.multiselect(
                        "Canal de contacto",
                        canales,
                        default=canales
                    )

                resultado = clientes.copy()

                if buscar:
                    texto = buscar.strip().lower()
                    mascara = (
                        resultado["ID cliente"].astype(str).str.lower().str.contains(texto, na=False)
                        | resultado["Nombre"].astype(str).str.lower().str.contains(texto, na=False)
                        | resultado["Apellido"].astype(str).str.lower().str.contains(texto, na=False)
                    )
                    resultado = resultado[mascara]

                resultado = resultado[
                    resultado["Estado"].isin(filtro_estado)
                    & resultado["Asesor asignado"].isin(filtro_asesora)
                    & resultado["Producto de interés"].isin(filtro_producto)
                    & resultado["Canal de contacto"].isin(filtro_canal)
                ]

                st.metric("Resultados encontrados", len(resultado))

                st.dataframe(
                    resultado,
                    use_container_width=True,
                    hide_index=True
                )

                st.download_button(
                    "Descargar resultados filtrados",
                    data=resultado.to_csv(index=False).encode("utf-8-sig"),
                    file_name="clientes_pastelpan.csv",
                    mime="text/csv"
                )

            # ----------------------------------
            # ATENCION AL CLIENTE Y VENTAS
            # ----------------------------------

            with tab2:

                st.subheader("Registro de atención al cliente y ventas")

                v1, v2, v3 = st.columns(3)

                v1.metric("Registros de venta", len(ventas))

                unidades = pd.to_numeric(
                    ventas["Cantidad (unid.)"], errors="coerce"
                ).sum()

                v2.metric("Unidades vendidas", f"{unidades:,.0f}")

                promedio = ventas["Calificación NPS (0-5)"].mean()

                v3.metric(
                    "Calificación promedio (0-5)",
                    f"{promedio:.2f}" if pd.notna(promedio) else "Sin datos"
                )

                st.divider()
                st.markdown("#### Filtros de ventas")

                fechas = ventas["Fecha de venta"].dropna()

                if not fechas.empty:

                    minimo = fechas.min().date()
                    maximo = fechas.max().date()

                    fecha_inicio, fecha_fin = st.date_input(
                        "Periodo de consulta",
                        value=(minimo, maximo),
                        min_value=minimo,
                        max_value=maximo
                    )

                    v4, v5, v6 = st.columns(3)

                    with v4:
                        productos_v = sorted(
                            ventas["Producto"].dropna().astype(str).unique()
                        )
                        producto_v = st.multiselect(
                            "Producto",
                            productos_v,
                            default=productos_v,
                            key="producto_venta"
                        )

                    with v5:
                        asesoras_v = sorted(
                            ventas["Asesor"].dropna().astype(str).unique()
                        )
                        asesora_v = st.multiselect(
                            "Asesora",
                            asesoras_v,
                            default=asesoras_v
                        )

                    with v6:
                        locales = sorted(
                            ventas["Punto de venta / atención"].dropna().astype(str).unique()
                        )
                        local_v = st.multiselect(
                            "Local",
                            locales,
                            default=locales
                        )

                    resultado_v = ventas[
                        ventas["Fecha de venta"].dt.date.between(
                            fecha_inicio, fecha_fin
                        )
                        & ventas["Producto"].isin(producto_v)
                        & ventas["Asesor"].isin(asesora_v)
                        & ventas["Punto de venta / atención"].isin(local_v)
                    ].copy()

                    st.metric(
                        "Ventas encontradas",
                        len(resultado_v)
                    )

                    st.metric(
                        "Unidades del periodo",
                        pd.to_numeric(
                            resultado_v["Cantidad (unid.)"],
                            errors="coerce"
                        ).sum()
                    )

                    st.dataframe(
                        resultado_v,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.download_button(
                        "Descargar ventas filtradas",
                        data=resultado_v.to_csv(index=False).encode("utf-8-sig"),
                        file_name="ventas_pastelpan.csv",
                        mime="text/csv"
                    )


# ==========================================
# PRODUCCION
# ==========================================

elif seccion == "Producción":

    st.title("🏭 Gestión de Producción")

    if excel is not None:

        hojas = [
            "Demanda y Pronóstico Vainilla",
            "Demanda y Pronóstico Chocolate",
            "PAP ",
            "PMP Vainilla",
            "PMP Chocolate",
            "Maestro de Materiales",
            "BOM",
            "MRP I",
            "MRP II AJUSTADO",
            "MRP II FINAL",
            "DRP"
        ]

        opcion = st.selectbox("Proceso productivo", hojas)
        datos = leer_hoja(opcion)

        st.dataframe(datos, use_container_width=True)

    else:
        st.warning("Carga el Excel maestro.")


# ==========================================
# INVENTARIOS
# ==========================================

elif seccion == "Inventarios":

    st.title("📦 Gestión de Inventarios")

    if excel is not None:

        datos = leer_hoja("Maestro de Materiales")

        st.subheader("Maestro de materiales")
        st.metric("Materiales registrados", len(datos))
        st.dataframe(datos, use_container_width=True)

    else:
        st.warning("Carga el Excel maestro.")


# ==========================================
# COMPRAS Y PROVEEDORES
# ==========================================

elif seccion == "Compras y Proveedores":

    st.title("🛒 Compras y Proveedores")

    if excel is not None:

        opcion = st.selectbox(
            "Selecciona una sección",
            ["Proveedores", "Órdenes de Compra y Servicio"]
        )

        datos = leer_hoja(opcion)

        st.metric("Registros encontrados", len(datos))
        st.dataframe(datos, use_container_width=True)

    else:
        st.warning("Carga el Excel maestro.")


# ==========================================
# DASHBOARD
# ==========================================

elif seccion == "Dashboard":

    st.title("📊 Dashboard de PASTELPAN")

    if excel is not None:
        st.success("Archivo maestro conectado.")
        st.info(
            "Los indicadores principales y gráficos se "
            "incorporarán después de validar los módulos."
        )
    else:
        st.warning("Carga el Excel maestro.")


# ==========================================
# PIE DE PAGINA
# ==========================================

st.divider()
st.caption("PASTELPAN ERP | Una tentación irresistible")
