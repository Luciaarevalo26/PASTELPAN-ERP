
import streamlit as st
import pandas as pd
from io import BytesIO

# ==========================================
# PASTELPAN ERP
# Una tentación irresistible
# ==========================================

st.set_page_config(
    page_title="PASTELPAN ERP",
    page_icon="🎂",
    layout="wide",
    initial_sidebar_state="expanded"
)

# PALETA DE COLORES
VERDE = "#A9D18E"
ROJO = "#C83E3E"
CORAL = "#E88F91"
BEIGE = "#F3E8D5"

# HOJAS DEL EXCEL MAESTRO
HOJAS_REQUERIDAS = [
    "Integrantes",
    "Presentación de la empresa ",
    "Definición del producto ",
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
    "DRP",
    "Clientes",
    "Atención al cliente",
    "Proveedores",
    "Órdenes de Compra y Servicio"
]

# DISEÑO
st.markdown("""
<style>
    .stApp {
        background-color: #FFFDF9;
    }
    [data-testid="stSidebar"] {
        background-color: #F3E8D5;
    }
    h1, h2, h3 {
        color: #C83E3E;
    }
    div.stButton > button {
        background-color: #C83E3E;
        color: white;
        border-radius: 9px;
        border: none;
    }
    div.stButton > button:hover {
        background-color: #A9D18E;
        color: #333333;
    }
</style>
""", unsafe_allow_html=True)


# ==========================================
# FUNCIONES DE CONEXIÓN CON EXCEL
# ==========================================

def obtener_excel():
    """Lee el archivo cargado sin modificar el original."""
    archivo = st.session_state.get("excel_bytes")

    if archivo is None:
        return None

    return pd.ExcelFile(BytesIO(archivo), engine="openpyxl")


def leer_hoja(nombre_hoja):
    """Permite consultar una hoja del Excel."""
    excel = obtener_excel()

    if excel is None or nombre_hoja not in excel.sheet_names:
        return None

    return pd.read_excel(excel, sheet_name=nombre_hoja)


# ==========================================
# ENCABEZADO Y MENÚ
# ==========================================

st.sidebar.markdown("# 🎂 PASTELPAN")
st.sidebar.caption("Una tentación irresistible")
st.sidebar.markdown("---")

st.sidebar.subheader("Archivo maestro")

archivo = st.sidebar.file_uploader(
    "Cargar Excel del proyecto",
    type=["xlsx"],
    help="Carga el archivo maestro definitivo de PASTELPAN."
)

if archivo is not None:
    st.session_state["excel_bytes"] = archivo.getvalue()
    st.session_state["excel_nombre"] = archivo.name

st.sidebar.markdown("---")

MODULOS = [
    "Inicio",
    "Información",
    "CRM",
    "Producción",
    "Inventarios",
    "Compras y Proveedores",
    "Dashboard"
]

seccion = st.sidebar.radio(
    "MENÚ PRINCIPAL",
    MODULOS,
    key="seccion"
)

st.sidebar.markdown("---")
st.sidebar.caption("Sistema de gestión empresarial")


# ==========================================
# ESTADO DEL ARCHIVO
# ==========================================

excel = obtener_excel()

if excel is not None:
    hojas_disponibles = excel.sheet_names
    hojas_faltantes = [
        hoja for hoja in HOJAS_REQUERIDAS
        if hoja not in hojas_disponibles
    ]

    if not hojas_faltantes:
        estado_excel = True
    else:
        estado_excel = False
else:
    hojas_disponibles = []
    hojas_faltantes = HOJAS_REQUERIDAS
    estado_excel = False


# ==========================================
# MÓDULO INICIO
# ==========================================

if seccion == "Inicio":

    st.title("🎂 PASTELPAN ERP")
    st.subheader("Una tentación irresistible")

    st.write(
        "Bienvenido al sistema de gestión empresarial "
        "de PASTELPAN."
    )

    st.markdown("---")

    st.subheader("Panel principal")

    if excel is None:
        st.warning(
            "Para comenzar, carga el archivo Excel maestro "
            "desde el menú lateral."
        )
    elif estado_excel:
        st.success(
            f"Archivo conectado correctamente: "
            f"{st.session_state['excel_nombre']}"
        )
        st.metric("Hojas reconocidas", len(hojas_disponibles))
    else:
        st.warning("El archivo se cargó, pero faltan algunas hojas.")
        for hoja in hojas_faltantes:
            st.write(f"- {hoja}")

    st.markdown("---")
    st.subheader("Módulos del sistema")

    columnas = st.columns(3)

    modulos = [
        ("🏢 Información", "Datos generales de la empresa."),
        ("👥 CRM", "Clientes y atención al cliente."),
        ("🏭 Producción", "Planificación y control productivo."),
        ("📦 Inventarios", "Ingredientes y materiales."),
        ("🛒 Compras y Proveedores", "Abastecimiento y proveedores."),
        ("📊 Dashboard", "Indicadores de gestión.")
    ]

    for i, (nombre, descripcion) in enumerate(modulos):
        with columnas[i % 3]:
            with st.container(border=True):
                st.markdown(f"### {nombre}")
                st.write(descripcion)


# ==========================================
# MÓDULO INFORMACIÓN
# ==========================================

elif seccion == "Información":

    st.title("🏢 Información de PASTELPAN")

    if excel is not None:
        datos = leer_hoja("Presentación de la empresa ")

        if datos is not None:
            st.write("Información obtenida del archivo maestro.")
            st.dataframe(datos, use_container_width=True)
    else:
        st.info("Carga el Excel maestro para consultar la información.")


# ==========================================
# MÓDULO CRM
# ==========================================

elif seccion == "CRM":

    st.title("👥 Gestión de clientes (CRM)")

    st.write("Administración de clientes y atención al cliente.")

    if excel is not None:
        opcion = st.selectbox(
            "Selecciona una sección",
            ["Clientes", "Atención al cliente"]
        )

        nombre_hoja = opcion

        datos = leer_hoja(nombre_hoja)

        if datos is not None:
            st.success(f"Hoja {opcion} reconocida.")
            st.metric("Registros encontrados", len(datos))
            st.info(
                "La consulta detallada y los filtros se "
                "implementarán en la siguiente etapa."
            )
    else:
        st.warning("Primero debes cargar el Excel maestro.")


# ==========================================
# MÓDULO PRODUCCIÓN
# ==========================================

elif seccion == "Producción":

    st.title("🏭 Gestión de Producción")

    hojas_produccion = [
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

    if excel is not None:

        opcion = st.selectbox(
            "Selecciona el proceso productivo",
            hojas_produccion
        )

        datos = leer_hoja(opcion)

        if datos is not None:
            st.success(f"Hoja {opcion} reconocida.")
            st.metric("Registros encontrados", len(datos))
            st.info(
                "La visualización de los planes y cálculos "
                "se desarrollará en la siguiente etapa."
            )
    else:
        st.warning("Primero debes cargar el Excel maestro.")


# ==========================================
# MÓDULO INVENTARIOS
# ==========================================

elif seccion == "Inventarios":

    st.title("📦 Gestión de Inventarios")

    st.write("Control de ingredientes y materiales.")

    if excel is not None:

        datos = leer_hoja("Maestro de Materiales")

        if datos is not None:
            st.success("Maestro de materiales reconocido.")
            st.metric("Registros encontrados", len(datos))
            st.info(
                "Aquí desarrollaremos la consulta de stock "
                "y el registro de nuevos materiales."
            )
    else:
        st.warning("Primero debes cargar el Excel maestro.")


# ==========================================
# MÓDULO COMPRAS Y PROVEEDORES
# ==========================================

elif seccion == "Compras y Proveedores":

    st.title("🛒 Compras y Proveedores")

    if excel is not None:

        opcion = st.selectbox(
            "Selecciona una sección",
            ["Proveedores", "Órdenes de Compra y Servicio"]
        )

        datos = leer_hoja(opcion)

        if datos is not None:
            st.success(f"Hoja {opcion} reconocida.")
            st.metric("Registros encontrados", len(datos))
            st.info(
                "La gestión de órdenes y proveedores "
                "se desarrollará en la siguiente etapa."
            )
    else:
        st.warning("Primero debes cargar el Excel maestro.")


# ==========================================
# MÓDULO DASHBOARD
# ==========================================

elif seccion == "Dashboard":

    st.title("📊 Dashboard de PASTELPAN")

    st.write("Panel de indicadores de gestión empresarial.")

    if excel is not None:
        st.success("Archivo maestro disponible.")
        st.info(
            "En esta sección construiremos los indicadores "
            "y gráficos utilizando los datos reales del proyecto."
        )
    else:
        st.warning("Primero debes cargar el Excel maestro.")


# ==========================================
# PIE DE PÁGINA
# ==========================================

st.markdown("---")
st.caption(
    "PASTELPAN ERP | Una tentación irresistible"
)


