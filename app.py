
import streamlit as st

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

# Colores de la marca
VERDE = "#A9D18E"
ROJO = "#C83E3E"
CORAL = "#E88F91"
BEIGE = "#F3E8D5"

# Diseño general
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
        border-radius: 10px;
        border: none;
        padding: 0.5rem 1rem;
    }
    div.stButton > button:hover {
        background-color: #A9D18E;
        color: #333333;
    }
</style>
""", unsafe_allow_html=True)


# Menú principal
MODULOS = [
    "Inicio",
    "Información",
    "CRM",
    "Producción",
    "Inventarios",
    "Compras y Proveedores",
    "Dashboard"
]


def volver_inicio():
    st.session_state["seccion"] = "Inicio"


# Navegación
st.sidebar.title("🎂 PASTELPAN")
st.sidebar.caption("Una tentación irresistible")
st.sidebar.markdown("---")

seccion = st.sidebar.radio(
    "MENÚ PRINCIPAL",
    MODULOS,
    key="seccion"
)

st.sidebar.markdown("---")
st.sidebar.caption("Sistema de gestión empresarial")


# Contenido de los módulos
if seccion == "Inicio":
    st.title("🎂 Bienvenido a PASTELPAN ERP")
    st.subheader("Una tentación irresistible")
    st.write(
        "Sistema de gestión integral para administrar "
        "las operaciones de nuestra empresa."
    )

    st.markdown("---")
    st.subheader("Módulos del sistema")

    columnas = st.columns(3)

    descripciones = [
        ("🏢 Información", "Conoce nuestra empresa."),
        ("👥 CRM", "Gestiona clientes y atención."),
        ("🏭 Producción", "Planifica y controla la producción."),
        ("📦 Inventarios", "Consulta y registra materiales."),
        ("🛒 Compras y Proveedores", "Gestiona compras y proveedores."),
        ("📊 Dashboard", "Visualiza los indicadores del negocio.")
    ]

    for i, (nombre, descripcion) in enumerate(descripciones):
        with columnas[i % 3]:
            with st.container(border=True):
                st.markdown(f"### {nombre}")
                st.write(descripcion)

elif seccion == "Información":
    st.title("🏢 Información de PASTELPAN")
    st.info("Aquí incorporaremos la presentación de la empresa, sus productos y procesos.")

elif seccion == "CRM":
    st.title("👥 Gestión de clientes (CRM)")
    st.info("Aquí desarrollaremos el registro de clientes, seguimiento comercial y atención al cliente.")

elif seccion == "Producción":
    st.title("🏭 Gestión de Producción")
    st.info("Aquí integraremos los pronósticos, PAP, PMP, BOM, MRP, MRP II, BOC, CRP y DRP.")

elif seccion == "Inventarios":
    st.title("📦 Gestión de Inventarios")
    st.info("Aquí desarrollaremos ingredientes y materiales, consulta de stock y registro de materiales.")

elif seccion == "Compras y Proveedores":
    st.title("🛒 Compras y Proveedores")
    st.info("Aquí integraremos las órdenes de compra y servicio y el registro de proveedores.")

elif seccion == "Dashboard":
    st.title("📊 Dashboard de Indicadores")
    st.info("Aquí construiremos los indicadores y gráficos de gestión de PASTELPAN.")


# Botón para regresar al inicio
if seccion != "Inicio":
    st.markdown("---")
    st.button(
        "🏠 Volver al menú principal",
        on_click=volver_inicio
    )
