
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(
    page_title="PASTELPAN ERP",
    page_icon="🎂",
    layout="wide"
)

# ==========================================
# PALETA DE MARCA
# ==========================================

ROJO = "#C83E3E"
VERDE = "#A9D18E"
CORAL = "#E88F91"
BEIGE = "#F3E8D5"

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
    font-weight: bold;
    min-height: 42px;
}

div.stButton > button:hover {
    background-color: #A82F35;
    color: white;
    border: none;
}

</style>
""", unsafe_allow_html=True)


# ==========================================
# CONEXIÓN CON EXCEL
# ==========================================

def obtener_excel():

    contenido = st.session_state.get("excel_bytes")

    if contenido is None:
        return None

    return pd.ExcelFile(
        BytesIO(contenido),
        engine="openpyxl"
    )


def leer_hoja(nombre, columnas=None):

    excel = obtener_excel()

    if excel is None or nombre not in excel.sheet_names:
        return pd.DataFrame()

    try:
        return pd.read_excel(
            excel,
            sheet_name=nombre,
            header=2,
            usecols=columnas
        )

    except Exception:
        return pd.DataFrame()


# ==========================================
# FUNCIÓN GENERAL PARA GRÁFICOS
# ==========================================

def mostrar_grafico_registros(datos, titulo):

    if datos is None or datos.empty:
        st.info("No hay información suficiente para generar el gráfico.")
        return

    datos = datos.copy()

    columnas_numericas = datos.select_dtypes(
        include="number"
    ).columns.tolist()

    columnas_categoricas = [
        columna for columna in datos.columns
        if columna not in columnas_numericas
        and 1 < datos[columna].nunique(dropna=True) <= 30
    ]

    st.markdown(f"#### {titulo}")

    if columnas_categoricas and columnas_numericas:

        categoria = columnas_categoricas[0]
        valor = columnas_numericas[0]

        grafico = (
            datos.groupby(categoria)[valor]
            .sum()
            .sort_values(ascending=False)
        )

        st.bar_chart(grafico)

    elif columnas_categoricas:

        categoria = columnas_categoricas[0]

        grafico = datos[categoria].value_counts()

        st.bar_chart(grafico)

    elif columnas_numericas:

        st.bar_chart(datos[columnas_numericas])

    else:

        st.info(
            "No se encontraron columnas adecuadas para representar gráficamente."
        )


# ==========================================
# NAVEGACIÓN
# ==========================================

MODULOS = [
    "Inicio",
    "Información",
    "CRM",
    "Producción",
    "Inventarios",
    "Compras y Proveedores",
    "Dashboard"
]


if "seccion_menu" not in st.session_state:
    st.session_state["seccion_menu"] = "Inicio"


def cambiar_seccion(nombre):
    st.session_state["seccion_menu"] = nombre


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


# ==========================================
# MENÚ LATERAL
# ==========================================

with st.sidebar:

    seccion = st.radio(
        "MENÚ PRINCIPAL",
        MODULOS,
        key="seccion_menu"
    )

    st.divider()

    st.caption("PASTELPAN ERP | Sistema de gestión")


# ==========================================
# BOTÓN VOLVER AL MENÚ PRINCIPAL
# ==========================================

if seccion != "Inicio":

    st.button(
        "← Volver al menú principal",
        use_container_width=True,
        on_click=cambiar_seccion,
        args=("Inicio",)
    )


# ==========================================
# INICIO
# ==========================================

if seccion == "Inicio":

    st.title("🎂 PASTELPAN ERP")
    st.subheader("Una tentación irresistible")

    st.write(
        "Sistema de gestión integral de PASTELPAN."
    )

    if excel is None:

        st.warning(
            "Carga el Excel maestro desde el menú lateral."
        )

    else:

        st.success(
            "Archivo maestro conectado correctamente."
        )

        st.metric(
            "Hojas disponibles",
            len(excel.sheet_names)
        )

    st.divider()

    st.subheader("Módulos del sistema")

    columnas = st.columns(3)

    modulos_inicio = [
        (
            "🏢 Información",
            "Información",
            "Datos generales de la empresa."
        ),
        (
            "👥 CRM",
            "CRM",
            "Gestión de clientes y atención."
        ),
        (
            "🏭 Producción",
            "Producción",
            "Planificación y control productivo."
        ),
        (
            "📦 Inventarios",
            "Inventarios",
            "Ingredientes y materiales."
        ),
        (
            "🛒 Compras",
            "Compras y Proveedores",
            "Proveedores y órdenes de compra."
        ),
        (
            "📊 Dashboard",
            "Dashboard",
            "Indicadores de gestión."
        )
    ]

    for i, (titulo, destino, descripcion) in enumerate(modulos_inicio):

        with columnas[i % 3]:

            with st.container(border=True):

                st.markdown(f"### {titulo}")

                st.write(descripcion)

                st.button(
                    "Ingresar →",
                    key=f"acceso_{destino}",
                    use_container_width=True,
                    on_click=cambiar_seccion,
                    args=(destino,)
                )


# ==========================================
# INFORMACIÓN
# ==========================================

elif seccion == "Información":

    st.title("🏢 Información de PASTELPAN")
    st.caption("Una tentación irresistible")

    st.markdown("""
    <div style="
        background: linear-gradient(110deg, #F3E8D5, #FCE8E5);
        padding: 25px;
        border-radius: 16px;
        text-align: center;
        margin-bottom: 20px;">
        <h1 style="color:#8B5145;">PASTELPAN</h1>
        <h4 style="color:#A65D54;">
            Empresa del sector pastelería
        </h4>
        <p style="color:#665047;">
            Elaboración y comercialización de tortas
            para diferentes ocasiones.
        </p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "🏢 Datos de la empresa",
        "🎂 Productos principales",
        "⚙️ Proceso de fabricación"
    ])

    # --------------------------------------
    # DATOS GENERALES
    # --------------------------------------

    with tab1:

        st.subheader("Datos generales de la empresa")

        datos_empresa = [
            ("Nombre comercial", "PASTELPAN"),
            ("Ubicación", "Jr. Leguía 151, Cajamarca"),
            ("Sector", "Pastelería"),
            (
                "Descripción",
                "Empresa dedicada a la elaboración y comercialización "
                "de productos de pastelería, con énfasis en tortas "
                "para celebraciones y ocasiones especiales."
            ),
            (
                "Productos principales",
                "Torta de vainilla y torta de chocolate"
            )
        ]

        for campo, valor in datos_empresa:

            with st.container(border=True):

                st.markdown(f"**{campo}**")
                st.write(valor)

        st.divider()

        st.subheader("📍 Sede principal de PASTELPAN")

        st.write(
            "Conoce la ubicación de nuestra sede principal en Cajamarca."
        )

        st.components.v1.html(
            """
            <iframe
                src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d1176.948056499078!2d-78.52059577910234!3d-7.153077470205614!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x91b25bdac5d30ed7%3A0x3a49f2db85b5efaa!2sPastelPan!5e0!3m2!1ses-419!2spe!4v1795561087245!5m2!1ses-419!2spe"
                width="100%"
                height="400"
                style="border:0; border-radius:15px;"
                allowfullscreen=""
                loading="lazy"
                referrerpolicy="strict-origin-when-cross-origin">
            </iframe>
            """,
            height=420
        )

    # --------------------------------------
    # PRODUCTOS PRINCIPALES
    # --------------------------------------

    with tab2:

        st.subheader("Características de los productos")

        col1, col2 = st.columns(2)

        with col1:

            st.markdown("""
            <div style="
                background:#FFF4E8;
                padding:20px;
                border-radius:15px;
                border:1px solid #EBD7C0;">
                <h3 style="color:#9B684D; text-align:center;">
                    🍰 Torta de vainilla
                </h3>
                <p style="color:#59483D; text-align:justify;">
                    Producto de pastelería elaborado principalmente
                    con harina, azúcar, huevos, mantequilla o aceite
                    y esencia de vainilla.
                </p>
                <p style="color:#59483D;">
                    <b>Características:</b>
                </p>
                <ul style="color:#59483D;">
                    <li>Sabor suave y característico a vainilla.</li>
                    <li>Textura esponjosa.</li>
                    <li>Puede incorporar diferentes rellenos.</li>
                    <li>Permite diversas decoraciones.</li>
                    <li>Se destina a celebraciones y reuniones.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

        with col2:

            st.markdown("""
            <div style="
                background:#FCEAEA;
                padding:20px;
                border-radius:15px;
                border:1px solid #EBCACA;">
                <h3 style="color:#A34D4D; text-align:center;">
                    🍫 Torta de chocolate
                </h3>
                <p style="color:#59483D; text-align:justify;">
                    Producto de pastelería elaborado principalmente
                    con harina, huevos, azúcar, cocoa o chocolate
                    y los demás ingredientes de la preparación.
                </p>
                <p style="color:#59483D;">
                    <b>Características:</b>
                </p>
                <ul style="color:#59483D;">
                    <li>Sabor característico a chocolate.</li>
                    <li>Textura suave y esponjosa.</li>
                    <li>Puede incorporar rellenos y coberturas.</li>
                    <li>Permite acabados personalizados.</li>
                    <li>Se destina a celebraciones y reuniones.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)

        st.caption(
            "Las características descritas corresponden a la "
            "presentación general de ambos productos."
        )

    # --------------------------------------
    # PROCESO DE FABRICACIÓN
    # --------------------------------------

    with tab3:

        st.subheader("Proceso de fabricación de las tortas")

        st.write("""
        La elaboración de las tortas de vainilla y chocolate
        comprende una secuencia de operaciones que permiten
        transformar los ingredientes en productos terminados.
        """)

        etapas = [
            (
                "01",
                "Selección y pesado de ingredientes",
                "Se seleccionan los insumos y se pesan de acuerdo "
                "con la formulación correspondiente."
            ),
            (
                "02",
                "Mezclado y preparación de la masa",
                "Se integran los ingredientes hasta obtener "
                "una mezcla homogénea."
            ),
            (
                "03",
                "Preparación de moldes",
                "Se acondicionan los moldes para recibir "
                "la mezcla preparada."
            ),
            (
                "04",
                "Horneado",
                "La mezcla se somete al proceso de cocción "
                "para obtener la estructura de la torta."
            ),
            (
                "05",
                "Enfriamiento",
                "Se deja enfriar el producto antes de continuar "
                "con el armado."
            ),
            (
                "06",
                "Rellenado y armado",
                "Se incorporan los rellenos y se ensamblan "
                "las capas de la torta."
            ),
            (
                "07",
                "Decoración y acabado",
                "Se aplican las coberturas y decoraciones "
                "según la presentación del producto."
            ),
            (
                "08",
                "Control de calidad y empaquetado",
                "Se verifica la presentación y condición del producto "
                "para su acondicionamiento y comercialización."
            )
        ]

        for numero, titulo, descripcion in etapas:

            st.markdown(f"""
            <div style="
                background:#F8F3EC;
                padding:15px;
                border-radius:12px;
                margin-bottom:10px;
                border-left:5px solid #A9D18E;">
                <span style="
                    color:#A65D54;
                    font-weight:bold;">
                    ETAPA {numero}
                </span>
                <h4 style="color:#6C594B; margin:5px 0;">
                    {titulo}
                </h4>
                <p style="color:#66584C; margin-bottom:0;">
                    {descripcion}
                </p>
            </div>
            """, unsafe_allow_html=True)

        st.divider()

        st.subheader("Diagrama del proceso de fabricación")

        st.graphviz_chart("""
        digraph Proceso {

            rankdir=LR;

            node [
                shape=box,
                style="rounded,filled",
                fillcolor="#F3E8D5",
                color="#A9D18E",
                fontcolor="#59483D"
            ];

            A [label="Selección y\\npesado"];
            B [label="Mezclado y\\npreparación"];
            C [label="Preparación\\nde moldes"];
            D [label="Horneado"];
            E [label="Enfriamiento"];
            F [label="Rellenado y\\narmado"];
            G [label="Decoración"];
            H [label="Control de calidad\\ny empaquetado"];

            A -> B -> C -> D -> E -> F -> G -> H;
        }
        """, use_container_width=True)

        st.success(
            "Resultado: tortas de vainilla y chocolate "
            "terminadas y listas para su comercialización."
        )


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

            clientes["Fecha de registro"] = pd.to_datetime(
                clientes["Fecha de registro"],
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

            ventas["Cantidad (unid.)"] = pd.to_numeric(
                ventas["Cantidad (unid.)"],
                errors="coerce"
            )

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

                c1.metric(
                    "Clientes registrados",
                    len(clientes)
                )

                c2.metric(
                    "Clientes con venta",
                    (clientes["Estado"] == "Venta").sum()
                )

                c3.metric(
                    "Clientes pendientes",
                    (clientes["Estado"] == "Pendiente").sum()
                )

                st.divider()

                st.subheader("Distribución de clientes por estado")

                grafico_estados = (
                    clientes["Estado"]
                    .dropna()
                    .value_counts()
                )

                st.bar_chart(grafico_estados)

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

                st.metric(
                    "Resultados encontrados",
                    len(resultado)
                )

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
            # ATENCIÓN AL CLIENTE Y VENTAS
            # ----------------------------------

            with tab2:

                st.subheader("Registro de atención al cliente y ventas")

                v1, v2, v3 = st.columns(3)

                v1.metric(
                    "Registros de venta",
                    len(ventas)
                )

                unidades = ventas["Cantidad (unid.)"].sum()

                v2.metric(
                    "Unidades vendidas",
                    f"{unidades:,.0f}"
                )

                promedio = ventas["Calificación NPS (0-5)"].mean()

                v3.metric(
                    "Calificación promedio (0-5)",
                    f"{promedio:.2f}" if pd.notna(promedio) else "Sin datos"
                )

                st.divider()

                st.subheader("Ventas por producto")

                grafico_productos = (
                    ventas.groupby("Producto")["Cantidad (unid.)"]
                    .sum()
                )

                st.bar_chart(grafico_productos)

                st.divider()

                st.markdown("#### Filtros de ventas")

                fechas = ventas["Fecha de venta"].dropna()

                if not fechas.empty:

                    minimo = fechas.min().date()
                    maximo = fechas.max().date()

                    periodo = st.date_input(
                        "Periodo de consulta",
                        value=(minimo, maximo),
                        min_value=minimo,
                        max_value=maximo
                    )

                    if isinstance(periodo, tuple) and len(periodo) == 2:

                        fecha_inicio, fecha_fin = periodo

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
                                fecha_inicio,
                                fecha_fin
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
# PRODUCCIÓN
# ==========================================

elif seccion == "Producción":

    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries

    st.title("🏭 Producción")

    st.markdown(
        "Consulta las tablas de planificación y control de producción de PASTELPAN."
    )

    archivo = st.session_state.get("excel_bytes")

    if not archivo:

        st.warning(
            "Primero debes cargar el archivo Excel de PASTELPAN desde el menú lateral."
        )

    else:

        try:

            wb = load_workbook(
                BytesIO(archivo),
                data_only=True
            )

            nombre_hoja = "PRODUCCION (RESUMEN)"

            if nombre_hoja not in wb.sheetnames:

                st.error(
                    f"No se encontró la hoja: {nombre_hoja}"
                )

            else:

                ws = wb[nombre_hoja]

                cuadros = {
                    "Demanda y pronóstico": [
                        ("Histórico de ventas - Vainilla", "B2:D26", 1),
                        ("Pronóstico - Vainilla", "F2:I14", 1),
                        ("Histórico de ventas - Chocolate", "B30:D54", 1),
                        ("Pronóstico - Chocolate", "F30:I42", 1)
                    ],

                    "Plan agregado (PAP)": [
                        ("Plan agregado de producción", "B58:F66", 1)
                    ],

                    "Programa maestro (PMP)": [
                        ("PMP - Torta de vainilla", "B70:Y78", 1),
                        ("PMP - Torta de chocolate", "B82:Y90", 1)
                    ],

                    "Maestro de materiales": [
                        ("Maestro de materiales", "B95:L115", 1)
                    ],

                    "Estructura de materiales (BOM)": [
                        ("BOM - Torta de vainilla", "B118:H134", 1),
                        ("BOM - Torta de chocolate", "B137:H149", 1)
                    ],

                    "TLO": [
                        ("Tabla de lanzamiento de órdenes", "B152:K172", 1)
                    ],

                    "BOC": [
                        ("Balance de operaciones y capacidades", "B176:K210", 2)
                    ],

                    "CRP": [
                        ("Planificación de requerimientos de capacidad", "B212:K241", 2)
                    ],

                    "Distribución (DRP)": [
                        ("DRP - Torta de vainilla", "B244:K251", 1),
                        ("DRP - Torta de chocolate", "B253:K260", 1)
                    ]
                }

                def aplicar_formato(data):

                    estilos = pd.DataFrame(
                        "",
                        index=data.index,
                        columns=data.columns
                    )

                    if len(data.columns) > 0:

                        estilos.iloc[:, 0] = (
                            "font-weight: bold; "
                            "background-color: #F5EBDD; "
                            "color: #51483F;"
                        )

                    return estilos

                def mostrar_cuadro(titulo, rango, filas_encabezado):

                    min_col, min_row, max_col, max_row = range_boundaries(rango)

                    datos = []

                    for fila in ws.iter_rows(
                        min_row=min_row,
                        max_row=max_row,
                        min_col=min_col,
                        max_col=max_col,
                        values_only=True
                    ):

                        datos.append(list(fila))

                    if not datos:
                        return

                    if filas_encabezado == 2 and len(datos) >= 2:

                        encabezado_superior = datos[0]
                        encabezado_inferior = datos[1]

                        encabezados = []

                        for superior, inferior in zip(
                            encabezado_superior,
                            encabezado_inferior
                        ):

                            if inferior is not None and str(inferior).strip():

                                encabezados.append(str(inferior).strip())

                            elif superior is not None and str(superior).strip():

                                encabezados.append(str(superior).strip())

                            else:

                                encabezados.append("")

                        filas = datos[2:]

                    else:

                        encabezados = [
                            str(valor).strip() if valor is not None else ""
                            for valor in datos[0]
                        ]

                        filas = datos[1:]

                    encabezados_finales = []
                    contador = {}

                    for i, encabezado in enumerate(encabezados):

                        nombre = encabezado if encabezado else f"Columna {i + 1}"

                        contador[nombre] = contador.get(nombre, 0) + 1

                        if contador[nombre] > 1:
                            nombre = f"{nombre} ({contador[nombre]})"

                        encabezados_finales.append(nombre)

                    df = pd.DataFrame(
                        filas,
                        columns=encabezados_finales
                    )

                    df = df.fillna("")

                    for columna in df.columns:

                        df[columna] = df[columna].apply(
                            lambda valor: round(valor, 1)
                            if isinstance(valor, (int, float))
                            and not isinstance(valor, bool)
                            else valor
                        )

                    df = df.loc[
                        ~df.apply(
                            lambda fila: all(
                                str(valor).strip() == ""
                                for valor in fila
                            ),
                            axis=1
                        )
                    ]

                    st.markdown(f"**{titulo}**")

                    tabla = (
                        df.style
                        .format(precision=1, na_rep="")
                        .apply(aplicar_formato, axis=None)
                        .set_table_styles([
                            {
                                "selector": "th",
                                "props": [
                                    ("background-color", "#A9D18E"),
                                    ("color", "#34452F"),
                                    ("font-weight", "bold"),
                                    ("text-align", "center"),
                                    ("border", "1px solid #D8E5D0"),
                                    ("padding", "9px")
                                ]
                            },
                            {
                                "selector": "td",
                                "props": [
                                    ("border", "1px solid #E8DED0"),
                                    ("padding", "8px"),
                                    ("color", "#51483F")
                                ]
                            }
                        ])
                    )

                    st.dataframe(
                        tabla,
                        use_container_width=True,
                        hide_index=True,
                        height=400
                    )

                    # GRÁFICO DE LA TABLA

                    columnas_numericas = df.select_dtypes(
                        include="number"
                    ).columns.tolist()

                    if columnas_numericas and len(df) > 0:

                        st.markdown("**Representación gráfica de los datos**")

                        grafico_df = df.copy()

                        if len(grafico_df.columns) > 0:

                            columna_categoria = grafico_df.columns[0]

                            grafico_df = grafico_df.set_index(
                                columna_categoria
                            )

                            grafico_df = grafico_df[
                                columnas_numericas
                            ]

                            st.bar_chart(
                                grafico_df
                            )

                    st.write("")

                categorias = list(cuadros.keys())

                categoria_seleccionada = st.selectbox(
                    "Selecciona una sección de producción:",
                    categorias
                )

                st.divider()

                for titulo, rango, filas_encabezado in cuadros[categoria_seleccionada]:

                    mostrar_cuadro(
                        titulo,
                        rango,
                        filas_encabezado
                    )

        except Exception as error:

            st.error(
                f"No se pudo cargar la información de Producción: {error}"
            )


# ==========================================
# INVENTARIOS
# ==========================================

elif seccion == "Inventarios":

    st.title("📦 Gestión de Inventarios")

    if excel is not None:

        datos = leer_hoja("Maestro de Materiales")

        if datos.empty:

            st.warning("No se encontraron datos en el Maestro de Materiales.")

        else:

            st.subheader("Maestro de materiales")

            st.metric(
                "Materiales registrados",
                len(datos)
            )

            st.dataframe(
                datos,
                use_container_width=True,
                hide_index=True
            )

            st.divider()

            mostrar_grafico_registros(
                datos,
                "Representación gráfica de materiales"
            )

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
            [
                "Proveedores",
                "Órdenes de Compra y Servicio"
            ]
        )

        datos = leer_hoja(opcion)

        if datos.empty:

            st.warning("No se encontraron registros en esta sección.")

        else:

            st.metric(
                "Registros encontrados",
                len(datos)
            )

            st.dataframe(
                datos,
                use_container_width=True,
                hide_index=True
            )

            st.divider()

            if opcion == "Proveedores":

                mostrar_grafico_registros(
                    datos,
                    "Distribución de registros de proveedores"
                )

            else:

                mostrar_grafico_registros(
                    datos,
                    "Representación gráfica de órdenes de compra y servicio"
                )

    else:

        st.warning("Carga el Excel maestro.")


# ==========================================
# DASHBOARD
# ==========================================

elif seccion == "Dashboard":

    st.title("📊 Dashboard de PASTELPAN")
    st.caption("Resumen general de gestión empresarial")

    if excel is None:

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
