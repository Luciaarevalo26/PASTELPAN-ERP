
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
        st.markdown("""
            <div style="
                background: linear-gradient(110deg, #F3E8D5, #FCE8E5);
                padding: 28px;
                border-radius: 18px;
                text-align: center;
                margin-bottom: 22px;">
                <h1 style="color:#8B5145; margin-bottom:5px;">
                    PASTELPAN
                </h1>
                <h4 style="color:#A65D54; margin-top:0;">
                    Una tentación irresistible
                </h4>
                <p style="color:#665047;">
                    Tradición, dedicación y sabor en cada preparación.
                </p>
            </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4 = st.tabs([
            "🏠 Conócenos",
            "🎂 Nuestros productos",
            "👩‍🍳 Nuestra planta",
            "📋 Proceso productivo"
        ])

        # -----------------------------------------
        # PESTAÑA 1: CONÓCENOS
        # -----------------------------------------
        with tab1:

            st.subheader("¿Quiénes somos?")

            st.write("""
            **PASTELPAN** es una empresa familiar dedicada a la
            elaboración de productos de panadería, pastelería y
            confitería. Su desarrollo comenzó como un emprendimiento
            impulsado por su fundador, quien inició sus actividades
            con un horno industrial.

            Con el crecimiento de sus clientes, la empresa amplió
            su oferta incorporando tortas personalizadas, bocaditos
            dulces y salados, y postres individuales. Asimismo,
            mejoró sus instalaciones e incorporó maquinaria y
            personal para fortalecer sus operaciones.
            """)

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("""
                <div style="background:#F3E8D5; padding:20px;
                border-radius:14px; min-height:145px;">
                    <h4 style="color:#8B5145;">📍 Ubicación</h4>
                    <p style="color:#59483D;">
                    Jr. Leguía 151, Cajamarca.
                    </p>
                </div>
                """, unsafe_allow_html=True)

            with col2:
                st.markdown("""
                <div style="background:#E8F0DF; padding:20px;
                border-radius:14px; min-height:145px;">
                    <h4 style="color:#557746;">🏢 Actividad empresarial</h4>
                    <p style="color:#43513B;">
                    Elaboración y comercialización de productos
                    de panadería, pastelería y confitería.
                    </p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("### Nuestro compromiso")

            st.info(
                "Ofrecer productos de calidad, procurando la "
                "satisfacción de nuestros clientes y fortaleciendo "
                "nuestra presencia en el mercado cajamarquino."
            )

        # -----------------------------------------
        # PESTAÑA 2: NUESTROS PRODUCTOS
        # -----------------------------------------
        with tab2:

            st.subheader("Nuestros productos principales")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("""
                <div style="background:#FFF4E8; padding:22px;
                border-radius:16px; border:1px solid #EBD7C0;">
                    <div style="font-size:48px; text-align:center;">
                        🍰
                    </div>
                    <h3 style="color:#9B684D; text-align:center;">
                        Torta de vainilla
                    </h3>
                    <p style="color:#59483D; text-align:justify;">
                        Preparada principalmente con harina, azúcar,
                        huevos, mantequilla o aceite y esencia de
                        vainilla. Se caracteriza por su sabor suave
                        y puede combinarse con diferentes rellenos,
                        cremas y decoraciones.
                    </p>
                    <p style="color:#9B684D; text-align:center;">
                        Ideal para celebraciones y reuniones.
                    </p>
                </div>
                """, unsafe_allow_html=True)

            with col2:
                st.markdown("""
                <div style="background:#FCEAEA; padding:22px;
                border-radius:16px; border:1px solid #EBCACA;">
                    <div style="font-size:48px; text-align:center;">
                        🍫
                    </div>
                    <h3 style="color:#A34D4D; text-align:center;">
                        Torta de chocolate
                    </h3>
                    <p style="color:#59483D; text-align:justify;">
                        Elaborada principalmente con harina, huevos,
                        azúcar, cocoa o chocolate y otros ingredientes.
                        Es una alternativa para cumpleaños,
                        aniversarios y reuniones familiares.
                    </p>
                    <p style="color:#A34D4D; text-align:center;">
                        Personalizable según la preferencia del cliente.
                    </p>
                </div>
                """, unsafe_allow_html=True)

            st.caption(
                "Los productos descritos corresponden a las "
                "dos líneas principales consideradas en el proyecto."
            )

        # -----------------------------------------
        # PESTAÑA 3: NUESTRA PLANTA
        # -----------------------------------------
        with tab3:

            st.subheader("Organización de las operaciones")

            st.write("""
            Para desarrollar las actividades productivas, PASTELPAN
            considera cuatro estaciones de trabajo, atendidas por
            cuatro operarios. La asignación de estaciones organiza
            las actividades, aunque un mismo operario puede participar
            en diferentes operaciones según las necesidades del proceso.
            """)

            estaciones = [
                ("01", "Preparación y pesado",
                 "Selección y pesaje de los ingredientes."),
                ("02", "Mezclado y preparación de la masa",
                 "Integración de los ingredientes para obtener la masa."),
                ("03", "Horneado y enfriamiento",
                 "Cocción del producto y posterior enfriamiento."),
                ("04", "Rellenado y decoración",
                 "Armado, acabado y presentación de las tortas.")
            ]

            for numero, nombre, descripcion in estaciones:
                st.markdown(f"""
                <div style="background:#F7F1E8; padding:15px;
                border-radius:12px; margin-bottom:10px;
                border-left:5px solid #A9D18E;">
                    <span style="color:#A65D54; font-weight:bold;">
                        ESTACIÓN {numero}
                    </span>
                    <h4 style="color:#6C594B; margin:5px 0;">
                        {nombre}
                    </h4>
                    <p style="color:#66584C; margin-bottom:0;">
                        {descripcion}
                    </p>
                </div>
                """, unsafe_allow_html=True)

            st.markdown("### Actividades antes de iniciar la producción")

            actividades = [
                "Registro de llegada y asignación de actividades.",
                "Lavado de manos y colocación de indumentaria de trabajo.",
                "Limpieza y desinfección de las estaciones.",
                "Revisión de ingredientes, materiales y equipos.",
                "Verificación de los pedidos y preparación de las recetas.",
                "Pesado de ingredientes y precalentamiento del horno."
            ]

            for actividad in actividades:
                st.markdown(f"- {actividad}")

        # -----------------------------------------
        # PESTAÑA 4: PROCESO PRODUCTIVO
        # -----------------------------------------
        with tab4:

            st.subheader("Elaboración de nuestras tortas")

            st.write("""
            El proceso productivo comprende las operaciones necesarias
            para transformar los ingredientes en un producto terminado,
            listo para su comercialización.
            """)

            pasos = [
                ("1. Selección y pesado",
                 "Se seleccionan y pesan los ingredientes necesarios."),
                ("2. Mezclado",
                 "Se integran los ingredientes para preparar la masa."),
                ("3. Preparación de moldes",
                 "Se acondicionan los moldes para recibir la mezcla."),
                ("4. Horneado",
                 "Se realiza la cocción de la preparación."),
                ("5. Enfriamiento",
                 "Se deja enfriar el producto después del horneado."),
                ("6. Rellenado y armado",
                 "Se incorpora el relleno y se ensamblan las partes."),
                ("7. Decoración y acabado",
                 "Se realizan los acabados y la presentación final."),
                ("8. Control de calidad y empaquetado",
                 "Se verifica el producto y se prepara para su venta.")
            ]

            for titulo, descripcion in pasos:
                st.markdown(f"""
                <div style="display:flex; gap:12px;
                background:#F8F3EC; padding:13px;
                border-radius:11px; margin-bottom:9px;">
                    <div style="min-width:100%;">
                        <strong style="color:#A65D54;">
                            {titulo}
                        </strong>
                        <p style="color:#66584C; margin:4px 0 0 0;">
                            {descripcion}
                        </p>
                    </div>
                </div>
                """, unsafe_allow_html=True)

            st.success(
                "Resultado: tortas elaboradas, acondicionadas y "
                "listas para su comercialización."
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

    import pandas as pd
    from io import BytesIO
    from openpyxl import load_workbook
    from openpyxl.utils import range_boundaries, get_column_letter

    st.title("🏭 Gestión de Producción")
    st.caption("Planificación y control productivo de PASTELPAN")

    archivo = st.session_state.get("excel_bytes")

    if not archivo:
        st.warning("Primero debes cargar el archivo Excel de PASTELPAN.")

    else:

        # Abrir el Excel original sin modificarlo
        wb = load_workbook(
            BytesIO(archivo),
            data_only=True
        )

        # Cuadros y rangos exactos indicados del proyecto
        cuadros_produccion = {

            "Demanda y pronóstico": [
                (
                    "Pronóstico de demanda - Torta de vainilla",
                    "Demanda y Pronóstico Vainilla",
                    "B32:E55"
                ),
                (
                    "Pronóstico de demanda - Torta de chocolate",
                    "Demanda y Pronóstico Chocolate",
                    "J31:O32"
                )
            ],

            "Plan agregado (PAP)": [
                (
                    "Plan agregado de producción",
                    "PAP ",
                    "A99:E108"
                )
            ],

            "Programa maestro (PMP)": [
                (
                    "PMP - Torta de vainilla",
                    "PMP Vainilla",
                    "A11:Y19"
                ),
                (
                    "PMP - Torta de chocolate",
                    "PMP Chocolate",
                    "A11:Y19"
                )
            ],

            "Maestro de materiales": [
                (
                    "Maestro de materiales",
                    "Maestro de Materiales",
                    "B4:L24"
                )
            ],

            "Estructura de materiales (BOM)": [
                (
                    "BOM - Torta de vainilla",
                    "BOM",
                    "D10:J26"
                ),
                (
                    "BOM - Torta de chocolate",
                    "BOM",
                    "Q10:W22"
                )
            ],

            "TLO": [
                (
                    "TLO - Tabla de planificación",
                    "MRP II FINAL",
                    "B240:K263"
                )
            ],

            "BOC": [
                (
                    "BOC - Balance de operaciones",
                    "MRP II FINAL",
                    "Q4:Z38"
                )
            ],

            "CRP": [
                (
                    "CRP - Planificación de capacidad",
                    "MRP II FINAL",
                    "Q46:Z75"
                )
            ],

            "Distribución (DRP)": [
                (
                    "DRP - Torta de vainilla",
                    "DRP",
                    "K95:T103"
                ),
                (
                    "DRP - Torta de chocolate",
                    "DRP",
                    "V95:AE103"
                )
            ]
        }

        # Selector de secciones
        seccion_prod = st.selectbox(
            "Selecciona el proceso productivo:",
            list(cuadros_produccion.keys())
        )

        st.markdown("---")

        # Mostrar exclusivamente los cuadros seleccionados
        for titulo, nombre_hoja, rango in cuadros_produccion[seccion_prod]:

            st.subheader(titulo)

            try:

                if nombre_hoja not in wb.sheetnames:
                    st.error(
                        f"No se encontró la hoja: {nombre_hoja}"
                    )
                    continue

                ws = wb[nombre_hoja]

                # Obtener límites del rango Excel
                min_col, min_row, max_col, max_row = (
                    range_boundaries(rango)
                )

                # Leer exactamente las celdas indicadas
                datos = []

                for fila in ws.iter_rows(
                    min_row=min_row,
                    max_row=max_row,
                    min_col=min_col,
                    max_col=max_col,
                    values_only=True
                ):
                    datos.append(list(fila))

                # Convertir a tabla
                columnas = [
                    get_column_letter(col)
                    for col in range(min_col, max_col + 1)
                ]

                df = pd.DataFrame(
                    datos,
                    columns=columnas
                )

                if not df.empty and not df.isna().all().all():

                    st.dataframe(
                        df,
                        use_container_width=True,
                        hide_index=True,
                        height=400
                    )

                else:
                    st.info(
                        "No hay valores disponibles en este rango."
                    )

            except Exception as e:

                st.error(
                    f"Error al leer {titulo}: {e}"
                )

            st.markdown("---")

        wb.close()

    st.caption("PASTELPAN ERP | Una tentación irresistible")


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
