import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
import re
import math

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="PASTELPAN ERP",
    page_icon="🎂",
    layout="wide",
    initial_sidebar_state="expanded"
)

ROJO = "#C83E3E"
CORAL = "#E88F91"
BEIGE = "#F3E8D5"
VERDE = "#A9D18E"
NARANJA = "#F4B183"
PALETA_PASTEL = ["#C83E3E", "#F4A261", "#8FBC8F", "#E88F91", "#F6C85F", "#91B7D9", "#B7A1D1"]
ARCHIVO_BASE = "PROYECTO FINAL PASTELPAN GESOOP(2).xlsx"

st.markdown("""
<style>
.stApp {background:#FFFDF9;}
[data-testid="stSidebar"] {background:#F3E8D5;}
h1,h2,h3 {color:#C83E3E;}
div.stButton > button, div.stDownloadButton > button {
    border-radius:9px; font-weight:600; min-height:40px;
}
div.stButton > button {
    background:#C83E3E; color:white; border:0;
}
div.stButton > button:hover {
    background:#A82F35; color:white; border:0;
}
[data-testid="stMetric"] {
    background:#FFFFFF; border:1px solid #F0D8D4;
    padding:14px; border-radius:12px;
}
.kpi-principal {
    background:linear-gradient(135deg,#C83E3E 0%,#E88F91 100%);
    color:white; border-radius:16px; padding:20px 18px;
    min-height:142px; box-shadow:0 5px 16px rgba(145,54,54,.13);
    border:1px solid rgba(255,255,255,.35);
}
.kpi-principal .kpi-nombre {font-size:14px;font-weight:650;opacity:.95;min-height:38px;}
.kpi-principal .kpi-valor {font-size:31px;font-weight:800;line-height:1.2;margin-top:9px;}
.kpi-principal .kpi-unidad {font-size:12px;opacity:.9;margin-top:4px;}
.interpretacion-kpi {
    background:#FFF8F5;border-left:3px solid #E8B8AE;
    border-radius:7px;padding:9px 11px;margin-top:7px;
    color:#6D514B;font-size:12px;min-height:58px;
}
.block-container {padding-top:1.5rem; padding-bottom:2rem;}
</style>
""", unsafe_allow_html=True)


# =========================================================
# CARGA DEL LIBRO Y FUNCIONES AUXILIARES
# =========================================================
def cargar_libro_bytes(contenido):
    return openpyxl.load_workbook(BytesIO(contenido), data_only=True)


def obtener_bytes_iniciales():
    carpeta = Path(__file__).resolve().parent
    ruta = carpeta / ARCHIVO_BASE
    if ruta.exists():
        return ruta.read_bytes()
    # Acepta también el nombre original sin sufijo de carga.
    candidatos = sorted(carpeta.glob("PROYECTO FINAL PASTELPAN GESOOP*.xlsx"))
    if candidatos:
        return candidatos[0].read_bytes()
    return None


if "excel_bytes" not in st.session_state:
    st.session_state["excel_bytes"] = obtener_bytes_iniciales()

if "excel_modificado" not in st.session_state:
    st.session_state["excel_modificado"] = None

if "inventario_editado" not in st.session_state:
    st.session_state["inventario_editado"] = None

if "movimientos_inventario" not in st.session_state:
    st.session_state["movimientos_inventario"] = []


def bytes_actuales():
    return st.session_state.get("excel_modificado") or st.session_state.get("excel_bytes")


def libro_actual(data_only=True):
    contenido = bytes_actuales()
    if not contenido:
        return None
    return openpyxl.load_workbook(BytesIO(contenido), data_only=data_only)


def guardar_libro_en_sesion(libro):
    salida = BytesIO()
    libro.save(salida)
    st.session_state["excel_modificado"] = salida.getvalue()


def hoja_df(nombre, fila_encabezado=1, fila_inicio=None, fila_fin=None,
            col_inicio=1, col_fin=None, data_only=True):
    libro = libro_actual(data_only=data_only)
    if libro is None or nombre not in libro.sheetnames:
        return pd.DataFrame()
    ws = libro[nombre]
    fila_inicio = fila_inicio or fila_encabezado + 1
    fila_fin = fila_fin or ws.max_row
    col_fin = col_fin or ws.max_column
    encabezados = [ws.cell(fila_encabezado, c).value for c in range(col_inicio, col_fin + 1)]
    filas = []
    for r in range(fila_inicio, fila_fin + 1):
        valores = [ws.cell(r, c).value for c in range(col_inicio, col_fin + 1)]
        if any(v is not None for v in valores):
            filas.append(valores)
    if not filas:
        return pd.DataFrame(columns=[str(x) if x is not None else f"Columna {i+1}" for i, x in enumerate(encabezados)])
    nombres = []
    for i, h in enumerate(encabezados):
        nombres.append(str(h).strip() if h is not None else f"Columna {i+1}")
    return pd.DataFrame(filas, columns=nombres)


def nombres_columnas_unicos(encabezados):
    """Genera encabezados no vacíos y únicos para evitar errores de PyArrow/Streamlit."""
    usados = set()
    nombres = []
    for i, encabezado in enumerate(encabezados):
        base = str(encabezado).strip() if encabezado is not None and str(encabezado).strip() else f"Campo {i+1}"
        nombre = base
        consecutivo = 2
        while nombre in usados:
            nombre = f"{base} ({consecutivo})"
            consecutivo += 1
        usados.add(nombre)
        nombres.append(nombre)
    return nombres


def rango_df(nombre, rango, encabezado=True):
    libro = libro_actual(data_only=True)
    if libro is None or nombre not in libro.sheetnames:
        return pd.DataFrame()
    ws = libro[nombre]
    filas = list(ws[rango])
    valores = [[celda.value for celda in fila] for fila in filas]
    if not valores:
        return pd.DataFrame()
    if encabezado:
        headers = nombres_columnas_unicos(valores[0])
        return pd.DataFrame(valores[1:], columns=headers).dropna(how="all")
    return pd.DataFrame(valores).dropna(how="all")


def limpiar_df(df):
    if df.empty:
        return df
    out = df.copy()
    out = out.dropna(axis=1, how="all").dropna(axis=0, how="all")
    out.columns = [str(c).strip() for c in out.columns]
    return out


def numero(valor):
    try:
        if valor is None or pd.isna(valor):
            return 0.0
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def formato_numero(valor):
    """Enteros sin decimales; valores no exactos con dos decimales."""
    try:
        if valor is None or pd.isna(valor):
            return "—"
        if isinstance(valor, (int, float)) and not isinstance(valor, bool):
            n = float(valor)
            if math.isfinite(n):
                return f"{n:.0f}" if abs(n - round(n)) < 1e-9 else f"{n:.2f}"
    except (TypeError, ValueError):
        pass
    return str(valor)


def formato_porcentaje(valor):
    return f"{formato_numero(float(valor) * 100)}%"


def mostrar_tabla(df, altura=420):
    if df is None or df.empty:
        st.info("No hay registros disponibles en este apartado.")
    else:
        # Protección adicional: algunas hojas contienen encabezados repetidos
        # por celdas combinadas o títulos duplicados en el Excel.
        df_vista = df.copy()
        df_vista.columns = nombres_columnas_unicos(df_vista.columns.tolist())
        formato = {col: formato_numero for col in df_vista.columns}
        estilo = (
            df_vista.style
            .format(formato, na_rep="—")
            .set_table_styles([
                {"selector": "th", "props": [
                    ("background-color", "#C83E3E"), ("color", "white"),
                    ("font-weight", "bold"), ("text-align", "center"),
                    ("border-bottom", "2px solid #A82F35")
                ]},
                {"selector": "td:first-child", "props": [
                    ("background-color", "#F3E8D5"), ("color", "#6D514B"),
                    ("font-weight", "bold")
                ]}
            ])
        )
        st.dataframe(estilo, use_container_width=True, hide_index=True, height=altura)


def cambiar_seccion(nombre):
    """Actualiza la navegación desde el callback, antes de crear el radio."""
    st.session_state["seccion_nav"] = nombre


def boton_volver_inicio():
    st.button(
        "← Volver al menú principal",
        key=f"volver_inicio_{st.session_state.get('seccion_nav', 'modulo')}",
        use_container_width=True,
        on_click=cambiar_seccion,
        args=("Inicio",)
    )


def crear_df_pmp(rango):
    # Muestra la matriz completa del PMP, con todas sus filas y columnas.
    return rango_df("PRODUCCION (RESUMEN)", rango, encabezado=True)


def datos_grafico_pmp(df):
    if df is None or df.empty:
        return pd.DataFrame()
    fila_semana = next((i for i in range(len(df))
                        if str(df.iloc[i, 0]).strip().lower() == "semanas"), None)
    fila_pmp = next((i for i in range(len(df))
                     if str(df.iloc[i, 0]).strip().lower() == "pmp"), None)
    if fila_semana is None or fila_pmp is None:
        return pd.DataFrame()
    registros = []
    for j in range(1, len(df.columns)):
        semana = df.iloc[fila_semana, j]
        cantidad = pd.to_numeric(pd.Series([df.iloc[fila_pmp, j]]), errors="coerce").iloc[0]
        if pd.notna(cantidad) and pd.notna(semana):
            registros.append({"Semana": str(semana), "Unidades": float(cantidad)})
    return pd.DataFrame(registros)


def mostrar_tarjetas_costos_pap(df):
    if df is None or df.empty:
        st.info("No se encontraron los costos totales de los planes PAP.")
        return
    fila_total = df[df.iloc[:, 0].astype(str).str.contains("Costo Total", case=False, na=False)]
    if fila_total.empty:
        st.info("No se encontró una fila identificada como Costo Total.")
        return
    registro = fila_total.iloc[0]
    valores = [(str(col), pd.to_numeric(registro[col], errors="coerce"))
               for col in fila_total.columns[1:]]
    valores = [(nombre, valor) for nombre, valor in valores if pd.notna(valor)]
    if not valores:
        st.info("No hay costos numéricos para mostrar.")
        return
    cols = st.columns(min(4, len(valores)))
    for i, (nombre, valor) in enumerate(valores):
        seleccionado = "plan 1" in nombre.lower()
        fondo = "linear-gradient(135deg,#C83E3E,#A82F35)" if seleccionado else "#FFF8F2"
        color_texto = "#FFFFFF" if seleccionado else "#6D514B"
        borde = "#A82F35" if seleccionado else "#E8D5C8"
        etiqueta = "PLAN SELECCIONADO" if seleccionado else "COSTO TOTAL"
        with cols[i % len(cols)]:
            tarjeta = f'''<div style="background:{fondo};border:2px solid {borde};border-radius:13px;
                padding:14px 10px;text-align:center;min-height:112px;box-shadow:0 3px 10px rgba(80,40,30,.08);">
                <div style="font-size:12px;font-weight:700;color:{color_texto};">{nombre}</div>
                <div style="font-size:10px;font-weight:650;color:{color_texto};margin-top:4px;">{etiqueta}</div>
                <div style="font-size:clamp(15px,1.6vw,21px);font-weight:800;line-height:1.2;
                    color:{color_texto};overflow-wrap:anywhere;word-break:break-word;margin-top:8px;">
                    S/ {formato_numero(float(valor))}
                </div></div>'''
            st.markdown(tarjeta, unsafe_allow_html=True)


def grafico_barras(df, x, y, titulo, color=None):
    if df is None or df.empty or x not in df.columns or y not in df.columns:
        st.info("No hay datos suficientes para generar este gráfico.")
        return
    fig = px.bar(df, x=x, y=y, color=color, title=titulo, text_auto=".2s", color_discrete_sequence=PALETA_PASTEL)
    fig.update_layout(
        title_x=0.02, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=15, r=15, t=55, b=15), legend_title_text=""
    )
    st.plotly_chart(fig, use_container_width=True)


def grafico_linea(df, x, y, titulo, color=None):
    if df is None or df.empty or x not in df.columns or y not in df.columns:
        st.info("No hay datos suficientes para generar este gráfico.")
        return
    fig = px.line(df, x=x, y=y, color=color, markers=True, title=titulo, color_discrete_sequence=PALETA_PASTEL)
    fig.update_layout(
        title_x=0.02, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=15, r=15, t=55, b=15), legend_title_text=""
    )
    st.plotly_chart(fig, use_container_width=True)


def filtro_opciones(df, columna, etiqueta):
    if columna not in df.columns:
        return None
    opciones = sorted([str(x) for x in df[columna].dropna().unique()])
    return st.selectbox(etiqueta, ["Todos"] + opciones, key=f"filtro_{etiqueta}")


def aplicar_filtro(df, columna, valor):
    if valor and valor != "Todos" and columna in df.columns:
        return df[df[columna].astype(str) == str(valor)]
    return df


def titulo_seccion(titulo, descripcion=None):
    st.title(titulo)
    if descripcion:
        st.caption(descripcion)


def obtener_indicadores():
    # La hoja INDICADORES ERP es la fuente central de los cuatro KPI.
    df = rango_df("INDICADORES ERP", "A4:G7", encabezado=False)
    if df.empty:
        return {}
    salida = {}
    for _, fila in df.iterrows():
        nombre = str(fila.iloc[0]) if pd.notna(fila.iloc[0]) else ""
        salida[nombre] = {
            "modulo": fila.iloc[1],
            "resultado": numero(fila.iloc[2]),
            "unidad": fila.iloc[3],
            "interpretacion": fila.iloc[4],
            "formula": fila.iloc[5],
            "fuente": fila.iloc[6]
        }
    return salida


# =========================================================
# MENÚ Y CARGA DEL EXCEL
# =========================================================
MODULOS = [
    "Inicio", "Información", "CRM", "Producción",
    "Inventarios", "Compras y Proveedores", "Dashboard"
]
if "seccion_nav" not in st.session_state:
    st.session_state["seccion_nav"] = "Inicio"

with st.sidebar:
    ruta_logo = Path(__file__).resolve().parent / "assets" / "pastelpan_logo.jpeg"
    if ruta_logo.exists():
        st.image(str(ruta_logo), use_container_width=True)
    st.markdown("## 🎂 PASTELPAN ERP")
    seccion = st.radio("MENÚ PRINCIPAL", MODULOS, key="seccion_nav")
    st.divider()
    st.markdown("### Archivo de datos")
    archivo = st.file_uploader("Cargar Excel maestro (.xlsx)", type=["xlsx"])
    if archivo is not None:
        st.session_state["excel_bytes"] = archivo.getvalue()
        st.session_state["excel_modificado"] = None
        st.session_state["inventario_editado"] = None
        st.success("Excel cargado correctamente.")
    if bytes_actuales():
        st.success("Fuente Excel disponible")
        if st.session_state.get("excel_modificado"):
            st.caption("Hay cambios pendientes de descargar.")
    else:
        st.warning("Sube el archivo Excel para iniciar.")

excel = libro_actual(data_only=True)


# =========================================================
# INICIO
# =========================================================
if seccion == "Inicio":
    st.markdown("<h1 style='text-align:center;'>🎂 PASTELPAN</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center;color:#6D514B;'>SISTEMA ERP – GESTIÓN EMPRESARIAL</h3>", unsafe_allow_html=True)
    st.markdown(
        "<p style='text-align:center;font-size:17px;'>Sistema integrado para consultar y gestionar "
        "la información comercial, productiva, logística y de abastecimiento de PASTELPAN.</p>",
        unsafe_allow_html=True
    )
    st.divider()
    tarjetas = [
        ("🏢", "Información", "Datos generales de la empresa y productos."),
        ("🤝", "CRM", "Gestión de clientes potenciales y atención al cliente."),
        ("🧁", "Producción", "Demanda, planificación, materiales y distribución."),
        ("📦", "Inventarios", "Consulta y actualización del maestro de materiales."),
        ("🛒", "Compras y Proveedores", "Directorio de proveedores y órdenes."),
        ("📊", "Dashboard", "Indicadores principales y complementarios.")
    ]
    cols = st.columns(3)
    for i, (icono, nombre, desc) in enumerate(tarjetas):
        with cols[i % 3]:
            st.markdown(
                f"""<div style="background:white;border:1px solid #F0D8D4;border-radius:14px;
                padding:18px 20px 10px;min-height:128px;margin-bottom:8px;">
                <div style="font-size:28px;">{icono}</div>
                <h3 style="margin:5px 0;">{nombre}</h3>
                <p style="color:#655A56;min-height:40px;">{desc}</p></div>""",
                unsafe_allow_html=True
            )
            st.button(
                f"Abrir {nombre} →",
                key=f"abrir_{i}",
                use_container_width=True,
                on_click=cambiar_seccion,
                args=(nombre,)
            )
    if bytes_actuales():
        st.success("El archivo maestro está conectado. Accede a cada módulo con su botón.")
    else:
        st.info("Para comenzar, carga el archivo PROYECTO FINAL PASTELPAN GESOOP.xlsx desde el menú lateral.")


# =========================================================
# INFORMACIÓN
# =========================================================
elif seccion == "Información":
    boton_volver_inicio()
    st.title("🏢 Información de PASTELPAN")
    st.caption("Una tentación irresistible")
    ruta_logo_info = Path(__file__).resolve().parent / "assets" / "pastelpan_logo.jpeg"
    ruta_tortas = Path(__file__).resolve().parent / "assets" / "pastelpan_tortas.jpeg"
    if ruta_logo_info.exists():
        st.image(str(ruta_logo_info), width=250)
    if ruta_tortas.exists():
        st.image(str(ruta_tortas), caption="Variedad de tortas elaboradas por PASTELPAN", use_container_width=True)

    st.markdown("""
    <div style="background:linear-gradient(110deg,#F3E8D5,#FCE8E5);
        padding:25px;border-radius:16px;text-align:center;margin-bottom:20px;">
        <h1 style="color:#8B5145;">PASTELPAN</h1>
        <h4 style="color:#A65D54;">Empresa del sector pastelería</h4>
        <p style="color:#665047;">Elaboración y comercialización de tortas para diferentes ocasiones.</p>
    </div>
    """, unsafe_allow_html=True)

    tab1, tab2, tab3 = st.tabs([
        "🏢 Datos de la empresa",
        "🎂 Productos principales",
        "⚙️ Proceso de fabricación"
    ])

    with tab1:
        st.subheader("Datos generales de la empresa")
        datos_empresa = [
            ("Nombre comercial", "PASTELPAN"),
            ("RUC", "10440919753"),
            ("Ubicación", "Jr. Leguía 151, Cajamarca"),
            ("Sector", "Pastelería"),
            ("Descripción", "Empresa familiar dedicada a la elaboración y comercialización de productos de panadería, pastelería y repostería, con énfasis en tortas para celebraciones y ocasiones especiales."),
            ("Productos principales", "Torta de vainilla y torta de chocolate")
        ]
        for campo, valor in datos_empresa:
            with st.container(border=True):
                st.markdown(f"**{campo}**")
                st.write(valor)

        st.subheader("Historia de PASTELPAN")
        st.write(
            "PASTELPAN nació como un emprendimiento familiar impulsado por su fundador, "
            "quien inició las actividades adquiriendo un horno industrial. En sus primeros "
            "años, el propietario asumía distintas labores, desde la producción y limpieza "
            "hasta la atención al cliente. Con el crecimiento de su clientela, la empresa "
            "amplió su oferta hacia tortas personalizadas, bocaditos dulces y salados, y "
            "postres individuales. También mejoró su infraestructura e incorporó maquinaria "
            "y colaboradores para fortalecer su capacidad productiva. Actualmente, busca "
            "mantener la calidad de sus productos, satisfacer las necesidades de sus clientes "
            "y consolidar su presencia en el mercado cajamarquino."
        )

        st.divider()
        st.subheader("📍 Sede principal de PASTELPAN")
        st.write("Conoce la ubicación de nuestra sede principal en Cajamarca.")
        st.components.v1.html(
            """
            <iframe
                src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d1176.948056499078!2d-78.52059577910234!3d-7.153077470205614!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x91b25bdac5d30ed7%3A0x3a49f2db85b5efaa!2sPastelPan!5e0!3m2!1ses-419!2spe!4v1795561087245!5m2!1ses-419!2spe"
                width="100%" height="400" style="border:0;border-radius:15px;"
                allowfullscreen="" loading="lazy"
                referrerpolicy="strict-origin-when-cross-origin">
            </iframe>
            """,
            height=420
        )

    with tab2:
        st.subheader("Características de los productos")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("""
            <div style="background:#FFF4E8;padding:20px;border-radius:15px;border:1px solid #EBD7C0;">
                <h3 style="color:#9B684D;text-align:center;">🍰 Torta de vainilla</h3>
                <p style="color:#59483D;text-align:justify;">Producto de pastelería elaborado principalmente con harina, azúcar, huevos, mantequilla o aceite y esencia de vainilla.</p>
                <p style="color:#59483D;"><b>Características:</b></p>
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
            <div style="background:#FCEAEA;padding:20px;border-radius:15px;border:1px solid #EBCACA;">
                <h3 style="color:#A34D4D;text-align:center;">🍫 Torta de chocolate</h3>
                <p style="color:#59483D;text-align:justify;">Producto de pastelería elaborado principalmente con harina, huevos, azúcar, cocoa o chocolate y los demás ingredientes de la preparación.</p>
                <p style="color:#59483D;"><b>Características:</b></p>
                <ul style="color:#59483D;">
                    <li>Sabor característico a chocolate.</li>
                    <li>Textura suave y esponjosa.</li>
                    <li>Puede incorporar rellenos y coberturas.</li>
                    <li>Permite acabados personalizados.</li>
                    <li>Se destina a celebraciones y reuniones.</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        st.caption("Las características descritas corresponden a la presentación general de ambos productos.")

    with tab3:
        st.subheader("Proceso de fabricación de las tortas")
        st.write("La elaboración de las tortas de vainilla y chocolate comprende una secuencia de operaciones que permiten transformar los ingredientes en productos terminados.")
        etapas = [
            ("01", "Selección y pesado de ingredientes", "Se seleccionan los insumos y se pesan de acuerdo con la formulación correspondiente."),
            ("02", "Mezclado y preparación de la masa", "Se integran los ingredientes hasta obtener una mezcla homogénea."),
            ("03", "Preparación de moldes", "Se acondicionan los moldes para recibir la mezcla preparada."),
            ("04", "Horneado", "La mezcla se somete al proceso de cocción para obtener la estructura de la torta."),
            ("05", "Enfriamiento", "Se deja enfriar el producto antes de continuar con el armado."),
            ("06", "Rellenado y armado", "Se incorporan los rellenos y se ensamblan las capas de la torta."),
            ("07", "Decoración y acabado", "Se aplican las coberturas y decoraciones según la presentación del producto."),
            ("08", "Control de calidad y empaquetado", "Se verifica la presentación y condición del producto para su acondicionamiento y comercialización.")
        ]
        for numero_etapa, titulo_etapa, descripcion in etapas:
            st.markdown(f"""
            <div style="background:#F8F3EC;padding:15px;border-radius:12px;margin-bottom:10px;border-left:5px solid #A9D18E;">
                <span style="color:#A65D54;font-weight:bold;">ETAPA {numero_etapa}</span>
                <h4 style="color:#6C594B;margin:5px 0;">{titulo_etapa}</h4>
                <p style="color:#66584C;margin-bottom:0;">{descripcion}</p>
            </div>
            """, unsafe_allow_html=True)
        st.divider()
        st.subheader("Diagramas de análisis del proceso (DAP)")
        carpeta_assets = Path(__file__).resolve().parent / "assets"
        col_dap_vainilla, col_dap_chocolate = st.columns(2)

        with col_dap_vainilla:
            st.markdown("#### DAP – Torta de vainilla")
            ruta_dap_vainilla = carpeta_assets / "DAP_Vainilla.jpg"
            if ruta_dap_vainilla.exists():
                st.image(
                    str(ruta_dap_vainilla),
                    caption="Diagrama de análisis del proceso de la torta de vainilla",
                    use_container_width=True
                )
            else:
                st.warning("No se encontró assets/DAP_Vainilla.jpg.")

        with col_dap_chocolate:
            st.markdown("#### DAP – Torta de chocolate")
            ruta_dap_chocolate = carpeta_assets / "DAP_Chocolate.jpg"
            if ruta_dap_chocolate.exists():
                st.image(
                    str(ruta_dap_chocolate),
                    caption="Diagrama de análisis del proceso de la torta de chocolate",
                    use_container_width=True
                )
            else:
                st.warning("No se encontró assets/DAP_Chocolate.jpg.")

        st.divider()
        st.subheader("Diagrama general del proceso de fabricación")
        st.graphviz_chart("""
        digraph Proceso {
            rankdir=LR;
            node [shape=box,style="rounded,filled",fillcolor="#F3E8D5",color="#A9D18E",fontcolor="#59483D"];
            A [label="Selección y\npesado"];
            B [label="Mezclado y\npreparación"];
            C [label="Preparación\nde moldes"];
            D [label="Horneado"];
            E [label="Enfriamiento"];
            F [label="Rellenado y\narmado"];
            G [label="Decoración"];
            H [label="Control de calidad\ny empaquetado"];
            A -> B -> C -> D -> E -> F -> G -> H;
        }
        """, use_container_width=True)
        st.success("Resultado: tortas de vainilla y chocolate terminadas y listas para su comercialización.")


# =========================================================
# CRM
# =========================================================
elif seccion == "CRM":
    boton_volver_inicio()
    titulo_seccion("🤝 CRM – Gestión de relaciones con clientes")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        opcion = st.selectbox("Selecciona un apartado", ["Leads / Clientes potenciales", "Atención al cliente", "Ventas"])
        if opcion == "Leads / Clientes potenciales":
            st.subheader("Gestión de Lead – Registro de clientes potenciales")
            columnas_leads = [
                "ID cliente", "Nombre", "Apellido", "Celular", "Correo electrónico",
                "Fecha de registro", "Canal de contacto", "Producto de interés",
                "Asesor asignado", "Estado"
            ]
            estados_lead = ["Pendiente", "Descartado", "Contactado", "Cotizado", "Venta"]

            # La base se carga una sola vez por sesión para permitir altas y ediciones.
            if "crm_leads" not in st.session_state:
                base_leads = limpiar_df(hoja_df("Clientes", 3, 4, 603, 1, 10))
                if base_leads.empty:
                    base_leads = pd.DataFrame(columns=columnas_leads)
                else:
                    for columna in columnas_leads:
                        if columna not in base_leads.columns:
                            base_leads[columna] = None
                    base_leads = base_leads[columnas_leads].copy()
                st.session_state["crm_leads"] = base_leads.reset_index(drop=True)

            df_leads = st.session_state["crm_leads"].copy()

            # Normaliza estados heredados y evita valores fuera de la lista solicitada.
            df_leads["Estado"] = df_leads["Estado"].fillna("Pendiente").astype(str)
            df_leads.loc[~df_leads["Estado"].isin(estados_lead), "Estado"] = "Pendiente"

            with st.expander("➕ Registrar un nuevo cliente potencial", expanded=True):
                with st.form("crm_form_nuevo_lead", clear_on_submit=True):
                    c1, c2, c3 = st.columns(3)
                    nombre_n = c1.text_input("Nombre *")
                    apellido_n = c2.text_input("Apellido *")
                    celular_n = c3.text_input("Celular")
                    c4, c5, c6 = st.columns(3)
                    correo_n = c4.text_input("Correo electrónico")
                    fecha_n = c5.date_input("Fecha de registro", value=pd.Timestamp.today().date())
                    canal_n = c6.selectbox(
                        "Canal de contacto",
                        ["WhatsApp", "Tienda física", "Facebook", "Instagram", "Llamada", "Otro"]
                    )
                    c7, c8, c9 = st.columns(3)
                    producto_n = c7.text_input("Producto de interés *", placeholder="Ej. Torta de vainilla de 1 kg")
                    asesor_n = c8.text_input("Asesor asignado *")
                    estado_n = c9.selectbox("Estado inicial", estados_lead, index=0)
                    registrar = st.form_submit_button("Registrar cliente potencial", use_container_width=True)

                    if registrar:
                        if not nombre_n.strip() or not apellido_n.strip() or not producto_n.strip() or not asesor_n.strip():
                            st.error("Completa los campos obligatorios: nombre, apellido, producto de interés y asesor.")
                        else:
                            ids_existentes = df_leads["ID cliente"].dropna().astype(str).tolist()
                            numeros_ids = [
                                int(re.search(r"(\d+)$", valor).group(1))
                                for valor in ids_existentes
                                if re.search(r"(\d+)$", valor)
                            ]
                            siguiente_id = max(numeros_ids, default=0) + 1
                            nuevo_id = f"CLI-{siguiente_id:04d}"
                            nuevo_lead = {
                                "ID cliente": nuevo_id,
                                "Nombre": nombre_n.strip(),
                                "Apellido": apellido_n.strip(),
                                "Celular": celular_n.strip(),
                                "Correo electrónico": correo_n.strip(),
                                "Fecha de registro": pd.Timestamp(fecha_n),
                                "Canal de contacto": canal_n,
                                "Producto de interés": producto_n.strip(),
                                "Asesor asignado": asesor_n.strip(),
                                "Estado": estado_n
                            }
                            st.session_state["crm_leads"] = pd.concat(
                                [df_leads, pd.DataFrame([nuevo_lead])],
                                ignore_index=True
                            )
                            st.success(f"Cliente potencial registrado correctamente con el ID {nuevo_id}.")
                            st.rerun()

            st.divider()
            st.subheader("Estado y seguimiento de clientes potenciales")
            st.caption("Edita directamente la información y el estado de cada lead. Los estados disponibles son Pendiente, Descartado, Contactado, Cotizado y Venta.")

            # Resumen calculado sobre todos los registros, no solo sobre los filtrados.
            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("Total de leads", len(df_leads))
            for columna_kpi, estado_kpi in zip([k2, k3, k4, k5], estados_lead[:4]):
                columna_kpi.metric(estado_kpi, int((df_leads["Estado"] == estado_kpi).sum()))
            st.metric("Leads en Venta", int((df_leads["Estado"] == "Venta").sum()))

            df_edicion = df_leads.copy()
            df_edicion["Cliente"] = (
                df_edicion["Nombre"].fillna("").astype(str).str.strip()
                + " " +
                df_edicion["Apellido"].fillna("").astype(str).str.strip()
            ).str.strip()
            orden_columnas = [
                "ID cliente", "Cliente", "Nombre", "Apellido", "Celular",
                "Correo electrónico", "Fecha de registro", "Canal de contacto",
                "Producto de interés", "Asesor asignado", "Estado"
            ]
            df_edicion = df_edicion[orden_columnas]

            st.markdown("#### Buscar y filtrar registros")
            f1, f2, f3 = st.columns(3)
            with f1:
                buscar_lead = st.text_input("Buscar por ID o nombre", key="crm_buscar_lead")
            with f2:
                filtro_estado_lead = st.selectbox("Filtrar por estado", ["Todos"] + estados_lead, key="crm_filtro_estado")
            with f3:
                asesores_lead = ["Todos"] + sorted(df_leads["Asesor asignado"].dropna().astype(str).unique().tolist())
                filtro_asesor_lead = st.selectbox("Filtrar por asesor", asesores_lead, key="crm_filtro_asesor")

            vista_leads = df_edicion.copy()
            if buscar_lead.strip():
                mascara = (
                    vista_leads["ID cliente"].astype(str).str.contains(buscar_lead.strip(), case=False, na=False)
                    | vista_leads["Cliente"].astype(str).str.contains(buscar_lead.strip(), case=False, na=False)
                )
                vista_leads = vista_leads[mascara]
            if filtro_estado_lead != "Todos":
                vista_leads = vista_leads[vista_leads["Estado"] == filtro_estado_lead]
            if filtro_asesor_lead != "Todos":
                vista_leads = vista_leads[vista_leads["Asesor asignado"].astype(str) == filtro_asesor_lead]

            st.caption("Puedes modificar los datos directamente en la tabla. El ID y el nombre consolidado son campos de solo lectura.")
            edited_leads = st.data_editor(
                vista_leads,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                key=f"crm_editor_leads_{len(df_leads)}",
                disabled=["ID cliente", "Cliente"],
                column_config={
                    "Estado": st.column_config.SelectboxColumn("Estado", options=estados_lead, required=True),
                    "Fecha de registro": st.column_config.DateColumn("Fecha de registro", format="DD/MM/YYYY"),
                    "Canal de contacto": st.column_config.SelectboxColumn(
                        "Canal de contacto",
                        options=["WhatsApp", "Tienda física", "Facebook", "Instagram", "Llamada", "Otro"]
                    )
                }
            )

            if st.button("Guardar cambios de la tabla", key="crm_guardar_edicion", use_container_width=True):
                actualizado = df_leads.copy()
                for indice in edited_leads.index:
                    for columna in columnas_leads:
                        if columna in edited_leads.columns and columna != "ID cliente":
                            actualizado.loc[indice, columna] = edited_leads.loc[indice, columna]
                st.session_state["crm_leads"] = actualizado.reset_index(drop=True)
                st.success("Los cambios de los clientes se actualizaron en la sesión.")
                st.rerun()

            st.divider()
            st.subheader("Distribución de leads")
            conteo_estados = (
                df_leads["Estado"].value_counts()
                .reindex(estados_lead, fill_value=0)
                .rename_axis("Estado").reset_index(name="Cantidad")
            )
            grafico_barras(conteo_estados, "Estado", "Cantidad", "Clientes potenciales por estado")
            if not df_leads.empty and "Asesor asignado" in df_leads.columns:
                conteo_asesores = (
                    df_leads["Asesor asignado"].fillna("Sin asesor")
                    .value_counts().rename_axis("Asesor").reset_index(name="Cantidad")
                )
                grafico_barras(conteo_asesores, "Asesor", "Cantidad", "Leads asignados por asesor")

            st.divider()
            if st.button("💾 Guardar gestión de leads en el Excel", key="crm_guardar_excel", use_container_width=True):
                try:
                    libro_leads = openpyxl.load_workbook(BytesIO(bytes_actuales()), data_only=False)
                    if "Clientes" not in libro_leads.sheetnames:
                        ws_leads = libro_leads.create_sheet("Clientes")
                        for j, encabezado in enumerate(columnas_leads, start=1):
                            ws_leads.cell(3, j).value = encabezado
                    else:
                        ws_leads = libro_leads["Clientes"]

                    # Actualiza la hoja Clientes respetando los encabezados y el formato existente.
                    ultima_fila = max(ws_leads.max_row, len(df_leads) + 3)
                    for fila_excel in range(4, ultima_fila + 1):
                        for col_excel in range(1, len(columnas_leads) + 1):
                            ws_leads.cell(fila_excel, col_excel).value = None

                    for i, (_, registro) in enumerate(df_leads.iterrows(), start=4):
                        for j, columna in enumerate(columnas_leads, start=1):
                            valor = registro[columna]
                            if pd.isna(valor):
                                valor = None
                            elif columna == "Fecha de registro" and valor is not None:
                                valor = pd.to_datetime(valor).to_pydatetime()
                            elif columna != "Fecha de registro" and valor is not None:
                                valor = str(valor)
                            ws_leads.cell(i, j).value = valor

                    guardar_libro_en_sesion(libro_leads)
                    st.success("La hoja Clientes del Excel se actualizó correctamente. Descarga el Excel actualizado para conservar los nuevos registros.")
                    st.download_button(
                        "⬇️ Descargar Excel con la gestión de leads",
                        data=bytes_actuales(),
                        file_name="PASTELPAN_ERP_con_leads_actualizados.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key="crm_descargar_excel"
                    )
                except Exception as error:
                    st.error(f"No se pudo actualizar el Excel: {error}")

            st.info("Importante: los registros y cambios se mantienen durante la sesión activa. Para conservarlos después de cerrar la aplicación o reiniciar Streamlit Cloud, descarga el Excel actualizado y vuelve a cargarlo en la siguiente sesión.")
        elif opcion == "Atención al cliente":
            st.subheader("Registro de atención al cliente")
            df = limpiar_df(hoja_df("Atención al cliente", 3, 4, 1839, 1, 10))
            if not df.empty:
                col_fecha = "Fecha de venta" if "Fecha de venta" in df.columns else None
                col_score = "Calificación NPS (0-5)" if "Calificación NPS (0-5)" in df.columns else None
                if col_fecha:
                    df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
                if col_score:
                    df[col_score] = pd.to_numeric(df[col_score], errors="coerce")
                f1, f2, f3 = st.columns(3)
                with f1:
                    asesores = filtro_opciones(df, "Asesor", "Asesor de atención")
                with f2:
                    scores = sorted(df[col_score].dropna().unique().tolist()) if col_score else []
                    score = st.selectbox("Calificación (0 a 5)", ["Todas"] + [str(x) for x in scores])
                with f3:
                    if col_fecha and df[col_fecha].notna().any():
                        fechas = df[col_fecha].dropna()
                        rango = st.date_input("Periodo de atención", value=(fechas.min().date(), fechas.max().date()))
                    else:
                        rango = None
                df = aplicar_filtro(df, "Asesor", asesores)
                if score != "Todas" and col_score:
                    df = df[df[col_score] == float(score)]
                if rango and col_fecha and isinstance(rango, tuple) and len(rango) == 2:
                    df = df[df[col_fecha].dt.date.between(rango[0], rango[1])]
                st.metric("Registros de atención", f"{len(df):,}")
                df_vista = df.copy()
                if col_fecha and col_fecha in df_vista.columns:
                    df_vista[col_fecha] = df_vista[col_fecha].dt.strftime("%d/%m/%Y")
                mostrar_tabla(df_vista)
                st.markdown("#### Calificaciones de atención al cliente")
                if col_score and not df.empty:
                    distrib = df[col_score].value_counts().sort_index().rename_axis("Calificación").reset_index(name="Cantidad")
                    grafico_barras(distrib, "Calificación", "Cantidad", "Distribución de calificaciones (0 a 5)")
                if col_score and not df.empty:
                    st.metric("Promedio del filtro actual", f"{formato_numero(df[col_score].mean())}/5")
                st.caption("La escala utilizada es de 0 a 5, conforme a la base del proyecto. No se interpreta como NPS tradicional.")
            else:
                st.warning("No se encontraron registros de atención al cliente.")
        else:
            st.subheader("Ventas registradas")
            df = limpiar_df(hoja_df("Atención al cliente", 3, 4, 1839, 1, 10))
            if not df.empty:
                if "Fecha de venta" in df.columns:
                    df["Fecha de venta"] = pd.to_datetime(df["Fecha de venta"], errors="coerce")
                if "Cantidad (unid.)" in df.columns:
                    df["Cantidad (unid.)"] = pd.to_numeric(df["Cantidad (unid.)"], errors="coerce")
                f1, f2 = st.columns(2)
                with f1:
                    producto = filtro_opciones(df, "Producto", "Producto vendido")
                with f2:
                    asesor = filtro_opciones(df, "Asesor", "Asesor de venta")
                df = aplicar_filtro(df, "Producto", producto)
                df = aplicar_filtro(df, "Asesor", asesor)
                df_vista = df.copy()
                if "Fecha de venta" in df_vista.columns:
                    df_vista["Fecha de venta"] = df_vista["Fecha de venta"].dt.strftime("%d/%m/%Y")
                mostrar_tabla(df_vista)
                if "Producto" in df.columns and "Cantidad (unid.)" in df.columns:
                    ventas_prod = df.groupby("Producto", dropna=False)["Cantidad (unid.)"].sum().reset_index()
                    grafico_barras(ventas_prod, "Producto", "Cantidad (unid.)", "Unidades vendidas por producto")


# =========================================================
# PRODUCCIÓN
# =========================================================
elif seccion == "Producción":
    boton_volver_inicio()
    titulo_seccion("🧁 Producción", "Consulta de planificación agregada, materiales, capacidad, rutas y distribución.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        apartados = [
            "Demanda y pronóstico", "PAP", "PMP – Vainilla", "PMP – Chocolate",
            "Maestro de materiales", "BOM – Vainilla", "BOM – Chocolate",
            "MRP", "DRP – Vainilla", "DRP – Chocolate"
        ]
        opcion = st.selectbox("Selecciona un apartado de producción", apartados)

        if opcion == "Demanda y pronóstico":
            st.subheader("Demanda histórica y pronóstico")
            especificaciones = [
                ("Demanda histórica – Vainilla", "B2:D26", "Año", "Unidades"),
                ("Pronóstico – Vainilla", "F2:I14", "Año", "Unidades"),
                ("Demanda histórica – Chocolate", "B30:D54", "Año", "Unidades"),
                ("Pronóstico – Chocolate", "F30:I42", "Año", "Unidades")
            ]
            for nombre, rango, _, _ in especificaciones:
                st.markdown(f"#### {nombre}")
                df = rango_df("PRODUCCION (RESUMEN)", rango, True)
                mostrar_tabla(df, 250)
                if not df.empty:
                    columnas = df.columns.tolist()
                    col_val = next((c for c in columnas if "UNIDADES" in c.upper()), None)
                    col_mes = next((c for c in columnas if "MES" in c.upper()), None)
                    col_anio = next((c for c in columnas if "AÑO" in c.upper() or "AÑOS" in c.upper()), None)
                    if col_val and col_mes:
                        graf = df.copy()
                        graf["Periodo"] = graf[col_mes].astype(str)
                        if col_anio:
                            graf["Periodo"] = graf[col_anio].astype(str) + " - " + graf[col_mes].astype(str)
                        graf[col_val] = pd.to_numeric(graf[col_val], errors="coerce")
                        graf = graf.dropna(subset=[col_val])
                        grafico_linea(graf, "Periodo", col_val, nombre)
        elif opcion == "PAP":
            st.subheader("Plan agregado de producción (PAP)")
            df = rango_df("PRODUCCION (RESUMEN)", "B58:F66", True)
            mostrar_tabla(df)
            if not df.empty:
                st.markdown("#### Costo total por plan")
                mostrar_tarjetas_costos_pap(df)
        elif opcion.startswith("PMP"):
            st.subheader(opcion)
            if "Vainilla" in opcion:
                df = crear_df_pmp("B70:Z78")
            else:
                df = crear_df_pmp("B82:Z90")
            mostrar_tabla(df, 420)
            datos = datos_grafico_pmp(df)
            if not datos.empty:
                grafico_linea(datos, "Semana", "Unidades", "Programa maestro de producción")
        elif opcion == "Maestro de materiales":
            st.subheader("Maestro de materiales")
            df = rango_df("PRODUCCION (RESUMEN)", "B95:L115", True)
            mostrar_tabla(df, 500)
            st.caption("La tabla se presenta desde la hoja PRODUCCION (RESUMEN), rango B95:L115.")
        elif opcion == "BOM – Vainilla":
            st.subheader("Lista de materiales – Torta de vainilla")
            ruta_bom_vainilla = Path(__file__).resolve().parent / "assets" / "bom_vainilla.jpg"
            if ruta_bom_vainilla.exists():
                st.image(str(ruta_bom_vainilla), caption="BOM – Torta de vainilla", use_container_width=True)
            df = rango_df("PRODUCCION (RESUMEN)", "B118:H134", True)
            mostrar_tabla(df)
        elif opcion == "BOM – Chocolate":
            st.subheader("Lista de materiales – Torta de chocolate")
            ruta_bom_chocolate = Path(__file__).resolve().parent / "assets" / "bom_chocolate.jpg"
            if ruta_bom_chocolate.exists():
                st.image(str(ruta_bom_chocolate), caption="BOM – Torta de chocolate", use_container_width=True)
            df = rango_df("PRODUCCION (RESUMEN)", "B137:H149", True)
            mostrar_tabla(df)
        elif opcion == "MRP":
            st.subheader("MRP – Planificación de requerimientos de materiales")
            submodulo_mrp = st.selectbox(
                "Selecciona un submódulo de MRP",
                ["TLO", "BOC", "CRP", "Lista de puestos de trabajo",
                 "Hoja de ruta – Vainilla", "Hoja de ruta – Chocolate"],
                key="submodulo_mrp"
            )
            if submodulo_mrp == "TLO":
                st.markdown("#### Tabla de Lanzamiento de Orden (TLO)")
                df = rango_df("PRODUCCION (RESUMEN)", "B152:K172", True)
                mostrar_tabla(df, 450)
            elif submodulo_mrp == "BOC":
                st.markdown("#### Balance de operaciones y carga (BOC)")
                df = rango_df("PRODUCCION (RESUMEN)", "B176:K210", False)
                mostrar_tabla(df, 500)
                st.caption("Se conservan los dos niveles de encabezado: estación de trabajo y tipo de recurso.")
            elif submodulo_mrp == "CRP":
                st.markdown("#### Planificación de requerimientos de capacidad (CRP)")
                df = rango_df("PRODUCCION (RESUMEN)", "B212:K241", False)
                mostrar_tabla(df, 560)
                st.info("El horizonte de planificación considerado en el proyecto corresponde a noviembre y diciembre, con nueve semanas: Sem. 00 a Sem. 08.")
            elif submodulo_mrp == "Lista de puestos de trabajo":
                st.markdown("#### Lista de puestos de trabajo")
                df = rango_df("PRODUCCION (RESUMEN)", "M213:U217", True)
                mostrar_tabla(df, 300)
            elif submodulo_mrp == "Hoja de ruta – Vainilla":
                st.markdown("#### Hoja de ruta – Torta de vainilla")
                df = rango_df("PRODUCCION (RESUMEN)", "M220:W229", True)
                mostrar_tabla(df, 400)
            elif submodulo_mrp == "Hoja de ruta – Chocolate":
                st.markdown("#### Hoja de ruta – Torta de chocolate")
                df = rango_df("PRODUCCION (RESUMEN)", "M232:W242", True)
                mostrar_tabla(df, 400)
        elif opcion == "DRP – Vainilla":
            st.subheader("Planificación de distribución – Torta de vainilla")
            df = rango_df("PRODUCCION (RESUMEN)", "B244:K251", True)
            mostrar_tabla(df)
            st.caption("Se presenta la planificación por seis sucursales y la fila Fábrica se conserva como referencia de producción.")
            ruta_drp = Path(__file__).resolve().parent / "assets" / "DRP.jpg"
            if ruta_drp.exists():
                st.image(str(ruta_drp), caption="Diagrama de planificación de distribución (DRP)", use_container_width=True)
            else:
                st.warning("No se encontró assets/DRP.jpg.")
        elif opcion == "DRP – Chocolate":
            st.subheader("Planificación de distribución – Torta de chocolate")
            df = rango_df("PRODUCCION (RESUMEN)", "B253:K260", True)
            mostrar_tabla(df)
            st.caption("Se presenta la planificación por seis sucursales y la fila Fábrica se conserva como referencia de producción.")
            ruta_drp = Path(__file__).resolve().parent / "assets" / "DRP.jpg"
            if ruta_drp.exists():
                st.image(str(ruta_drp), caption="Diagrama de planificación de distribución (DRP)", use_container_width=True)
            else:
                st.warning("No se encontró assets/DRP.jpg.")


# =========================================================
# INVENTARIOS
# =========================================================
elif seccion == "Inventarios":
    boton_volver_inicio()
    titulo_seccion("📦 Inventarios", "Consulta, registro de materiales y actualización de existencias.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        st.subheader("Maestro de materiales")
        df_base = rango_df("PRODUCCION (RESUMEN)", "B95:L115", True)
        if df_base.empty:
            st.warning("No se encontró el maestro de materiales.")
        else:
            df_base.columns = [str(c).strip() for c in df_base.columns]
            if st.session_state["inventario_editado"] is None:
                st.session_state["inventario_editado"] = df_base.copy()
            df_inv = st.session_state["inventario_editado"].copy()
            c1, c2 = st.columns([2, 1])
            with c1:
                busqueda = st.text_input("Buscar por código o material")
            with c2:
                tipo = st.selectbox("Tipo de material", ["Todos"] + sorted(df_inv["Tipo"].dropna().astype(str).unique().tolist()))
            vista = df_inv.copy()
            if busqueda:
                vista = vista[
                    vista["Código"].astype(str).str.contains(busqueda, case=False, na=False)
                    | vista["Material"].astype(str).str.contains(busqueda, case=False, na=False)
                ]
            if tipo != "Todos":
                vista = vista[vista["Tipo"].astype(str) == tipo]
            st.metric("Materiales registrados", len(df_inv))
            edited = st.data_editor(
                vista,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=[c for c in vista.columns if c not in ["Stock actual"]],
                column_config={"Stock actual": st.column_config.NumberColumn("Stock actual", min_value=0, step=1, format="%g")}
            )
            if st.button("Aplicar cambios de stock", key="guardar_stock"):
                for idx in edited.index:
                    df_inv.loc[idx, "Stock actual"] = edited.loc[idx, "Stock actual"]
                st.session_state["inventario_editado"] = df_inv
                st.success("Cambios aplicados en la sesión. Descarga el Excel actualizado para conservarlos.")
            st.divider()
            with st.expander("Registrar nuevo material"):
                with st.form("nuevo_material"):
                    cols = st.columns(3)
                    codigo = cols[0].text_input("Código (único)")
                    material = cols[1].text_input("Material")
                    tipo_n = cols[2].selectbox("Tipo", ["Materia prima", "Insumos", "Material de empaque", "Producto terminado"])
                    cols2 = st.columns(3)
                    unidad = cols2[0].text_input("Unidad")
                    stock = cols2[1].number_input("Stock inicial", min_value=0.0, step=1.0)
                    seguridad = cols2[2].number_input("Stock de seguridad", min_value=0.0, step=1.0)
                    cols3 = st.columns(3)
                    lead = cols3[0].text_input("Lead time", value="1 sem.")
                    lote = cols3[1].text_input("Lote mínimo")
                    costo = cols3[2].number_input("Costo unitario", min_value=0.0, step=0.1)
                    costo_unidad = st.text_input("Unidad de costo", value=unidad)
                    proveedor = st.text_input("Proveedor")
                    enviar = st.form_submit_button("Agregar material")
                    if enviar:
                        if not codigo.strip() or not material.strip():
                            st.error("El código y el nombre del material son obligatorios.")
                        elif codigo.strip() in df_inv["Código"].astype(str).values:
                            st.error("Ese código ya existe.")
                        else:
                            nueva = {
                                "Código": codigo.strip(), "Material": material.strip(), "Tipo": tipo_n,
                                "Unidad": unidad, "Stock actual": stock, "Stock seguridad": seguridad,
                                "Lead time": lead, "Lote mínimo": lote, "Costo unitario": costo,
                                "Unidad de costo": costo_unidad, "Proveedor": proveedor
                            }
                            st.session_state["inventario_editado"] = pd.concat([df_inv, pd.DataFrame([nueva])], ignore_index=True)
                            st.success("Material agregado en la sesión.")
            with st.expander("Registrar entrada o salida de inventario"):
                codigos = st.session_state["inventario_editado"]["Código"].astype(str).tolist()
                with st.form("movimiento_stock"):
                    codigo_mov = st.selectbox("Material", codigos)
                    tipo_mov = st.radio("Movimiento", ["Entrada", "Salida"], horizontal=True)
                    cantidad = st.number_input("Cantidad", min_value=0.01, step=1.0)
                    motivo = st.text_input("Motivo / referencia")
                    ok = st.form_submit_button("Registrar movimiento")
                    if ok:
                        tabla = st.session_state["inventario_editado"].copy()
                        idx = tabla.index[tabla["Código"].astype(str) == codigo_mov][0]
                        stock_actual = numero(tabla.loc[idx, "Stock actual"])
                        nuevo_stock = stock_actual + cantidad if tipo_mov == "Entrada" else stock_actual - cantidad
                        if nuevo_stock < 0:
                            st.error("La salida supera el stock disponible.")
                        else:
                            tabla.loc[idx, "Stock actual"] = nuevo_stock
                            st.session_state["inventario_editado"] = tabla
                            st.session_state["movimientos_inventario"].append({
                                "Código": codigo_mov, "Movimiento": tipo_mov,
                                "Cantidad": cantidad, "Motivo": motivo
                            })
                            st.success(f"Movimiento registrado. Nuevo stock: {formato_numero(nuevo_stock)}")
            if st.session_state["movimientos_inventario"]:
                st.markdown("#### Movimientos registrados en esta sesión")
                mostrar_tabla(pd.DataFrame(st.session_state["movimientos_inventario"]), 220)
            if st.button("Guardar inventario en el Excel descargable", key="guardar_inv_excel"):
                libro_edit = openpyxl.load_workbook(BytesIO(bytes_actuales()), data_only=False)
                ws_res = libro_edit["PRODUCCION (RESUMEN)"]
                ws_maestro = libro_edit["Maestro de Materiales"]
                datos = st.session_state["inventario_editado"]
                # Actualiza el rango B95:L115 del resumen y la tabla B4:L...
                for i, fila in datos.iterrows():
                    r = 96 + i
                    for j, col in enumerate(["Código", "Material", "Tipo", "Unidad", "Stock actual", "Stock seguridad", "Lead time", "Lote mínimo", "Costo unitario", "Unidad de costo", "Proveedor"], start=2):
                        ws_res.cell(r, j).value = fila[col]
                    r2 = 5 + i
                    for j, col in enumerate(["Código", "Material", "Tipo", "Unidad", "Stock actual", "Stock seguridad", "Lead time", "Lote mínimo", "Costo unitario", "Unidad de costo", "Proveedor"], start=2):
                        ws_maestro.cell(r2, j).value = fila[col]
                guardar_libro_en_sesion(libro_edit)
                st.success("El Excel en memoria se actualizó. Descárgalo desde el botón de la barra lateral inferior.")
            if bytes_actuales():
                st.download_button(
                    "⬇️ Descargar Excel con los cambios",
                    data=bytes_actuales(),
                    file_name="PASTELPAN_ERP_actualizado.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            st.caption("Los cambios se mantienen en la sesión actual y se incorporan al Excel al guardarlo y descargarlo. Para persistencia automática entre sesiones en Streamlit Cloud se requiere conectar una base de datos externa.")


# =========================================================
# COMPRAS Y PROVEEDORES
# =========================================================
elif seccion == "Compras y Proveedores":
    boton_volver_inicio()
    titulo_seccion("🛒 Compras y Proveedores", "Directorio de proveedores y control de órdenes.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        opcion = st.selectbox("Selecciona un apartado", ["Proveedores", "Órdenes de Compra y Servicio"])
        if opcion == "Proveedores":
            st.subheader("Directorio de proveedores")
            df = limpiar_df(hoja_df("Proveedores", 3, 4, 19, 1, 6))
            if not df.empty:
                cal_col = "Calificación" if "Calificación" in df.columns else None
                if cal_col:
                    df[cal_col] = pd.to_numeric(df[cal_col], errors="coerce")
                mostrar_tabla(df)
                if cal_col:
                    grafico_barras(df, "Nombre del proveedor", cal_col, "Calificación de proveedores (0 a 5)")
                    st.metric("Calificación promedio", f"{formato_numero(df[cal_col].mean())}/5")
        else:
            st.subheader("Órdenes de Compra y Servicio")
            df = limpiar_df(hoja_df("Órdenes de Compra y Servicio", 3, 4, 23, 1, 7))
            if not df.empty:
                if "Monto (S/)" in df.columns:
                    df["Monto (S/)"] = pd.to_numeric(df["Monto (S/)"], errors="coerce")
                estado = filtro_opciones(df, "Estado", "Estado de la orden")
                tipo = filtro_opciones(df, "Tipo", "Tipo de orden")
                df = aplicar_filtro(df, "Estado", estado)
                df = aplicar_filtro(df, "Tipo", tipo)
                st.metric("Órdenes encontradas", len(df))
                mostrar_tabla(df)
                if "Estado" in df.columns:
                    conteo = df["Estado"].fillna("Sin estado").value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                    grafico_barras(conteo, "Estado", "Cantidad", "Distribución de órdenes por estado")
                if "Monto (S/)" in df.columns:
                    st.metric("Monto de las órdenes filtradas", f"S/ {formato_numero(df['Monto (S/)'].sum())}")


# =========================================================
# DASHBOARD
# =========================================================
elif seccion == "Dashboard":
    boton_volver_inicio()
    titulo_seccion("📊 Dashboard de PASTELPAN", "Resumen ejecutivo del desempeño comercial, productivo y de abastecimiento.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        indicadores = obtener_indicadores()
        if indicadores:
            st.markdown("### Indicadores principales")
            nombres = list(indicadores.keys())[:4]
            kpi_cols = st.columns(4)
            for i, nombre in enumerate(nombres):
                dato = indicadores[nombre]
                valor = dato["resultado"]
                unidad = str(dato["unidad"]).lower()
                if "porcentaje" in unidad or "%" in unidad:
                    valor_txt = formato_porcentaje(valor)
                elif "puntos" in unidad:
                    valor_txt = f"{formato_numero(valor)}/5"
                else:
                    valor_txt = formato_numero(valor)
                unidad_txt = str(dato["unidad"]) if dato["unidad"] is not None else ""
                kpi_cols[i].markdown(
                    f'''<div class="kpi-principal">
                        <div class="kpi-nombre">{nombre}</div>
                        <div class="kpi-valor">{valor_txt}</div>
                        <div class="kpi-unidad">{unidad_txt}</div>
                    </div>''',
                    unsafe_allow_html=True
                )
            interpret_cols = st.columns(4)
            for i, nombre in enumerate(nombres):
                dato = indicadores[nombre]
                with interpret_cols[i]:
                    with st.popover("ⓘ Ver interpretación", use_container_width=True):
                        st.markdown(f"**{nombre}**")
                        interpretaciones_cortas = {
                            "Calificación promedio de atención al cliente":
                                "La atención al cliente obtiene una valoración promedio de 4.59 sobre 5.",
                            "Utilización de capacidad productiva":
                                "El plan utiliza el 82.23% de las horas de capacidad disponibles.",
                            "Volumen total de distribución planificada":
                                "Se han programado 1,818 tortas para distribuir entre las seis sucursales.",
                            "Calificación promedio de proveedores":
                                "Los proveedores alcanzan una calificación promedio de 4.54 sobre 5."
                        }
                        st.write(interpretaciones_cortas.get(nombre, dato["interpretacion"]))
            st.caption("Los cuatro indicadores principales se obtienen directamente de la hoja INDICADORES ERP.")
        else:
            st.warning("No se encontraron los cuatro indicadores principales en INDICADORES ERP.")

        st.divider()
        st.divider()
        st.subheader("Análisis visual por módulo")
        clientes = limpiar_df(hoja_df("Clientes", 3, 4, 603, 1, 10))
        atencion = limpiar_df(hoja_df("Atención al cliente", 3, 4, 1839, 1, 10))
        proveedores = limpiar_df(hoja_df("Proveedores", 3, 4, 19, 1, 6))
        ordenes = limpiar_df(hoja_df("Órdenes de Compra y Servicio", 3, 4, 23, 1, 7))

        st.markdown("#### Filtros de los gráficos")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            filtro_lead = filtro_opciones(clientes, "Estado", "Estado de leads") if not clientes.empty else None
        with f2:
            filtro_atencion = filtro_opciones(atencion, "Calificación NPS (0-5)", "Calificación de atención") if not atencion.empty else None
        with f3:
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
