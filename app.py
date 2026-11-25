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
        headers = [str(x).strip() if x is not None else f"Campo {i+1}" for i, x in enumerate(valores[0])]
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


def mostrar_tabla(df, altura=420):
    if df is None or df.empty:
        st.info("No hay registros disponibles en este apartado.")
    else:
        st.dataframe(df, use_container_width=True, hide_index=True, height=altura)


def grafico_barras(df, x, y, titulo, color=None):
    if df is None or df.empty or x not in df.columns or y not in df.columns:
        st.info("No hay datos suficientes para generar este gráfico.")
        return
    fig = px.bar(df, x=x, y=y, color=color, title=titulo, text_auto=".2s")
    fig.update_layout(
        title_x=0.02, plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(l=15, r=15, t=55, b=15), legend_title_text=""
    )
    st.plotly_chart(fig, use_container_width=True)


def grafico_linea(df, x, y, titulo, color=None):
    if df is None or df.empty or x not in df.columns or y not in df.columns:
        st.info("No hay datos suficientes para generar este gráfico.")
        return
    fig = px.line(df, x=x, y=y, color=color, markers=True, title=titulo)
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

with st.sidebar:
    st.markdown("## 🎂 PASTELPAN ERP")
    seccion = st.radio("MENÚ PRINCIPAL", MODULOS, index=0)
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
                padding:22px;min-height:150px;margin-bottom:12px;">
                <div style="font-size:30px;">{icono}</div>
                <h3 style="margin:5px 0;">{nombre}</h3>
                <p style="color:#655A56;">{desc}</p></div>""",
                unsafe_allow_html=True
            )
    if bytes_actuales():
        st.success("El archivo maestro está conectado. Selecciona un módulo en el menú lateral.")
    else:
        st.info("Para comenzar, carga el archivo PROYECTO FINAL PASTELPAN GESOOP.xlsx desde el menú lateral.")


# =========================================================
# INFORMACIÓN
# =========================================================
elif seccion == "Información":
    titulo_seccion("🏢 Información de PASTELPAN", "Información empresarial y descripción de los productos.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        df_empresa = hoja_df("Presentación de la empresa ", 1, 1, 79, 1, 12)
        # Presentación de la empresa conserva su estructura original.
        st.subheader("Presentación de la empresa")
        mostrar_tabla(df_empresa, 300)
        st.divider()
        st.subheader("Productos principales")
        df_productos = hoja_df("Definición del producto ", 1, 1, 59, 1, 25)
        mostrar_tabla(df_productos, 380)
        c1, c2 = st.columns(2)
        with c1:
            st.markdown("#### 🎂 Torta de vainilla de 1 kg")
            st.write("Producto principal elaborado con harina de trigo, azúcar, huevos, aceite, leche, esencia de vainilla y materiales de decoración y empaque, según la estructura de materiales del proyecto.")
        with c2:
            st.markdown("#### 🍫 Torta de chocolate de 1 kg")
            st.write("Producto elaborado con premezcla doble chocolate, huevos, aceite, leche, chantilly, decoración de chocolate y materiales de empaque, según la estructura de materiales del proyecto.")
        st.caption("Los detalles técnicos y las cantidades por producto se consultan en las secciones BOM y Maestro de Materiales del módulo Producción.")


# =========================================================
# CRM
# =========================================================
elif seccion == "CRM":
    titulo_seccion("🤝 CRM – Gestión de relaciones con clientes")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        opcion = st.selectbox("Selecciona un apartado", ["Leads / Clientes potenciales", "Atención al cliente", "Ventas"])
        if opcion == "Leads / Clientes potenciales":
            st.subheader("Base de clientes potenciales")
            df = limpiar_df(hoja_df("Clientes", 3, 4, 603, 1, 10))
            if not df.empty:
                nombre_cols = [c for c in ["Nombre", "Apellido"] if c in df.columns]
                if nombre_cols:
                    df["Cliente"] = df[nombre_cols].fillna("").astype(str).agg(" ".join, axis=1).str.strip()
                filtros = st.columns(4)
                with filtros[0]:
                    estado = filtro_opciones(df, "Estado", "Estado del lead")
                with filtros[1]:
                    asesor = filtro_opciones(df, "Asesor asignado", "Asesor")
                with filtros[2]:
                    producto = filtro_opciones(df, "Producto de interés", "Producto")
                with filtros[3]:
                    cliente = filtro_opciones(df, "Cliente", "Cliente")
                for col, val in [("Estado", estado), ("Asesor asignado", asesor), ("Producto de interés", producto), ("Cliente", cliente)]:
                    df = aplicar_filtro(df, col, val)
                st.metric("Leads encontrados", f"{len(df):,}")
                mostrar_tabla(df)
                st.markdown("#### Distribución de leads por estado")
                if "Estado" in df.columns:
                    conteo = df["Estado"].fillna("Sin estado").value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                    grafico_barras(conteo, "Estado", "Cantidad", "Leads por estado")
                if "Asesor asignado" in df.columns:
                    st.markdown("#### Leads por asesor")
                    conteo = df["Asesor asignado"].value_counts().rename_axis("Asesor").reset_index(name="Cantidad")
                    grafico_barras(conteo, "Asesor", "Cantidad", "Leads asignados por asesor")
            else:
                st.warning("No se encontraron registros en la hoja Clientes.")
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
                mostrar_tabla(df)
                st.markdown("#### Calificaciones de atención al cliente")
                if col_score and not df.empty:
                    distrib = df[col_score].value_counts().sort_index().rename_axis("Calificación").reset_index(name="Cantidad")
                    grafico_barras(distrib, "Calificación", "Cantidad", "Distribución de calificaciones (0 a 5)")
                if col_score and not df.empty:
                    st.metric("Promedio del filtro actual", f"{df[col_score].mean():.2f}/5")
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
                mostrar_tabla(df)
                if "Producto" in df.columns and "Cantidad (unid.)" in df.columns:
                    ventas_prod = df.groupby("Producto", dropna=False)["Cantidad (unid.)"].sum().reset_index()
                    grafico_barras(ventas_prod, "Producto", "Cantidad (unid.)", "Unidades vendidas por producto")


# =========================================================
# PRODUCCIÓN
# =========================================================
elif seccion == "Producción":
    titulo_seccion("🧁 Producción", "Consulta de planificación agregada, materiales, capacidad, rutas y distribución.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        apartados = [
            "Demanda y pronóstico", "PAP", "PMP – Vainilla", "PMP – Chocolate",
            "Maestro de materiales", "BOM – Vainilla", "BOM – Chocolate",
            "TLO", "BOC", "CRP", "Puestos de trabajo",
            "Hoja de ruta – Vainilla", "Hoja de ruta – Chocolate",
            "MRP II FINAL", "DRP – Vainilla", "DRP – Chocolate"
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
                fila_total = df[df.iloc[:, 0].astype(str).str.contains("Costo Total", case=False, na=False)]
                if not fila_total.empty:
                    long = fila_total.melt(id_vars=fila_total.columns[0], var_name="Plan", value_name="Costo")
                    long["Costo"] = pd.to_numeric(long["Costo"], errors="coerce")
                    grafico_barras(long, "Plan", "Costo", "Costo total por estrategia PAP")
        elif opcion.startswith("PMP"):
            rango = "B70:Y78" if "Vainilla" in opcion else "B82:Y90"
            st.subheader(opcion)
            df = rango_df("PRODUCCION (RESUMEN)", rango, True)
            mostrar_tabla(df, 400)
            if not df.empty:
                fila = df[df.iloc[:, 0].astype(str).str.strip().str.upper() == "PMP"]
                if not fila.empty:
                    datos = fila.iloc[0, 1:].reset_index()
                    datos.columns = ["Semana", "Unidades"]
                    datos["Unidades"] = pd.to_numeric(datos["Unidades"], errors="coerce")
                    datos["Semana"] = range(1, len(datos) + 1)
                    grafico_linea(datos, "Semana", "Unidades", "Programa maestro de producción")
        elif opcion == "Maestro de materiales":
            st.subheader("Maestro de materiales")
            df = rango_df("PRODUCCION (RESUMEN)", "B95:L115", True)
            mostrar_tabla(df, 500)
            st.caption("La tabla se presenta desde la hoja PRODUCCION (RESUMEN), rango B95:L115.")
        elif opcion == "BOM – Vainilla":
            st.subheader("Lista de materiales – Torta de vainilla")
            df = rango_df("PRODUCCION (RESUMEN)", "B118:H134", True)
            mostrar_tabla(df)
        elif opcion == "BOM – Chocolate":
            st.subheader("Lista de materiales – Torta de chocolate")
            df = rango_df("PRODUCCION (RESUMEN)", "B137:H149", True)
            mostrar_tabla(df)
        elif opcion == "TLO":
            st.subheader("Tabla de lanzamiento de órdenes (TLO)")
            df = rango_df("PRODUCCION (RESUMEN)", "B152:K172", True)
            mostrar_tabla(df, 450)
        elif opcion == "BOC":
            st.subheader("Balance de operaciones y carga (BOC)")
            df = rango_df("PRODUCCION (RESUMEN)", "B176:K210", False)
            mostrar_tabla(df, 500)
            st.caption("Se conservan los dos niveles de encabezado: estación de trabajo y tipo de recurso.")
        elif opcion == "CRP":
            st.subheader("Planificación de requerimientos de capacidad (CRP)")
            df = rango_df("PRODUCCION (RESUMEN)", "B212:K241", False)
            mostrar_tabla(df, 560)
            st.info("El horizonte de planificación considerado en el proyecto corresponde a noviembre y diciembre, con nueve semanas: Sem. 00 a Sem. 08.")
        elif opcion == "Puestos de trabajo":
            st.subheader("Lista de puestos de trabajo")
            df = rango_df("MRP I", "A38:I42", False)
            mostrar_tabla(df)
        elif opcion == "Hoja de ruta – Vainilla":
            st.subheader("Hoja de ruta – Torta de vainilla")
            df = rango_df("MRP I", "A56:K65", False)
            mostrar_tabla(df, 400)
        elif opcion == "Hoja de ruta – Chocolate":
            st.subheader("Hoja de ruta – Torta de chocolate")
            df = rango_df("MRP I", "A71:K80", False)
            mostrar_tabla(df, 400)
        elif opcion == "MRP II FINAL":
            st.subheader("Planificación de recursos de manufactura – MRP II FINAL")
            st.caption("Se consulta exclusivamente la hoja MRP II FINAL.")
            libro = libro_actual(data_only=True)
            ws = libro["MRP II FINAL"]
            # Se muestran las secciones pobladas de MRP II FINAL sin incorporar MRP II AJUSTADO.
            for rango, nombre in [
                ("B2:K20", "Lanzamiento y planificación"),
                ("B22:K42", "Planificación de recursos"),
                ("B44:K62", "Programación y requerimientos"),
                ("B64:K82", "Resumen de planificación")
            ]:
                st.markdown(f"#### {nombre}")
                df = rango_df("MRP II FINAL", rango, False)
                mostrar_tabla(df, 320)
        elif opcion == "DRP – Vainilla":
            st.subheader("Planificación de distribución – Torta de vainilla")
            df = rango_df("PRODUCCION (RESUMEN)", "B244:K251", True)
            mostrar_tabla(df)
            st.caption("Se presenta la planificación por seis sucursales y la fila Fábrica se conserva como referencia de producción.")
        elif opcion == "DRP – Chocolate":
            st.subheader("Planificación de distribución – Torta de chocolate")
            df = rango_df("PRODUCCION (RESUMEN)", "B253:K260", True)
            mostrar_tabla(df)
            st.caption("Se presenta la planificación por seis sucursales y la fila Fábrica se conserva como referencia de producción.")


# =========================================================
# INVENTARIOS
# =========================================================
elif seccion == "Inventarios":
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
                column_config={"Stock actual": st.column_config.NumberColumn("Stock actual", min_value=0, step=1, format="%.2f")}
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
                            st.success(f"Movimiento registrado. Nuevo stock: {nuevo_stock:,.2f}")
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
                    st.metric("Calificación promedio", f"{df[cal_col].mean():.2f}/5")
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
                    st.metric("Monto de las órdenes filtradas", f"S/ {df['Monto (S/)'].sum():,.2f}")


# =========================================================
# DASHBOARD
# =========================================================
elif seccion == "Dashboard":
    titulo_seccion("📊 Dashboard de PASTELPAN", "Indicadores principales y complementarios del ERP.")
    if excel is None:
        st.warning("Carga el archivo Excel maestro.")
    else:
        indicadores = obtener_indicadores()
        if not indicadores:
            st.warning("No se encontraron los cuatro indicadores principales en INDICADORES ERP.")
        else:
            nombres = list(indicadores.keys())
            cols = st.columns(4)
            for i, nombre in enumerate(nombres[:4]):
                dato = indicadores[nombre]
                valor = dato["resultado"]
                unidad = str(dato["unidad"]).lower()
                if "porcentaje" in unidad:
                    texto = f"{valor*100:.2f}%"
                elif "puntos" in unidad:
                    texto = f"{valor:.2f}/5"
                else:
                    texto = f"{valor:,.0f}"
                cols[i].metric(nombre, texto)
                with cols[i].popover("Interpretación"):
                    st.write(dato["interpretacion"])
                    st.caption(f"Fuente: {dato['fuente']}")
            st.divider()
            st.subheader("Interpretación del indicador de capacidad")
            st.write(
                "El 82.23% corresponde al horizonte de planificación de noviembre y diciembre, "
                "considerando nueve semanas (Sem. 00 a Sem. 08). El indicador relaciona las horas "
                "requeridas acumuladas del plan con las 4,032 horas de capacidad consignadas para "
                "el periodo en el CRP del proyecto."
            )
            st.caption("El Dashboard toma los resultados principales de la hoja INDICADORES ERP para mantener coherencia con el archivo Excel.")
            st.divider()

        # Complementarios, se leen de la hoja de indicadores desde fila 13.
        df_comp = rango_df("INDICADORES ERP", "A12:G48", True)
        if not df_comp.empty:
            st.subheader("Indicadores complementarios")
            df_comp = df_comp.dropna(how="all")
            mostrar_tabla(df_comp, 360)

        st.divider()
        st.subheader("Análisis por módulo")
        c1, c2 = st.columns(2)
        with c1:
            clientes = limpiar_df(hoja_df("Clientes", 3, 4, 603, 1, 10))
            if not clientes.empty and "Estado" in clientes.columns:
                conteo = clientes["Estado"].fillna("Sin estado").value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                grafico_barras(conteo, "Estado", "Cantidad", "Leads por estado")
            proveedores = limpiar_df(hoja_df("Proveedores", 3, 4, 19, 1, 6))
            if not proveedores.empty and "Calificación" in proveedores.columns:
                proveedores["Calificación"] = pd.to_numeric(proveedores["Calificación"], errors="coerce")
                grafico_barras(proveedores, "Nombre del proveedor", "Calificación", "Calificación por proveedor")
        with c2:
            atencion = limpiar_df(hoja_df("Atención al cliente", 3, 4, 1839, 1, 10))
            if not atencion.empty and "Calificación NPS (0-5)" in atencion.columns:
                atencion["Calificación NPS (0-5)"] = pd.to_numeric(atencion["Calificación NPS (0-5)"], errors="coerce")
                distrib = atencion["Calificación NPS (0-5)"].value_counts().sort_index().rename_axis("Calificación").reset_index(name="Cantidad")
                grafico_barras(distrib, "Calificación", "Cantidad", "Atención al cliente – escala 0 a 5")
            ordenes = limpiar_df(hoja_df("Órdenes de Compra y Servicio", 3, 4, 23, 1, 7))
            if not ordenes.empty and "Estado" in ordenes.columns:
                conteo = ordenes["Estado"].value_counts().rename_axis("Estado").reset_index(name="Cantidad")
                grafico_barras(conteo, "Estado", "Cantidad", "Órdenes por estado")


# =========================================================
# PIE DE PÁGINA Y DESCARGA GENERAL
# =========================================================
st.divider()
col_a, col_b = st.columns([3, 1])
with col_a:
    st.caption("PASTELPAN ERP | Una tentación irresistible")
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
