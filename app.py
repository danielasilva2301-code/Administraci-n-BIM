import io
from pathlib import Path

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.utils import get_column_letter


st.set_page_config(page_title="Control de Partidas", layout="wide")
st.title("🏗️ Control de Partidas, Recursos y Tiempos")

archivo_local = Path(__file__).resolve().with_name(
    "7.21.5_Anexo_N°_4_Presupuesto (1).xlsx"
)
archivo_subido = st.file_uploader("Sube otro Excel (.xlsx)", type=["xlsx"])

if archivo_subido is not None:
    contenido = archivo_subido.getvalue()
elif archivo_local.exists():
    contenido = archivo_local.read_bytes()
    st.caption(f"Archivo utilizado: {archivo_local.name}")
else:
    st.error(f"No se encontró el archivo Excel esperado: {archivo_local.name}")
    st.stop()

wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True, read_only=True)
hoja_predeterminada = (
    wb.sheetnames.index("PPTO MODIFICADO")
    if "PPTO MODIFICADO" in wb.sheetnames
    else 0
)
hoja = st.selectbox("Selecciona la hoja", wb.sheetnames, index=hoja_predeterminada)
datos = list(wb[hoja].values)

if not datos:
    st.error("La hoja seleccionada no contiene datos.")
    st.stop()

fila_inicio = (
    st.number_input(
        "Fila de encabezados",
        min_value=1,
        max_value=len(datos),
        value=min(10, len(datos)),
        step=1,
    )
    - 1
)

encabezados = []
for indice, valor in enumerate(datos[fila_inicio], start=1):
    nombre = str(valor).strip() if valor is not None else ""
    encabezados.append(nombre or f"Columna {get_column_letter(indice)}")

vistos = {}
for indice, nombre in enumerate(encabezados):
    vistos[nombre] = vistos.get(nombre, 0) + 1
    if vistos[nombre] > 1:
        encabezados[indice] = f"{nombre} ({vistos[nombre]})"

df = pd.DataFrame(datos[fila_inicio + 1 :], columns=encabezados).dropna(how="all")
columnas = list(df.columns)

if not columnas or df.empty:
    st.error("No se encontraron filas de datos debajo de los encabezados.")
    st.stop()


def predeterminado(opciones, nombres, fallback):
    for nombre in nombres:
        for opcion in opciones:
            if opcion.strip().casefold() == nombre.casefold():
                return opcion
    return fallback


c_partida = st.selectbox(
    "Columna Partida",
    columnas,
    index=columnas.index(
        predeterminado(columnas, ["DESIGNACIÓN", "DESIGNACION", "ITEM"], columnas[0])
    ),
)
c_cant = st.selectbox(
    "Columna Cantidad",
    columnas,
    index=columnas.index(predeterminado(columnas, ["CANT", "CANTIDAD"], columnas[0])),
)
c_rend = st.selectbox(
    "Columna Rendimiento",
    columnas,
    index=columnas.index(
        predeterminado(columnas, ["RENDIMIENTOS", "RENDIMIENTO"], columnas[0])
    ),
)

sin_agrupacion = "Sin agrupación"
opciones_color = [sin_agrupacion] + columnas
unidad = predeterminado(columnas, ["U", "UNIDAD"], sin_agrupacion)
c_color = st.selectbox(
    "Color / agrupación del gráfico",
    opciones_color,
    index=opciones_color.index(unidad),
)

st.info(
    "La hoja no tiene una columna de recurso independiente. "
    "El gráfico puede agruparse por unidad (U) u otra columna, o mostrarse sin agrupación."
)

modo_rendimiento = st.selectbox(
    "Unidad del rendimiento",
    ["Días por unidad (según el Excel)", "Horas por unidad"],
    index=0,
)
horas_por_dia = st.number_input(
    "Horas por jornada",
    min_value=1.0,
    max_value=24.0,
    value=8.0,
    step=0.5,
)

cantidad = pd.to_numeric(df[c_cant], errors="coerce")
rendimiento = pd.to_numeric(df[c_rend], errors="coerce")
filas_validas = cantidad.notna() & rendimiento.notna() & df[c_partida].notna()
resultado = df.loc[filas_validas].copy()
cantidad = cantidad.loc[filas_validas]
rendimiento = rendimiento.loc[filas_validas]

if modo_rendimiento == "Días por unidad (según el Excel)":
    resultado["Dias_Ocupacion"] = cantidad * rendimiento
    resultado["Horas_Ocupacion"] = resultado["Dias_Ocupacion"] * horas_por_dia
else:
    resultado["Horas_Ocupacion"] = cantidad * rendimiento
    resultado["Dias_Ocupacion"] = resultado["Horas_Ocupacion"] / horas_por_dia

if resultado.empty:
    st.warning("No hay partidas con cantidad y rendimiento numéricos.")
    st.stop()

st.plotly_chart(
    px.bar(
        resultado.loc[resultado["Dias_Ocupacion"] > 0],
        x="Dias_Ocupacion",
        y=c_partida,
        color=None if c_color == sin_agrupacion else c_color,
        orientation="h",
        title="Tiempo de Ocupación por Partida (Días)",
        labels={"Dias_Ocupacion": "Días de ocupación", c_partida: "Partida"},
    ),
    use_container_width=True,
)

columnas_tabla = [c_partida]
if c_color != sin_agrupacion and c_color != c_partida:
    columnas_tabla.append(c_color)
columnas_tabla.extend([c_cant, c_rend, "Horas_Ocupacion", "Dias_Ocupacion"])
st.dataframe(resultado[columnas_tabla], use_container_width=True)
