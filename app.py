import streamlit as st
import pandas as pd
from datetime import time


# =========================================================
# CONFIGURACIÓN
# =========================================================

st.set_page_config(
    page_title="Control Geovictoria",
    page_icon="⏱️",
    layout="wide"
)


# =========================================================
# FUNCIONES
# =========================================================

def limpiar_texto(valor):
    if pd.isna(valor):
        return ""

    return str(valor).strip()


def normalizar_columna(nombre):
    texto = str(nombre).strip().lower()

    texto = texto.replace("\n", " ")
    texto = " ".join(texto.split())

    return texto


def buscar_columna(columnas, nombre_buscado):
    objetivo = normalizar_columna(nombre_buscado)

    for columna in columnas:

        if normalizar_columna(columna) == objetivo:
            return columna

    return None


def convertir_minutos(valor):

    if pd.isna(valor):
        return 0

    # Timedelta de pandas
    if isinstance(valor, pd.Timedelta):

        return int(
            round(
                valor.total_seconds() / 60
            )
        )

    # Hora de Python
    if isinstance(valor, time):

        return (
            valor.hour * 60
            + valor.minute
            + round(valor.second / 60)
        )

    # Números
    if isinstance(valor, (int, float)):

        if valor == 0:
            return 0

        # Excel puede representar una hora como
        # fracción de día
        if abs(valor) < 1:

            return int(
                round(
                    valor * 24 * 60
                )
            )

        return int(round(valor))

    texto = str(valor).strip()

    if texto == "":
        return 0

    # Casos como:
    # -03:35
    # -00:02
    # 00:00
    # 03:35
    negativo = texto.startswith("-")

    texto = texto.replace("-", "").strip()

    partes = texto.split(":")

    try:

        if len(partes) == 2:

            horas = int(partes[0])
            minutos = int(partes[1])

            total = (
                horas * 60
                + minutos
            )

            if negativo:
                total = -total

            return total

        if len(partes) == 3:

            horas = int(partes[0])
            minutos = int(partes[1])
            segundos = int(
                float(partes[2])
            )

            total = (
                horas * 60
                + minutos
                + round(segundos / 60)
            )

            if negativo:
                total = -total

            return total

    except Exception:
        return 0

    return 0


def minutos_a_texto(minutos):

    minutos = int(
        round(minutos)
    )

    signo = ""

    if minutos < 0:
        signo = "-"

    minutos = abs(minutos)

    horas = minutos // 60
    mins = minutos % 60

    return (
        f"{signo}{horas} h {mins:02d} min"
    )


def obtener_semana(fecha):

    if pd.isna(fecha):
        return None

    dia = fecha.day

    if dia <= 6:
        return "Semana 1"

    if dia <= 13:
        return "Semana 2"

    if dia <= 20:
        return "Semana 3"

    if dia <= 27:
        return "Semana 4"

    return "Semana 5"


# =========================================================
# LEER EXCEL
# =========================================================

def leer_excel(archivo):

    # -----------------------------------------------------
    # PRIMERO: buscar automáticamente la fila de encabezados
    # -----------------------------------------------------

    vista = pd.read_excel(
        archivo,
        header=None,
        nrows=20
    )

    fila_encabezado = None

    for i in range(len(vista)):

        valores = [
            normalizar_columna(x)
            for x in vista.iloc[i].tolist()
        ]

        tiene_nombre = (
            "nombre" in valores
        )

        tiene_apellidos = (
            "apellidos" in valores
        )

        tiene_grupo = (
            "grupo" in valores
        )

        tiene_fecha = (
            "fecha" in valores
        )

        tiene_diferencia = (
            "diferencia del día" in valores
        )

        if (
            tiene_nombre
            and tiene_apellidos
            and tiene_grupo
            and tiene_fecha
            and tiene_diferencia
        ):

            fila_encabezado = i

            break

    # -----------------------------------------------------
    # SI NO ENCUENTRA LA FILA
    # -----------------------------------------------------

    if fila_encabezado is None:

        st.error(
            "No pude encontrar automáticamente "
            "los encabezados del Excel."
        )

        st.write(
            "Estas son las primeras filas del archivo:"
        )

        st.dataframe(
            vista,
            use_container_width=True
        )

        st.stop()

    # -----------------------------------------------------
    # LEER EL EXCEL CON LA FILA CORRECTA
    # -----------------------------------------------------

    df = pd.read_excel(
        archivo,
        header=fila_encabezado
    )

    columnas = list(df.columns)

    # -----------------------------------------------------
    # BUSCAR COLUMNAS
    # -----------------------------------------------------

    col_nombre = buscar_columna(
        columnas,
        "Nombre"
    )

    col_apellidos = buscar_columna(
        columnas,
        "Apellidos"
    )

    col_grupo = buscar_columna(
        columnas,
        "Grupo"
    )

    col_fecha = buscar_columna(
        columnas,
        "Fecha"
    )

    col_diferencia = buscar_columna(
        columnas,
        "Diferencia del día"
    )

    # -----------------------------------------------------
    # VERIFICAR COLUMNAS
    # -----------------------------------------------------

    faltantes_columnas = []

    if col_nombre is None:
        faltantes_columnas.append(
            "Nombre"
        )

    if col_apellidos is None:
        faltantes_columnas.append(
            "Apellidos"
        )

    if col_grupo is None:
        faltantes_columnas.append(
            "Grupo"
        )

    if col_fecha is None:
        faltantes_columnas.append(
            "Fecha"
        )

    if col_diferencia is None:
        faltantes_columnas.append(
            "Diferencia del día"
        )

    if faltantes_columnas:

        st.error(
            "No se encontraron estas columnas: "
            + ", ".join(
                faltantes_columnas
            )
        )

        st.write(
            "Columnas encontradas:"
        )

        st.write(columnas)

        st.stop()

    # -----------------------------------------------------
    # CREAR EJECUTIVO
    # -----------------------------------------------------

    df["Ejecutivo"] = (
        df[col_nombre]
        .apply(limpiar_texto)
        + " "
        + df[col_apellidos]
        .apply(limpiar_texto)
    ).str.strip()

    # -----------------------------------------------------
    # TIENDA
    # -----------------------------------------------------

    df["Tienda"] = (
        df[col_grupo]
        .apply(limpiar_texto)
    )

    # -----------------------------------------------------
    # FECHA
    # -----------------------------------------------------

    df["Fecha"] = pd.to_datetime(
        df[col_fecha],
        errors="coerce",
        dayfirst=True
    )

    # -----------------------------------------------------
    # DIFERENCIA DEL DÍA
    #
    # IMPORTANTE:
    # SOLO usamos esta columna para calcular faltantes
    # y horas extras.
    #
    # NO usamos "Diferencia de ingreso REAL".
    # -----------------------------------------------------

    df["Diferencia Minutos"] = (
        df[col_diferencia]
        .apply(convertir_minutos)
    )

    # -----------------------------------------------------
    # SEMANA AUTOMÁTICA
    # -----------------------------------------------------

    df["Semana"] = (
        df["Fecha"]
        .apply(obtener_semana)
    )

    # -----------------------------------------------------
    # HORAS FALTANTES
    #
    # Las diferencias negativas se convierten a positivas
    # para poder acumularlas como tiempo faltante.
    # -----------------------------------------------------

    df["Faltantes Minutos"] = (
        df["Diferencia Minutos"]
        .apply(
            lambda x:
                abs(x)
                if x < 0
                else 0
        )
    )

    # -----------------------------------------------------
    # HORAS EXTRAS
    #
    # Se mantienen completamente separadas.
    #
    # NO compensan las horas faltantes.
    # -----------------------------------------------------

    df["Extras Minutos"] = (
        df["Diferencia Minutos"]
        .apply(
            lambda x:
                x
                if x > 0
                else 0
        )
    )

    return df


# =========================================================
# TÍTULO
# =========================================================

st.title(
    "⏱️ Control de Marcaciones Geovictoria"
)

st.write(
    "Carga el Excel mensual de Geovictoria "
    "para revisar horas faltantes, horas extras "
    "y descuentos estimados."
)


# =========================================================
# CARGAR ARCHIVO
# =========================================================

archivo = st.file_uploader(
    "📂 Cargar Excel mensual",
    type=["xlsx"]
)


if archivo is None:

    st.info(
        "👆 Primero carga el archivo Excel "
        "de Geovictoria."
    )

    st.stop()


# =========================================================
# LEER ARCHIVO
# =========================================================

try:

    df = leer_excel(
        archivo
    )

except Exception as e:

    st.error(
        "Ocurrió un problema al leer el Excel."
    )

    st.exception(e)

    st.stop()


# =========================================================
# FILTROS
# =========================================================

st.subheader(
    "🔎 Filtros"
)

col1, col2 = st.columns(2)


with col1:

    tiendas = [
        "Todas"
    ] + sorted(
        [
            str(x)
            for x in df["Tienda"]
            .dropna()
            .unique()
            if str(x).strip() != ""
        ]
    )

    tienda_seleccionada = st.selectbox(
        "Tienda",
        tiendas
    )


with col2:

    semanas = [
        "Todas",
        "Semana 1",
        "Semana 2",
        "Semana 3",
        "Semana 4",
        "Semana 5"
    ]

    semana_seleccionada = st.selectbox(
        "Semana",
        semanas
    )


# =========================================================
# APLICAR FILTROS
# =========================================================

df_filtrado = df.copy()


if tienda_seleccionada != "Todas":

    df_filtrado = df_filtrado[
        df_filtrado["Tienda"]
        == tienda_seleccionada
    ]


if semana_seleccionada != "Todas":

    df_filtrado = df_filtrado[
        df_filtrado["Semana"]
        == semana_seleccionada
    ]


# =========================================================
# RESUMEN
# =========================================================

st.subheader(
    "📊 Resumen"
)


total_ejecutivos = (
    df_filtrado["Ejecutivo"]
    .nunique()
)


total_faltantes = (
    df_filtrado["Faltantes Minutos"]
    .sum()
)


total_extras = (
    df_filtrado["Extras Minutos"]
    .sum()
)


# Casos críticos:
# 1 hora o más de tiempo faltante acumulado

casos_criticos = (
    df_filtrado
    .groupby("Ejecutivo")
    ["Faltantes Minutos"]
    .sum()
)


casos_criticos = (
    casos_criticos[
        casos_criticos >= 60
    ]
    .count()
)


# =========================================================
# CÁLCULO DESCUENTO
# =========================================================

# Sueldo base:
# $539.000
#
# Fórmula definida:
#
# 539.000 / 30 * 7 / 40
#
# = valor hora

valor_hora = (
    539000 / 30 * 7
) / 40


valor_minuto = (
    valor_hora / 60
)


descuento_total = (
    total_faltantes
    * valor_minuto
)


# =========================================================
# TARJETAS DE RESUMEN
# =========================================================

c1, c2, c3, c4, c5 = st.columns(5)


with c1:

    st.metric(
        "👥 Ejecutivos",
        total_ejecutivos
    )


with c2:

    st.metric(
        "⏰ Horas faltantes",
        minutos_a_texto(
            total_faltantes
        )
    )


with c3:

    st.metric(
        "➕ Horas extras",
        minutos_a_texto(
            total_extras
        )
    )


with c4:

    st.metric(
        "💰 Descuento aprox.",
        f"${descuento_total:,.0f}"
        .replace(",", ".")
    )


with c5:

    st.metric(
        "⚠️ Casos críticos",
        casos_criticos
    )


# =========================================================
# TABLA DE HORAS FALTANTES
# =========================================================

st.subheader(
    "⏰ Horas faltantes por ejecutivo"
)


semanas_tabla = [
    "Semana 1",
    "Semana 2",
    "Semana 3",
    "Semana 4",
    "Semana 5"
]


if len(df_filtrado) > 0:

    faltantes = (
        df_filtrado
        .groupby(
            [
                "Ejecutivo",
                "Tienda",
                "Semana"
            ]
        )["Faltantes Minutos"]
        .sum()
        .reset_index()
    )

    tabla = (
        faltantes
        .pivot_table(
            index=[
                "Ejecutivo",
                "Tienda"
            ],
            columns="Semana",
            values="Faltantes Minutos",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

else:

    tabla = pd.DataFrame(
        columns=[
            "Ejecutivo",
            "Tienda"
        ]
    )


# ---------------------------------------------------------
# ASEGURAR LAS 5 SEMANAS
# ---------------------------------------------------------

for semana in semanas_tabla:

    if semana not in tabla.columns:

        tabla[semana] = 0


for semana in semanas_tabla:

    tabla[semana] = (
        tabla[semana]
        .fillna(0)
    )


# ---------------------------------------------------------
# TOTAL NEGATIVO
# ---------------------------------------------------------

tabla["Total Negativo"] = (
    tabla[semanas_tabla]
    .sum(axis=1)
)


# ---------------------------------------------------------
# DESCUENTO ESTIMADO
# ---------------------------------------------------------

tabla["Descuento Estimado"] = (
    tabla["Total Negativo"]
    * valor_minuto
)


# =========================================================
# TABLA PARA MOSTRAR
# =========================================================

tabla_mostrar = tabla.copy()


for semana in semanas_tabla:

    tabla_mostrar[semana] = (
        tabla_mostrar[semana]
        .apply(minutos_a_texto)
    )


tabla_mostrar["Total Negativo"] = (
    tabla["Total Negativo"]
    .apply(minutos_a_texto)
)


tabla_mostrar["Descuento Estimado"] = (
    tabla["Descuento Estimado"]
    .apply(
        lambda x:
            f"${x:,.0f}"
            .replace(",", ".")
    )
)


columnas_tabla = [
    "Ejecutivo",
    "Tienda",
    "Semana 1",
    "Semana 2",
    "Semana 3",
    "Semana 4",
    "Semana 5",
    "Total Negativo",
    "Descuento Estimado"
]


tabla_mostrar = (
    tabla_mostrar[
        columnas_tabla
    ]
)


st.dataframe(
    tabla_mostrar,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# HORAS EXTRAS
# =========================================================

st.subheader(
    "➕ Horas extras"
)


if len(df_filtrado) > 0:

    extras = (
        df_filtrado
        .groupby(
            [
                "Ejecutivo",
                "Tienda"
            ]
        )["Extras Minutos"]
        .sum()
        .reset_index()
    )

else:

    extras = pd.DataFrame(
        columns=[
            "Ejecutivo",
            "Tienda",
            "Extras Minutos"
        ]
    )


extras["Horas Extras"] = (
    extras["Extras Minutos"]
    .apply(minutos_a_texto)
)


extras_mostrar = extras[
    [
        "Ejecutivo",
        "Tienda",
        "Horas Extras"
    ]
]


st.dataframe(
    extras_mostrar,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# DETALLE DIARIO
# =========================================================

st.subheader(
    "📅 Detalle diario"
)


detalle = df_filtrado[
    [
        "Fecha",
        "Semana",
        "Ejecutivo",
        "Tienda",
        "Diferencia Minutos"
    ]
].copy()


detalle["Diferencia del día"] = (
    detalle["Diferencia Minutos"]
    .apply(minutos_a_texto)
)


detalle = detalle.drop(
    columns=[
        "Diferencia Minutos"
    ]
)


detalle = detalle.sort_values(
    [
        "Fecha",
        "Ejecutivo"
    ]
)


st.dataframe(
    detalle,
    use_container_width=True,
    hide_index=True
)


# =========================================================
# DESCARGAR EXCEL
# =========================================================

st.subheader(
    "📥 Descargar resultados"
)


nombre_archivo = (
    "Control_Geovictoria_Resultados.xlsx"
)


with pd.ExcelWriter(
    nombre_archivo,
    engine="openpyxl"
) as writer:

    tabla_mostrar.to_excel(
        writer,
        sheet_name="Faltantes",
        index=False
    )

    extras_mostrar.to_excel(
        writer,
        sheet_name="Horas Extras",
        index=False
    )

    detalle.to_excel(
        writer,
        sheet_name="Detalle Diario",
        index=False
    )


with open(
    nombre_archivo,
    "rb"
) as archivo_excel:

    st.download_button(
        label="⬇️ Descargar Excel",
        data=archivo_excel,
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
