import streamlit as st
import pandas as pd
from datetime import time, datetime


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Control Geovictoria",
    page_icon="⏱️",
    layout="wide"
)


# ============================================================
# DATOS DEL CÁLCULO
# ============================================================

SUELDO_BASE = 539000
JORNADA_SEMANAL = 40

VALOR_HORA = (
    SUELDO_BASE / 30 * 7
) / JORNADA_SEMANAL

VALOR_MINUTO = VALOR_HORA / 60


# ============================================================
# FUNCIONES DE TEXTO
# ============================================================

def limpiar_texto(valor):

    if pd.isna(valor):
        return ""

    return str(valor).strip()


def normalizar_texto(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor)

    texto = texto.replace("\n", " ")
    texto = texto.replace("\xa0", " ")

    texto = " ".join(
        texto.strip().split()
    )

    return texto.lower()


def buscar_columna(columnas, nombre):

    objetivo = normalizar_texto(nombre)

    for columna in columnas:

        if normalizar_texto(columna) == objetivo:
            return columna

    return None


# ============================================================
# CONVERSIÓN DE DIFERENCIAS A MINUTOS
# ============================================================

def convertir_minutos(valor):

    if valor is None:
        return 0

    try:

        if pd.isna(valor):
            return 0

    except Exception:
        pass


    # --------------------------------------------------------
    # Timedelta de pandas
    # --------------------------------------------------------

    if isinstance(valor, pd.Timedelta):

        return int(
            round(
                valor.total_seconds() / 60
            )
        )


    # --------------------------------------------------------
    # time
    # --------------------------------------------------------

    if isinstance(valor, time):

        total = (
            valor.hour * 60
            + valor.minute
        )

        if valor.second >= 30:
            total += 1

        return total


    # --------------------------------------------------------
    # datetime
    # --------------------------------------------------------

    if isinstance(valor, datetime):

        return (
            valor.hour * 60
            + valor.minute
        )


    # --------------------------------------------------------
    # Números
    # --------------------------------------------------------

    if isinstance(valor, (int, float)):

        if valor == 0:
            return 0

        # Excel guarda horas como fracción de día
        if abs(valor) < 1:

            return int(
                round(
                    valor * 24 * 60
                )
            )

        return int(
            round(valor)
        )


    # --------------------------------------------------------
    # Texto
    # --------------------------------------------------------

    texto = str(valor).strip()

    if texto == "":
        return 0


    # Normalizar posibles espacios
    texto = texto.replace(
        "\xa0",
        " "
    ).strip()


    # --------------------------------------------------------
    # Detectar signo negativo
    # --------------------------------------------------------

    negativo = texto.startswith("-")

    texto = texto.lstrip("-").strip()


    # --------------------------------------------------------
    # Formato HH:MM
    # Formato HH:MM:SS
    # --------------------------------------------------------

    partes = texto.split(":")


    try:

        if len(partes) == 2:

            horas = int(
                partes[0]
            )

            minutos = int(
                partes[1]
            )

            total = (
                horas * 60
                + minutos
            )


            if negativo:
                total = -total


            return total


        if len(partes) == 3:

            horas = int(
                partes[0]
            )

            minutos = int(
                partes[1]
            )

            segundos = int(
                float(partes[2])
            )


            total = (
                horas * 60
                + minutos
            )


            if segundos >= 30:
                total += 1


            if negativo:
                total = -total


            return total


    except Exception:
        return 0


    return 0


# ============================================================
# MINUTOS A TEXTO
# ============================================================

def minutos_a_texto(minutos):

    minutos = int(
        round(minutos)
    )

    if minutos == 0:
        return "0 h 00 min"


    signo = ""

    if minutos < 0:
        signo = "-"


    minutos = abs(minutos)

    horas = minutos // 60
    minutos_restantes = minutos % 60


    return (
        f"{signo}"
        f"{horas} h "
        f"{minutos_restantes:02d} min"
    )


# ============================================================
# DETERMINAR SEMANA
# ============================================================

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


# ============================================================
# BUSCAR FILA DE ENCABEZADOS
# ============================================================

def encontrar_fila_encabezados(archivo):

    vista = pd.read_excel(
        archivo,
        header=None,
        nrows=30
    )


    for fila in range(
        len(vista)
    ):

        valores = [
            normalizar_texto(x)
            for x in vista.iloc[fila].tolist()
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
            "diferencia del día"
            in valores
        )


        if (
            tiene_nombre
            and tiene_apellidos
            and tiene_grupo
            and tiene_fecha
            and tiene_diferencia
        ):

            return fila


    return None


# ============================================================
# LEER EXCEL
# ============================================================

def leer_excel(archivo):

    fila_encabezados = (
        encontrar_fila_encabezados(
            archivo
        )
    )


    if fila_encabezados is None:

        st.error(
            "No pude encontrar la fila "
            "de encabezados del Excel."
        )

        st.stop()


    # --------------------------------------------------------
    # Leer Excel con encabezado correcto
    # --------------------------------------------------------

    df = pd.read_excel(
        archivo,
        header=fila_encabezados
    )


    # --------------------------------------------------------
    # Eliminar columnas completamente vacías
    # --------------------------------------------------------

    df = df.dropna(
        axis=1,
        how="all"
    )


    columnas = list(
        df.columns
    )


    # --------------------------------------------------------
    # Buscar columnas
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # Verificación
    # --------------------------------------------------------

    faltantes = []


    if col_nombre is None:
        faltantes.append("Nombre")


    if col_apellidos is None:
        faltantes.append("Apellidos")


    if col_grupo is None:
        faltantes.append("Grupo")


    if col_fecha is None:
        faltantes.append("Fecha")


    if col_diferencia is None:
        faltantes.append(
            "Diferencia del día"
        )


    if faltantes:

        st.error(
            "No se encontraron estas columnas: "
            + ", ".join(faltantes)
        )

        st.write(
            "Columnas encontradas:"
        )

        st.write(
            columnas
        )

        st.stop()


    # ========================================================
    # EJECUTIVO
    # ========================================================

    df["Ejecutivo"] = (
        df[col_nombre]
        .apply(limpiar_texto)
        + " "
        + df[col_apellidos]
        .apply(limpiar_texto)
    ).str.strip()


    # ========================================================
    # TIENDA
    # ========================================================

    df["Tienda"] = (
        df[col_grupo]
        .apply(limpiar_texto)
    )


    # ========================================================
    # FECHA
    # ========================================================

    df["Fecha"] = pd.to_datetime(
        df[col_fecha],
        errors="coerce",
        dayfirst=True
    )


    # ========================================================
    # DIFERENCIA DEL DÍA
    #
    # ESTA ES LA ÚNICA COLUMNA UTILIZADA PARA:
    # - HORAS FALTANTES
    # - HORAS EXTRAS
    #
    # NO SE USA:
    # "Diferencia de ingreso REAL"
    # ========================================================

    df["Diferencia Minutos"] = (
        df[col_diferencia]
        .apply(convertir_minutos)
    )


    # ========================================================
    # SEMANA
    # ========================================================

    df["Semana"] = (
        df["Fecha"]
        .apply(obtener_semana)
    )


    # ========================================================
    # FALTANTES
    # ========================================================

    df["Faltantes Minutos"] = (
        df["Diferencia Minutos"]
        .apply(
            lambda x:
                abs(x)
                if x < 0
                else 0
        )
    )


    # ========================================================
    # HORAS EXTRAS
    # ========================================================

    df["Extras Minutos"] = (
        df["Diferencia Minutos"]
        .apply(
            lambda x:
                x
                if x > 0
                else 0
        )
    )


    # ========================================================
    # ELIMINAR FILAS SIN EJECUTIVO
    # ========================================================

    df = df[
        df["Ejecutivo"].str.strip() != ""
    ].copy()


    return df


# ============================================================
# TÍTULO
# ============================================================

st.title(
    "⏱️ Control de Marcaciones Geovictoria"
)


st.write(
    "Carga el Excel mensual de Geovictoria "
    "para revisar horas faltantes, horas extras "
    "y descuentos estimados."
)


# ============================================================
# CARGAR ARCHIVO
# ============================================================

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


# ============================================================
# PROCESAR
# ============================================================

try:

    df = leer_excel(
        archivo
    )

except Exception as error:

    st.error(
        "Ocurrió un error al procesar el Excel."
    )

    st.exception(
        error
    )

    st.stop()


# ============================================================
# FILTROS
# ============================================================

st.subheader(
    "🔎 Filtros"
)


col1, col2 = st.columns(2)


with col1:

    lista_tiendas = sorted(
        [
            x
            for x in df["Tienda"]
            .dropna()
            .unique()
            if str(x).strip() != ""
        ]
    )


    tiendas = [
        "Todas"
    ] + lista_tiendas


    tienda_seleccionada = (
        st.selectbox(
            "Tienda",
            tiendas
        )
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


    semana_seleccionada = (
        st.selectbox(
            "Semana",
            semanas
        )
    )


# ============================================================
# APLICAR FILTROS
# ============================================================

df_filtrado = df.copy()


if (
    tienda_seleccionada
    != "Todas"
):

    df_filtrado = df_filtrado[
        df_filtrado["Tienda"]
        == tienda_seleccionada
    ]


if (
    semana_seleccionada
    != "Todas"
):

    df_filtrado = df_filtrado[
        df_filtrado["Semana"]
        == semana_seleccionada
    ]


# ============================================================
# RESUMEN
# ============================================================

st.subheader(
    "📊 Resumen"
)


total_ejecutivos = (
    df_filtrado["Ejecutivo"]
    .nunique()
)


total_faltantes = int(
    df_filtrado[
        "Faltantes Minutos"
    ].sum()
)


total_extras = int(
    df_filtrado[
        "Extras Minutos"
    ].sum()
)


# ============================================================
# CASOS CRÍTICOS
# ============================================================

faltantes_por_ejecutivo = (
    df_filtrado
    .groupby(
        "Ejecutivo"
    )["Faltantes Minutos"]
    .sum()
)


casos_criticos = int(
    (
        faltantes_por_ejecutivo
        >= 60
    ).sum()
)


# ============================================================
# DESCUENTO
# ============================================================

descuento_total = (
    total_faltantes
    * VALOR_MINUTO
)


# ============================================================
# TARJETAS
# ============================================================

c1, c2, c3, c4, c5 = (
    st.columns(5)
)


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


# ============================================================
# TABLA FALTANTES
# ============================================================

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


# ------------------------------------------------------------
# Agregar semanas que todavía no existen
# ------------------------------------------------------------

for semana in semanas_tabla:

    if semana not in tabla.columns:

        tabla[semana] = 0


# ------------------------------------------------------------
# Asegurar valores numéricos
# ------------------------------------------------------------

for semana in semanas_tabla:

    tabla[semana] = pd.to_numeric(
        tabla[semana],
        errors="coerce"
    ).fillna(0)


# ============================================================
# TOTAL NEGATIVO
# ============================================================

tabla["Total Negativo Minutos"] = (
    tabla[semanas_tabla]
    .sum(axis=1)
)


# ============================================================
# DESCUENTO INDIVIDUAL
# ============================================================

tabla["Descuento Estimado Num"] = (
    tabla[
        "Total Negativo Minutos"
    ]
    * VALOR_MINUTO
)


# ============================================================
# TABLA PARA MOSTRAR
# ============================================================

tabla_mostrar = tabla[
    [
        "Ejecutivo",
        "Tienda",
        "Semana 1",
        "Semana 2",
        "Semana 3",
        "Semana 4",
        "Semana 5",
        "Total Negativo Minutos",
        "Descuento Estimado Num"
    ]
].copy()


for semana in semanas_tabla:

    tabla_mostrar[semana] = (
        tabla_mostrar[semana]
        .apply(minutos_a_texto)
    )


tabla_mostrar[
    "Total Negativo Minutos"
] = (
    tabla_mostrar[
        "Total Negativo Minutos"
    ]
    .apply(minutos_a_texto)
)


tabla_mostrar[
    "Descuento Estimado Num"
] = (
    tabla_mostrar[
        "Descuento Estimado Num"
    ]
    .apply(
        lambda x:
            f"${x:,.0f}"
            .replace(",", ".")
    )
)


tabla_mostrar = (
    tabla_mostrar.rename(
        columns={
            "Total Negativo Minutos":
                "Total Negativo",
            "Descuento Estimado Num":
                "Descuento Estimado"
        }
    )
)


st.dataframe(
    tabla_mostrar,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# HORAS EXTRAS
# ============================================================

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
].copy()


st.dataframe(
    extras_mostrar,
    use_container_width=True,
    hide_index=True
)


# ============================================================
# DETALLE DIARIO
# ============================================================

st.subheader(
    "📅 Detalle diario"
)


detalle = df_filtrado[
    [
        "Fecha",
        "Semana",
        "Ejecutivo",
        "Tienda",
        "Diferencia Minutos",
        "Faltantes Minutos",
        "Extras Minutos"
    ]
].copy()


detalle["Diferencia del día"] = (
    detalle[
        "Diferencia Minutos"
    ]
    .apply(minutos_a_texto)
)


detalle["Horas faltantes"] = (
    detalle[
        "Faltantes Minutos"
    ]
    .apply(minutos_a_texto)
)


detalle["Horas extras"] = (
    detalle[
        "Extras Minutos"
    ]
    .apply(minutos_a_texto)
)


detalle = detalle.drop(
    columns=[
        "Diferencia Minutos",
        "Faltantes Minutos",
        "Extras Minutos"
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


# ============================================================
# INFORMACIÓN DEL CÁLCULO
# ============================================================

with st.expander(
    "ℹ️ Ver cálculo del descuento"
):

    st.write(
        f"Sueldo base: ${SUELDO_BASE:,.0f}"
        .replace(",", ".")
    )

    st.write(
        f"Jornada semanal: "
        f"{JORNADA_SEMANAL} horas"
    )

    st.write(
        f"Valor hora aproximado: "
        f"${VALOR_HORA:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )

    st.write(
        f"Valor minuto aproximado: "
        f"${VALOR_MINUTO:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# ============================================================
# DESCARGAR EXCEL
# ============================================================

st.subheader(
    "📥 Descargar resultados"
)


nombre_archivo = (
    "Control_Geovictoria_Resultados.xlsx"
)


# ------------------------------------------------------------
# Crear archivo Excel en memoria
# ------------------------------------------------------------

import io


salida = io.BytesIO()


with pd.ExcelWriter(
    salida,
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


salida.seek(0)


st.download_button(
    label="⬇️ Descargar Excel",
    data=salida,
    file_name=nombre_archivo,
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    )
)
