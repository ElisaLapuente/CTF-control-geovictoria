import streamlit as st
import pandas as pd
import re
from io import BytesIO

# ---------------------------------------------------
# CONFIGURACIÓN
# ---------------------------------------------------

st.set_page_config(
    page_title="Control Geovictoria",
    page_icon="📊",
    layout="wide"
)

SUELDO_BASE = 539000
JORNADA_SEMANAL = 40

VALOR_HORA = SUELDO_BASE / 30 * 7 / JORNADA_SEMANAL
VALOR_MINUTO = VALOR_HORA / 60


# ---------------------------------------------------
# FUNCIONES
# ---------------------------------------------------

def convertir_a_minutos(valor):
    """
    Convierte valores como:
    -1 h 25 min
    -0 h 40 min
    1:25
    -00:40
    números negativos
    """

    if pd.isna(valor):
        return 0

    texto = str(valor).strip().lower()

    if texto in ["", "0", "0 h 00 min", "0:00", "nan"]:
        return 0

    # Formato: -1 h 25 min
    patron_horas = re.search(
        r"(-?\d+)\s*h\s*(\d+)\s*min",
        texto
    )

    if patron_horas:
        horas = int(patron_horas.group(1))
        minutos = int(patron_horas.group(2))

        signo = -1 if "-" in texto else 1

        return signo * (abs(horas) * 60 + minutos)

    # Formato: -01:25
    patron_hora = re.search(
        r"(-?)(\d+):(\d+)",
        texto
    )

    if patron_hora:
        signo = -1 if patron_hora.group(1) == "-" else 1
        horas = int(patron_hora.group(2))
        minutos = int(patron_hora.group(3))

        return signo * (horas * 60 + minutos)

    # Si es un número
    try:
        return float(valor)
    except:
        return 0


def minutos_a_texto(minutos):
    """
    Convierte minutos a formato:
    -2 h 18 min
    0 h 00 min
    """

    minutos = int(round(minutos))

    signo = "-" if minutos < 0 else ""

    minutos_absolutos = abs(minutos)
    horas = minutos_absolutos // 60
    minutos_restantes = minutos_absolutos % 60

    return f"{signo}{horas} h {minutos_restantes:02d} min"


def calcular_descuento(minutos_negativos):
    """
    Calcula el descuento estimado según:
    sueldo base / 30 * 7 / 40 horas
    """

    minutos_negativos = abs(minutos_negativos)

    return minutos_negativos * VALOR_MINUTO


def buscar_columna(df, nombres_posibles):
    """
    Busca una columna aunque tenga diferencias
    de mayúsculas, espacios o tildes.
    """

    columnas = list(df.columns)

    for columna in columnas:
        columna_limpia = (
            str(columna)
            .strip()
            .lower()
            .replace("ó", "o")
            .replace("í", "i")
            .replace("é", "e")
            .replace("á", "a")
            .replace("ú", "u")
        )

        for nombre in nombres_posibles:
            nombre_limpio = (
                nombre.lower()
                .replace("ó", "o")
                .replace("í", "i")
                .replace("é", "e")
                .replace("á", "a")
                .replace("ú", "u")
            )

            if nombre_limpio in columna_limpia:
                return columna

    return None


def procesar_archivo(archivo, numero_semana):
    """
    Lee y procesa un archivo Excel.
    """

    try:
        df = pd.read_excel(archivo)

        # Eliminar filas completamente vacías
        df = df.dropna(how="all")

        columna_ejecutivo = buscar_columna(
            df,
            [
                "ejecutivo",
                "colaborador",
                "nombre",
                "trabajador",
                "empleado"
            ]
        )

        columna_tienda = buscar_columna(
            df,
            [
                "tienda",
                "pdv",
                "local",
                "punto de venta"
            ]
        )

        columna_diferencia = buscar_columna(
            df,
            [
                "diferencia del día",
                "diferencia del dia",
                "total negativo",
                "diferencia",
                "saldo"
            ]
        )

        if columna_ejecutivo is None:
            st.warning(
                f"No se encontró la columna del ejecutivo en {archivo.name}. "
                f"Columnas encontradas: {list(df.columns)}"
            )
            return pd.DataFrame()

        if columna_tienda is None:
            df["Tienda"] = "Sin tienda"
            columna_tienda = "Tienda"

        if columna_diferencia is None:
            st.warning(
                f"No se encontró la columna de diferencia en {archivo.name}. "
                f"Columnas encontradas: {list(df.columns)}"
            )
            return pd.DataFrame()

        resultado = pd.DataFrame()

        resultado["Ejecutivo"] = (
            df[columna_ejecutivo]
            .fillna("Sin nombre")
            .astype(str)
            .str.strip()
        )

        resultado["Tienda"] = (
            df[columna_tienda]
            .fillna("Sin tienda")
            .astype(str)
            .str.strip()
        )

        resultado[f"Semana {numero_semana}"] = (
            df[columna_diferencia]
            .apply(convertir_a_minutos)
        )

        # Agrupar por ejecutivo y tienda
        resultado = (
            resultado
            .groupby(["Ejecutivo", "Tienda"], as_index=False)
            [f"Semana {numero_semana}"]
            .sum()
        )

        return resultado

    except Exception as error:
        st.error(
            f"No se pudo procesar el archivo {archivo.name}: {error}"
        )
        return pd.DataFrame()


# ---------------------------------------------------
# TÍTULO
# ---------------------------------------------------

st.title("📊 Control de Diferencias Geovictoria")

st.write(
    "Carga los archivos Excel de cada semana para acumular "
    "las diferencias de asistencia y calcular el descuento estimado."
)

st.info(
    f"Valor estimado por hora: ${VALOR_HORA:,.2f} | "
    f"Valor estimado por minuto: ${VALOR_MINUTO:,.2f}"
)


# ---------------------------------------------------
# CARGA DE ARCHIVOS
# ---------------------------------------------------

st.subheader("📁 Cargar archivos semanales")

archivos = st.file_uploader(
    "Selecciona uno o varios archivos Excel",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

if archivos:

    resultados_semanales = []

    for archivo in archivos:

        nombre_archivo = archivo.name.lower()

        # Detectar el número de semana por el nombre
        coincidencia = re.search(
            r"s(?:emana)?\s*([1-5])",
            nombre_archivo
        )

        if coincidencia:
            numero_semana = int(coincidencia.group(1))
        else:
            st.warning(
                f"No pude detectar la semana en {archivo.name}. "
                "Se asignará como Semana 1."
            )
            numero_semana = 1

        datos_semana = procesar_archivo(
            archivo,
            numero_semana
        )

        if not datos_semana.empty:
            resultados_semanales.append(datos_semana)

    if resultados_semanales:

        # Unir todas las semanas
        datos_completos = resultados_semanales[0]

        for datos in resultados_semanales[1:]:
            datos_completos = pd.merge(
                datos_completos,
                datos,
                on=["Ejecutivo", "Tienda"],
                how="outer"
            )

        # Crear siempre las columnas de Semana 1 a Semana 5
        for semana in range(1, 6):
            columna = f"Semana {semana}"

            if columna not in datos_completos.columns:
                datos_completos[columna] = 0

        columnas_semanas = [
            f"Semana {semana}" for semana in range(1, 6)
        ]

        # Reemplazar vacíos por cero
        datos_completos[columnas_semanas] = (
            datos_completos[columnas_semanas]
            .fillna(0)
        )

        # Total acumulado
        datos_completos["Total Negativo Minutos"] = (
            datos_completos[columnas_semanas]
            .sum(axis=1)
        )

        # Texto visible del total
        datos_completos["Total Negativo"] = (
            datos_completos["Total Negativo Minutos"]
            .apply(minutos_a_texto)
        )

        # Descuento estimado
        datos_completos["Descuento Estimado"] = (
            datos_completos["Total Negativo Minutos"]
            .apply(calcular_descuento)
        )

        # Ordenar desde el mayor negativo
        datos_completos = datos_completos.sort_values(
            by="Total Negativo Minutos"
        )

        # ---------------------------------------------------
        # FILTROS
        # ---------------------------------------------------

        st.subheader("🔎 Filtros")

        solo_criticos = st.checkbox(
            "Mostrar solo casos críticos",
            value=False
        )

        if solo_criticos:
            datos_mostrar = datos_completos[
                datos_completos["Total Negativo Minutos"] <= -60
            ].copy()
        else:
            datos_mostrar = datos_completos.copy()

        # ---------------------------------------------------
        # RESUMEN
        # ---------------------------------------------------

        st.subheader("📌 Resumen general")

        total_ejecutivos = len(datos_mostrar)

        total_minutos = abs(
            datos_mostrar["Total Negativo Minutos"]
            .sum()
        )

        total_descuentos = datos_mostrar[
            "Descuento Estimado"
        ].sum()

        casos_criticos = len(
            datos_completos[
                datos_completos["Total Negativo Minutos"] <= -60
            ]
        )

        col1, col2, col3, col4 = st.columns(4)

        col1.metric(
            "Total ejecutivos",
            total_ejecutivos
        )

        col2.metric(
            "Horas negativas",
            minutos_a_texto(-total_minutos)
        )

        col3.metric(
            "Descuento estimado",
            f"${total_descuentos:,.0f}"
        )

        col4.metric(
            "Casos críticos",
            casos_criticos
        )

        # ---------------------------------------------------
        # TABLA FINAL
        # ---------------------------------------------------

        st.subheader("📋 Detalle por ejecutivo")

        columnas_mostrar = [
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

        tabla_final = datos_mostrar[columnas_mostrar].copy()

        for semana in range(1, 6):
            columna = f"Semana {semana}"

            tabla_final[columna] = (
                tabla_final[columna]
                .apply(minutos_a_texto)
            )

        tabla_final["Descuento Estimado"] = (
            tabla_final["Descuento Estimado"]
            .apply(lambda x: f"${x:,.0f}")
        )

        st.dataframe(
            tabla_final,
            use_container_width=True,
            hide_index=True
        )

        # ---------------------------------------------------
        # DESCARGAR RESULTADO
        # ---------------------------------------------------

        st.subheader("📥 Descargar resultado")

        archivo_salida = BytesIO()

        with pd.ExcelWriter(
            archivo_salida,
            engine="openpyxl"
        ) as escritor:

            tabla_final.to_excel(
                escritor,
                index=False,
                sheet_name="Control Geovictoria"
            )

        st.download_button(
            label="Descargar Excel acumulado",
            data=archivo_salida.getvalue(),
            file_name="Control_Geovictoria_Acumulado.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    else:
        st.warning(
            "No se pudo procesar ningún archivo. "
            "Revisa que el Excel tenga columnas de ejecutivo, tienda "
            "y diferencia."
        )

else:
    st.write(
        "👆 Sube los archivos Excel de las semanas que quieras revisar."
    )
