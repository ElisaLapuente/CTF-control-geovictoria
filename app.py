import streamlit as st
import pandas as pd
import re
from io import BytesIO
from datetime import datetime, time

# ============================================================

# CONFIGURACIÓN

# ============================================================

st.set_page_config(
page_title="Control Geovictoria",
page_icon="📊",
layout="wide"
)

SUELDO_BASE = 539000
JORNADA_SEMANAL = 40

# Fórmula definida:

# sueldo / 30 * 7 / 40

VALOR_HORA = SUELDO_BASE / 30 * 7 / JORNADA_SEMANAL
VALOR_MINUTO = VALOR_HORA / 60

# ============================================================

# FUNCIONES GENERALES

# ============================================================

def limpiar_texto(valor):
"""Limpia textos y evita valores NaN."""
if pd.isna(valor):
return ""
return str(valor).strip()

def normalizar_columna(nombre):
"""Normaliza nombres de columnas para encontrarlos fácilmente."""
texto = str(nombre).strip().lower()

```
reemplazos = {
    "á": "a",
    "é": "e",
    "í": "i",
    "ó": "o",
    "ú": "u",
    "ü": "u",
    "ñ": "n"
}

for original, nuevo in reemplazos.items():
    texto = texto.replace(original, nuevo)

texto = re.sub(r"\s+", " ", texto)

return texto
```

def encontrar_columna(df, nombres):
"""Busca una columna por diferentes nombres posibles."""

```
columnas_normalizadas = {
    normalizar_columna(col): col
    for col in df.columns
}

for nombre in nombres:
    nombre_normalizado = normalizar_columna(nombre)

    if nombre_normalizado in columnas_normalizadas:
        return columnas_normalizadas[nombre_normalizado]

# Búsqueda parcial
for columna in df.columns:
    columna_normalizada = normalizar_columna(columna)

    for nombre in nombres:
        nombre_normalizado = normalizar_columna(nombre)

        if nombre_normalizado in columna_normalizada:
            return columna

return None
```

# ============================================================

# CONVERSIÓN DE HORAS

# ============================================================

def valor_a_minutos(valor):
"""
Convierte diferentes formatos de hora a minutos.

```
Ejemplos:
-03:35 -> -215
-00:02 -> -2
03:35 -> 215
-3 h 35 min -> -215
"""

if pd.isna(valor):
    return 0

if isinstance(valor, time):
    return valor.hour * 60 + valor.minute + valor.second / 60

if isinstance(valor, pd.Timedelta):
    return valor.total_seconds() / 60

if isinstance(valor, datetime):
    return valor.hour * 60 + valor.minute + valor.second / 60

texto = str(valor).strip().lower()

if texto in [
    "",
    "nan",
    "none",
    "nat",
    "-",
    "sin marca",
    "00:00",
    "0:00",
    "00:00:00",
    "0"
]:
    return 0

# ----------------------------------------
# Formato HH:MM:SS
# ----------------------------------------

patron_hms = re.match(
    r"^(-)?(\d+):(\d+):(\d+)$",
    texto
)

if patron_hms:
    signo = -1 if patron_hms.group(1) else 1

    horas = int(patron_hms.group(2))
    minutos = int(patron_hms.group(3))
    segundos = int(patron_hms.group(4))

    total = horas * 60 + minutos + segundos / 60

    return signo * total

# ----------------------------------------
# Formato HH:MM
# ----------------------------------------

patron_hm = re.match(
    r"^(-)?(\d+):(\d+)$",
    texto
)

if patron_hm:
    signo = -1 if patron_hm.group(1) else 1

    horas = int(patron_hm.group(2))
    minutos = int(patron_hm.group(3))

    total = horas * 60 + minutos

    return signo * total

# ----------------------------------------
# Formato -3 h 35 min
# ----------------------------------------

patron_texto = re.search(
    r"(-)?\s*(\d+)\s*h(?:oras?)?\s*(\d+)?\s*min",
    texto
)

if patron_texto:
    signo = -1 if patron_texto.group(1) else 1

    horas = int(patron_texto.group(2))
    minutos = int(patron_texto.group(3) or 0)

    return signo * (horas * 60 + minutos)

# ----------------------------------------
# Intentar número
# ----------------------------------------

try:
    return float(valor)
except:
    return 0
```

def minutos_a_texto(minutos):
"""Convierte minutos a formato legible."""

```
if pd.isna(minutos):
    minutos = 0

minutos = int(round(minutos))

signo = "-" if minutos < 0 else ""

minutos_abs = abs(minutos)

horas = minutos_abs // 60
minutos_restantes = minutos_abs % 60

return f"{signo}{horas} h {minutos_restantes:02d} min"
```

def minutos_positivos_a_texto(minutos):
"""Muestra horas positivas."""

```
minutos = abs(int(round(minutos)))

horas = minutos // 60
minutos_restantes = minutos % 60

return f"{horas} h {minutos_restantes:02d} min"
```

# ============================================================

# DESCUENTO

# ============================================================

def calcular_descuento(minutos):
"""
Calcula descuento SOLO sobre minutos negativos.

```
Las horas extras NO compensan las horas faltantes.
"""

minutos_negativos = abs(min(minutos, 0))

return minutos_negativos * VALOR_MINUTO
```

# ============================================================

# DETECCIÓN DE SEMANA

# ============================================================

def obtener_semana(fecha):
"""
Determina la semana dentro del mes.

```
Semana 1: días 1 al 6
Semana 2: días 7 al 13
Semana 3: días 14 al 20
Semana 4: días 21 al 27
Semana 5: días 28 al final del mes
"""

if pd.isna(fecha):
    return None

dia = fecha.day

if dia <= 6:
    return 1
elif dia <= 13:
    return 2
elif dia <= 20:
    return 3
elif dia <= 27:
    return 4
else:
    return 5
```

# ============================================================

# PROCESAMIENTO DEL EXCEL

# ============================================================

def leer_excel(archivo):
"""
Lee el Excel real de Geovictoria.

```
Busca automáticamente:
- Nombre
- Apellidos
- Grupo
- Fecha
- Diferencia del día
"""

try:

    # El Excel de Geovictoria tiene encabezados en la fila 3
    df = pd.read_excel(
        archivo,
        header=2
    )

    df = df.dropna(how="all")

    # ----------------------------------------
    # Buscar columnas
    # ----------------------------------------

    col_nombre = encontrar_columna(
        df,
        ["Nombre"]
    )

    col_apellidos = encontrar_columna(
        df,
        ["Apellidos", "Apellido"]
    )

    col_grupo = encontrar_columna(
        df,
        ["Grupo"]
    )

    col_fecha = encontrar_columna(
        df,
        ["Fecha"]
    )

    col_diferencia = encontrar_columna(
        df,
        [
            "Diferencia del día",
            "Diferencia del dia"
        ]
    )

    # ----------------------------------------
    # Verificar columnas
    # ----------------------------------------

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
        faltantes.append("Diferencia del día")

    if faltantes:

        st.error(
            "No se encontraron estas columnas: "
            + ", ".join(faltantes)
        )

        st.write(
            "Columnas encontradas en el archivo:"
        )

        st.write(list(df.columns))

        return pd.DataFrame()

    # ----------------------------------------
    # Crear tabla limpia
    # ----------------------------------------

    datos = pd.DataFrame()

    datos["Nombre"] = (
        df[col_nombre]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    datos["Apellidos"] = (
        df[col_apellidos]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    datos["Ejecutivo"] = (
        datos["Nombre"]
        + " "
        + datos["Apellidos"]
    ).str.strip()

    datos["Tienda"] = (
        df[col_grupo]
        .fillna("")
        .astype(str)
        .str.strip()
    )

    datos["Fecha"] = pd.to_datetime(
        df[col_fecha],
        errors="coerce",
        dayfirst=True
    )

    datos["Diferencia Minutos"] = (
        df[col_diferencia]
        .apply(valor_a_minutos)
    )

    # ----------------------------------------
    # Eliminar filas que no corresponden
    # ----------------------------------------

    datos = datos[
        (datos["Ejecutivo"] != "")
        &
        (datos["Tienda"] != "")
        &
        (datos["Fecha"].notna())
    ].copy()

    # ----------------------------------------
    # Semana automática
    # ----------------------------------------

    datos["Semana"] = datos["Fecha"].apply(
        obtener_semana
    )

    datos["Semana"] = datos["Semana"].astype("Int64")

    return datos

except Exception as error:

    st.error(
        f"No fue posible leer el archivo: {error}"
    )

    return pd.DataFrame()
```

# ============================================================

# TÍTULO

# ============================================================

st.title("📊 Control de Diferencias Geovictoria")

st.write(
"Carga tu archivo mensual de Geovictoria y el sistema "
"se encargará de separar automáticamente las horas "
"faltantes y las horas extras."
)

st.info(
f"💰 Sueldo base: ${SUELDO_BASE:,.0f} | "
f"⏱️ Jornada: {JORNADA_SEMANAL} horas semanales | "
f"💵 Valor hora estimado: ${VALOR_HORA:,.2f} | "
f"💵 Valor minuto: ${VALOR_MINUTO:,.2f}"
)

# ============================================================

# CARGA DEL ARCHIVO

# ============================================================

st.subheader("📁 Cargar archivo de Geovictoria")

archivo = st.file_uploader(
"Selecciona el Excel mensual actualizado",
type=["xlsx", "xls"]
)

if archivo is not None:

```
datos = leer_excel(archivo)

if not datos.empty:

    # ====================================================
    # RESUMEN GENERAL DE DATOS
    # ====================================================

    # Horas faltantes
    datos["Minutos Faltantes"] = datos[
        "Diferencia Minutos"
    ].apply(
        lambda x: abs(min(x, 0))
    )

    # Horas extras
    datos["Minutos Extras"] = datos[
        "Diferencia Minutos"
    ].apply(
        lambda x: max(x, 0)
    )

    # ====================================================
    # FILTROS
    # ====================================================

    st.subheader("🔎 Filtros")

    col_filtro1, col_filtro2 = st.columns(2)

    with col_filtro1:

        tiendas = sorted(
            datos["Tienda"]
            .dropna()
            .unique()
            .tolist()
        )

        tienda_seleccionada = st.multiselect(
            "Seleccionar tienda",
            tiendas
        )

    with col_filtro2:

        semanas_disponibles = sorted(
            datos["Semana"]
            .dropna()
            .astype(int)
            .unique()
            .tolist()
        )

        semana_seleccionada = st.multiselect(
            "Seleccionar semana",
            semanas_disponibles
        )

    datos_filtrados = datos.copy()

    if tienda_seleccionada:

        datos_filtrados = datos_filtrados[
            datos_filtrados["Tienda"].isin(
                tienda_seleccionada
            )
        ]

    if semana_seleccionada:

        datos_filtrados = datos_filtrados[
            datos_filtrados["Semana"].isin(
                semana_seleccionada
            )
        ]

    # ====================================================
    # RESUMEN
    # ====================================================

    st.subheader("📌 Resumen general")

    resumen_ejecutivos = (
        datos_filtrados
        .groupby(
            ["Ejecutivo", "Tienda"],
            as_index=False
        )
        .agg(
            Faltantes=(
                "Minutos Faltantes",
                "sum"
            ),
            Extras=(
                "Minutos Extras",
                "sum"
            )
        )
    )

    total_ejecutivos = len(
        resumen_ejecutivos
    )

    total_faltantes = resumen_ejecutivos[
        "Faltantes"
    ].sum()

    total_extras = resumen_ejecutivos[
        "Extras"
    ].sum()

    total_descuento = (
        total_faltantes
        * VALOR_MINUTO
    )

    casos_criticos = len(
        resumen_ejecutivos[
            resumen_ejecutivos["Faltantes"] >= 60
        ]
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "👥 Ejecutivos",
        total_ejecutivos
    )

    c2.metric(
        "🔴 Horas faltantes",
        minutos_positivos_a_texto(
            total_faltantes
        )
    )

    c3.metric(
        "🟢 Horas extras",
        minutos_positivos_a_texto(
            total_extras
        )
    )

    c4.metric(
        "💰 Descuento aprox.",
        f"${total_descuento:,.0f}"
    )

    st.write(
        f"⚠️ **Casos con 1 hora o más de tiempo faltante:** "
        f"{casos_criticos}"
    )

    # ====================================================
    # TABLA DE FALTANTES Y DESCUENTOS
    # ====================================================

    st.subheader(
        "🔴 Horas faltantes y descuentos aproximados"
    )

    acumulado = (
        datos_filtrados
        .groupby(
            ["Ejecutivo", "Tienda", "Semana"],
            as_index=False
        )
        .agg(
            Faltantes=(
                "Minutos Faltantes",
                "sum"
            )
        )
    )

    tabla_faltantes = (
        acumulado
        .pivot_table(
            index=["Ejecutivo", "Tienda"],
            columns="Semana",
            values="Faltantes",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    # Crear semanas 1 a 5
    for semana in range(1, 6):

        if semana not in tabla_faltantes.columns:

            tabla_faltantes[semana] = 0

    tabla_faltantes = tabla_faltantes[
        [
            "Ejecutivo",
            "Tienda",
            1,
            2,
            3,
            4,
            5
        ]
    ]

    tabla_faltantes = tabla_faltantes.rename(
        columns={
            1: "Semana 1",
            2: "Semana 2",
            3: "Semana 3",
            4: "Semana 4",
            5: "Semana 5"
        }
    )

    columnas_semana = [
        "Semana 1",
        "Semana 2",
        "Semana 3",
        "Semana 4",
        "Semana 5"
    ]

    tabla_faltantes["Total Negativo Minutos"] = (
        tabla_faltantes[columnas_semana]
        .sum(axis=1)
    )

    tabla_faltantes["Descuento Estimado"] = (
        tabla_faltantes[
            "Total Negativo Minutos"
        ]
        * VALOR_MINUTO
    )

    tabla_faltantes = tabla_faltantes[
        tabla_faltantes[
            "Total Negativo Minutos"
        ] > 0
    ].copy()

    tabla_faltantes = tabla_faltantes.sort_values(
        "Total Negativo Minutos",
        ascending=False
    )

    tabla_faltantes["Total Negativo"] = (
        tabla_faltantes[
            "Total Negativo Minutos"
        ]
        .apply(
            lambda x: "-"
            + minutos_positivos_a_texto(x)
        )
    )

    # Formatear semanas
    for columna in columnas_semana:

        tabla_faltantes[columna] = (
            tabla_faltantes[columna]
            .apply(
                lambda x:
                "-"
                + minutos_positivos_a_texto(x)
                if x > 0
                else "0 h 00 min"
            )
        )

    tabla_faltantes["Descuento Estimado"] = (
        tabla_faltantes[
            "Descuento Estimado"
        ]
        .apply(
            lambda x: f"${x:,.0f}"
        )
    )

    tabla_mostrar = tabla_faltantes[
        [
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
    ]

    st.dataframe(
        tabla_mostrar,
        use_container_width=True,
        hide_index=True
    )

    # ====================================================
    # HORAS EXTRAS
    # ====================================================

    st.subheader(
        "🟢 Horas extras registradas"
    )

    extras = (
        datos_filtrados[
            datos_filtrados["Minutos Extras"] > 0
        ]
        .groupby(
            ["Ejecutivo", "Tienda"],
            as_index=False
        )
        .agg(
            Horas_Extras=(
                "Minutos Extras",
                "sum"
            )
        )
    )

    if extras.empty:

        st.info(
            "No se registran horas extras en "
            "el período seleccionado."
        )

    else:

        extras["Horas Extras"] = (
            extras["Horas_Extras"]
            .apply(
                minutos_positivos_a_texto
            )
        )

        extras = extras[
            [
                "Ejecutivo",
                "Tienda",
                "Horas Extras"
            ]
        ]

        st.dataframe(
            extras,
            use_container_width=True,
            hide_index=True
        )

    # ====================================================
    # DETALLE DIARIO
    # ====================================================

    with st.expander(
        "📅 Ver detalle diario"
    ):

        detalle = datos_filtrados.copy()

        detalle["Semana"] = (
            detalle["Semana"]
            .apply(
                lambda x:
                f"Semana {int(x)}"
                if pd.notna(x)
                else ""
            )
        )

        detalle["Diferencia del día"] = (
            detalle["Diferencia Minutos"]
            .apply(
                lambda x:
                minutos_a_texto(x)
            )
        )

        detalle["Horas faltantes"] = (
            detalle["Minutos Faltantes"]
            .apply(
                minutos_positivos_a_texto
            )
        )

        detalle["Horas extras"] = (
            detalle["Minutos Extras"]
            .apply(
                minutos_positivos_a_texto
            )
        )

        detalle["Descuento aprox."] = (
            detalle["Minutos Faltantes"]
            .apply(
                lambda x:
                f"${x * VALOR_MINUTO:,.0f}"
            )
        )

        detalle = detalle[
            [
                "Fecha",
                "Semana",
                "Ejecutivo",
                "Tienda",
                "Diferencia del día",
                "Horas faltantes",
                "Horas extras",
                "Descuento aprox."
            ]
        ]

        detalle = detalle.sort_values(
            [
                "Ejecutivo",
                "Fecha"
            ]
        )

        st.dataframe(
            detalle,
            use_container_width=True,
            hide_index=True
        )

    # ====================================================
    # DESCARGAR EXCEL
    # ====================================================

    st.subheader(
        "📥 Descargar resultados"
    )

    archivo_salida = BytesIO()

    with pd.ExcelWriter(
        archivo_salida,
        engine="openpyxl"
    ) as writer:

        tabla_mostrar.to_excel(
            writer,
            index=False,
            sheet_name="Faltantes y descuentos"
        )

        extras.to_excel(
            writer,
            index=False,
            sheet_name="Horas extras"
        )

        detalle.to_excel(
            writer,
            index=False,
            sheet_name="Detalle diario"
        )

    st.download_button(
        label="📥 Descargar Excel de resultados",
        data=archivo_salida.getvalue(),
        file_name=(
            "Control_Geovictoria_"
            "Resultados.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )

else:

    st.warning(
        "No se encontraron datos válidos "
        "en el archivo."
    )
```

else:

```
st.write(
    "👆 Sube tu archivo mensual de Geovictoria "
    "para comenzar."
)
```
