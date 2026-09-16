import streamlit as st
import pandas as pd
import re
from io import BytesIO

st.set_page_config(
page_title="Control Geovictoria",
page_icon="📊",
layout="wide"
)

SUELDO_BASE = 539000
JORNADA_SEMANAL = 40

VALOR_HORA = SUELDO_BASE / 30 * 7 / JORNADA_SEMANAL
VALOR_MINUTO = VALOR_HORA / 60

def limpiar_texto(valor):
if pd.isna(valor):
return ""
return str(valor).strip()

def normalizar_columna(nombre):
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
for columna in df.columns:
columna_limpia = normalizar_columna(columna)

```
    for nombre in nombres:
        nombre_limpio = normalizar_columna(nombre)

        if columna_limpia == nombre_limpio:
            return columna

for columna in df.columns:
    columna_limpia = normalizar_columna(columna)

    for nombre in nombres:
        nombre_limpio = normalizar_columna(nombre)

        if nombre_limpio in columna_limpia:
            return columna

return None
```

def convertir_minutos(valor):
if pd.isna(valor):
return 0

```
if isinstance(valor, pd.Timedelta):
    return valor.total_seconds() / 60

texto = str(valor).strip().lower()

if texto in ["", "nan", "none", "nat", "-", "sin marca"]:
    return 0

patron = re.match(
    r"^(-)?(\d+):(\d+):(\d+)$",
    texto
)

if patron:
    signo = -1 if patron.group(1) else 1
    horas = int(patron.group(2))
    minutos = int(patron.group(3))
    segundos = int(patron.group(4))

    return signo * (
        horas * 60
        + minutos
        + segundos / 60
    )

patron = re.match(
    r"^(-)?(\d+):(\d+)$",
    texto
)

if patron:
    signo = -1 if patron.group(1) else 1
    horas = int(patron.group(2))
    minutos = int(patron.group(3))

    return signo * (
        horas * 60 + minutos
    )

patron = re.search(
    r"(-)?\s*(\d+)\s*h\s*(\d+)?\s*min",
    texto
)

if patron:
    signo = -1 if patron.group(1) else 1
    horas = int(patron.group(2))
    minutos = int(patron.group(3) or 0)

    return signo * (
        horas * 60 + minutos
    )

try:
    return float(valor)
except:
    return 0
```

def minutos_a_texto(minutos):
minutos = int(round(minutos))

```
signo = "-" if minutos < 0 else ""

minutos = abs(minutos)

horas = minutos // 60
minutos_restantes = minutos % 60

return (
    f"{signo}{horas} h "
    f"{minutos_restantes:02d} min"
)
```

def horas_positivas_a_texto(minutos):
minutos = int(round(abs(minutos)))

```
horas = minutos // 60
minutos_restantes = minutos % 60

return (
    f"{horas} h "
    f"{minutos_restantes:02d} min"
)
```

def obtener_semana(fecha):
if pd.isna(fecha):
return None

```
dia = fecha.day

if dia <= 6:
    return 1

if dia <= 13:
    return 2

if dia <= 20:
    return 3

if dia <= 27:
    return 4

return 5
```

def leer_excel(archivo):
try:
df = pd.read_excel(
archivo,
header=2
)

```
    df = df.dropna(how="all")

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
            "No se encontraron: "
            + ", ".join(faltantes)
        )

        st.write(
            "Columnas encontradas:"
        )

        st.write(list(df.columns))

        return pd.DataFrame()

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
        .apply(convertir_minutos)
    )

    datos = datos[
        (datos["Ejecutivo"] != "")
        & (datos["Tienda"] != "")
        & (datos["Fecha"].notna())
    ].copy()

    datos["Semana"] = datos["Fecha"].apply(
        obtener_semana
    )

    return datos

except Exception as error:
    st.error(
        f"Error al leer el Excel: {error}"
    )

    return pd.DataFrame()
```

st.title("📊 Control de Diferencias Geovictoria")

st.write(
"Carga tu Excel mensual de Geovictoria "
"para controlar horas faltantes, horas "
"extras y descuentos aproximados."
)

st.info(
f"Sueldo base: ${SUELDO_BASE:,.0f} | "
f"Jornada: {JORNADA_SEMANAL} horas | "
f"Valor hora: ${VALOR_HORA:,.2f} | "
f"Valor minuto: ${VALOR_MINUTO:,.2f}"
)

st.subheader("📁 Cargar Excel")

archivo = st.file_uploader(
"Selecciona el archivo mensual de Geovictoria",
type=["xlsx", "xls"]
)

if archivo is not None:

```
datos = leer_excel(archivo)

if not datos.empty:

    datos["Faltantes"] = (
        datos["Diferencia Minutos"]
        .apply(
            lambda x: abs(min(x, 0))
        )
    )

    datos["Extras"] = (
        datos["Diferencia Minutos"]
        .apply(
            lambda x: max(x, 0)
        )
    )

    st.success(
        f"Archivo procesado correctamente. "
        f"Se encontraron {len(datos):,} registros."
    )

    st.subheader("🔎 Filtros")

    col1, col2 = st.columns(2)

    with col1:

        tiendas = sorted(
            datos["Tienda"]
            .unique()
            .tolist()
        )

        filtro_tienda = st.multiselect(
            "Tienda",
            tiendas
        )

    with col2:

        semanas = sorted(
            datos["Semana"]
            .dropna()
            .astype(int)
            .unique()
            .tolist()
        )

        filtro_semana = st.multiselect(
            "Semana",
            semanas
        )

    datos_filtrados = datos.copy()

    if filtro_tienda:
        datos_filtrados = datos_filtrados[
            datos_filtrados["Tienda"].isin(
                filtro_tienda
            )
        ]

    if filtro_semana:
        datos_filtrados = datos_filtrados[
            datos_filtrados["Semana"].isin(
                filtro_semana
            )
        ]

    st.subheader("📌 Resumen")

    resumen = (
        datos_filtrados
        .groupby(
            ["Ejecutivo", "Tienda"],
            as_index=False
        )
        .agg(
            Faltantes=("Faltantes", "sum"),
            Extras=("Extras", "sum")
        )
    )

    total_ejecutivos = len(resumen)

    total_faltantes = resumen[
        "Faltantes"
    ].sum()

    total_extras = resumen[
        "Extras"
    ].sum()

    descuento_total = (
        total_faltantes
        * VALOR_MINUTO
    )

    casos_criticos = len(
        resumen[
            resumen["Faltantes"] >= 60
        ]
    )

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "👥 Ejecutivos",
        total_ejecutivos
    )

    c2.metric(
        "🔴 Horas faltantes",
        horas_positivas_a_texto(
            total_faltantes
        )
    )

    c3.metric(
        "🟢 Horas extras",
        horas_positivas_a_texto(
            total_extras
        )
    )

    c4.metric(
        "💰 Descuento aprox.",
        f"${descuento_total:,.0f}"
    )

    st.write(
        f"⚠️ Casos críticos "
        f"(1 hora o más): {casos_criticos}"
    )

    st.subheader(
        "🔴 Horas faltantes y descuento aproximado"
    )

    faltantes = (
        datos_filtrados[
            datos_filtrados["Faltantes"] > 0
        ]
        .groupby(
            ["Ejecutivo", "Tienda", "Semana"],
            as_index=False
        )["Faltantes"]
        .sum()
    )

    if faltantes.empty:

        st.info(
            "No existen horas faltantes "
            "en la selección."
        )

    else:

        tabla = (
            faltantes
            .pivot_table(
                index=[
                    "Ejecutivo",
                    "Tienda"
                ],
                columns="Semana",
                values="Faltantes",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        for semana in range(1, 6):

            if semana not in tabla.columns:
                tabla[semana] = 0

        tabla = tabla[
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

        tabla = tabla.rename(
            columns={
                1: "Semana 1",
                2: "Semana 2",
                3: "Semana 3",
                4: "Semana 4",
                5: "Semana 5"
            }
        )

        semanas_columnas = [
            "Semana 1",
            "Semana 2",
            "Semana 3",
            "Semana 4",
            "Semana 5"
        ]

        tabla["Total Minutos"] = (
            tabla[semanas_columnas]
            .sum(axis=1)
        )

        tabla["Descuento"] = (
            tabla["Total Minutos"]
            * VALOR_MINUTO
        )

        tabla = tabla.sort_values(
            "Total Minutos",
            ascending=False
        )

        tabla["Total Negativo"] = (
            tabla["Total Minutos"]
            .apply(
                lambda x:
                "-"
                + horas_positivas_a_texto(x)
            )
        )

        for columna in semanas_columnas:

            tabla[columna] = (
                tabla[columna]
                .apply(
                    lambda x:
                    "-"
                    + horas_positivas_a_texto(x)
                    if x > 0
                    else "0 h 00 min"
                )
            )

        tabla["Descuento aprox."] = (
            tabla["Descuento"]
            .apply(
                lambda x:
                f"${x:,.0f}"
            )
        )

        tabla_final = tabla[
            [
                "Ejecutivo",
                "Tienda",
                "Semana 1",
                "Semana 2",
                "Semana 3",
                "Semana 4",
                "Semana 5",
                "Total Negativo",
                "Descuento aprox."
            ]
        ]

        st.dataframe(
            tabla_final,
            use_container_width=True,
            hide_index=True
        )

    st.subheader(
        "🟢 Horas extras"
    )

    extras = (
        datos_filtrados[
            datos_filtrados["Extras"] > 0
        ]
        .groupby(
            ["Ejecutivo", "Tienda"],
            as_index=False
        )["Extras"]
        .sum()
    )

    if extras.empty:

        st.info(
            "No existen horas extras "
            "en la selección."
        )

    else:

        extras["Horas extras"] = (
            extras["Extras"]
            .apply(
                horas_positivas_a_texto
            )
        )

        extras_final = extras[
            [
                "Ejecutivo",
                "Tienda",
                "Horas extras"
            ]
        ]

        st.dataframe(
            extras_final,
            use_container_width=True,
            hide_index=True
        )

    st.subheader(
        "📅 Detalle diario"
    )

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
        .apply(minutos_a_texto)
    )

    detalle["Horas faltantes"] = (
        detalle["Faltantes"]
        .apply(horas_positivas_a_texto)
    )

    detalle["Horas extras"] = (
        detalle["Extras"]
        .apply(horas_positivas_a_texto)
    )

    detalle["Descuento aprox."] = (
        detalle["Faltantes"]
        .apply(
            lambda x:
            f"${x * VALOR_MINUTO:,.0f}"
        )
    )

    detalle_final = detalle[
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
    ].sort_values(
        ["Ejecutivo", "Fecha"]
    )

    with st.expander(
        "Ver detalle por día"
    ):

        st.dataframe(
            detalle_final,
            use_container_width=True,
            hide_index=True
        )

    st.subheader(
        "📥 Descargar resultados"
    )

    archivo_salida = BytesIO()

    with pd.ExcelWriter(
        archivo_salida,
        engine="openpyxl"
    ) as writer:

        tabla_final.to_excel(
            writer,
            index=False,
            sheet_name="Faltantes"
        )

        extras_final.to_excel(
            writer,
            index=False,
            sheet_name="Horas extras"
        )

        detalle_final.to_excel(
            writer,
            index=False,
            sheet_name="Detalle diario"
        )

    st.download_button(
        "📥 Descargar Excel",
        data=archivo_salida.getvalue(),
        file_name="Control_Geovictoria_Resultados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        )
    )
```

else:

```
st.info(
    "👆 Sube tu Excel mensual para comenzar."
)
```
