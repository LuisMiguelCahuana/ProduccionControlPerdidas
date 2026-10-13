
import io
from datetime import datetime

import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

st.set_page_config(
    page_title="Procesador de Producción y Costos",
    page_icon="📊",
    layout="wide"
)

# Google Sheets de destino
SPREADSHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"

# Pestaña donde se guardarán los datos finales
HOJA_DESTINO = "Produccion_Costos"

# Hoja del archivo de producción
HOJA_PRODUCCION = "11-12-13-14 Actividad-CP"

# Zonas válidas
ZONAS_VALIDAS = ["UB", "UR", "R"]


# ============================================================
# ENCABEZADO
# ============================================================

st.title("📊 PROCESADOR DE PRODUCCIÓN Y COSTOS")

st.markdown(
    """
    Aplicación para procesar los archivos Excel de producción y costos,
    calcular los costos totales y exportar el resultado a Google Sheets.
    """
)

st.divider()


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def limpiar_texto(valor):
    """Convierte valores a texto sin espacios innecesarios."""
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    # Evitar códigos como 123.0 cuando Excel los lee como número
    if texto.endswith(".0") and texto[:-2].isdigit():
        texto = texto[:-2]

    return texto


def normalizar_codigos(serie):
    """Normaliza los códigos para facilitar el cruce."""
    return serie.map(limpiar_texto)


def convertir_a_numero(serie):
    """Convierte una columna a valores numéricos."""
    return pd.to_numeric(serie, errors="coerce")


# ============================================================
# PROCESAR ARCHIVOS EXCEL
# ============================================================

def procesar_archivos(
    archivo_produccion,
    archivo_costos,
    hoja_costos
):
    # --------------------------------------------------------
    # 1. LEER ARCHIVO DE PRODUCCIÓN
    # --------------------------------------------------------

    try:
        archivo_produccion.seek(0)

        df = pd.read_excel(
            archivo_produccion,
            sheet_name=HOJA_PRODUCCION,
            header=None
        )

    except ValueError as error:
        raise ValueError(
            f"No se encontró la hoja '{HOJA_PRODUCCION}' "
            "en el archivo de producción."
        ) from error

    except Exception as error:
        raise ValueError(
            f"No se pudo leer el archivo de producción: {error}"
        ) from error

    # --------------------------------------------------------
    # 2. ESTRUCTURA DEL ARCHIVO ORIGINAL
    # --------------------------------------------------------

    fila_zona = 2
    fila_codigo = 3
    fila_datos = 5

    col_fecha = 1
    col_suministro = 3
    col_cuadrilla = 6

    if df.shape[0] <= fila_datos:
        raise ValueError(
            "El archivo de producción no contiene suficientes filas."
        )

    if df.shape[1] <= col_cuadrilla:
        raise ValueError(
            "El archivo de producción no contiene las columnas esperadas."
        )

    # Última fila válida según la columna de fecha
    ultima_fila = df[col_fecha].last_valid_index()

    if ultima_fila is None or ultima_fila < fila_datos:
        raise ValueError(
            "No se encontraron fechas en la columna B del archivo."
        )

    # --------------------------------------------------------
    # 3. DETECTAR COLUMNAS DE PRODUCCIÓN
    # --------------------------------------------------------

    columnas_validas = []
    nombres_columnas = []

    for col in range(df.shape[1]):

        zona = limpiar_texto(
            df.iloc[fila_zona, col]
        ).upper()

        codigo = limpiar_texto(
            df.iloc[fila_codigo, col]
        )

        if zona in ZONAS_VALIDAS and codigo != "":
            columnas_validas.append(col)
            nombres_columnas.append(
                f"{zona}_{codigo}"
            )

    if not columnas_validas:
        raise ValueError(
            "No se encontraron columnas de producción "
            "con zonas UB, UR o R y código válido."
        )

    # --------------------------------------------------------
    # 4. EXTRAER DATOS DE PRODUCCIÓN
    # --------------------------------------------------------

    rango = slice(
        fila_datos,
        ultima_fila + 1
    )

    df_prod = df.iloc[
        rango,
        columnas_validas
    ].copy()

    df_prod.columns = nombres_columnas

    # --------------------------------------------------------
    # 5. CREAR COLUMNAS BASE
    # --------------------------------------------------------

    df_base = pd.DataFrame(
        {
            "fecha": pd.to_datetime(
                df.iloc[rango, col_fecha],
                errors="coerce"
            ),

            "suministro": df.iloc[
                rango, col_suministro
            ].map(limpiar_texto),

            "cuadrilla": df.iloc[
                rango, col_cuadrilla
            ].map(limpiar_texto),
        },
        index=df_prod.index
    )

    df_prod = pd.concat(
        [df_base, df_prod],
        axis=1
    )

    # --------------------------------------------------------
    # 6. CONVERTIR COLUMNAS A FILAS
    # --------------------------------------------------------

    df_long = (
        df_prod
        .set_index(
            ["fecha", "suministro", "cuadrilla"]
        )
        .stack()
        .reset_index()
    )

    df_long.columns = [
        "fecha",
        "suministro",
        "cuadrilla",
        "zona_codigo",
        "cantidad"
    ]

    # Convertir cantidad a número
    df_long["cantidad"] = convertir_a_numero(
        df_long["cantidad"]
    )

    # Eliminar cantidades nulas y ceros
    df_long = df_long[
        df_long["cantidad"].notna()
        & (df_long["cantidad"] != 0)
    ].copy()

    # --------------------------------------------------------
    # 7. SEPARAR ZONA Y CÓDIGO
    # --------------------------------------------------------

    partes = (
        df_long["zona_codigo"]
        .astype(str)
        .str.split("_", n=1, expand=True)
    )

    df_long["zona"] = partes[0].str.strip().str.upper()
    df_long["codigo"] = partes[1].map(limpiar_texto)

    df_final = df_long.drop(
        columns=["zona_codigo"]
    )

    # --------------------------------------------------------
    # 8. LEER ARCHIVO DE COSTOS
    # --------------------------------------------------------

    try:
        archivo_costos.seek(0)

        df_costos = pd.read_excel(
            archivo_costos,
            sheet_name=hoja_costos
        )

    except Exception as error:
        raise ValueError(
            f"No se pudo leer la hoja de costos: {error}"
        ) from error

    # Limpiar nombres de columnas
    df_costos.columns = [
        str(col).strip()
        for col in df_costos.columns
    ]

    # Normalizar nombres esperados
    renombres = {}

    for columna in df_costos.columns:

        nombre = str(columna).strip().upper()

        if nombre == "ITEM":
            renombres[columna] = "codigo"

        elif nombre == "ACTIVIDAD/SUBACTIVIDAD":
            renombres[columna] = "actividad"

    df_costos = df_costos.rename(
        columns=renombres
    )

    # Validar columnas obligatorias
    columnas_requeridas = [
        "codigo",
        "actividad",
        "UB",
        "UR",
        "R"
    ]

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_costos.columns
    ]

    if faltantes:
        raise ValueError(
            "El archivo de costos no contiene estas columnas: "
            + ", ".join(faltantes)
            + ". Verifica que tenga ITEM, "
              "ACTIVIDAD/SUBACTIVIDAD, UB, UR y R."
        )

    # --------------------------------------------------------
    # 9. CONVERTIR COSTOS UB, UR Y R A FILAS
    # --------------------------------------------------------

    df_costos = df_costos.melt(
        id_vars=["codigo", "actividad"],
        value_vars=ZONAS_VALIDAS,
        var_name="zona",
        value_name="costo_unitario"
    )

    df_costos["codigo"] = normalizar_codigos(
        df_costos["codigo"]
    )

    df_costos["zona"] = (
        df_costos["zona"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df_costos["actividad"] = (
        df_costos["actividad"]
        .map(limpiar_texto)
    )

    df_costos["costo_unitario"] = convertir_a_numero(
        df_costos["costo_unitario"]
    )

    df_costos = df_costos.dropna(
        subset=["costo_unitario"]
    ).copy()

    df_costos = df_costos[
        df_costos["codigo"] != ""
    ].copy()

    # --------------------------------------------------------
    # 10. NORMALIZAR CLAVES PARA EL CRUCE
    # --------------------------------------------------------

    df_final["codigo"] = normalizar_codigos(
        df_final["codigo"]
    )

    df_final["zona"] = (
        df_final["zona"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # Comprobar claves duplicadas en la tabla de costos
    duplicados = df_costos.duplicated(
        subset=["codigo", "zona"],
        keep=False
    )

    cantidad_duplicados = int(duplicados.sum())

    # Conservar una sola fila por código y zona.
    # Esto evita multiplicar registros durante el cruce.
    df_costos = df_costos.drop_duplicates(
        subset=["codigo", "zona"],
        keep="first"
    )

    # --------------------------------------------------------
    # 11. CRUZAR PRODUCCIÓN CON COSTOS
    # --------------------------------------------------------

    df_final = df_final.merge(
        df_costos[
            [
                "codigo",
                "zona",
                "actividad",
                "costo_unitario"
            ]
        ],
        on=["codigo", "zona"],
        how="left",
        validate="many_to_one"
    )

    # --------------------------------------------------------
    # 12. CALCULAR COSTO TOTAL
    # --------------------------------------------------------

    df_final["costo_total"] = (
        df_final["cantidad"]
        * df_final["costo_unitario"]
    )

    # --------------------------------------------------------
    # 13. ORDENAR COLUMNAS FINALES
    # --------------------------------------------------------

    columnas_finales = [
        "fecha",
        "suministro",
        "cuadrilla",
        "zona",
        "codigo",
        "cantidad",
        "actividad",
        "costo_unitario",
        "costo_total"
    ]

    df_final = df_final[
        columnas_finales
    ].copy()

    return df_final, cantidad_duplicados


# ============================================================
# CONECTAR CON GOOGLE SHEETS
# ============================================================

def conectar_google_sheets():

    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "No se encontraron las credenciales de Google. "
            "Configura gcp_service_account en "
            "Streamlit Cloud > Settings > Secrets."
        )

    credenciales_dict = dict(
        st.secrets["gcp_service_account"]
    )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    credenciales = Credentials.from_service_account_info(
        credenciales_dict,
        scopes=scopes
    )

    cliente = gspread.authorize(
        credenciales
    )

    return cliente


# ============================================================
# EXPORTAR RESULTADO A GOOGLE SHEETS
# ============================================================

def exportar_google_sheets(df):

    cliente = conectar_google_sheets()

    libro = cliente.open_by_key(
        SPREADSHEET_ID
    )

    # Buscar la pestaña de destino.
    # Si no existe, crearla.
    try:
        hoja = libro.worksheet(
            HOJA_DESTINO
        )

    except gspread.WorksheetNotFound:
        hoja = libro.add_worksheet(
            title=HOJA_DESTINO,
            rows=max(len(df) + 1, 100),
            cols=max(len(df.columns), 10)
        )

    # Preparar datos para exportación
    df_exportar = df.copy()

    if "fecha" in df_exportar.columns:
        df_exportar["fecha"] = (
            pd.to_datetime(
                df_exportar["fecha"],
                errors="coerce"
            )
            .dt.strftime("%Y-%m-%d")
            .fillna("")
        )

    # Reemplazar valores vacíos por cadenas vacías
    df_exportar = df_exportar.astype(object)
    df_exportar = df_exportar.where(
        pd.notna(df_exportar),
        ""
    )

    encabezados = df_exportar.columns.tolist()

    filas = df_exportar.values.tolist()

    # El documento recibe encabezados + registros
    valores = [encabezados] + filas

    # Limpiar únicamente la pestaña de destino
    hoja.clear()

    # Ajustar tamaño de la hoja
    hoja.resize(
        rows=max(len(valores), 1),
        cols=max(len(encabezados), 1)
    )

    # Enviar los datos
    hoja.update(
        range_name="A1",
        values=valores,
        value_input_option="USER_ENTERED"
    )

    return libro.url


# ============================================================
# INTERFAZ STREAMLIT
# ============================================================

st.subheader("1. Cargar archivos Excel")

col1, col2 = st.columns(2)

with col1:

    archivo_produccion = st.file_uploader(
        "📁 Archivo Excel de producción",
        type=["xlsx", "xls", "xlsm"],
        key="produccion"
    )

    st.caption(
        "Debe contener la hoja: "
        "11-12-13-14 Actividad-CP"
    )

with col2:

    archivo_costos = st.file_uploader(
        "📁 Archivo Excel de costos",
        type=["xlsx", "xls", "xlsm"],
        key="costos"
    )

    st.caption(
        "Debe contener las columnas ITEM, "
        "ACTIVIDAD/SUBACTIVIDAD, UB, UR y R."
    )


# ============================================================
# SELECCIONAR HOJA DEL ARCHIVO DE COSTOS
# ============================================================

hoja_costos = None

if archivo_costos is not None:

    try:
        archivo_costos.seek(0)

        libro_costos = pd.ExcelFile(
            archivo_costos
        )

        hojas_costos = libro_costos.sheet_names

        hoja_costos = st.selectbox(
            "Selecciona la hoja del archivo de costos",
            options=hojas_costos
        )

    except Exception as error:
        st.error(
            f"No se pudo identificar la hoja de costos: {error}"
        )


# ============================================================
# PROCESAR ARCHIVOS
# ============================================================

if st.button(
    "⚙️ PROCESAR PRODUCCIÓN Y COSTOS",
    type="primary",
    use_container_width=True
):

    if archivo_produccion is None:
        st.error(
            "Primero debes cargar el archivo de producción."
        )

    elif archivo_costos is None:
        st.error(
            "Primero debes cargar el archivo de costos."
        )

    elif hoja_costos is None:
        st.error(
            "Selecciona una hoja del archivo de costos."
        )

    else:

        try:

            with st.spinner(
                "Procesando archivos y calculando costos..."
            ):

                archivo_produccion.seek(0)
                archivo_costos.seek(0)

                df_resultado, duplicados = procesar_archivos(
                    archivo_produccion,
                    archivo_costos,
                    hoja_costos
                )

                st.session_state["resultado"] = df_resultado
                st.session_state["duplicados"] = duplicados

            st.success(
                "Los archivos se procesaron correctamente."
            )

        except Exception as error:

            st.error(
                f"Error al procesar los archivos: {error}"
            )


# ============================================================
# MOSTRAR RESULTADOS
# ============================================================

if "resultado" in st.session_state:

    df_resultado = st.session_state["resultado"]

    cantidad_duplicados = st.session_state.get(
        "duplicados",
        0
    )

    if cantidad_duplicados > 0:

        st.warning(
            f"Se encontraron {cantidad_duplicados} filas "
            "involucradas en claves código/zona duplicadas "
            "en el archivo de costos. Se conservó la primera "
            "coincidencia de cada clave para evitar duplicar "
            "registros de producción."
        )

    st.divider()

    st.subheader("2. Resumen del procesamiento")

    registros_sin_costo = int(
        df_resultado["costo_unitario"].isna().sum()
    )

    costo_total = df_resultado[
        "costo_total"
    ].sum()

    cantidad_total = df_resultado[
        "cantidad"
    ].sum()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Registros procesados",
        f"{len(df_resultado):,}"
    )

    c2.metric(
        "Cantidad total",
        f"{cantidad_total:,.2f}"
    )

    c3.metric(
        "Costo total",
        f"{costo_total:,.2f}"
    )

    c4.metric(
        "Registros sin costo",
        f"{registros_sin_costo:,}"
    )

    st.subheader("3. Vista previa")

    st.dataframe(
        df_resultado.head(500),
        use_container_width=True,
        hide_index=True
    )

    if len(df_resultado) > 500:
        st.caption(
            "Se muestran los primeros 500 registros. "
            "La exportación incluye todos los registros."
        )

    # ========================================================
    # DESCARGAR EXCEL FINAL
    # ========================================================

    salida_excel = io.BytesIO()

    with pd.ExcelWriter(
        salida_excel,
        engine="openpyxl"
    ) as writer:

        df_resultado.to_excel(
            writer,
            index=False,
            sheet_name="Resultado"
        )

    nombre_archivo = (
        "produccion_costos_"
        + datetime.now().strftime("%Y%m%d_%H%M%S")
        + ".xlsx"
    )

    st.download_button(
        label="⬇️ DESCARGAR RESULTADO EN EXCEL",
        data=salida_excel.getvalue(),
        file_name=nombre_archivo,
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True
    )

    # ========================================================
    # EXPORTAR A GOOGLE SHEETS
    # ========================================================

    st.divider()

    st.subheader("4. Exportar a Google Sheets")

    st.warning(
        f"Se actualizará únicamente la pestaña "
        f"'{HOJA_DESTINO}'. Su contenido anterior será "
        "reemplazado por los resultados actuales."
    )

    if st.button(
        "📤 EXPORTAR A GOOGLE SHEETS",
        type="primary",
        use_container_width=True
    ):

        try:

            with st.spinner(
                "Exportando registros a Google Sheets..."
            ):

                url_google = exportar_google_sheets(
                    df_resultado
                )

            st.success(
                f"Se exportaron {len(df_resultado):,} registros "
                f"a la pestaña '{HOJA_DESTINO}'."
            )

            st.markdown(
                f"[📄 ABRIR GOOGLE SHEETS]({url_google})"
            )

        except Exception as error:

            st.error(
                "No se pudo exportar a Google Sheets. "
                "Verifica que la cuenta de servicio tenga "
                "permiso de Editor y que los Secrets estén "
                f"configurados correctamente. Detalle: {error}"
            )


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.divider()

st.caption(
    "Procesador de Producción y Costos | "
    "La exportación solo modifica la pestaña "
    "'Produccion_Costos' del documento de destino."
)
