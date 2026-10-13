import io
import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ============================================================
# CONFIGURACIÓN
# ============================================================
SHEET_ID = "1gNq2nuNGfogcgcq7gsIqmQ5BZmMuWiSgAbDLpytJTIY"
HOJA_PRODUCCION = "11-12-13-14 Actividad-CP"

st.set_page_config(
    page_title="Procesador de Producción y Costos",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Procesador de Producción y Costos")
st.caption("Carga la producción, cruza los códigos con la tabla de costos y descarga el resultado.")

# ============================================================
# CONEXIÓN CON GOOGLE SHEETS
# ============================================================
@st.cache_resource
def conectar_google_sheets():
    if "gcp_service_account" not in st.secrets:
        raise RuntimeError(
            "No se encontró [gcp_service_account] en Secrets. "
            "Configura las credenciales en Streamlit Cloud."
        )

    cuenta = dict(st.secrets["gcp_service_account"])
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    credenciales = Credentials.from_service_account_info(cuenta, scopes=scopes)
    cliente = gspread.authorize(credenciales)
    return cliente

@st.cache_data(ttl=300)
def obtener_nombres_hojas(sheet_id):
    cliente = conectar_google_sheets()
    libro = cliente.open_by_key(sheet_id)
    return [hoja.title for hoja in libro.worksheets()]

@st.cache_data(ttl=300)
def cargar_costos(sheet_id, nombre_hoja):
    cliente = conectar_google_sheets()
    libro = cliente.open_by_key(sheet_id)
    hoja = libro.worksheet(nombre_hoja)
    registros = hoja.get_all_values()
    if not registros:
        raise ValueError("La hoja seleccionada está vacía.")

    encabezados = [str(x).strip() for x in registros[0]]
    df = pd.DataFrame(registros[1:], columns=encabezados)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def normalizar_texto(serie):
    return serie.astype("string").str.strip()

def convertir_numero(valor):
    """Convierte números con decimal punto o coma y separador de miles habitual."""
    if pd.isna(valor):
        return pd.NA
    if isinstance(valor, (int, float)):
        return valor
    texto = str(valor).strip().replace("S/", "").replace("$", "").replace(" ", "")
    if texto == "":
        return pd.NA

    # Si hay coma y punto, el último separador se considera decimal.
    if "," in texto and "." in texto:
        if texto.rfind(",") > texto.rfind("."):
            texto = texto.replace(".", "").replace(",", ".")
        else:
            texto = texto.replace(",", "")
    elif "," in texto:
        partes = texto.split(",")
        if len(partes) == 2 and len(partes[1]) in (1, 2):
            texto = texto.replace(",", ".")
        else:
            texto = texto.replace(",", "")
    try:
        return float(texto)
    except (TypeError, ValueError):
        return pd.NA

def procesar_produccion(archivo_excel, df_costos):
    try:
        df = pd.read_excel(archivo_excel, sheet_name=HOJA_PRODUCCION, header=None)
    except ValueError as exc:
        raise ValueError(
            f"No se encontró la hoja '{HOJA_PRODUCCION}' en el Excel de producción. "
            "Verifica el nombre de la hoja."
        ) from exc

    # Filas y columnas de la estructura original (índices desde cero).
    fila_zona = 2
    fila_codigo = 3
    fila_datos = 5
    col_fecha = 1
    col_suministro = 3
    col_cuadrilla = 6

    if df.shape[0] <= fila_datos or df.shape[1] <= max(col_fecha, col_suministro, col_cuadrilla):
        raise ValueError("El archivo de producción no tiene la estructura esperada.")

    ultima_fila = df[col_fecha].last_valid_index()
    if ultima_fila is None or ultima_fila < fila_datos:
        raise ValueError("No se encontraron fechas de producción en la columna esperada.")

    zonas_validas = ["UB", "UR", "R"]
    columnas_validas = []
    nombres_columnas = []

    for col in range(df.shape[1]):
        zona = str(df.iloc[fila_zona, col]).strip() if pd.notna(df.iloc[fila_zona, col]) else ""
        codigo = df.iloc[fila_codigo, col]
        if zona in zonas_validas and pd.notna(codigo) and str(codigo).strip():
            columnas_validas.append(col)
            nombres_columnas.append(f"{zona}_{str(codigo).strip()}")

    if not columnas_validas:
        raise ValueError(
            "No se detectaron columnas de producción válidas. "
            "Se esperan zonas UB, UR o R en la fila 3 y códigos en la fila 4."
        )

    df_prod = df.iloc[fila_datos:ultima_fila + 1, columnas_validas].copy()
    df_prod.columns = nombres_columnas

    df_base = pd.DataFrame({
        "fecha": pd.to_datetime(df.iloc[fila_datos:ultima_fila + 1, col_fecha], errors="coerce").values,
        "suministro": df.iloc[fila_datos:ultima_fila + 1, col_suministro].values,
        "cuadrilla": df.iloc[fila_datos:ultima_fila + 1, col_cuadrilla].values,
    })

    df_prod = pd.concat(
        [df_base.reset_index(drop=True), df_prod.reset_index(drop=True)],
        axis=1
    )

    df_long = df_prod.melt(
        id_vars=["fecha", "suministro", "cuadrilla"],
        var_name="zona_codigo",
        value_name="cantidad"
    )

    df_long["cantidad"] = df_long["cantidad"].map(convertir_numero)
    df_long = df_long[df_long["cantidad"].notna()]
    df_long = df_long[df_long["cantidad"] != 0].copy()

    partes = df_long["zona_codigo"].astype(str).str.split("_", n=1, expand=True)
    if partes.shape[1] != 2:
        raise ValueError("No se pudieron separar las zonas y códigos de producción.")

    df_long["zona"] = partes[0].str.strip()
    df_long["codigo"] = partes[1].str.strip()
    df_final = df_long.drop(columns=["zona_codigo"])

    # Validar encabezados de la tabla de costos.
    requeridas = ["ITEM", "ACTIVIDAD/SUBACTIVIDAD", "UB", "UR", "R"]
    faltantes = [c for c in requeridas if c not in df_costos.columns]
    if faltantes:
        raise ValueError(
            "La hoja de costos debe tener estas columnas: ITEM, "
            "ACTIVIDAD/SUBACTIVIDAD, UB, UR y R. "
            f"Faltan: {', '.join(faltantes)}. "
            "Si los encabezados están escritos de otra manera, corrígelos en Google Sheets."
        )

    costos = df_costos[requeridas].copy()
    costos = costos.rename(columns={
        "ITEM": "codigo",
        "ACTIVIDAD/SUBACTIVIDAD": "actividad"
    })

    costos["codigo"] = normalizar_texto(costos["codigo"])
    df_final["codigo"] = normalizar_texto(df_final["codigo"])
    costos = costos.melt(
        id_vars=["codigo", "actividad"],
        value_vars=["UB", "UR", "R"],
        var_name="zona",
        value_name="costo_unitario"
    )
    costos["costo_unitario"] = costos["costo_unitario"].map(convertir_numero)
    costos = costos.dropna(subset=["costo_unitario"])
    costos["zona"] = normalizar_texto(costos["zona"])

    df_final = df_final.merge(
        costos,
        on=["codigo", "zona"],
        how="left"
    )
    df_final["costo_total"] = (
        pd.to_numeric(df_final["cantidad"], errors="coerce")
        * pd.to_numeric(df_final["costo_unitario"], errors="coerce")
    )

    # Orden de columnas solicitado por el flujo original.
    orden = [
        "fecha", "suministro", "cuadrilla", "zona", "codigo",
        "cantidad", "actividad", "costo_unitario", "costo_total"
    ]
    df_final = df_final[[c for c in orden if c in df_final.columns]]
    return df_final

# ============================================================
# INTERFAZ
# ============================================================
with st.sidebar:
    st.header("⚙️ Configuración")
    st.text("Origen de costos: Google Sheets")
    st.code(SHEET_ID, language=None)
    st.caption("La cuenta de servicio debe tener permiso de lectura en este archivo.")

try:
    nombres_hojas = obtener_nombres_hojas(SHEET_ID)
    if not nombres_hojas:
        st.error("El Google Sheets no contiene hojas.")
        st.stop()

    hoja_costos = st.selectbox(
        "Selecciona la hoja que contiene la tabla de costos",
        options=nombres_hojas
    )

    archivo_produccion = st.file_uploader(
        "Sube el archivo Excel de producción (.xlsx o .xlsm)",
        type=["xlsx", "xlsm"],
        help=f"Debe contener la hoja '{HOJA_PRODUCCION}'."
    )

    with st.expander("Vista previa de la tabla de costos"):
        try:
            df_costos_preview = cargar_costos(SHEET_ID, hoja_costos)
            st.dataframe(df_costos_preview.head(20), use_container_width=True)
            st.caption(f"Registros cargados: {len(df_costos_preview):,}")
        except Exception as exc:
            st.error(f"No se pudo leer la hoja de costos: {exc}")

    if archivo_produccion is not None:
        if st.button("🚀 Procesar producción y calcular costos", type="primary", use_container_width=True):
            try:
                with st.spinner("Procesando archivo y cruzando costos..."):
                    archivo_bytes = io.BytesIO(archivo_produccion.getvalue())
                    df_costos = cargar_costos(SHEET_ID, hoja_costos)
                    resultado = procesar_produccion(archivo_bytes, df_costos)

                st.success("Procesamiento completado.")
                c1, c2, c3 = st.columns(3)
                c1.metric("Registros resultantes", f"{len(resultado):,}")
                c2.metric("Costo total calculado", f"{resultado['costo_total'].sum(skipna=True):,.2f}")
                c3.metric("Registros sin costo", f"{resultado['costo_unitario'].isna().sum():,}")

                st.subheader("Vista previa del resultado")
                st.dataframe(resultado.head(500), use_container_width=True)

                salida = io.BytesIO()
                with pd.ExcelWriter(salida, engine="openpyxl") as writer:
                    resultado.to_excel(writer, index=False, sheet_name="Produccion_Costos")
                    sin_costo = resultado[resultado["costo_unitario"].isna()]
                    if not sin_costo.empty:
                        sin_costo.to_excel(writer, index=False, sheet_name="Sin_Costo")
                salida.seek(0)

                st.download_button(
                    label="⬇️ Descargar Excel procesado",
                    data=salida.getvalue(),
                    file_name="produccion_con_costos.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            except Exception as exc:
                st.error(f"No se pudo procesar el archivo: {exc}")
                st.info(
                    "Verifica que la hoja de producción tenga la estructura esperada "
                    "y que la hoja seleccionada en Google Sheets contenga los encabezados requeridos."
                )
except Exception as exc:
    st.error(f"No se pudo conectar con Google Sheets: {exc}")
    st.markdown(
        "Revisa que hayas configurado `gcp_service_account` en Streamlit Secrets, "
        "que el JSON de la cuenta de servicio sea válido y que hayas compartido "
        "el Google Sheets con el correo de la cuenta de servicio."
    )
