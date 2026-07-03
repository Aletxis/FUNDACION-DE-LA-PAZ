"""
Dashboard exploratorio - Reporte Consolidado de Movilidad Humana
Ejecutar localmente con:  streamlit run app.py
"""

import base64
import io
import re
import unicodedata
import pandas as pd
import numpy as np
import plotly.express as px
import streamlit as st
from pathlib import Path

# ----------------------------------------------------------------------------
# Configuración general
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="Observatorio - Movilidad Humana",
    page_icon="🌎",
    layout="wide",
)

DATA_FILENAME = "REPORTE_CONSOLIDADO_MOVILIDAD_HUMANA.xlsx"


def _require_uploaded_file() -> bytes:
    """Exige que el usuario suba el .xlsx en CADA sesión de navegador, vía
    st.file_uploader. El archivo vive solo en memoria de esa sesión (RAM del
    proceso de Streamlit) y nunca se guarda en disco ni se sube a ningún
    repositorio o link externo. Diseño intencional: estos datos contienen PII
    de personas en movilidad humana, incluyendo menores (NNA), y no pueden
    depender de un link "compartido" ni vivir en Git bajo ninguna circunstancia."""
    if "uploaded_xlsx_bytes" not in st.session_state:
        st.session_state["uploaded_xlsx_bytes"] = None

    if st.session_state["uploaded_xlsx_bytes"] is None:
        st.title("🌎 Observatorio - Movilidad Humana")
        st.warning(
            f"Por seguridad, sube tu archivo **{DATA_FILENAME}** para esta sesión. "
            "No se guarda en el servidor ni se comparte; solo vive en tu navegador "
            "mientras uses el dashboard."
        )
        uploaded = st.file_uploader("Archivo de datos (.xlsx)", type=["xlsx"])
        if uploaded is None:
            st.stop()
        st.session_state["uploaded_xlsx_bytes"] = uploaded.getvalue()
        st.rerun()

    return st.session_state["uploaded_xlsx_bytes"]

# Rutas de los logos (ajusta el nombre/extensión si es necesario)
LOGO1_PATH = Path(__file__).parent / "logo1.jpeg"   # Fundación Mensajeros de la Paz
LOGO2_PATH = Path(__file__).parent / "logo2.jpeg"   # Tec.Azuay

# Coordenadas aproximadas de las zonas de planificación presentes en los datos
# (se usan para el mapa, ya que la base solo trae Zona/Provincia/Ciudad, no coordenadas exactas)
ZONA_COORDS = {
    "ZONA 6": {"lat": -2.9001, "lon": -79.0059, "label": "Zona 6 (Cuenca - Austro)"},
    "ZONA 7": {"lat": -3.9931, "lon": -79.2042, "label": "Zona 7 (Loja - Sur)"},
    "ZONA 8": {"lat": -2.1710, "lon": -79.9224, "label": "Zona 8 (Guayaquil)"},
}

SI_NO_COLS = [
    "atencion_emergente", "kit_aseo", "kit_salud", "kit_escolar",
    "enfermedad_catastrofica", "tiene_discapacidad", "embarazo", "estudiando",
    "atencion_trabajo_social", "atencion_psicologica", "atencion_legal",
    "serv_salud", "serv_educacion", "serv_junta_cantonal",
    "serv_reunificacion_familiar", "serv_eti", "serv_acogimiento_institucional",
    "serv_apoyo_custodia_familiar", "serv_discapacidades", "serv_adulto_mayor",
    "serv_cdi", "serv_cnh", "part_talleres_capacitacion",
    "part_talleres_sensibilizacion", "part_encuentros_comunitarios",
    "part_talleres_nna", "part_redes_comunitarias",
]

FRIENDLY_NAMES = {
    "atencion_trabajo_social": "Trabajo social",
    "atencion_psicologica": "Atención psicológica",
    "atencion_legal": "Atención legal",
    "serv_salud": "Salud",
    "serv_educacion": "Educación",
    "serv_junta_cantonal": "Junta cantonal",
    "serv_reunificacion_familiar": "Reunificación familiar",
    "serv_eti": "ETI",
    "serv_acogimiento_institucional": "Acogimiento institucional",
    "serv_apoyo_custodia_familiar": "Apoyo y custodia familiar",
    "serv_discapacidades": "Discapacidades",
    "serv_adulto_mayor": "Adulto mayor",
    "serv_cdi": "CDI",
    "serv_cnh": "CNH",
    "kit_aseo": "Kit de aseo",
    "kit_salud": "Kit de salud",
    "kit_escolar": "Kit escolar",
    "part_talleres_capacitacion": "Talleres de capacitación",
    "part_talleres_sensibilizacion": "Talleres de sensibilización",
    "part_encuentros_comunitarios": "Encuentros comunitarios",
    "part_talleres_nna": "Talleres lúdicos NNA",
    "part_redes_comunitarias": "Inserción a redes comunitarias",
}

# Tarjetas de categoría que se muestran en la portada (icono opcional + texto)
PORTADA_CARDS = [
    "Perfil de la Población",
    "Vulnerabilidades",
    "Intervenciones Técnicas",
    "Situación Migratoria",
    "Asistencia Humanitaria",
    "Integración Comunitaria",
]


def load_data() -> pd.DataFrame:
    file_bytes = _require_uploaded_file()
    try:
        return _process_uploaded_bytes(file_bytes)
    except ImportError:
        st.error("Falta instalar 'openpyxl' para leer el archivo Excel. Ejecuta: pip install openpyxl")
        st.stop()


@st.cache_data(show_spinner="Procesando archivo...")
def _process_uploaded_bytes(file_bytes: bytes) -> pd.DataFrame:
    """Cacheado por Streamlit en memoria, con hash de los bytes subidos. Así
    no se reprocesa el Excel completo en cada rerun de la sesión (ej. al
    cambiar de sección del dashboard), sin tocar el disco."""
    df = _read_source_excel(file_bytes)
    return _prepare_excel_data(df)


def si_pct(series: pd.Series) -> float:
    """% de 'SI' sobre los valores no nulos de una columna SI/NO."""
    s = series.dropna()
    if len(s) == 0:
        return 0.0
    return 100 * (s == "SI").sum() / len(s)


def _img_to_base64(path: Path) -> str | None:
    """Convierte una imagen local a base64 para poder incrustarla en HTML."""
    if not path.exists():
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


MESES_ES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre",
}

MESES_NUM = {v.upper(): k for k, v in MESES_ES.items()}

COLUMN_MAP = {
    "ZONA": "zona",
    "PROVINCIA": "provincia",
    "DISTRITO": "distrito",
    "CIUDAD": "ciudad",
    "FECHA_DE_NACIMIENTO": "fecha_nacimiento",
    "EDAD_ANOS": "edad_anios",
    "EDAD_MESES": "edad_meses",
    "RANGO_DE_EDAD": "rango_edad",
    "SEXO": "sexo",
    "GENERO": "genero",
    "NACIONALIDAD": "nacionalidad",
    "ETNIA": "etnia",
    "SITUACION_DE_MOVILIDAD": "situacion_movilidad",
    "FORMA_DE_INGRESO_AL_ECUADOR": "forma_ingreso",
    "ATENCION_EMERGENTE": "atencion_emergente",
    "KIT_DE_ASEO": "kit_aseo",
    "KIT_DE_SALUD": "kit_salud",
    "KIT_ESCOLAR": "kit_escolar",
    "TIENE_ENFERMEDAD_CATASTROFICA": "enfermedad_catastrofica",
    "TIENE_DISCAPACIDAD": "tiene_discapacidad",
    "EMBARAZO_SOLO_PARA_MUJERES": "embarazo",
    "ACTUALMENTE_ESTA_ESTUDIANDO": "estudiando",
    "ATENCION_DE_TRABAJO_SOCIAL": "atencion_trabajo_social",
    "ATENCION_PSICOLOGICA": "atencion_psicologica",
    "ATENCION_LEGAL": "atencion_legal",
    "SALUD": "serv_salud",
    "EDUCACION": "serv_educacion",
    "JUNTA_CANTONAL": "serv_junta_cantonal",
    "REUNIFICACION_FAMILIAR": "serv_reunificacion_familiar",
    "ETI": "serv_eti",
    "ACOGIMINETO_INSTITUCIONAL": "serv_acogimiento_institucional",
    "ACOGIMIENTO_INSTITUCIONAL": "serv_acogimiento_institucional",
    "APOYO_Y_CUSTODIA_FAMILIAR": "serv_apoyo_custodia_familiar",
    "DISCAPACIDADES": "serv_discapacidades",
    "ADULTO_MAYOR": "serv_adulto_mayor",
    "CDI": "serv_cdi",
    "CNH": "serv_cnh",
    "PARTICIPACION_A_TALLERES_DE_CAPACITACION": "part_talleres_capacitacion",
    "PARTICIPACION_A_TALLERES_DE_SENSIBILIZACION": "part_talleres_sensibilizacion",
    "PARTICIPACION_EN_ENCUENTROS_COMUNITARIOS": "part_encuentros_comunitarios",
    "PARTICIPACION_EN_TALLERES_Y_ESPACIOS_LUDICO_DE_RECREACION_DE_NNA": "part_talleres_nna",
    "INSERCION_A_REDES_COMUNITARIAS": "part_redes_comunitarias",
    "SITUACION_MIGRATORIA": "situacion_migratoria",
    "NOMBRE_DE_LA_HOJA": "nombre_hoja",
    "NOMBRE_DEL_ARCHIVO": "nombre_archivo",
}


def _norm_text(value) -> str:
    text = "" if pd.isna(value) else str(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    return text.strip().upper()


def _norm_col(value) -> str:
    text = _norm_text(value)
    text = re.sub(r"[^A-Z0-9]+", "_", text)
    return text.strip("_")


def _read_source_excel(file_bytes: bytes) -> pd.DataFrame:
    return pd.read_excel(
        io.BytesIO(file_bytes),
        dtype=str,
        usecols=lambda col: _norm_col(col) in COLUMN_MAP,
    )


def _parse_period_from_file(value) -> pd.Timestamp:
    text = _norm_text(value)
    match = re.search(
        r"\b(ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)[\s_]+(20\d{2})(?=\D|$)",
        text,
    )
    if not match:
        return pd.NaT
    return pd.Timestamp(year=int(match.group(2)), month=MESES_NUM[match.group(1)], day=1)


def _clean_yes_no_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in SI_NO_COLS:
        if col in df.columns:
            df[col] = df[col].map(_norm_text).replace({"": np.nan, "N/D": np.nan})
    return df


def _clean_dimension_columns(df: pd.DataFrame) -> pd.DataFrame:
    dims = [
        "zona", "provincia", "distrito", "ciudad", "rango_edad", "sexo", "genero",
        "nacionalidad", "etnia", "situacion_movilidad", "forma_ingreso",
        "situacion_migratoria",
    ]
    for col in dims:
        if col in df.columns:
            df[col] = df[col].map(_norm_text).replace({"": np.nan, "N/D": np.nan})
    if "zona" in df.columns:
        df["zona"] = df["zona"].str.replace(r"\s+", " ", regex=True).str.strip()
    return df


def _prepare_excel_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.rename(columns={col: COLUMN_MAP.get(_norm_col(col), _norm_col(col).lower()) for col in df.columns})

    if "nombre_archivo" in df.columns:
        df["periodo"] = df["nombre_archivo"].apply(_parse_period_from_file)
    else:
        df["periodo"] = pd.NaT

    if "fecha_nacimiento" in df.columns:
        df["fecha_nacimiento"] = pd.to_datetime(df["fecha_nacimiento"], errors="coerce")

    if "edad_anios" in df.columns:
        df["edad_anios"] = df["edad_anios"].astype(str).str.extract(r"(\d+)", expand=False).astype(float)

    df = _clean_dimension_columns(df)
    df = _clean_yes_no_columns(df)
    df = df.dropna(subset=["periodo"])
    if df.empty:
        st.error("No se pudo identificar el periodo mensual desde la columna 'Nombre del Archivo' del Excel.")
        st.stop()
    df["anio"] = df["periodo"].dt.year.astype("Int64")
    df["mes_num"] = df["periodo"].dt.month.astype("Int64")

    expected_cols = [
        "periodo", "anio", "mes_num", "zona", "provincia", "distrito", "ciudad",
        "rango_edad", "sexo", "genero", "nacionalidad", "etnia",
        "situacion_movilidad", "forma_ingreso", "situacion_migratoria", "edad_anios",
        *SI_NO_COLS,
    ]
    for col in expected_cols:
        if col not in df.columns:
            df[col] = np.nan
    return df[expected_cols]


def formato_mes_anio(ts: pd.Timestamp) -> str:
    return f"{MESES_ES[ts.month]} {ts.year}"


# ----------------------------------------------------------------------------
# Portada / pantalla de bienvenida
# ----------------------------------------------------------------------------
if "ingresado" not in st.session_state:
    st.session_state["ingresado"] = False
if "seccion" not in st.session_state:
    st.session_state["seccion"] = PORTADA_CARDS[0]


def mostrar_portada():
    df_preview = load_data()
    rango_fechas = f"{formato_mes_anio(df_preview['periodo'].min())} - {formato_mes_anio(df_preview['periodo'].max())}"

    logo1_b64 = _img_to_base64(LOGO1_PATH)
    logo2_b64 = _img_to_base64(LOGO2_PATH)

    logo1_html = (
        f'<img src="data:image/jpeg;base64,{logo1_b64}" class="portada-logo" />'
        if logo1_b64 else ""
    )
    logo2_html = (
        f'<img src="data:image/jpeg;base64,{logo2_b64}" class="portada-logo" />'
        if logo2_b64 else ""
    )

    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{display: none;}}
        [data-testid="stAppViewContainer"] > .main {{
            background: #ffffff;
        }}
        .portada-wrap {{
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            padding-top: 3vh;
        }}
        .portada-logos {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 40px;
            margin-bottom: 1.8rem;
        }}
        .portada-logo {{
            height: 100px;
            object-fit: contain;
        }}
        .portada-titulo-row {{
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 22px;
            margin-bottom: 1.6rem;
        }}
        .portada-titulo {{
            font-size: 4rem;
            font-weight: 800;
            color: #14208a;
            letter-spacing: 1px;
            margin: 0;
        }}
        .portada-barra {{
            width: 6px;
            height: 60px;
            background: #b9b9b9;
            border-radius: 3px;
        }}
        .portada-info-box {{
            border: 1px solid #d9d9d9;
            border-radius: 8px;
            padding: 10px 28px;
            font-size: 1.15rem;
            font-weight: 700;
            color: #14208a;
            margin-bottom: 10px;
        }}
        .portada-fecha-box {{
            border: 1px solid #d9d9d9;
            border-radius: 8px;
            padding: 8px 28px;
            font-size: 1.05rem;
            font-weight: 700;
            color: #e0a800;
            margin-bottom: 2.2rem;
        }}
        div[data-testid="stButton"] > button {{
            border: 1px solid #d9d9d9;
            border-radius: 8px;
            padding: 22px 14px;
            font-size: 1.05rem;
            font-weight: 700;
            color: #4a4a4a;
            min-height: 80px;
            background: #ffffff;
        }}
        div[data-testid="stButton"] > button:hover {{
            border-color: #14208a;
            color: #14208a;
            background: #f7f8ff;
        }}
        div[data-testid="stButton"] > button[kind="primary"] {{
            display: none;
        }}
        @media (max-width: 900px) {{
            .portada-titulo {{ font-size: 2.6rem; }}
        }}
        </style>

        <div class="portada-wrap">
            <div class="portada-logos">
                {logo1_html}
                {logo2_html}
            </div>
            <div class="portada-titulo-row">
                <div class="portada-titulo">OBSERVATORIO</div>
                <div class="portada-barra"></div>
            </div>
            <div class="portada-info-box">Datos disponibles</div>
            <div class="portada-fecha-box">{rango_fechas}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    for row_start in range(0, len(PORTADA_CARDS), 3):
        cols = st.columns(3)
        for col, card in zip(cols, PORTADA_CARDS[row_start:row_start + 3]):
            with col:
                if st.button(card, use_container_width=True):
                    st.session_state["seccion"] = card
                    st.session_state["ingresado"] = True
                    st.rerun()

    col_a, col_b, col_c = st.columns([1, 1, 1])
    with col_b:
        if st.button("Ingresar ➜", use_container_width=True, type="primary"):
            st.session_state["ingresado"] = True
            st.rerun()

    st.caption(
        "Fuente: Reporte Consolidado de Movilidad Humana. "
        "Este panel no expone nombres, direcciones, teléfonos, correos ni documentos de identificación."
    )


if not st.session_state["ingresado"]:
    mostrar_portada()
    st.stop()

# ----------------------------------------------------------------------------
# Carga de datos
# ----------------------------------------------------------------------------
df = load_data()

top_col1, top_col2 = st.columns([6, 1])
with top_col1:
    st.title("🌎 Dashboard de Movilidad Humana")
    st.caption(
        "Reporte consolidado de caracterización del Servicio de Movilidad Humana. "
        "Los datos mostrados son agregados y no incluyen información de identificación "
        "personal (nombres, direcciones, teléfonos, correos ni documentos)."
    )
with top_col2:
    if st.button("⟵ Salir"):
        st.session_state["ingresado"] = False
        st.rerun()

# ----------------------------------------------------------------------------
# Filtros (sidebar)
# ----------------------------------------------------------------------------
st.sidebar.header("Filtros")

min_periodo, max_periodo = df["periodo"].min(), df["periodo"].max()
periodo_range = st.sidebar.slider(
    "Periodo (mes)",
    min_value=min_periodo.to_pydatetime(),
    max_value=max_periodo.to_pydatetime(),
    value=(min_periodo.to_pydatetime(), max_periodo.to_pydatetime()),
    format="MMM YYYY",
)

def multiselect_filter(label, col):
    options = sorted(df[col].dropna().unique().tolist())
    return st.sidebar.multiselect(label, options, default=[])

zona_sel = multiselect_filter("Zona", "zona")
nacionalidad_sel = multiselect_filter("Nacionalidad", "nacionalidad")
sexo_sel = multiselect_filter("Sexo", "sexo")
rango_edad_sel = multiselect_filter("Rango de edad", "rango_edad")
etnia_sel = multiselect_filter("Etnia", "etnia")
situacion_migratoria_sel = multiselect_filter("Situación migratoria", "situacion_migratoria")
forma_ingreso_sel = multiselect_filter("Forma de ingreso", "forma_ingreso")

# Aplicar filtros
f = df[(df["periodo"] >= pd.Timestamp(periodo_range[0])) & (df["periodo"] <= pd.Timestamp(periodo_range[1]))]

def apply_filter(frame, col, selection):
    if selection:
        return frame[frame[col].isin(selection)]
    return frame

f = apply_filter(f, "zona", zona_sel)
f = apply_filter(f, "nacionalidad", nacionalidad_sel)
f = apply_filter(f, "sexo", sexo_sel)
f = apply_filter(f, "rango_edad", rango_edad_sel)
f = apply_filter(f, "etnia", etnia_sel)
f = apply_filter(f, "situacion_migratoria", situacion_migratoria_sel)
f = apply_filter(f, "forma_ingreso", forma_ingreso_sel)

st.sidebar.markdown(f"**Registros filtrados:** {len(f):,}".replace(",", "."))

f = df.copy()

if len(f) == 0:
    st.warning("No hay registros para los filtros seleccionados.")
    st.stop()

f = df.copy()
seccion_actual = st.session_state.get("seccion", PORTADA_CARDS[0])
st.markdown(
    """
    <style>
    [data-testid="stSidebar"] {display: none;}
    [data-testid="collapsedControl"] {display: none;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# KPIs
# ----------------------------------------------------------------------------
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total de registros", f"{len(f):,}".replace(",", "."))
col2.metric("% Mujeres", f"{100 * (f['sexo'] == 'MUJER').sum() / f['sexo'].notna().sum():.1f}%" if f['sexo'].notna().sum() else "N/D")
col3.metric("Nacionalidades distintas", f["nacionalidad"].nunique())
irregular_pct = 100 * (f["situacion_migratoria"] == "IRREGULAR").sum() / f["situacion_migratoria"].notna().sum() if f["situacion_migratoria"].notna().sum() else 0
col4.metric("% Situación irregular", f"{irregular_pct:.1f}%")
nna_pct = 100 * f["rango_edad"].isin(["NN", "ADOLESCENTE"]).sum() / f["rango_edad"].notna().sum() if f["rango_edad"].notna().sum() else 0
col5.metric("% NNA (niños/adolescentes)", f"{nna_pct:.1f}%")

st.divider()

st.subheader(seccion_actual)


def plot_si_bars(cols: list[str], title: str, height: int = 420) -> None:
    data = []
    for col in cols:
        if col in f.columns and f[col].notna().sum() > 0:
            data.append({
                "indicador": FRIENDLY_NAMES.get(col, col),
                "pct_si": si_pct(f[col]),
                "n_valido": f[col].notna().sum(),
            })
    if not data:
        st.info("No hay datos disponibles para esta seccion.")
        return
    chart_df = pd.DataFrame(data).sort_values("pct_si", ascending=True)
    fig = px.bar(chart_df, x="pct_si", y="indicador", orientation="h", title=title)
    fig.update_layout(height=height, xaxis_title="% SI", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True)


if seccion_actual == PORTADA_CARDS[0]:
    st.markdown("**Evolucion mensual**")
    ts = f.groupby("periodo").size().reset_index(name="registros")
    fig_ts = px.line(ts, x="periodo", y="registros", markers=True)
    fig_ts.update_layout(height=380, xaxis_title="Mes", yaxis_title="Numero de registros")
    st.plotly_chart(fig_ts, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        nac = f["nacionalidad"].value_counts().head(10).reset_index()
        nac.columns = ["nacionalidad", "registros"]
        fig_nac = px.bar(nac.sort_values("registros"), x="registros", y="nacionalidad", orientation="h",
                          title="Top 10 nacionalidades")
        fig_nac.update_layout(height=380)
        st.plotly_chart(fig_nac, use_container_width=True)
    with c2:
        edad = f["rango_edad"].value_counts().reset_index()
        edad.columns = ["rango_edad", "registros"]
        fig_edad = px.pie(edad, names="rango_edad", values="registros", title="Distribucion por rango de edad", hole=0.4)
        fig_edad.update_layout(height=380)
        st.plotly_chart(fig_edad, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        genero = f["genero"].value_counts().reset_index()
        genero.columns = ["genero", "registros"]
        fig_genero = px.bar(genero, x="genero", y="registros", title="Distribucion por genero")
        fig_genero.update_layout(height=350)
        st.plotly_chart(fig_genero, use_container_width=True)
    with c4:
        ingreso = f["forma_ingreso"].value_counts().reset_index()
        ingreso.columns = ["forma_ingreso", "registros"]
        fig_ing = px.bar(ingreso, x="forma_ingreso", y="registros", title="Forma de ingreso al Ecuador")
        fig_ing.update_layout(height=350)
        st.plotly_chart(fig_ing, use_container_width=True)

    if f["edad_anios"].notna().sum() > 20:
        fig_hist = px.histogram(f.dropna(subset=["edad_anios"]), x="edad_anios", nbins=30,
                                title="Distribucion de edad en anios")
        fig_hist.update_layout(height=320, xaxis_title="Edad", yaxis_title="Numero de registros")
        st.plotly_chart(fig_hist, use_container_width=True)
    st.stop()

if seccion_actual == PORTADA_CARDS[1]:
    plot_si_bars(
        ["tiene_discapacidad", "enfermedad_catastrofica", "embarazo"],
        "Indicadores de vulnerabilidad (% SI)",
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[2]:
    plot_si_bars(
        [
            "atencion_trabajo_social", "atencion_psicologica", "atencion_legal",
            "serv_salud", "serv_educacion", "serv_junta_cantonal",
            "serv_reunificacion_familiar", "serv_eti", "serv_acogimiento_institucional",
            "serv_apoyo_custodia_familiar", "serv_discapacidades", "serv_adulto_mayor",
            "serv_cdi", "serv_cnh",
        ],
        "Intervenciones tecnicas y servicios recibidos (% SI)",
        height=560,
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[3]:
    c1, c2 = st.columns(2)
    with c1:
        mig = f["situacion_migratoria"].value_counts().reset_index()
        mig.columns = ["situacion_migratoria", "registros"]
        fig_mig = px.pie(mig, names="situacion_migratoria", values="registros", title="Situacion migratoria", hole=0.4)
        fig_mig.update_layout(height=380)
        st.plotly_chart(fig_mig, use_container_width=True)
    with c2:
        mov = f["situacion_movilidad"].value_counts().head(10).reset_index()
        mov.columns = ["situacion_movilidad", "registros"]
        fig_mov = px.bar(mov.sort_values("registros"), x="registros", y="situacion_movilidad", orientation="h",
                         title="Situacion de movilidad")
        fig_mov.update_layout(height=380)
        st.plotly_chart(fig_mov, use_container_width=True)
    st.stop()

if seccion_actual == PORTADA_CARDS[4]:
    plot_si_bars(
        ["atencion_emergente", "kit_aseo", "kit_salud", "kit_escolar"],
        "Asistencia humanitaria entregada (% SI)",
    )
    st.stop()

if seccion_actual == PORTADA_CARDS[5]:
    plot_si_bars(
        [
            "part_talleres_capacitacion", "part_talleres_sensibilizacion",
            "part_encuentros_comunitarios", "part_talleres_nna",
            "part_redes_comunitarias",
        ],
        "Participacion e integracion comunitaria (% SI)",
    )
    st.stop()

# ----------------------------------------------------------------------------
# Serie de tiempo
# ----------------------------------------------------------------------------
st.subheader("📈 Evolución mensual")

ts_dim = st.radio(
    "Desagregar serie de tiempo por:",
    ["(sin desagregar)", "sexo", "nacionalidad", "situacion_migratoria", "rango_edad"],
    horizontal=True,
)

if ts_dim == "(sin desagregar)":
    ts = f.groupby("periodo").size().reset_index(name="registros")
    fig_ts = px.line(ts, x="periodo", y="registros", markers=True)
else:
    top_cats = f[ts_dim].value_counts().head(6).index
    ts_data = f[f[ts_dim].isin(top_cats)]
    ts = ts_data.groupby(["periodo", ts_dim]).size().reset_index(name="registros")
    fig_ts = px.line(ts, x="periodo", y="registros", color=ts_dim, markers=True)

fig_ts.update_layout(height=420, xaxis_title="Mes", yaxis_title="N° de registros", legend_title=ts_dim)
st.plotly_chart(fig_ts, use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Mapa
# ----------------------------------------------------------------------------
st.subheader("🗺️ Distribución geográfica (por zona de planificación)")
st.caption(
    "La base de datos original solo registra Zona/Provincia/Ciudad de atención "
    "(no coordenadas exactas por persona), por lo que el mapa agrega los registros "
    "a nivel de zona de planificación."
)

zona_counts = f["zona"].value_counts().reset_index()
zona_counts.columns = ["zona", "registros"]
zona_counts["zona_norm"] = zona_counts["zona"].str.replace(r"\s+", " ", regex=True).str.strip()

map_rows = []
for _, row in zona_counts.iterrows():
    key = row["zona_norm"]
    if key in ZONA_COORDS:
        c = ZONA_COORDS[key]
        map_rows.append({"zona": c["label"], "lat": c["lat"], "lon": c["lon"], "registros": row["registros"]})

if map_rows:
    map_df = pd.DataFrame(map_rows)
    fig_map = px.scatter_mapbox(
        map_df, lat="lat", lon="lon", size="registros", color="registros",
        hover_name="zona", hover_data={"registros": True, "lat": False, "lon": False},
        zoom=5.5, height=450, color_continuous_scale="Reds",
    )
    fig_map.update_layout(mapbox_style="open-street-map", margin=dict(l=0, r=0, t=0, b=0))
    st.plotly_chart(fig_map, use_container_width=True)
else:
    st.info("No hay zonas reconocidas para graficar en el mapa con los filtros actuales.")

st.divider()

# ----------------------------------------------------------------------------
# Distribuciones demográficas
# ----------------------------------------------------------------------------
st.subheader("👥 Perfil demográfico")

c1, c2 = st.columns(2)

with c1:
    nac = f["nacionalidad"].value_counts().head(10).reset_index()
    nac.columns = ["nacionalidad", "registros"]
    fig_nac = px.bar(nac.sort_values("registros"), x="registros", y="nacionalidad", orientation="h",
                      title="Top 10 nacionalidades")
    fig_nac.update_layout(height=380)
    st.plotly_chart(fig_nac, use_container_width=True)

with c2:
    edad = f["rango_edad"].value_counts().reset_index()
    edad.columns = ["rango_edad", "registros"]
    fig_edad = px.pie(edad, names="rango_edad", values="registros", title="Distribución por rango de edad", hole=0.4)
    fig_edad.update_layout(height=380)
    st.plotly_chart(fig_edad, use_container_width=True)

c3, c4 = st.columns(2)

with c3:
    genero = f["genero"].value_counts().reset_index()
    genero.columns = ["genero", "registros"]
    fig_genero = px.bar(genero, x="genero", y="registros", title="Distribución por género")
    fig_genero.update_layout(height=350)
    st.plotly_chart(fig_genero, use_container_width=True)

with c4:
    ingreso = f["forma_ingreso"].value_counts().reset_index()
    ingreso.columns = ["forma_ingreso", "registros"]
    fig_ing = px.bar(ingreso, x="forma_ingreso", y="registros", title="Forma de ingreso al Ecuador")
    fig_ing.update_layout(height=350)
    st.plotly_chart(fig_ing, use_container_width=True)

if f["edad_anios"].notna().sum() > 20:
    st.markdown("**Distribución de edad (años) - solo registros con dato disponible**")
    fig_hist = px.histogram(f.dropna(subset=["edad_anios"]), x="edad_anios", nbins=30)
    fig_hist.update_layout(height=320, xaxis_title="Edad (años)", yaxis_title="N° de registros")
    st.plotly_chart(fig_hist, use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Servicios y atención brindada
# ----------------------------------------------------------------------------
st.subheader("🩺 Servicios y atención brindada (% que recibió 'SI')")

svc_data = []
for col in SI_NO_COLS:
    if col in f.columns:
        pct = si_pct(f[col])
        n = f[col].notna().sum()
        if n > 0:
            svc_data.append({"servicio": FRIENDLY_NAMES.get(col, col), "pct_si": pct, "n_valido": n})

svc_df = pd.DataFrame(svc_data).sort_values("pct_si", ascending=True)
fig_svc = px.bar(svc_df, x="pct_si", y="servicio", orientation="h",
                  title="% de casos con atención/servicio recibido (SI)")
fig_svc.update_layout(height=650, xaxis_title="% SI", yaxis_title="")
st.plotly_chart(fig_svc, use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Vulnerabilidad y situación migratoria
# ----------------------------------------------------------------------------
st.subheader("⚠️ Situación migratoria y vulnerabilidad")

c5, c6 = st.columns(2)
with c5:
    mig = f["situacion_migratoria"].value_counts().reset_index()
    mig.columns = ["situacion_migratoria", "registros"]
    fig_mig = px.pie(mig, names="situacion_migratoria", values="registros", title="Situación migratoria", hole=0.4)
    fig_mig.update_layout(height=380)
    st.plotly_chart(fig_mig, use_container_width=True)

with c6:
    vuln_cols = ["tiene_discapacidad", "enfermedad_catastrofica", "embarazo"]
    vuln_data = [{"indicador": FRIENDLY_NAMES.get(c, c), "pct_si": si_pct(f[c])} for c in vuln_cols if f[c].notna().sum() > 0]
    if vuln_data:
        fig_vuln = px.bar(pd.DataFrame(vuln_data), x="indicador", y="pct_si", title="Indicadores de vulnerabilidad (% SI)")
        fig_vuln.update_layout(height=380, yaxis_title="% SI")
        st.plotly_chart(fig_vuln, use_container_width=True)

st.divider()

# ----------------------------------------------------------------------------
# Tabla agregada (sin datos personales)
# ----------------------------------------------------------------------------
st.subheader("📋 Tabla resumen (agregada por periodo, nacionalidad y sexo)")
tabla = f.groupby(["periodo", "nacionalidad", "sexo"]).size().reset_index(name="registros")
tabla = tabla.sort_values(["periodo", "registros"], ascending=[False, False])
st.dataframe(tabla, use_container_width=True, height=350)

csv_export = tabla.to_csv(index=False).encode("utf-8")
st.download_button("⬇️ Descargar tabla resumen (CSV)", csv_export, "resumen_movilidad_humana.csv", "text/csv")

st.caption(
    "Fuente: Reporte Consolidado de Movilidad Humana. "
    "Este panel no expone nombres, direcciones, teléfonos, correos ni documentos de identificación."
)
