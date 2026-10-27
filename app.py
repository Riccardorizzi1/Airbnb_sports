# AIRBNB&SPORTS
# app.py — Streamlit + Kepler.gl
#
# Dipendenze:
# streamlit
# pandas
# plotly
# openpyxl
# keplergl

from pathlib import Path
from urllib.request import Request, urlopen
import base64
import html
import json
import math
import re
import tempfile
import unicodedata

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components

try:
    from keplergl import KeplerGl
except ImportError:
    KeplerGl = None


# ============================================================
# CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Airbnb&Sports",
    page_icon="🏟️",
    layout="wide",
)

ROOT = Path(__file__).resolve().parent
VERSIONE = "v4.1 · Eventi e simulazione"
SITO_AZIENDA = "https://www.sblconsultancy.it/"

SCENARI = ["Prudente", "Intermedio", "Espansivo"]
COLORI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C",
}

PAGINE = [
    "Panorama nazionale",
    "Esplora un territorio",
    "Eventi e simulatore",
    "Forecast 2026–2030",
    "Dati e metodo",
]


def file_dati(nome):
    for cartella in (ROOT, ROOT / "output_eurostat"):
        percorso = cartella / nome
        if percorso.exists():
            return percorso
    return ROOT / nome


FILES = {
    "reg": file_dati("eurostat_master_final.csv"),
    "prov": file_dati("eurostat_nuts3_annuale_2018_2025.csv"),
    "fr": file_dati("forecast_nuts2_2026_2030_3_scenari.csv"),
    "fp": file_dati("forecast_nuts3_2026_2030_3_scenari.csv"),
    "eventi": file_dati(
        "AirbnbSports_Event_Master_Simone_Visitatori.xlsx"
    ),
}

SCREENING = [
    "screening_eventi_completo.csv",
    "screening_regioni_mesi_completo.csv",
    "screening_eventi_equivalenti.csv",
    "screening_regioni_mesi_equivalenti.csv",
    "screening_pernottamenti_5_scenari_notti.csv",
    "screening_pernottamenti_5_scenari.csv",
    "screening_2pct.csv",
    "event_exposure_monthly.csv",
]


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def normalizza(valore):
    testo = unicodedata.normalize("NFKD", str(valore))
    testo = "".join(c for c in testo if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", testo.lower())


def testo(valore):
    if pd.isna(valore):
        return ""
    return str(valore).strip()


def numero(valore, decimali=0):
    if pd.isna(valore):
        return "n.d."
    return (
        f"{float(valore):,.{decimali}f}"
        .replace(",", "§")
        .replace(".", ",")
        .replace("§", ".")
    )


def percentuale(valore):
    return "n.d." if pd.isna(valore) else f"{numero(valore, 1)}%"


def numero_dato(valore):
    """Legge numeri; non inventa valori da intervalli o descrizioni."""
    if pd.isna(valore):
        return float("nan")

    if isinstance(valore, (int, float)):
        return float(valore) if math.isfinite(float(valore)) else float("nan")

    s = str(valore).strip().replace("\u00a0", "").replace(" ", "")
    s = s.replace("'", "")

    if not re.fullmatch(r"[+-]?\d+(?:[.,]\d+)*", s):
        return float("nan")

    if "." in s and "," in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+", s):
        s = s.replace(".", "")
    elif "," in s:
        s = s.replace(",", ".")

    try:
        risultato = float(s)
        return risultato if math.isfinite(risultato) else float("nan")
    except ValueError:
        return float("nan")


def data_singola(valore):
    if pd.isna(valore):
        return pd.NaT

    if isinstance(valore, (int, float)):
        if 20000 <= float(valore) <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(float(valore), unit="D")
            ).normalize()
        return pd.NaT

    s = str(valore).strip()
    iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", s))
    d = pd.to_datetime(s, errors="coerce", dayfirst=not iso)

    if pd.isna(d):
        return pd.NaT
    if d.tzinfo is not None:
        d = d.tz_localize(None)

    return d.normalize()


def trova_colonna(df, *nomi):
    disponibili = {normalizza(c): c for c in df.columns}
    for nome in nomi:
        if normalizza(nome) in disponibili:
            return disponibili[normalizza(nome)]
    return None


@st.cache_data(show_spinner=False)
def leggi_csv(percorso, modifica):
    df = pd.read_csv(percorso, encoding="utf-8-sig")
    if len(df.columns) == 1 and ";" in str(df.columns[0]):
        df = pd.read_csv(percorso, encoding="utf-8-sig", sep=";")
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def fogli_excel(percorso, modifica):
    with pd.ExcelFile(percorso) as libro:
        return libro.sheet_names


@st.cache_data(show_spinner=False)
def leggi_excel(percorso, modifica, foglio):
    df = pd.read_excel(percorso, sheet_name=foglio)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def carica_csv(percorso):
    return leggi_csv(str(percorso), percorso.stat().st_mtime)


def prepara_tabella(df):
    risultato = df.copy()

    for col in risultato.columns:
        if pd.api.types.is_datetime64_any_dtype(risultato[col]):
            risultato[col] = risultato[col].dt.strftime("%d/%m/%Y")
        elif risultato[col].dtype == object:
            risultato[col] = risultato[col].map(
                lambda x: pd.NA
                if isinstance(x, str)
                and x.strip().lower() in ("", "nan", "none", "nat", "null")
                else x
            )

            if normalizza(col) in (
                "datainizio", "datafine", "startdate", "enddate", "data"
            ):
                originali = risultato[col]
                date = originali.map(data_singola)
                risultato[col] = originali.where(
                    date.isna(), date.dt.strftime("%d/%m/%Y")
                )

    return risultato.dropna(axis=1, how="all")


def tabella(df, nome, altezza=320):
    pulita = prepara_tabella(df)

    if pulita.empty:
        st.info("Nessuna riga per questa selezione.")
        return

    st.dataframe(
        pulita,
        use_container_width=True,
        hide_index=True,
        height=min(altezza, 38 + 35 * len(pulita)),
    )

    st.download_button(
        "Scarica CSV",
        pulita.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=nome,
        mime="text/csv",
        key=f"download_{nome}",
    )


def grafico(fig, altezza=340):
    fig.update_layout(
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TESTO, size=12),
        margin=dict(l=12, r=12, t=30, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, automargin=True)
    fig.update_yaxes(gridcolor=BORDO, automargin=True)
    st.plotly_chart(fig, use_container_width=True)


def somma_completa(serie):
    """Un aggregato con valori mancanti resta non disponibile."""
    if serie.empty or serie.isna().any():
        return float("nan")
    return serie.sum()


def aggrega(df, gruppi, campo):
    return (
        df.groupby(gruppi, dropna=False)[campo]
        .agg(somma_completa)
        .reset_index()
    )


# ============================================================
# LOGO E STILE
# ============================================================

def mostra_logo():
    cartelle = [
        ROOT,
        ROOT / "assets",
        ROOT / "images",
        ROOT / "img",
        ROOT / "output_eurostat",
    ]
    formati = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }

    for cartella in cartelle:
        if not cartella.exists():
            continue

        for percorso in sorted(cartella.iterdir()):
            if (
                percorso.is_file()
                and percorso.stem.lower() == "logo_sbl"
                and percorso.suffix.lower() in formati
            ):
                dati = base64.b64encode(percorso.read_bytes()).decode()
                mime = formati[percorso.suffix.lower()]
                st.markdown(
                    f"""
                    <a href="{SITO_AZIENDA}" target="_blank"
                       rel="noopener noreferrer"
                       aria-label="Visita il sito SBL Consultancy">
                        <img src="data:{mime};base64,{dati}"
                             alt="SBL Consultancy"
                             style="display:block;width:116px;
                                    max-width:100%;height:auto;
                                    margin:0 0 12px 0;">
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
                return

    st.markdown(f"[**SBL CONSULTANCY**]({SITO_AZIENDA})")


with st.sidebar:
    mostra_logo()
    st.markdown("### Airbnb&Sports")
    st.caption("Tourism Demand · Italia")
    pagina = st.radio("Sezione", PAGINE)
    dark = st.checkbox("Tema scuro", value=False)
    st.caption(VERSIONE)

if dark:
    TESTO = "#F3FAFC"
    SECONDARIO = "#D3E5EC"
    BORDO = "#315462"
    SFONDO = "linear-gradient(100deg,#071822 0%,#0C2B39 47%,#10475B 100%)"
    SIDE = "linear-gradient(180deg,#071F31 0%,#0E4058 52%,#126982 100%)"
    CARD = "#102D3B"
else:
    TESTO = "#082B46"
    SECONDARIO = "#254A5D"
    BORDO = "#B5D8E1"
    SFONDO = "linear-gradient(100deg,#C7F1F6 0%,#9FE5F0 47%,#5BC6E2 100%)"
    SIDE = "linear-gradient(180deg,#153F57 0%,#17627E 52%,#259BBA 100%)"
    CARD = "#FDFEFF"

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {SFONDO};
        color: {TESTO};
    }}
    .block-container {{
        max-width: 1320px;
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        padding-left: 1.6rem;
        padding-right: 1.6rem;
    }}
    [data-testid="stSidebar"] {{
        background: {SIDE};
        min-width: 260px;
        max-width: 280px;
    }}
    [data-testid="stSidebar"] [data-testid="stSidebarContent"] {{
        padding-top: 0.5rem;
    }}
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] {{
        color: #FFFFFF !important;
    }}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #E2F0F5 !important;
    }}
    h1 {{font-size: 1.8rem !important; padding-bottom: 0.3rem !important;}}
    h2 {{font-size: 1.35rem !important;}}
    h3 {{font-size: 1.08rem !important;}}
    p, label, .stMarkdown {{font-size: 0.92rem;}}
    [data-testid="stMain"] h1,
    [data-testid="stMain"] h2,
    [data-testid="stMain"] h3,
    [data-testid="stMain"] p,
    [data-testid="stMain"] label {{
        color: {TESTO};
    }}
    [data-testid="stCaptionContainer"] p {{
        color: {SECONDARIO} !important;
        font-size: 0.82rem !important;
    }}
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stMultiSelect"] [data-baseweb="select"] > div {{
        background: #F7FAFC !important;
        border-color: #C8DDE5 !important;
        color: #082B46 !important;
        min-height: 36px !important;
        border-radius: 8px !important;
    }}
    [data-testid="stSelectbox"] [data-baseweb="select"] *,
    [data-testid="stMultiSelect"] [data-baseweb="select"] * {{
        color: #082B46 !important;
        -webkit-text-fill-color: #082B46 !important;
        font-size: 0.86rem !important;
    }}
    [data-baseweb="popover"] *,
    [role="listbox"] *,
    [role="option"] * {{
        color: #082B46 !important;
        -webkit-text-fill-color: #082B46 !important;
    }}
    [data-baseweb="popover"] [role="listbox"],
    [role="option"] {{
        background-color: #F7FAFC !important;
    }}
    [data-baseweb="tag"] {{
        background-color: #D8EDF3 !important;
        border-radius: 5px !important;
    }}
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stDateInput"] input {{
        color: #082B46 !important;
        -webkit-text-fill-color: #082B46 !important;
        background: #F7FAFC !important;
        font-size: 0.86rem !important;
    }}
    [data-testid="stMetric"] {{
        background: {CARD};
        border: 1px solid {BORDO};
        border-radius: 12px;
        padding: 12px 14px;
    }}
    [data-testid="stMetricValue"] {{
        color: {TESTO};
        font-size: 1.65rem !important;
    }}
    [data-testid="stMetricLabel"] p {{
        font-size: 0.82rem !important;
    }}
    [data-testid="stVerticalBlock"] {{
        gap: 0.8rem;
    }}
    .stButton > button,
    .stDownloadButton > button {{
        min-height: 34px;
        color: #082B46 !important;
        background: #F7FAFC;
        border: 1px solid #B5D8E1;
        border-radius: 8px;
    }}
    .stButton > button p,
    .stDownloadButton > button p {{
        color: #082B46 !important;
        font-size: 0.84rem !important;
    }}
    button[data-baseweb="tab"] {{
        font-size: 0.87rem;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# CARICAMENTO DATI TURISTICI
# ============================================================

SCHEMI = {
    "reg": [
        "year", "month", "nuts2_code", "nuts2_name", "guest_nights_total"
    ],
    "prov": [
        "year", "nuts3_code", "nuts3_name", "nuts2_code", "guest_nights_total"
    ],
    "fr": [
        "year", "month", "nuts2_code",
        "gn_prudente", "gn_intermedio", "gn_espansivo"
    ],
    "fp": [
        "year", "nuts3_code", "nuts2_code",
        "forecast_prudente", "forecast_intermedio", "forecast_espansivo"
    ],
}

CHIAVI = {
    "reg": ["nuts2_code", "year", "month"],
    "prov": ["nuts3_code", "year"],
    "fr": ["nuts2_code", "year", "month"],
    "fp": ["nuts3_code", "year"],
}

DATI = {}
ERRORI = {}

for chiave, colonne in SCHEMI.items():
    try:
        df = carica_csv(FILES[chiave])
        mancanti = [c for c in colonne if c not in df.columns]
        if mancanti:
            raise ValueError("Colonne mancanti: " + ", ".join(mancanti))

        for col in df.columns:
            if col in ("nuts2_code", "nuts3_code"):
                df[col] = df[col].astype("string").str.strip().str.upper()
            elif (
                col in ("year", "month")
                or col.startswith("guest_nights")
                or col.startswith("gn_")
                or col.startswith("forecast_")
            ):
                df[col] = pd.to_numeric(df[col], errors="coerce")
                df[col] = df[col].replace(
                    [float("inf"), -float("inf")], float("nan")
                )

        if df[CHIAVI[chiave]].isna().any().any():
            raise ValueError("Sono presenti chiavi territoriali o temporali vuote.")

        if df.duplicated(CHIAVI[chiave]).any():
            raise ValueError(
                "Chiavi duplicate: correggere il file prima di aggregare."
            )

        if "month" in df and (
            ~df["month"].between(1, 12)
            | df["month"].mod(1).ne(0)
        ).any():
            raise ValueError("Sono presenti mesi non validi.")

        if df["year"].mod(1).ne(0).any():
            raise ValueError("Sono presenti anni non interi.")

        df["year"] = df["year"].astype(int)

        if "month" in df:
            df["month"] = df["month"].astype(int)
            df["data"] = pd.to_datetime(
                dict(year=df["year"], month=df["month"], day=1),
                errors="coerce",
            )
            if df["data"].isna().any():
                raise ValueError("Sono presenti date non valide.")

        DATI[chiave] = df

    except Exception as exc:
        ERRORI[chiave] = str(exc)
        DATI[chiave] = pd.DataFrame()

reg = DATI["reg"]
prov = DATI["prov"]
fr = DATI["fr"]
fp = DATI["fp"]

nomi_reg = (
    reg.drop_duplicates("nuts2_code")
    .set_index("nuts2_code")["nuts2_name"].to_dict()
    if not reg.empty else {}
)

nomi_prov = (
    prov.drop_duplicates("nuts3_code")
    .set_index("nuts3_code")["nuts3_name"].to_dict()
    if not prov.empty else {}
)

with st.sidebar:
    st.divider()
    st.markdown("#### Territorio")

    regione = st.selectbox(
        "Regione",
        ["Tutte"] + sorted(nomi_reg, key=lambda k: str(nomi_reg[k])),
        format_func=lambda k: (
            "Tutta Italia" if k == "Tutte" else str(nomi_reg[k])
        ),
    )

    disponibili_prov = prov.copy()
    if not prov.empty and regione != "Tutte":
        disponibili_prov = prov[prov["nuts2_code"].eq(regione)]

    codici_prov = (
        sorted(
            disponibili_prov["nuts3_code"].unique(),
            key=lambda k: str(nomi_prov.get(k, k)),
        )
        if not disponibili_prov.empty else []
    )

    provincia = st.selectbox(
        "Provincia",
        ["Tutte"] + codici_prov,
        format_func=lambda k: (
            "Tutte le province" if k == "Tutte" else str(nomi_prov.get(k, k))
        ),
    )

    st.caption(
        "Il simulatore usa regione e mese. "
        "Il filtro provinciale riguarda le serie annuali."
    )


def richiedi(chiave):
    if DATI[chiave].empty:
        st.error(
            f"{FILES[chiave].name}: "
            f"{ERRORI.get(chiave, 'nessun dato disponibile')}"
        )
        st.stop()


def filtra_territorio(df):
    risultato = df.copy()
    if regione != "Tutte" and "nuts2_code" in risultato:
        risultato = risultato[risultato["nuts2_code"].eq(regione)]
    if provincia != "Tutte" and "nuts3_code" in risultato:
        risultato = risultato[risultato["nuts3_code"].eq(provincia)]
    return risultato


# ============================================================
# KEPLER.GL — CONFINI CLASSICI
# ============================================================

@st.cache_data(show_spinner=False, ttl=86400)
def confini_nuts(livello, versione):
    nome = f"NUTS_RG_20M_{versione}_4326_LEVL_{livello}.geojson"
    locale = file_dati(nome)

    if locale.exists():
        contenuto = json.loads(locale.read_text(encoding="utf-8"))
    else:
        url = (
            "https://gisco-services.ec.europa.eu/"
            f"distribution/v2/nuts/geojson/{nome}"
        )
        richiesta = Request(url, headers={"User-Agent": "AirbnbSports/4.1"})
        with urlopen(richiesta, timeout=30) as risposta:
            contenuto = json.load(risposta)

    return [
        f for f in contenuto["features"]
        if str(f["properties"].get("NUTS_ID", "")).startswith("IT")
    ]


def coordinate(oggetto):
    if (
        isinstance(oggetto, (list, tuple))
        and len(oggetto) >= 2
        and isinstance(oggetto[0], (int, float))
    ):
        yield oggetto[0], oggetto[1]
    elif isinstance(oggetto, (list, tuple)):
        for elemento in oggetto:
            yield from coordinate(elemento)


def mappa_kepler(df, livello, codice, campo, titolo):
    st.subheader(titolo)

    if df.empty:
        st.info("Nessun dato geografico per questa selezione.")
        return

    if KeplerGl is None:
        st.error(
            "Kepler.gl non è installato. "
            "Aggiungi keplergl al file requirements.txt."
        )
        return

    c1, c2 = st.columns([2, 1])
    with c1:
        versione = st.selectbox(
            "Versione dei confini NUTS",
            [2024, 2021, 2016],
            key=f"confini_{titolo}",
        )
    with c2:
        strumenti = st.checkbox(
            "Strumenti mappa", value=False, key=f"strumenti_{titolo}"
        )

    st.caption(
        "Confini NUTS: scegli la versione coerente con i codici dei dati. "
        "Altezza della mappa: 430 px."
    )

    try:
        valori = df.dropna(subset=[codice, campo]).copy()
        if valori.duplicated(codice).any():
            raise ValueError("Più valori per lo stesso territorio.")

        lookup = valori.set_index(codice)[campo].to_dict()
        nomi = nomi_reg if livello == 2 else nomi_prov
        features = []
        punti = []
        abbinati = set()

        for feature in confini_nuts(livello, versione):
            nuts = feature["properties"]["NUTS_ID"]
            if nuts not in lookup:
                continue

            abbinati.add(nuts)
            features.append({
                "type": "Feature",
                "geometry": feature["geometry"],
                "properties": {
                    "Territorio": str(nomi.get(nuts, nuts)),
                    "NUTS": nuts,
                    "Valore": float(lookup[nuts]),
                    "Pernottamenti": numero(lookup[nuts]),
                },
            })
            punti.extend(coordinate(feature["geometry"]["coordinates"]))

        non_abbinati = set(lookup) - abbinati
        if non_abbinati:
            st.warning(
                "Codici senza confine corrispondente: "
                + ", ".join(sorted(non_abbinati))
            )

        if not features:
            raise ValueError("Nessun territorio abbinato ai confini selezionati.")

        xs, ys = zip(*punti)
        ampiezza = max(max(xs) - min(xs), 1.4 * (max(ys) - min(ys)), 0.1)
        zoom = max(3.0, min(9.0, math.log2(360 / ampiezza) - 0.9))
        data_id = "territori"

        # Basemap chiara/scura senza token Mapbox.
        stile_url = (
            "https://basemaps.cartocdn.com/gl/dark-matter-gl-style/style.json"
            if dark else
            "https://basemaps.cartocdn.com/gl/positron-gl-style/style.json"
        )

        config = {
            "version": "v1",
            "config": {
                "mapState": {
                    "latitude": (min(ys) + max(ys)) / 2,
                    "longitude": (min(xs) + max(xs)) / 2,
                    "zoom": zoom,
                    "pitch": 0,
                    "bearing": 0,
                    "dragRotate": False,
                },
                "mapStyle": {
                    "styleType": "sbl_base",
                    "mapStyles": {
                        "sbl_base": {
                            "id": "sbl_base",
                            "label": "SBL · Base cartografica",
                            "url": stile_url,
                        }
                    },
                },
                "visState": {
                    "filters": [],
                    "layers": [{
                        "id": "confini_nuts",
                        "type": "geojson",
                        "config": {
                            "dataId": data_id,
                            "label": f"Confini NUTS{livello}",
                            "color": [18, 105, 130],
                            "columns": {"geojson": "_geojson"},
                            "isVisible": True,
                            "visConfig": {
                                "opacity": 0.70,
                                "strokeOpacity": 1,
                                "thickness": 0.9,
                                "strokeColor": [8, 43, 70],
                                "stroked": True,
                                "filled": True,
                                "enable3d": False,
                                "wireframe": False,
                                "colorRange": {
                                    "name": "SBL Blu",
                                    "type": "sequential",
                                    "category": "Custom",
                                    "colors": [
                                        "#C2F0F6",
                                        "#88D6E7",
                                        "#4BAFC9",
                                        "#126982",
                                        "#084A68",
                                    ],
                                },
                            },
                        },
                        "visualChannels": {
                            "colorField": {"name": "Valore", "type": "real"},
                            "colorScale": "quantile",
                            "strokeColorField": None,
                            "strokeColorScale": "quantile",
                            "sizeField": None,
                            "sizeScale": "linear",
                            "heightField": None,
                            "heightScale": "linear",
                        },
                    }],
                    "interactionConfig": {
                        "tooltip": {
                            "enabled": True,
                            "fieldsToShow": {
                                data_id: [
                                    {"name": "Territorio", "format": None},
                                    {"name": "NUTS", "format": None},
                                    {"name": "Pernottamenti", "format": None},
                                ]
                            },
                        }
                    },
                },
            },
        }

        mappa = KeplerGl(
            height=430,
            show_docs=False,
            data={
                data_id: {
                    "type": "FeatureCollection",
                    "features": features,
                }
            },
            config=config,
        )

        with tempfile.TemporaryDirectory() as cartella:
            percorso = Path(cartella) / "kepler.html"
            mappa.save_to_html(
                file_name=str(percorso),
                read_only=not strumenti,
            )
            pagina_html = percorso.read_text(encoding="utf-8")

        components.html(pagina_html, height=440, scrolling=False)
        st.caption(
            "Confini: Eurostat/GISCO. Base: CARTO · © OpenStreetMap contributors. "
            "La visualizzazione richiede WebGL e accesso ai servizi cartografici."
        )

    except Exception as exc:
        st.warning(f"Mappa non disponibile: {exc}")

    with st.expander("Valori della mappa"):
        tabella(df, f"mappa_nuts{livello}.csv")


# ============================================================
# EVENT MASTER: LETTURA SENZA ESCLUSIONI SILENZIOSE
# ============================================================

ALIAS = {
    "Evento": ["Evento originale", "Evento", "event_name", "nome_evento"],
    "ID": ["event_id", "id_evento", "ID evento"],
    "Sport": ["Sport", "disciplina"],
    "Inizio": ["Data inizio", "data_inizio", "start_date", "inizio"],
    "Fine": ["Data fine", "data_fine", "end_date", "fine"],
    "Territorio": ["Codici NUTS2", "nuts2_code", "NUTS2", "codice_nuts2"],
    "Persone": [
        "Visitatori unici",
        "visitatori_unici",
        "Visitatori unici stimati",
        "visitatori_unici_stimati",
        "unique_visitors",
    ],
    "Attendance": [
        "Presenze stimate", "Presenze", "Presenze totali",
        "presenze_totali", "attendance", "valore_presenze",
    ],
    "Tipo": [
        "Tipo indicatore", "tipo_indicatore", "tipo_base",
        "tipo_indicatore_presenze", "attendance_type",
    ],
    "Equivalenti": [
        "base_equivalente", "presenze_equivalenti",
        "visitatori_equivalenti", "equivalenti",
    ],
}

NATURE = [
    "Da verificare",
    "Persone",
    "Presenze cumulative",
    "Equivalenti",
    "Biglietti",
]


def mappatura_master(master):
    mappa = {}
    with st.expander("Collegamento delle colonne del master", expanded=False):
        st.caption(
            "Il riconoscimento è automatico. "
            "Puoi correggere qui il collegamento di una colonna."
        )
        colonne_ui = st.columns(2)
        for i, (nome, alias) in enumerate(ALIAS.items()):
            automatico = trova_colonna(master, *alias)
            opzioni = ["— Non disponibile —"] + list(master.columns)
            with colonne_ui[i % 2]:
                scelta = st.selectbox(
                    nome,
                    opzioni,
                    index=opzioni.index(automatico) if automatico else 0,
                    key=f"colonna_master_{nome}",
                )
            mappa[nome] = None if scelta == opzioni[0] else scelta
    return mappa


def codici_nuts(valore):
    s = testo(valore).upper()
    return list(dict.fromkeys(re.findall(r"\bIT[A-Z0-9]{2}\b", s)))


def natura_indicatore(valore):
    s = normalizza(valore)

    if any(x in s for x in ("bigliett", "ticket")):
        return "Biglietti"
    if "equivalent" in s:
        return "Equivalenti"
    if any(x in s for x in ("cumul", "ingress", "spettatorigiorno")):
        return "Presenze cumulative"
    if any(x in s for x in ("unic", "person", "partecipant", "visitator")):
        return "Persone"

    # "Presenze" da solo è ambiguo: non viene convertito automaticamente.
    return "Da verificare"


def prepara_eventi(master, mappa):
    righe = []

    for posizione, (_, originale) in enumerate(master.iterrows()):
        def valore(campo):
            col = mappa.get(campo)
            return originale[col] if col else None

        nome = testo(valore("Evento")) or f"Riga master {posizione + 1}"
        sport = testo(valore("Sport")) or "Sport non indicato"
        inizio = data_singola(valore("Inizio"))
        fine = data_singola(valore("Fine"))

        # Data fine assente: convenzione esplicita di evento di un giorno.
        fine_assente = pd.isna(fine)
        if fine_assente and pd.notna(inizio):
            fine = inizio

        persone = numero_dato(valore("Persone"))
        equivalenti = numero_dato(valore("Equivalenti"))
        attendance = numero_dato(valore("Attendance"))
        natura = natura_indicatore(valore("Tipo"))

        if pd.notna(persone) and persone >= 0:
            volume = persone
            natura = "Persone"
        elif pd.notna(equivalenti) and equivalenti >= 0:
            volume = equivalenti
            natura = "Equivalenti"
        else:
            volume = attendance

        identificativo = testo(valore("ID"))
        if not identificativo:
            identificativo = (
                f"{normalizza(nome)}|{inizio}|{fine}|{normalizza(sport)}"
            )

        territori = codici_nuts(valore("Territorio"))
        if not territori:
            territori = [testo(valore("Territorio"))]

        # Una riga multiregionale viene resa esplicita.
        # Il riparto deve essere confermato prima di usarla nel calcolo.
        for nuts in territori:
            righe.append({
                "Riga master": posizione + 1,
                "ID evento": identificativo,
                "Evento": nome,
                "Sport": sport,
                "Inizio": inizio,
                "Fine": fine,
                "NUTS2": nuts,
                "Volume": volume,
                "Natura": natura,
                "Indicatore originale": testo(valore("Tipo")),
                "Attendance originale": testo(valore("Attendance")),
                "Fine assente nel master": fine_assente,
            })

    if not righe:
        return pd.DataFrame()

    eventi = pd.DataFrame(righe)
    dimensione = eventi.groupby("ID evento")["ID evento"].transform("size")
    eventi["Quota territoriale"] = 1.0 / dimensione
    eventi["Riparto confermato"] = dimensione.eq(1)

    return eventi


def verifica_eventi(eventi):
    risultato = eventi.copy()
    risultato["Inizio"] = risultato["Inizio"].map(data_singola)
    risultato["Fine"] = risultato["Fine"].map(data_singola)
    risultato["Volume"] = pd.to_numeric(risultato["Volume"], errors="coerce")
    risultato["Quota territoriale"] = pd.to_numeric(
        risultato["Quota territoriale"], errors="coerce"
    )

    somme = risultato.groupby("ID evento")["Quota territoriale"].agg(
        somma_completa
    )
    conferme = risultato.groupby("ID evento")["Riparto confermato"].all()

    basi, categorie, motivi = [], [], []

    for _, r in risultato.iterrows():
        problemi = []
        base = float("nan")
        categoria = ""

        if normalizza(r["Sport"]) == "scidifondo":
            problemi.append("Sci di fondo fuori dal perimetro analitico")

        if pd.isna(r["Inizio"]) or pd.isna(r["Fine"]):
            problemi.append("Data da verificare")
            durata = None
        elif r["Fine"] < r["Inizio"]:
            problemi.append("Data fine precedente all'inizio")
            durata = None
        else:
            durata = (r["Fine"] - r["Inizio"]).days + 1

        if r["NUTS2"] not in nomi_reg:
            problemi.append("Codice NUTS2 non abbinato ai dati turistici")

        if (
            pd.isna(r["Volume"])
            or not math.isfinite(float(r["Volume"]))
            or r["Volume"] < 0
        ):
            problemi.append("Volume mancante o non valido")

        if r["Natura"] == "Persone":
            base = r["Volume"]
            categoria = "PERSONE"
        elif r["Natura"] == "Equivalenti":
            base = r["Volume"]
            categoria = "EQUIVALENTI"
        elif r["Natura"] == "Presenze cumulative":
            if durata:
                base = r["Volume"] / durata
            categoria = "EQUIVALENTI"
        elif r["Natura"] == "Biglietti":
            problemi.append("Biglietti: manca una conversione documentata")
        else:
            problemi.append("Natura dell'indicatore da verificare")

        quota = r["Quota territoriale"]
        if pd.isna(quota) or not 0 <= quota <= 1:
            problemi.append("Quota territoriale non valida")

        totale = somme.get(r["ID evento"], float("nan"))
        if pd.isna(totale) or abs(totale - 1) > 0.00001:
            problemi.append("Le quote dello stesso evento devono sommare a 1")

        if not conferme.get(r["ID evento"], False):
            problemi.append("Riparto o righe ripetute da confermare")

        basi.append(base)
        categorie.append(categoria)
        motivi.append("; ".join(problemi))

    risultato["Base"] = basi
    risultato["Categoria"] = categorie
    risultato["Verifica"] = motivi
    risultato["Calcolabile"] = risultato["Verifica"].eq("")
    return risultato


def filtro_compatto(etichetta, opzioni, chiave, format_func=None):
    """Tutti inclusi senza visualizzare centinaia di etichette."""
    tutti = st.checkbox(
        f"{etichetta}: tutti",
        value=True,
        key=f"{chiave}_tutti",
    )
    if tutti:
        return list(opzioni)

    kwargs = {}
    if format_func is not None:
        kwargs["format_func"] = format_func

    return st.multiselect(
        etichetta,
        list(opzioni),
        default=[],
        key=chiave,
        **kwargs,
    )


def espandi_mesi(eventi):
    righe = []

    for _, r in eventi[eventi["Calcolabile"]].iterrows():
        durata = (r["Fine"] - r["Inizio"]).days + 1

        for mese in pd.period_range(r["Inizio"], r["Fine"], freq="M"):
            a = max(r["Inizio"], mese.start_time.normalize())
            b = min(r["Fine"], mese.end_time.normalize())
            giorni = (b - a).days + 1

            righe.append({
                "ID evento": r["ID evento"],
                "Evento": r["Evento"],
                "Sport": r["Sport"],
                "nuts2_code": r["NUTS2"],
                "Regione": nomi_reg.get(r["NUTS2"], r["NUTS2"]),
                "year": mese.year,
                "month": mese.month,
                "data": mese.start_time,
                "Durata": durata,
                "Giorni nel mese": giorni,
                "Peso temporale": giorni / durata,
                "Quota territoriale": r["Quota territoriale"],
                "Base": r["Base"],
                "Categoria": r["Categoria"],
            })

    if not righe:
        return pd.DataFrame()

    risultato = pd.DataFrame(righe)

    if reg.empty:
        risultato["guest_nights_total"] = float("nan")
        return risultato

    return risultato.merge(
        reg[["nuts2_code", "year", "month", "guest_nights_total"]],
        on=["nuts2_code", "year", "month"],
        how="left",
        validate="many_to_one",
    )


def simula(mesi, parametri, quota_uniforme=None):
    risultato = mesi.copy()

    risultato["Quota fuori territorio %"] = risultato["Sport"].map(
        lambda s: parametri[s]["fuori"]
    )
    risultato["Quota pernottanti %"] = (
        quota_uniforme
        if quota_uniforme is not None
        else risultato["Sport"].map(lambda s: parametri[s]["quota"])
    )
    richieste = risultato["Sport"].map(lambda s: parametri[s]["notti"])
    risultato["Notti applicate"] = pd.concat(
        [richieste.rename("richieste"), risultato["Durata"]], axis=1
    ).min(axis=1).clip(lower=1, upper=5)

    risultato["Pernottamenti simulati"] = (
        risultato["Base"]
        * risultato["Quota territoriale"]
        * risultato["Quota fuori territorio %"] / 100
        * risultato["Quota pernottanti %"] / 100
        * risultato["Notti applicate"]
        * risultato["Peso temporale"]
    )

    denominatore = risultato["guest_nights_total"].where(
        risultato["guest_nights_total"] > 0
    )

    risultato["Rapporto con domanda osservata %"] = (
        100 * risultato["Pernottamenti simulati"] / denominatore
    )
    return risultato


def riepilogo_simulazione(df):
    risultato = (
        df.groupby(
            ["nuts2_code", "Regione", "year", "month", "data", "Categoria"],
            dropna=False,
        )
        .agg(
            Pernottamenti_simulati=("Pernottamenti simulati", "sum"),
            Domanda_osservata=("guest_nights_total", "first"),
            Eventi=("ID evento", "nunique"),
        )
        .reset_index()
    )

    risultato["Rapporto %"] = (
        100 * risultato["Pernottamenti_simulati"]
        / risultato["Domanda_osservata"].where(
            risultato["Domanda_osservata"] > 0
        )
    )
    risultato["Soglia esplorativa 2%"] = risultato["Rapporto %"].ge(2)
    return risultato


# ============================================================
# PANORAMA NAZIONALE
# ============================================================

st.title(pagina)

if pagina == "Panorama nazionale":
    richiedi("reg")
    st.caption("Quadro nazionale. I filtri territoriali si applicano alle altre sezioni.")

    annuale = aggrega(reg, ["year"], "guest_nights_total")
    annuale = annuale.sort_values("year")
    ultimo = int(annuale["year"].max())
    corrente = annuale.loc[
        annuale["year"].eq(ultimo), "guest_nights_total"
    ].iloc[0]

    mesi_ultimo = set(reg.loc[reg["year"].eq(ultimo), "month"])
    confrontabile = reg[reg["month"].isin(mesi_ultimo)]
    precedente = somma_completa(
        confrontabile.loc[
            confrontabile["year"].eq(ultimo - 1), "guest_nights_total"
        ]
    )

    crescita = (
        100 * (corrente / precedente - 1)
        if pd.notna(precedente) and precedente > 0 else float("nan")
    )

    a, b, c = st.columns(3)
    a.metric(f"Pernottamenti {ultimo}", numero(corrente))
    b.metric("Variazione sugli stessi mesi", percentuale(crescita))
    c.metric("Regioni / territori NUTS2", numero(reg["nuts2_code"].nunique()))

    st.caption(
        f"Anno {ultimo}: {len(mesi_ultimo)} mesi presenti. "
        "La copertura territoriale e i valori mancanti vanno verificati in Dati e metodo."
    )

    fig = px.bar(
        annuale,
        x="year",
        y="guest_nights_total",
        labels={"year": "Anno", "guest_nights_total": "Pernottamenti"},
        color_discrete_sequence=["#126982"],
    )
    grafico(fig)

    anno = st.selectbox(
        "Anno della mappa",
        sorted(reg["year"].unique(), reverse=True),
    )
    dati_mappa = aggrega(
        reg[reg["year"].eq(anno)], ["nuts2_code"], "guest_nights_total"
    )
    mappa_kepler(
        dati_mappa, 2, "nuts2_code", "guest_nights_total",
        f"Pernottamenti per regione · {anno}",
    )


# ============================================================
# TERRITORIO
# ============================================================

elif pagina == "Esplora un territorio":
    if provincia != "Tutte":
        richiedi("prov")
        dati = filtra_territorio(prov)
        serie = aggrega(dati, ["year"], "guest_nights_total")
        st.subheader(str(nomi_prov.get(provincia, provincia)))
        st.caption("Serie provinciale annuale. Nessuna mensilizzazione artificiale.")
        fig = px.line(
            serie.sort_values("year"),
            x="year", y="guest_nights_total", markers=True,
            labels={"year": "Anno", "guest_nights_total": "Pernottamenti"},
            color_discrete_sequence=["#126982"],
        )
        grafico(fig)
        tabella(dati, "territorio_provincia.csv")

    else:
        richiedi("reg")
        dati = filtra_territorio(reg)
        st.subheader(
            "Italia" if regione == "Tutte" else str(nomi_reg[regione])
        )
        serie = aggrega(dati, ["data"], "guest_nights_total")
        fig = px.line(
            serie.sort_values("data"),
            x="data", y="guest_nights_total",
            labels={"data": "Mese", "guest_nights_total": "Pernottamenti"},
            color_discrete_sequence=["#126982"],
        )
        grafico(fig)

        stagioni = aggrega(
            dati, ["year", "month"], "guest_nights_total"
        )
        stagioni["Anno"] = stagioni["year"].astype(str)
        fig = px.line(
            stagioni.sort_values(["year", "month"]),
            x="month", y="guest_nights_total", color="Anno",
            labels={"month": "Mese", "guest_nights_total": "Pernottamenti"},
        )
        st.subheader("Stagionalità")
        grafico(fig)
        tabella(dati, "territorio_regione.csv")


# ============================================================
# EVENTI E SIMULATORE
# ============================================================

elif pagina == "Eventi e simulatore":
    percorso = FILES["eventi"]

    if not percorso.exists():
        candidati = [
            p for cartella in (ROOT, ROOT / "output_eurostat")
            if cartella.exists()
            for p in cartella.glob("*.xlsx")
            if "eventmaster" in normalizza(p.stem)
        ]
        if len(candidati) == 1:
            percorso = candidati[0]
        else:
            st.error("File Excel dell'Event Master non trovato.")
            st.stop()

    try:
        fogli = fogli_excel(str(percorso), percorso.stat().st_mtime)
        foglio = st.selectbox("Foglio Event Master", fogli)
        master = leggi_excel(
            str(percorso), percorso.stat().st_mtime, foglio
        )
    except Exception as exc:
        st.error(f"Impossibile leggere l'Event Master: {exc}")
        st.stop()

    if master.empty:
        st.info("Il foglio selezionato non contiene righe.")
        st.stop()

    mappa = mappatura_master(master)
    standard = prepara_eventi(master, mappa)

    catalogo_tab, verifica_tab, simulatore_tab, confronto_tab, salvati_tab = st.tabs([
        "Catalogo completo",
        "Verifica dati",
        "Simulatore",
        "Contesto turistico",
        "Elaborazioni salvate",
    ])

    with catalogo_tab:
        st.caption(
            "Qui sono conservate tutte le righe del foglio Excel selezionato, "
            "comprese quelle non ancora utilizzabili nel simulatore."
        )

        tutte_regioni = st.checkbox(
            "Mostra il master di tutte le regioni",
            value=True,
            key="catalogo_tutte_regioni",
        )

        a, b = st.columns(2)
        ricerca = a.text_input("Cerca nel master")
        sport_catalogo = b.selectbox(
            "Sport nel catalogo",
            ["Tutti"] + sorted(standard["Sport"].unique()),
        )

        selezionate = standard.copy()

        if not tutte_regioni and regione != "Tutte":
            selezionate = selezionate[selezionate["NUTS2"].eq(regione)]

        if sport_catalogo != "Tutti":
            selezionate = selezionate[
                selezionate["Sport"].eq(sport_catalogo)
            ]

        righe_master = sorted(selezionate["Riga master"].unique())
        visibile = master.iloc[[int(i) - 1 for i in righe_master]].copy()

        if ricerca.strip():
            maschera = visibile.astype(str).apply(
                lambda col: col.str.contains(
                    ricerca.strip(), case=False, regex=False, na=False
                )
            ).any(axis=1)
            visibile = visibile.loc[maschera]

        a, b, c = st.columns(3)
        a.metric("Righe nel foglio", numero(len(master)))
        b.metric("Righe mostrate", numero(len(visibile)))
        c.metric("Sport nel foglio", numero(standard["Sport"].nunique()))

        tabella(visibile, "event_master_catalogo.csv", altezza=380)

    with verifica_tab:
        st.write(
            "Le righe non riconosciute restano nel catalogo. "
            "Questa tabella permette di correggere date, territorio, "
            "volume e natura dell'indicatore per il calcolo."
        )
        st.caption(
            "Le modifiche valgono per la sessione e non riscrivono l'Excel. "
            "Scarica la tabella verificata per conservarle. "
            "Una data finale assente è inizialmente trattata come evento di un giorno."
        )

        ripetuti = standard["ID evento"].duplicated(keep=False)
        if ripetuti.any():
            st.info(
                "Alcuni eventi compaiono su più righe o territori. "
                "Le quote proposte sono uniformi, ma non vengono applicate "
                "finché il riparto non viene confermato. "
                "Controlla anche se il volume è già territoriale."
            )

        conferma_uniforme = st.checkbox(
            "Confermo che, per gli eventi ripetuti, il volume è il totale "
            "dell'evento ripetuto sulle righe e accetto il riparto uniforme",
            value=False,
        )

        base_editor = standard.copy()
        if conferma_uniforme:
            base_editor.loc[ripetuti, "Riparto confermato"] = True

        modificabili = {
            "Inizio", "Fine", "NUTS2", "Volume",
            "Natura", "Quota territoriale", "Riparto confermato"
        }

        modificati = st.data_editor(
            base_editor,
            hide_index=True,
            use_container_width=True,
            height=350,
            disabled=[
                c for c in base_editor.columns if c not in modificabili
            ],
            column_config={
                "Inizio": st.column_config.DateColumn(
                    "Inizio", format="DD/MM/YYYY"
                ),
                "Fine": st.column_config.DateColumn(
                    "Fine", format="DD/MM/YYYY"
                ),
                "Natura": st.column_config.SelectboxColumn(
                    "Natura", options=NATURE, required=True
                ),
                "Volume": st.column_config.NumberColumn(
                    "Volume", min_value=0.0
                ),
                "Quota territoriale": st.column_config.NumberColumn(
                    "Quota territoriale", min_value=0.0,
                    max_value=1.0, step=0.01, format="%.4f"
                ),
                "Riparto confermato": st.column_config.CheckboxColumn(
                    "Riparto confermato"
                ),
            },
            key=f"editor_eventi_{foglio}_{tuple(mappa.values())}_{conferma_uniforme}",
        )

        eventi = verifica_eventi(modificati)
        calcolabili = eventi[eventi["Calcolabile"]]
        sospesi = eventi[~eventi["Calcolabile"]]

        a, b, c = st.columns(3)
        a.metric("Righe territoriali", numero(len(eventi)))
        b.metric("Calcolabili", numero(len(calcolabili)))
        c.metric("Da verificare / fuori perimetro", numero(len(sospesi)))

        if not sospesi.empty:
            st.subheader("Motivo per ogni riga")
            tabella(
                sospesi[
                    [
                        "Riga master", "Evento", "Sport", "Inizio",
                        "Fine", "NUTS2", "Volume", "Natura", "Verifica"
                    ]
                ],
                "eventi_da_verificare.csv",
            )

        with st.expander("Scarica la base verificata"):
            tabella(eventi, "eventi_base_verificata.csv")

    with simulatore_tab:
        st.write(
            "Simula quanta domanda di pernottamento potrebbe essere associata "
            "agli eventi. I risultati sono scenari, non impatti causali misurati."
        )

        # I filtri mostrano tutti gli sport del master:
        # non vengono costruiti solo sulle righe già calcolabili.
        with st.expander("Filtri della simulazione", expanded=False):
            sport_scelti = filtro_compatto(
                "Sport",
                sorted(eventi["Sport"].unique()),
                "sim_sport",
            )

            etichette = eventi.drop_duplicates("ID evento").set_index(
                "ID evento"
            )["Evento"].to_dict()

            id_disponibili = eventi.loc[
                eventi["Sport"].isin(sport_scelti), "ID evento"
            ].drop_duplicates().tolist()

            id_scelti = filtro_compatto(
                "Eventi",
                id_disponibili,
                "sim_eventi",
                format_func=lambda k: f"{etichette.get(k, k)} · {k}",
            )

        selezione = eventi[
            eventi["Sport"].isin(sport_scelti)
            & eventi["ID evento"].isin(id_scelti)
        ].copy()

        if regione != "Tutte":
            selezione = selezione[selezione["NUTS2"].eq(regione)]

        st.caption(
            "Territorio: "
            + ("Italia" if regione == "Tutte" else str(nomi_reg[regione]))
            + ". Gli eventi con dati turistici mancanti mantengono la "
              "simulazione, ma non il rapporto percentuale."
        )

        buone = selezione[selezione["Calcolabile"]]
        problematiche = selezione[~selezione["Calcolabile"]]

        a, b, c = st.columns(3)
        a.metric("Eventi selezionati", numero(selezione["ID evento"].nunique()))
        b.metric("Eventi con righe calcolabili", numero(buone["ID evento"].nunique()))
        c.metric("Righe da verificare", numero(len(problematiche)))

        if not problematiche.empty:
            with st.expander("Perché alcune righe non vengono calcolate"):
                st.dataframe(
                    prepara_tabella(
                        problematiche[["Evento", "Sport", "NUTS2", "Verifica"]]
                    ),
                    hide_index=True,
                    use_container_width=True,
                )

        mesi = espandi_mesi(selezione)

        if mesi.empty:
            st.info(
                "Nessuna riga calcolabile per questa selezione. "
                "Apri Verifica dati: tutti gli eventi restano nel catalogo."
            )
        else:
            with st.expander("Periodo della simulazione", expanded=False):
                anni_scelti = filtro_compatto(
                    "Anni", sorted(mesi["year"].unique()), "sim_anni"
                )
                mesi_scelti = filtro_compatto(
                    "Mesi", list(range(1, 13)), "sim_mesi"
                )

            mesi = mesi[
                mesi["year"].isin(anni_scelti)
                & mesi["month"].isin(mesi_scelti)
            ]

            if mesi.empty:
                st.info("Nessun evento nel periodo selezionato.")
            else:
                sport_attivi = sorted(mesi["Sport"].unique())
                a, b = st.columns([2, 1])
                preset = a.selectbox(
                    "Quota pernottanti: preset comune",
                    [5, 15, 25, 40, 60],
                    index=2,
                    format_func=lambda q: f"{q}%",
                )

                if b.button("Applica agli sport"):
                    for sport in sport_attivi:
                        st.session_state[f"quota_{sport}"] = int(preset)

                st.caption(
                    "La quota 0–100% indica quanti, tra i non residenti "
                    "considerati nello scenario, pernottano. "
                    "0% = nessuno; 100% = tutti. "
                    "La quota di persone provenienti da fuori è un'ipotesi separata."
                )

                parametri = {}

                for sport in sport_attivi:
                    with st.expander(
                        f"{sport} · parametri",
                        expanded=len(sport_attivi) <= 2,
                    ):
                        c1, c2, c3 = st.columns(3)

                        fuori = c1.slider(
                            "Provenienti da fuori territorio (%)",
                            0, 100, 100,
                            key=f"fuori_{sport}",
                            help=(
                                "100% è un'ipotesi di scenario, non un dato osservato. "
                                "Se la base contiene già solo non residenti, lascia 100."
                            ),
                        )
                        quota = c2.slider(
                            "Tra questi: quota pernottanti (%)",
                            0, 100, 25,
                            key=f"quota_{sport}",
                        )
                        notti = c3.slider(
                            "Notti per pernottante",
                            1, 5, 1,
                            key=f"notti_{sport}",
                        )

                        st.caption(
                            "Le notti vengono limitate alla durata del singolo "
                            "evento, fino a un massimo di 5. "
                            "Non sono inclusi soggiorni prima o dopo l'evento."
                        )

                        parametri[sport] = {
                            "fuori": fuori,
                            "quota": quota,
                            "notti": notti,
                        }

                risultato = simula(mesi, parametri)
                mensile = riepilogo_simulazione(risultato)

                a, b = st.columns(2)
                persone = risultato.loc[
                    risultato["Categoria"].eq("PERSONE"),
                    "Pernottamenti simulati",
                ].sum()
                equivalenti = risultato.loc[
                    risultato["Categoria"].eq("EQUIVALENTI"),
                    "Pernottamenti simulati",
                ].sum()

                a.metric("Pernottamenti · base persone", numero(persone))
                b.metric("Pernottamenti · base equivalente", numero(equivalenti))

                st.caption(
                    "Le due basi restano separate. Gli equivalenti non sono "
                    "visitatori unici. La somma tra eventi può includere "
                    "persone o soggiorni presenti in più eventi."
                )

                n_senza = risultato.loc[
                    risultato["guest_nights_total"].isna()
                    | risultato["guest_nights_total"].le(0),
                    ["nuts2_code", "year", "month"],
                ].drop_duplicates()

                if not n_senza.empty:
                    st.info(
                        f"{len(n_senza)} combinazioni regione–mese "
                        "senza denominatore turistico positivo: "
                        "il rapporto percentuale resta non disponibile."
                    )

                andamento = (
                    mensile.groupby(["data", "Categoria"])[
                        "Pernottamenti_simulati"
                    ].sum().reset_index()
                )
                fig = px.bar(
                    andamento,
                    x="data", y="Pernottamenti_simulati",
                    color="Categoria", barmode="group",
                    color_discrete_map={
                        "PERSONE": "#126982",
                        "EQUIVALENTI": "#DC813F",
                    },
                    labels={
                        "data": "Mese",
                        "Pernottamenti_simulati": "Pernottamenti simulati",
                    },
                )
                grafico(fig)

                st.subheader("Confronto con la domanda osservata")
                solo_soglia = st.checkbox(
                    "Mostra solo rapporti almeno pari al 2%",
                    value=False,
                )
                vista = (
                    mensile[mensile["Soglia esplorativa 2%"]]
                    if solo_soglia else mensile
                )
                tabella(vista, "simulazione_regioni_mesi.csv")

                with st.expander("Confronto dei cinque scenari"):
                    confronti = []
                    for q in [5, 15, 25, 40, 60]:
                        scenario = simula(mesi, parametri, quota_uniforme=q)
                        totali = scenario.groupby("Categoria")[
                            "Pernottamenti simulati"
                        ].sum()
                        for categoria, totale in totali.items():
                            confronti.append({
                                "Quota pernottanti": f"{q}%",
                                "Categoria": categoria,
                                "Pernottamenti simulati": totale,
                            })

                    scenari_df = pd.DataFrame(confronti)
                    st.caption(
                        "Cambia solo la quota pernottanti. "
                        "Origine, notti e selezione degli eventi restano invariati."
                    )
                    fig = px.bar(
                        scenari_df,
                        x="Quota pernottanti",
                        y="Pernottamenti simulati",
                        color="Categoria",
                        barmode="group",
                        color_discrete_map={
                            "PERSONE": "#126982",
                            "EQUIVALENTI": "#DC813F",
                        },
                    )
                    grafico(fig)
                    tabella(scenari_df, "simulazione_cinque_scenari.csv")

                with st.expander("Dettaglio evento–regione–mese"):
                    tabella(risultato, "simulazione_dettaglio.csv")

                with st.expander("Come leggere il calcolo"):
                    st.markdown(
                        """
                        **Pernottamenti simulati = base × quota territoriale
                        × quota da fuori × quota pernottanti × notti
                        × peso temporale.**

                        Le percentuali entrano nella formula divise per 100.

                        Il peso temporale è il rapporto tra i giorni dell'evento
                        nel mese e la sua durata complessiva. È un'ipotesi di
                        ripartizione uniforme, non una distribuzione osservata.

                        Il rapporto con Eurostat misura una **dimensione relativa**.
                        Non dimostra che quella quota di turismo sia stata causata
                        dall'evento. Il 2% è una soglia esplorativa.
                        """
                    )

    with confronto_tab:
        if reg.empty:
            st.info("Dati regionali non disponibili.")
        else:
            codici = sorted(nomi_reg, key=lambda k: str(nomi_reg[k]))
            indice = codici.index(regione) if regione in codici else 0
            scelta = st.selectbox(
                "Regione del confronto",
                codici,
                index=indice,
                format_func=lambda k: str(nomi_reg[k]),
            )

            serie = reg[reg["nuts2_code"].eq(scelta)].copy()
            serie = serie.sort_values("data")
            precedenti = serie[
                ["year", "month", "guest_nights_total"]
            ].copy()
            precedenti["year"] += 1
            precedenti = precedenti.rename(
                columns={"guest_nights_total": "Anno precedente"}
            )
            serie = serie.merge(
                precedenti, on=["year", "month"], how="left",
                validate="one_to_one",
            )
            serie["Variazione stesso mese %"] = (
                100 * (
                    serie["guest_nights_total"]
                    / serie["Anno precedente"].where(
                        serie["Anno precedente"] > 0
                    ) - 1
                )
            )

            fig = px.line(
                serie,
                x="data", y="guest_nights_total",
                color_discrete_sequence=["#126982"],
                labels={"data": "Mese", "guest_nights_total": "Pernottamenti"},
            )
            grafico(fig)

            st.caption(
                "Confronto descrittivo con lo stesso mese dell'anno precedente. "
                "Stagionalità, festività, prezzi e altri eventi possono influire. "
                "Questo grafico non identifica un effetto causale."
            )
            tabella(serie, "contesto_turistico_regionale.csv")

    with salvati_tab:
        disponibili = [n for n in SCREENING if file_dati(n).exists()]

        if not disponibili:
            st.info("Nessuna elaborazione salvata trovata.")
        else:
            scelta = st.selectbox("Elaborazione", disponibili)
            st.caption(
                "Output già presente nel repository. "
                "Non viene modificato dai cursori del simulatore."
            )
            try:
                tabella(carica_csv(file_dati(scelta)), f"copia_{scelta}")
            except Exception as exc:
                st.error(f"Impossibile leggere il file: {exc}")


# ============================================================
# FORECAST
# ============================================================

elif pagina == "Forecast 2026–2030":
    provinciale = provincia != "Tutte"
    chiave = "fp" if provinciale else "fr"
    richiedi(chiave)

    dati = filtra_territorio(DATI[chiave])
    prefisso = "forecast_" if provinciale else "gn_"
    codice = "nuts3_code" if provinciale else "nuts2_code"
    livello = 3 if provinciale else 2

    st.caption(
        "Scenari letti dai CSV del progetto. "
        "I parametri del simulatore eventi non modificano questi forecast."
    )

    scenari_scelti = st.multiselect(
        "Scenari",
        SCENARI,
        default=SCENARI,
    )

    if not scenari_scelti:
        st.info("Seleziona almeno uno scenario.")
    elif dati.empty:
        st.info("Nessun forecast per il territorio selezionato.")
    else:
        righe = []
        for scenario in scenari_scelti:
            campo = prefisso + scenario.lower()
            annuale = aggrega(dati, ["year"], campo)
            annuale = annuale.rename(columns={campo: "Pernottamenti"})
            annuale["Scenario"] = scenario
            righe.append(annuale)

        lungo = pd.concat(righe, ignore_index=True)
        fig = px.line(
            lungo.sort_values("year"),
            x="year", y="Pernottamenti",
            color="Scenario", markers=True,
            color_discrete_map=COLORI,
            labels={"year": "Anno"},
        )
        grafico(fig)

        a, b = st.columns(2)
        anno = a.selectbox(
            "Anno della mappa",
            sorted(dati["year"].unique(), reverse=True),
        )
        scenario = b.selectbox("Scenario della mappa", scenari_scelti)
        campo = prefisso + scenario.lower()

        mappa_df = aggrega(
            dati[dati["year"].eq(anno)], [codice], campo
        )
        mappa_kepler(
            mappa_df, livello, codice, campo,
            f"Forecast {scenario.lower()} · {anno}",
        )

        tabella(lungo, "forecast_riepilogo.csv")

        if provinciale:
            st.caption(
                "Se il forecast provinciale è ottenuto con quote territoriali "
                "costanti, non rappresenta una previsione indipendente "
                "della dinamica di ciascuna provincia."
            )


# ============================================================
# DATI E METODO
# ============================================================

else:
    lettura, fonti, metodo, previsione, qualita = st.tabs([
        "Come leggere",
        "Fonti",
        "Metodo eventi",
        "Forecast",
        "Qualità e limiti",
    ])

    with lettura:
        st.markdown(
            """
            ### Tre livelli distinti

            **Dati osservati.** Descrivono i pernottamenti registrati
            nell'ambito coperto dalla fonte Eurostat.

            **Simulazione degli eventi.** Traduce una base di partecipazione
            in pernottamenti teorici mediante ipotesi modificabili.
            Non misura direttamente i pernottamenti realmente generati.

            **Forecast.** Mostra i tre scenari già elaborati nel progetto.
            I pernottamenti simulati degli eventi non vengono aggiunti
            automaticamente alle previsioni.

            ### Domanda di ricerca

            La dashboard permette di individuare eventi e mesi nei quali
            la domanda teorica potrebbe essere rilevante rispetto
            alla domanda turistica osservata.

            Non consente, da sola, di concludere che un evento abbia causato
            un certo incremento del turismo regionale.
            """
        )

    with fonti:
        st.markdown(
            """
            ### Dati turistici

            **`tour_ce_omn12`** è il riferimento per l'analisi mensile
            regionale NUTS2 del progetto.

            **`tour_ce_oan3`** è il riferimento per le serie annuali NUTS3.

            I pernottamenti degli ospiti sono una misura diversa da arrivi,
            persone uniche o notti di occupazione dell'alloggio:
            tre ospiti per due notti corrispondono a sei pernottamenti.

            La copertura riguarda gli alloggi e le piattaforme comprese
            nella fonte. Non coincide automaticamente con tutto il turismo,
            con il solo mercato Airbnb o con tutta la ricettività regionale.

            ### Event Master

            Il master è la fonte del catalogo degli eventi.
            Il catalogo mostra tutte le righe del foglio selezionato,
            anche quando una riga non è ancora utilizzabile nel calcolo.

            Il numero di righe può essere diverso dal numero di eventi:
            un evento può comparire in più territori o su più righe.

            ### Riferimento

            [Metadati Eurostat sulla domanda tramite piattaforme]
            (https://ec.europa.eu/eurostat/cache/metadata/en/tour_ce_esms.htm)
            """
        )

    with metodo:
        st.markdown(
            """
            ### 1. Individuazione della base

            - **Persone:** visitatori o partecipanti unici documentati
              oppure stimati come tali dalla fonte.
            - **Presenze cumulative:** vengono divise per la durata
              dell'evento per costruire una base equivalente.
            - **Equivalenti:** proxy già disponibile nel dataset.
            - **Biglietti:** non vengono automaticamente trasformati
              in persone o pernottamenti.
            - **Indicatore ambiguo:** resta da verificare.

            Il rapporto tra presenze cumulative e giorni è una media
            giornaliera equivalente. **Non stima le persone uniche.**

            Un campo chiamato soltanto “presenze” non basta a stabilire
            se siano persone, ingressi ripetuti o pernottamenti.

            ### 2. Provenienza e pernottamento

            La quota proveniente da fuori territorio e la quota che pernotta
            sono parametri distinti.

            Il cursore pernottanti va da 0 a 100:
            a zero nessuno dei non residenti considerati pernotta;
            a cento pernottano tutti.

            Se la provenienza non è documentata, il parametro resta
            un'ipotesi di scenario. Il valore iniziale del 100% non significa
            che tutti i partecipanti provengano effettivamente da fuori.

            La provenienza da fuori regione non coincide con la nazionalità:
            un italiano può essere non residente nella regione
            e uno straniero può risiedervi.

            ### 3. Durata del soggiorno

            Le notti applicate sono il minimo tra:
            notti scelte, durata dell'evento e cinque notti.

            Una data finale assente viene inizialmente trattata come
            evento di un giorno, ma è modificabile in Verifica dati.

            Questa convenzione non include eventuali soggiorni
            prima o dopo l'evento.

            ### 4. Territorio e mese

            Per eventi su più territori il totale non deve essere
            duplicato integralmente in ogni regione.

            Il riparto uniforme proposto deve essere confermato.
            Se i volumi sono già regionali, occorre invece impostare
            quote coerenti con la loro definizione: non confermare
            automaticamente il riparto di un totale nazionale.

            Gli eventi a cavallo di più mesi sono distribuiti in proporzione
            ai giorni di esposizione. È una convenzione uniforme,
            non un andamento giornaliero osservato.
            """
        )

        st.latex(
            r"N_{e,r,m}=B_e\cdot w_{e,r}\cdot f_e"
            r"\cdot q_e\cdot n_e\cdot t_{e,m}"
        )
        st.markdown(
            """
            Dove **B** è la base, **w** la quota territoriale,
            **f** la quota da fuori territorio, **q** la quota pernottanti,
            **n** le notti e **t** il peso temporale.

            Le percentuali sono espresse nella formula come frazioni
            comprese tra zero e uno.
            """
        )
        st.latex(
            r"R_{r,m}=100\cdot"
            r"\frac{\text{pernottamenti simulati}_{r,m}}"
            r"{\text{pernottamenti osservati}_{r,m}}"
        )
        st.markdown(
            """
            ### 5. Interpretazione del rapporto

            Il rapporto misura la dimensione dello scenario rispetto
            alla domanda osservata. Non è la quota di turismo
            attribuibile causalmente agli eventi.

            Il numeratore può includere pernottamenti in strutture
            non coperte dal denominatore Eurostat: in questo caso
            il rapporto è un confronto esplorativo tra perimetri diversi.

            La soglia del **2%** serve a selezionare casi da approfondire.
            Non è una soglia di significatività statistica.

            Le basi persone ed equivalenti restano separate.
            Anche all'interno della stessa base, somme tra eventi
            possono includere partecipanti o soggiorni ripetuti.

            Lo sci di fondo resta visibile nel catalogo,
            ma è escluso dal perimetro della simulazione.
            """
        )

    with previsione:
        st.markdown(
            """
            ### Cosa esegue questa dashboard

            L'app legge i CSV di forecast già prodotti.
            Non riaddestra il modello e non ricalcola i parametri
            del procedimento che ha generato quei file.

            Prudente, Intermedio ed Espansivo sono scenari.
            Non sono automaticamente intervalli di confidenza
            o probabilità che il risultato reale cada in un intervallo.

            ### Come documentare il modello a monte

            La documentazione del forecast deve riportare:

            - periodo di stima e trattamento degli anni pandemici;
            - anno base e gestione di eventuali mesi mancanti;
            - formule e parametri di crescita;
            - conservazione o modifica della stagionalità;
            - metodo di riparto territoriale;
            - validazione su periodi non usati nella stima.

            Questi elementi non possono essere verificati
            osservando soltanto i CSV finali.

            ### Validazione consigliata

            Confrontare il modello con almeno una previsione stagionale
            semplice: lo stesso mese dell'anno precedente.

            Usare finestre temporali successive senza impiegare
            informazioni future nella stima. Valutare errore assoluto,
            WAPE e bias sullo stesso campione per tutti i modelli.

            Per un riparto provinciale a quote costanti,
            la somma delle province deve riconciliare con la regione.
            Le quote costanti non catturano cambiamenti
            nella distribuzione territoriale della domanda.
            """
        )

        st.latex(
            r"\mathrm{WAPE}=100\cdot"
            r"\frac{\sum_t|y_t-\hat{y}_t|}{\sum_t|y_t|}"
        )
        st.latex(
            r"\mathrm{Bias}=100\cdot"
            r"\frac{\sum_t(\hat{y}_t-y_t)}{\sum_t y_t}"
        )
        st.caption(
            "Le metriche non sono definite quando il rispettivo "
            "denominatore è zero. Qui non vengono riportati risultati "
            "di validazione non ricalcolati dai dati di test."
        )

    with qualita:
        st.subheader("Controlli dei file caricati")

        controlli = []

        for chiave in SCHEMI:
            df = DATI[chiave]
            if df.empty:
                controlli.append({
                    "File": FILES[chiave].name,
                    "Stato": ERRORI.get(chiave, "Vuoto"),
                    "Righe": 0,
                    "Valori mancanti": None,
                    "Valori negativi": None,
                })
                continue

            numeriche = [
                c for c in SCHEMI[chiave]
                if c.startswith(("guest_nights", "gn_", "forecast_"))
            ]
            controlli.append({
                "File": FILES[chiave].name,
                "Stato": "Letto; chiavi valide e non duplicate",
                "Righe": len(df),
                "Valori mancanti": int(df[numeriche].isna().sum().sum()),
                "Valori negativi": int(df[numeriche].lt(0).sum().sum()),
            })

        tabella(pd.DataFrame(controlli), "controlli_file.csv")

        if not reg.empty:
            copertura = (
                reg.groupby(["nuts2_code", "year"])
                .agg(
                    Mesi_presenti=("month", "nunique"),
                    Mesi_con_valore=("guest_nights_total", "count"),
                )
                .reset_index()
            )
            incompleti = copertura[
                copertura["Mesi_presenti"].ne(12)
                | copertura["Mesi_con_valore"].ne(12)
            ]

            with st.expander("Copertura annuale regionale"):
                st.caption(
                    "Un anno con meno di 12 mesi non è un totale annuale completo. "
                    "L'assenza di un'intera regione deve essere verificata anche "
                    "rispetto al perimetro atteso della fonte."
                )
                tabella(
                    incompleti if not incompleti.empty else copertura,
                    "copertura_regioni_anni.csv",
                )

        if not fr.empty and not fp.empty:
            verifiche = []

            for scenario in SCENARI:
                regionale = aggrega(
                    fr, ["nuts2_code", "year"],
                    "gn_" + scenario.lower(),
                ).rename(
                    columns={"gn_" + scenario.lower(): "Totale regionale"}
                )
                provinciale = aggrega(
                    fp, ["nuts2_code", "year"],
                    "forecast_" + scenario.lower(),
                ).rename(
                    columns={
                        "forecast_" + scenario.lower(): "Somma province"
                    }
                )
                confronto = regionale.merge(
                    provinciale,
                    on=["nuts2_code", "year"],
                    how="outer",
                    validate="one_to_one",
                )
                confronto["Scenario"] = scenario
                confronto["Scarto"] = (
                    confronto["Somma province"]
                    - confronto["Totale regionale"]
                )
                verifiche.append(confronto)

            with st.expander("Riconciliazione forecast regioni–province"):
                tabella(
                    pd.concat(verifiche, ignore_index=True),
                    "riconciliazione_forecast.csv",
                )

        st.markdown(
            """
            ### Limiti dell'interpretazione

            Un aumento dei pernottamenti durante un evento può dipendere
            anche da stagionalità, festività, prezzi, trasporti,
            condizioni economiche o altri eventi.

            L'evento può spostare domanda nel tempo o sostituire altri turisti.
            Alcuni partecipanti avrebbero visitato comunque la regione.

            Per stimare un effetto causale servono un controfattuale credibile,
            territori e periodi comparabili, verifica delle tendenze precedenti
            e controlli per eventi contemporanei.

            Il confronto prima/dopo o con lo stesso mese dell'anno precedente,
            da solo, non soddisfa questi requisiti.
            """
        )


st.divider()
st.caption(
    f"Airbnb&Sports · {VERSIONE} · "
    f"[SBL Consultancy]({SITO_AZIENDA})"
)
