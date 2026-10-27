# AIRBNB&SPORTS
# app.py
#
# Dipendenze in requirements.txt:
# streamlit
# pandas
# plotly
# openpyxl
# keplergl

from pathlib import Path
from urllib.request import Request, urlopen
import base64
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
# 1. CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Airbnb&Sports",
    page_icon="🏟️",
    layout="wide",
)

ROOT = Path(__file__).resolve().parent
VERSIONE = "v4.2 · Basi eventi e proxy"
SITO = "https://www.sblconsultancy.it/"

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


def percorso(nome):
    for cartella in (ROOT, ROOT / "output_eurostat"):
        p = cartella / nome
        if p.exists():
            return p
    return ROOT / nome


FILES = {
    "reg": percorso("eurostat_master_final.csv"),
    "prov": percorso("eurostat_nuts3_annuale_2018_2025.csv"),
    "fr": percorso("forecast_nuts2_2026_2030_3_scenari.csv"),
    "fp": percorso("forecast_nuts3_2026_2030_3_scenari.csv"),
    "eventi": percorso(
        "AirbnbSports_Event_Master_Simone_Visitatori.xlsx"
    ),
}

SCREENING = [
    "screening_eventi_completo.csv",
    "screening_eventi_equivalenti.csv",
    "screening_regioni_mesi_completo.csv",
    "screening_regioni_mesi_equivalenti.csv",
    "screening_pernottamenti_5_scenari_notti.csv",
    "screening_pernottamenti_5_scenari.csv",
    "screening_2pct.csv",
    "event_exposure_monthly.csv",
]


# ============================================================
# 2. FUNZIONI GENERALI
# ============================================================

def norm(valore):
    s = unicodedata.normalize("NFKD", str(valore))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]", "", s.lower())


def testo(valore):
    if valore is None or pd.isna(valore):
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


def pct(valore):
    if pd.isna(valore):
        return "n.d."
    return f"{numero(valore, 1)}%"


def leggi_numero(valore):
    """
    Accetta numeri, separatori italiani/internazionali e un solo
    valore accompagnato da testo.

    Non sceglie arbitrariamente un estremo o il punto medio
    di un intervallo. Il testo originale viene conservato.
    """
    if valore is None or pd.isna(valore):
        return float("nan")

    if isinstance(valore, (int, float)):
        v = float(valore)
        return v if math.isfinite(v) else float("nan")

    s = str(valore).strip().replace("\u00a0", " ")

    # Evita di leggere percentuali come quantità di persone.
    if "%" in s:
        return float("nan")

    # Riunisce separatori delle migliaia espressi con spazi.
    s = re.sub(r"(?<=\d)[ '\u2019](?=\d{3}(?:\D|$))", "", s)

    trovati = re.findall(r"[-+]?\d+(?:[.,]\d+)*", s)
    if len(trovati) != 1:
        return float("nan")

    v = trovati[0]

    # Notazione scientifica testuale non interpretata parzialmente.
    if re.search(r"\d[eE][+-]?\d", s):
        return float("nan")

    # Evita di leggere "1,2 milioni" come 1,2 persone.
    if re.search(r"\b(mila|milion\w*|thousand|million)\b", s.lower()):
        return float("nan")

    if "." in v and "," in v:
        if v.rfind(",") > v.rfind("."):
            v = v.replace(".", "").replace(",", ".")
        else:
            v = v.replace(",", "")
    elif re.fullmatch(r"[-+]?\d{1,3}(?:\.\d{3})+", v):
        v = v.replace(".", "")
    elif re.fullmatch(r"[-+]?\d{1,3}(?:,\d{3})+", v):
        v = v.replace(",", "")
    else:
        v = v.replace(",", ".")

    try:
        risultato = float(v)
        return risultato if math.isfinite(risultato) else float("nan")
    except ValueError:
        return float("nan")


def leggi_data(valore):
    if valore is None or pd.isna(valore):
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


def trova(df, *nomi):
    dizionario = {norm(c): c for c in df.columns}
    for nome in nomi:
        if norm(nome) in dizionario:
            return dizionario[norm(nome)]
    return None


@st.cache_data(show_spinner=False)
def csv_cache(nome, modifica):
    df = pd.read_csv(nome, encoding="utf-8-sig")
    if len(df.columns) == 1 and ";" in str(df.columns[0]):
        df = pd.read_csv(nome, encoding="utf-8-sig", sep=";")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def csv(p):
    return csv_cache(str(p), p.stat().st_mtime)


@st.cache_data(show_spinner=False)
def excel_fogli(nome, modifica):
    with pd.ExcelFile(nome) as libro:
        return libro.sheet_names


@st.cache_data(show_spinner=False)
def excel_cache(nome, modifica, foglio):
    df = pd.read_excel(nome, sheet_name=foglio)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def pulisci_tabella(df):
    risultato = df.copy()

    for col in risultato:
        if pd.api.types.is_datetime64_any_dtype(risultato[col]):
            risultato[col] = risultato[col].dt.strftime("%d/%m/%Y")
        elif risultato[col].dtype == object:
            risultato[col] = risultato[col].map(
                lambda x: pd.NA
                if isinstance(x, str)
                and x.strip().lower() in ("", "nan", "none", "nat", "null")
                else x
            )

            if norm(col) in (
                "datainizio", "datafine", "startdate",
                "enddate", "inizio", "fine", "data"
            ):
                originali = risultato[col]
                date = originali.map(leggi_data)
                risultato[col] = originali.where(
                    date.isna(), date.dt.strftime("%d/%m/%Y")
                )

    return risultato.dropna(axis=1, how="all")


def tabella(df, nome, altezza=330):
    if df.empty:
        st.info("Nessuna riga per questa selezione.")
        return

    mostra = pulisci_tabella(df)

    st.dataframe(
        mostra,
        hide_index=True,
        use_container_width=True,
        height=min(altezza, 40 + len(mostra) * 35),
    )
    st.download_button(
        "Scarica CSV",
        mostra.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=nome,
        mime="text/csv",
        key="download_" + nome,
    )


def somma_valida(s):
    if s.empty or s.isna().any():
        return float("nan")
    return s.sum()


def aggrega(df, gruppi, campo):
    return (
        df.groupby(gruppi, dropna=False)[campo]
        .agg(somma_valida)
        .reset_index()
    )


def grafico(fig, altezza=330):
    fig.update_layout(
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TESTO, size=12),
        margin=dict(l=10, r=10, t=30, b=20),
        legend=dict(orientation="h", y=1.04, x=0),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, automargin=True)
    fig.update_yaxes(gridcolor=BORDO, automargin=True)
    st.plotly_chart(fig, use_container_width=True)


# ============================================================
# 3. LOGO, NAVIGAZIONE E COLORI
# ============================================================

def logo():
    mime = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }

    for cartella in (
        ROOT, ROOT / "assets", ROOT / "images",
        ROOT / "img", ROOT / "output_eurostat"
    ):
        if not cartella.exists():
            continue

        for p in sorted(cartella.iterdir()):
            if (
                p.is_file()
                and p.stem.lower() == "logo_sbl"
                and p.suffix.lower() in mime
            ):
                dati = base64.b64encode(p.read_bytes()).decode()
                st.markdown(
                    f"""
                    <a href="{SITO}" target="_blank"
                       rel="noopener noreferrer">
                      <img
                        src="data:{mime[p.suffix.lower()]};base64,{dati}"
                        alt="SBL Consultancy"
                        style="width:108px;max-width:100%;height:auto;
                               display:block;margin:0 0 10px 0;">
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
                return

    st.markdown(f"[**SBL CONSULTANCY**]({SITO})")


with st.sidebar:
    logo()
    st.markdown("### Airbnb&Sports")
    st.caption("Tourism Demand · Italia")
    pagina = st.radio("Sezione", PAGINE)
    dark = st.checkbox("Tema scuro", value=False)
    st.caption(VERSIONE)

TESTO = "#F3FAFC" if dark else "#082B46"
SECONDARIO = "#D3E5EC" if dark else "#254A5D"
BORDO = "#315462" if dark else "#B5D8E1"
CARD = "#102D3B" if dark else "#FDFEFF"

SFONDO = (
    "linear-gradient(100deg,#071822 0%,#0C2B39 47%,#10475B 100%)"
    if dark else
    "linear-gradient(100deg,#C7F1F6 0%,#9FE5F0 47%,#5BC6E2 100%)"
)

SIDE = (
    "linear-gradient(180deg,#071F31 0%,#0E4058 52%,#126982 100%)"
    if dark else
    "linear-gradient(180deg,#153F57 0%,#17627E 52%,#259BBA 100%)"
)

st.markdown(
    f"""
    <style>
    .stApp {{background:{SFONDO};color:{TESTO};}}
    .block-container {{
        max-width:1280px;
        padding:1.4rem 1.5rem 2rem 1.5rem;
    }}
    [data-testid="stSidebar"] {{
        background:{SIDE};
        min-width:255px;
        max-width:275px;
    }}
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h3 {{
        color:#FFFFFF !important;
    }}
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color:#E2F0F5 !important;
    }}
    h1 {{font-size:1.8rem !important;}}
    h2 {{font-size:1.3rem !important;}}
    h3 {{font-size:1.06rem !important;}}
    p,label {{font-size:0.92rem;}}
    [data-testid="stMain"] p,
    [data-testid="stMain"] label,
    [data-testid="stMain"] h1,
    [data-testid="stMain"] h2,
    [data-testid="stMain"] h3 {{
        color:{TESTO};
    }}
    [data-testid="stCaptionContainer"] p {{
        color:{SECONDARIO} !important;
        font-size:0.82rem !important;
    }}
    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stMultiSelect"] [data-baseweb="select"] > div {{
        background:#F7FAFC !important;
        border-color:#B5D8E1 !important;
        min-height:36px !important;
        border-radius:8px !important;
    }}
    [data-testid="stSelectbox"] [data-baseweb="select"] *,
    [data-testid="stMultiSelect"] [data-baseweb="select"] *,
    [data-baseweb="popover"] *,
    [role="listbox"] *,
    [role="option"] * {{
        color:#082B46 !important;
        -webkit-text-fill-color:#082B46 !important;
        font-size:0.87rem !important;
    }}
    [role="listbox"],[role="option"] {{
        background:#F7FAFC !important;
    }}
    [data-baseweb="tag"] {{
        background:#D8EDF3 !important;
        border-radius:5px !important;
    }}
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {{
        background:#F7FAFC !important;
        color:#082B46 !important;
        -webkit-text-fill-color:#082B46 !important;
    }}
    [data-testid="stMetric"] {{
        background:{CARD};
        border:1px solid {BORDO};
        border-radius:12px;
        padding:11px 13px;
    }}
    [data-testid="stMetricValue"] {{
        color:{TESTO};
        font-size:1.55rem !important;
    }}
    [data-testid="stMetricLabel"] p {{font-size:0.8rem !important;}}
    [data-testid="stVerticalBlock"] {{gap:0.8rem;}}
    .stButton > button,.stDownloadButton > button {{
        background:#F7FAFC;
        color:#082B46 !important;
        border:1px solid #B5D8E1;
        border-radius:8px;
    }}
    .stButton > button p,.stDownloadButton > button p {{
        color:#082B46 !important;
        font-size:0.84rem !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 4. DATI EUROSTAT E FORECAST
# ============================================================

SCHEMI = {
    "reg": [
        "year", "month", "nuts2_code", "nuts2_name", "guest_nights_total"
    ],
    "prov": [
        "year", "nuts3_code", "nuts3_name", "nuts2_code",
        "guest_nights_total"
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

for chiave, richieste in SCHEMI.items():
    try:
        df = csv(FILES[chiave])
        mancanti = [c for c in richieste if c not in df]

        if mancanti:
            raise ValueError("Colonne mancanti: " + ", ".join(mancanti))

        for c in df:
            if c in ("nuts2_code", "nuts3_code"):
                df[c] = df[c].astype("string").str.strip().str.upper()
            elif (
                c in ("year", "month")
                or c.startswith(("guest_nights", "gn_", "forecast_"))
            ):
                df[c] = pd.to_numeric(df[c], errors="coerce")
                df[c] = df[c].replace(
                    [float("inf"), -float("inf")], float("nan")
                )

        if df[CHIAVI[chiave]].isna().any().any():
            raise ValueError("Chiavi territoriali o temporali mancanti.")

        if df.duplicated(CHIAVI[chiave]).any():
            raise ValueError("Chiavi duplicate: aggregazione sospesa.")

        if df["year"].mod(1).ne(0).any():
            raise ValueError("Anno non intero.")

        df["year"] = df["year"].astype(int)

        if "month" in df:
            if (
                ~df["month"].between(1, 12)
                | df["month"].mod(1).ne(0)
            ).any():
                raise ValueError("Mesi non validi.")

            df["month"] = df["month"].astype(int)
            df["data"] = pd.to_datetime(
                dict(year=df["year"], month=df["month"], day=1),
                errors="coerce",
            )

            if df["data"].isna().any():
                raise ValueError("Date non valide.")

        DATI[chiave] = df

    except Exception as exc:
        DATI[chiave] = pd.DataFrame()
        ERRORI[chiave] = str(exc)

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


def richiedi(chiave):
    if DATI[chiave].empty:
        st.error(
            f"{FILES[chiave].name}: "
            f"{ERRORI.get(chiave, 'nessun dato disponibile')}"
        )
        st.stop()


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

    p_disponibili = prov.copy()
    if not prov.empty and regione != "Tutte":
        p_disponibili = prov[prov["nuts2_code"].eq(regione)]

    p_codici = (
        sorted(
            p_disponibili["nuts3_code"].unique(),
            key=lambda k: str(nomi_prov.get(k, k)),
        )
        if not p_disponibili.empty else []
    )

    provincia = st.selectbox(
        "Provincia",
        ["Tutte"] + p_codici,
        format_func=lambda k: (
            "Tutte le province"
            if k == "Tutte" else str(nomi_prov.get(k, k))
        ),
    )

    st.caption(
        "Gli eventi sono confrontati con regione e mese. "
        "Le province hanno dati annuali."
    )


def filtra_territorio(df):
    d = df.copy()
    if regione != "Tutte" and "nuts2_code" in d:
        d = d[d["nuts2_code"].eq(regione)]
    if provincia != "Tutte" and "nuts3_code" in d:
        d = d[d["nuts3_code"].eq(provincia)]
    return d


# ============================================================
# 5. MAPPA KEPLER CON SFONDO FISICO E CONFINI BIANCHI
# ============================================================

@st.cache_data(show_spinner=False, ttl=86400)
def confini(livello, anno):
    nome = f"NUTS_RG_20M_{anno}_4326_LEVL_{livello}.geojson"
    p = percorso(nome)

    if p.exists():
        contenuto = json.loads(p.read_text(encoding="utf-8"))
    else:
        url = (
            "https://gisco-services.ec.europa.eu/"
            f"distribution/v2/nuts/geojson/{nome}"
        )
        richiesta = Request(url, headers={"User-Agent": "AirbnbSports/4.2"})
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


def mappa(df, livello, codice, campo, titolo, chiave):
    st.subheader(titolo)

    if df.empty:
        st.info("Nessun dato da rappresentare.")
        return

    if KeplerGl is None:
        st.error("Aggiungi keplergl al file requirements.txt.")
        return

    with st.expander("Impostazioni cartografiche"):
        versione = st.selectbox(
            "Confini NUTS", [2024, 2021, 2016], key="nuts_" + chiave
        )
        editor = st.checkbox(
            "Mostra strumenti Kepler", key="editor_" + chiave
        )
        st.caption(
            "La versione dei confini deve corrispondere ai codici dei dati."
        )

    try:
        d = df.dropna(subset=[codice, campo])
        if d.duplicated(codice).any():
            raise ValueError("Più valori per lo stesso territorio.")

        lookup = d.set_index(codice)[campo].to_dict()
        nomi = nomi_reg if livello == 2 else nomi_prov

        features = []
        punti = []
        abbinati = set()

        for feature in confini(livello, versione):
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

        mancanti = set(lookup) - abbinati
        if mancanti:
            st.warning(
                "Codici non abbinati ai confini: "
                + ", ".join(sorted(mancanti))
            )

        if not features:
            raise ValueError("Nessun confine corrispondente ai dati.")

        xs, ys = zip(*punti)
        ampiezza = max(max(xs) - min(xs), 1.4 * (max(ys) - min(ys)), 0.1)
        zoom = max(3, min(9, math.log2(360 / ampiezza) - 1.05))

        # Base fisica raster senza etichette.
        # Il servizio World Physical Map ha dettaglio massimo limitato:
        # a zoom elevati il raster viene ingrandito.
        stile_fisico = {
            "version": 8,
            "name": "SBL Physical",
            "sources": {
                "physical": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/"
                        "services/World_Physical_Map/MapServer/"
                        "tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                    "maxzoom": 8,
                    "attribution": "Tiles © Esri — Source: US National Park Service",
                }
            },
            "layers": [
                {
                    "id": "background",
                    "type": "background",
                    "paint": {"background-color": "#FAF9F6"},
                },
                {
                    "id": "physical",
                    "type": "raster",
                    "source": "physical",
                    "paint": {
                        "raster-opacity": 0.48,
                        "raster-saturation": -0.55,
                    },
                },
            ],
        }

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
                    "styleType": "sbl_physical",
                    "mapStyles": {
                        "sbl_physical": {
                            "id": "sbl_physical",
                            "label": "Sfondo fisico",
                            "custom": True,
                            "style": stile_fisico,
                        }
                    },
                },
                "visState": {
                    "filters": [],
                    "layers": [{
                        "id": "nuts",
                        "type": "geojson",
                        "config": {
                            "dataId": "territori",
                            "label": f"NUTS{livello}",
                            "columns": {"geojson": "_geojson"},
                            "isVisible": True,
                            "color": [75, 175, 201],
                            "visConfig": {
                                "opacity": 0.85,
                                "strokeOpacity": 1,
                                "strokeColor": [255, 255, 255],
                                "thickness": 1,
                                "stroked": True,
                                "filled": True,
                                "enable3d": False,
                                "colorRange": {
                                    "name": "SBL",
                                    "type": "sequential",
                                    "category": "Custom",
                                    "colors": [
                                        "#D1EFF2",
                                        "#B4DEE6",
                                        "#8CC6D2",
                                        "#599CAC",
                                        "#276579",
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
                                "territori": [
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

        k = KeplerGl(
            height=420,
            show_docs=False,
            data={
                "territori": {
                    "type": "FeatureCollection",
                    "features": features,
                }
            },
            config=config,
        )

        with tempfile.TemporaryDirectory() as cartella:
            p = Path(cartella) / "mappa.html"
            k.save_to_html(file_name=str(p), read_only=not editor)
            contenuto = p.read_text(encoding="utf-8")

        components.html(contenuto, height=430, scrolling=False)
        st.caption(
            "Confini: Eurostat/GISCO. Sfondo: Esri, US National Park Service. "
            "La mappa richiede WebGL e accesso ai servizi cartografici."
        )

    except Exception as exc:
        st.warning(f"Mappa non disponibile: {exc}")

    with st.expander("Dati della mappa"):
        tabella(df, f"mappa_{chiave}.csv")


# ============================================================
# 6. LETTURA DEL MASTER: PERSONE, PARTECIPANTI E PROXY
# ============================================================

ALIAS = {
    "evento": [
        "Evento originale", "evento_originale", "Evento",
        "event_name", "nome_evento"
    ],
    "id": ["event_id", "id_evento", "ID evento"],
    "sport": ["Sport", "disciplina"],
    "inizio": ["Data inizio", "data_inizio", "start_date", "inizio"],
    "fine": ["Data fine", "data_fine", "end_date", "fine"],
    "nuts": ["Codici NUTS2", "nuts2_code", "NUTS2", "codice_nuts2"],
    "tipo": [
        "tipo_base", "Tipo indicatore", "tipo_indicatore",
        "tipo_indicatore_presenze", "attendance_type"
    ],
    "persone": [
        "Visitatori unici", "visitatori_unici",
        "Visitatori unici stimati", "visitatori_unici_stimati",
        "Persone", "persone_stimate", "unique_visitors"
    ],
    "partecipanti": [
        "Partecipanti", "partecipanti_stimati",
        "Numero partecipanti", "participants"
    ],
    "ingressi": [
        "Ingressi cumulativi", "ingressi_cumulativi",
        "Presenze cumulative", "presenze_cumulative",
        "attendance_cumulativa", "total_attendance"
    ],
    "biglietti": [
        "Biglietti", "Biglietti stimati", "biglietti_stimati",
        "biglietti_venduti", "tickets", "tickets_sold"
    ],
    "attendance": [
        "Presenze stimate", "presenze_stimate", "Presenze",
        "Presenze totali", "presenze_totali", "Attendance",
        "attendance_total", "attendance_stimata",
        "Valore presenze", "valore_presenze",
        "Valore indicatore", "valore_indicatore",
    ],
    "equivalenti": [
        "base_equivalente", "presenze_equivalenti",
        "visitatori_equivalenti"
    ],
}


def colonne_ruolo(df, ruolo):
    """Solo corrispondenze esplicite: nessuna ricerca in fonti o note."""
    risultato = []
    for alias in ALIAS[ruolo]:
        c = trova(df, alias)
        if c and c not in risultato:
            risultato.append(c)
    return risultato


def primo_testo(riga, colonne):
    for c in colonne:
        v = testo(riga[c])
        if v:
            return v
    return ""


def primo_numero(riga, colonne):
    """
    Cerca il primo valore numerico disponibile fra le colonne del ruolo.
    Restituisce anche la colonna e il contenuto originale.
    """
    primo_originale = ""
    prima_colonna = ""

    for c in colonne:
        originale = testo(riga[c])
        if originale and not primo_originale:
            primo_originale = originale
            prima_colonna = c

        v = leggi_numero(riga[c])
        if pd.notna(v):
            return v, c, originale

    return float("nan"), prima_colonna, primo_originale


def classe_tipo(valore):
    s = norm(valore)

    if "equivalent" in s:
        return "EQUIVALENTI"

    if "bigliett" in s or "ticket" in s:
        return "BIGLIETTI_STIMATI" if "stimat" in s else "BIGLIETTI"

    if "partecipant" in s or "participant" in s:
        return "PARTECIPANTI"

    if "cumul" in s:
        return "INGRESSI_CUMULATIVI"

    if "ingress" in s:
        return "INGRESSI_STIMATI" if "stimat" in s else "INGRESSI_CUMULATIVI"

    if any(x in s for x in ("person", "unic", "unique", "visitator")):
        return "PERSONE_STIMATE" if "stimat" in s else "PERSONE"

    return ""


def prepara_master(master, ruoli):
    risultati = []

    for posizione, (_, r) in enumerate(master.iterrows()):
        nome = primo_testo(r, ruoli["evento"]) or f"Riga {posizione + 1}"
        sport = primo_testo(r, ruoli["sport"]) or "Non indicato"
        inizio = leggi_data(primo_testo(r, ruoli["inizio"]))
        fine_raw = primo_testo(r, ruoli["fine"])
        fine = leggi_data(fine_raw)

        note = []
        problemi = []

        if not fine_raw and pd.notna(inizio):
            fine = inizio
            note.append("Data fine assente: durata convenzionale di 1 giorno")

        durata = (
            (fine - inizio).days + 1
            if pd.notna(inizio) and pd.notna(fine) and fine >= inizio
            else None
        )

        if durata is None:
            problemi.append("Date mancanti o non valide")

        nuts_raw = primo_testo(r, ruoli["nuts"]).upper()
        codici = list(dict.fromkeys(re.findall(r"\bIT[A-Z0-9]{2}\b", nuts_raw)))
        nuts = codici[0] if len(codici) == 1 else nuts_raw

        if len(codici) != 1:
            problemi.append("Serve una singola regione per riga territoriale")
        elif nomi_reg and nuts not in nomi_reg:
            problemi.append("NUTS2 non abbinato a Eurostat")

        tipo_originale = primo_testo(r, ruoli["tipo"])
        classe = classe_tipo(tipo_originale)

        valori = {
            ruolo: primo_numero(r, ruoli[ruolo])
            for ruolo in (
                "persone", "partecipanti", "ingressi",
                "biglietti", "attendance", "equivalenti"
            )
        }

        # La classificazione esplicita ha precedenza.
        # I valori alternativi vengono cercati nelle colonne pertinenti.
        ordini = {
            "PERSONE": ["persone", "attendance", "partecipanti"],
            "PERSONE_STIMATE": ["persone", "attendance", "partecipanti"],
            "PARTECIPANTI": ["partecipanti", "attendance", "persone"],
            "INGRESSI_CUMULATIVI": ["ingressi", "attendance"],
            "INGRESSI_STIMATI": ["ingressi", "attendance"],
            "BIGLIETTI": ["biglietti", "attendance"],
            "BIGLIETTI_STIMATI": ["biglietti", "attendance"],
            "EQUIVALENTI": ["equivalenti"],
        }

        if not classe:
            if pd.notna(valori["persone"][0]):
                classe = "PERSONE"
            elif pd.notna(valori["partecipanti"][0]):
                classe = "PARTECIPANTI"
            elif pd.notna(valori["equivalenti"][0]):
                classe = "EQUIVALENTI"
            elif pd.notna(valori["ingressi"][0]):
                classe = "INGRESSI_CUMULATIVI"
            elif pd.notna(valori["biglietti"][0]):
                classe = "BIGLIETTI"
            elif norm(sport) in (
                "f1", "formula1", "formulauno", "motogp", "superbike"
            ):
                # Regola già documentata nel progetto per questi eventi.
                classe = "INGRESSI_CUMULATIVI"
                note.append("Classificazione motoristica concordata nel progetto")

        volume = float("nan")
        colonna = ""
        originale = ""

        for ruolo in ordini.get(classe, ["attendance"]):
            v, c, o = valori[ruolo]
            if not originale and o:
                originale, colonna = o, c
            if pd.notna(v):
                volume, colonna, originale = v, c, o
                break

        base = float("nan")
        categoria = ""
        metodo = ""

        if pd.isna(volume) or volume < 0:
            problemi.append("Valore numerico non recuperato: vedere colonna e testo originali")
        elif classe in ("PERSONE", "PERSONE_STIMATE", "PARTECIPANTI"):
            base = volume
            categoria = "PERSONE"
            metodo = "Valore della riga territoriale"
        elif classe == "EQUIVALENTI":
            base = volume
            categoria = "PROXY"
            metodo = "Base equivalente già presente nel master"
        elif classe in (
            "INGRESSI_CUMULATIVI", "INGRESSI_STIMATI",
            "BIGLIETTI", "BIGLIETTI_STIMATI"
        ):
            categoria = "PROXY"
            if durata:
                base = volume / durata
                metodo = "Ingressi o biglietti / durata: equivalente giornaliero"
        else:
            problemi.append("Tipo indicatore non riconosciuto")

        if norm(sport) == "scidifondo":
            problemi.append("Sci di fondo escluso dal perimetro")

        identificativo = primo_testo(r, ruoli["id"])
        if not identificativo:
            identificativo = f"{norm(nome)}|{inizio}|{norm(sport)}"

        risultati.append({
            "Riga master": posizione + 1,
            "ID evento": identificativo,
            "Evento": nome,
            "Sport": sport,
            "Inizio": inizio,
            "Fine": fine,
            "Durata": durata,
            "NUTS2": nuts,
            "Tipo originale": tipo_originale,
            "Classe": classe or "NON RICONOSCIUTA",
            "Categoria": categoria,
            "Colonna usata": colonna,
            "Valore originale": originale,
            "Volume": volume,
            "Base": base,
            "Metodo": metodo,
            "Note": "; ".join(note),
            "Problema": "; ".join(problemi),
        })

    eventi = pd.DataFrame(risultati)

    if eventi.empty:
        return eventi

    # Nessuna divisione per il numero di regioni:
    # il master del progetto contiene righe già territorializzate.
    duplicati = eventi.duplicated(
        ["ID evento", "NUTS2", "Inizio", "Fine"], keep=False
    )

    for i in eventi.index[duplicati]:
        messaggio = "Più righe per lo stesso evento-territorio: verificare duplicazione"
        precedente = eventi.at[i, "Problema"]
        eventi.at[i, "Problema"] = (
            precedente + "; " + messaggio if precedente else messaggio
        )

    eventi["Calcolabile"] = eventi["Problema"].eq("")
    return eventi


def mesi_eventi(eventi):
    righe = []

    for _, r in eventi[eventi["Calcolabile"]].iterrows():
        for periodo in pd.period_range(r["Inizio"], r["Fine"], freq="M"):
            a = max(r["Inizio"], periodo.start_time.normalize())
            b = min(r["Fine"], periodo.end_time.normalize())
            giorni = (b - a).days + 1

            righe.append({
                "Riga master": r["Riga master"],
                "ID evento": r["ID evento"],
                "Evento": r["Evento"],
                "Sport": r["Sport"],
                "nuts2_code": r["NUTS2"],
                "Regione": nomi_reg.get(r["NUTS2"], r["NUTS2"]),
                "year": periodo.year,
                "month": periodo.month,
                "data": periodo.start_time,
                "Categoria": r["Categoria"],
                "Classe": r["Classe"],
                "Base": r["Base"],
                "Durata": r["Durata"],
                "Peso mese": giorni / r["Durata"],
            })

    if not righe:
        return pd.DataFrame()

    d = pd.DataFrame(righe)

    if reg.empty:
        d["guest_nights_total"] = float("nan")
    else:
        d = d.merge(
            reg[["nuts2_code", "year", "month", "guest_nights_total"]],
            on=["nuts2_code", "year", "month"],
            how="left",
            validate="many_to_one",
        )

    return d


def simula(mesi, parametri, preset=None):
    d = mesi.copy()

    d["Quota pernottanti %"] = (
        preset if preset is not None
        else d["Sport"].map(lambda s: parametri[s]["quota"])
    )
    notti = d["Sport"].map(lambda s: parametri[s]["notti"])
    d["Notti applicate"] = pd.concat(
        [notti.rename("scelte"), d["Durata"]], axis=1
    ).min(axis=1).clip(1, 5)

    d["Pernottamenti simulati"] = (
        d["Base"]
        * d["Quota pernottanti %"] / 100
        * d["Notti applicate"]
        * d["Peso mese"]
    )

    d["Rapporto %"] = (
        100 * d["Pernottamenti simulati"]
        / d["guest_nights_total"].where(d["guest_nights_total"] > 0)
    )
    return d


def riepilogo(d):
    r = (
        d.groupby(
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
    r["Rapporto %"] = (
        100 * r["Pernottamenti_simulati"]
        / r["Domanda_osservata"].where(r["Domanda_osservata"] > 0)
    )
    return r


def filtro_tutti(label, opzioni, key, formato=None):
    if st.checkbox(f"{label}: tutti", value=True, key=key + "_all"):
        return list(opzioni)

    kwargs = {"format_func": formato} if formato else {}
    return st.multiselect(label, list(opzioni), key=key, **kwargs)


# ============================================================
# 7. PAGINE
# ============================================================

st.title(pagina)

if pagina == "Panorama nazionale":
    richiedi("reg")
    st.caption("Quadro nazionale. I filtri territoriali si applicano alle altre sezioni.")

    # Solo anni con 12 mesi valorizzati per tutti i territori presenti nel panel.
    copertura = reg.groupby(["year", "nuts2_code"]).agg(
        mesi=("month", "nunique"),
        validi=("guest_nights_total", "count"),
    )
    anni_completi = []

    for anno in sorted(reg["year"].unique()):
        c = copertura.loc[anno]
        if (
            len(c) == reg["nuts2_code"].nunique()
            and c["mesi"].eq(12).all()
            and c["validi"].eq(12).all()
        ):
            anni_completi.append(anno)

    if not anni_completi:
        st.info("Non risultano anni completi per tutto il panel.")
    else:
        annuale = aggrega(
            reg[reg["year"].isin(anni_completi)],
            ["year"], "guest_nights_total"
        )
        ultimo = max(anni_completi)
        valori = annuale.set_index("year")["guest_nights_total"]
        attuale = valori.loc[ultimo]
        precedente = valori.get(ultimo - 1, float("nan"))
        variazione = (
            100 * (attuale / precedente - 1)
            if pd.notna(precedente) and precedente > 0 else float("nan")
        )

        a, b, c = st.columns(3)
        a.metric(f"Pernottamenti {ultimo}", numero(attuale))
        b.metric("Variazione annua", pct(variazione))
        c.metric("Territori NUTS2", numero(reg["nuts2_code"].nunique()))

        grafico(px.bar(
            annuale, x="year", y="guest_nights_total",
            color_discrete_sequence=["#126982"],
            labels={"year": "Anno", "guest_nights_total": "Pernottamenti"},
        ))

    anno = st.selectbox(
        "Anno cartografico",
        sorted(reg["year"].unique(), reverse=True),
        index=(
            sorted(reg["year"].unique(), reverse=True).index(max(anni_completi))
            if anni_completi else 0
        ),
    )

    st.caption(
        "Per anni incompleti la mappa rappresenta i soli mesi disponibili."
    )
    d = aggrega(
        reg[reg["year"].eq(anno)], ["nuts2_code"], "guest_nights_total"
    )
    mappa(
        d, 2, "nuts2_code", "guest_nights_total",
        f"Pernottamenti regionali · {anno}", "nazionale"
    )


elif pagina == "Esplora un territorio":
    if provincia != "Tutte":
        richiedi("prov")
        d = filtra_territorio(prov)
        st.subheader(str(nomi_prov.get(provincia, provincia)))
        st.caption("Serie annuale provinciale, senza mensilizzazione artificiale.")
        serie = aggrega(d, ["year"], "guest_nights_total")
        asse = "year"
    else:
        richiedi("reg")
        d = filtra_territorio(reg)
        st.subheader("Italia" if regione == "Tutte" else str(nomi_reg[regione]))
        serie = aggrega(d, ["data"], "guest_nights_total")
        asse = "data"

    grafico(px.line(
        serie.sort_values(asse),
        x=asse, y="guest_nights_total",
        color_discrete_sequence=["#126982"],
        labels={asse: "Periodo", "guest_nights_total": "Pernottamenti"},
    ))

    if provincia == "Tutte":
        st.subheader("Stagionalità")
        stagioni = aggrega(d, ["year", "month"], "guest_nights_total")
        stagioni["Anno"] = stagioni["year"].astype(str)
        grafico(px.line(
            stagioni.sort_values(["year", "month"]),
            x="month", y="guest_nights_total", color="Anno",
            labels={"month": "Mese", "guest_nights_total": "Pernottamenti"},
        ))

    tabella(d, "territorio.csv")


elif pagina == "Eventi e simulatore":
    p = FILES["eventi"]

    if not p.exists():
        candidati = [
            f for cartella in (ROOT, ROOT / "output_eurostat")
            if cartella.exists()
            for f in cartella.glob("*.xlsx")
            if "eventmaster" in norm(f.stem)
        ]
        if len(candidati) == 1:
            p = candidati[0]
        else:
            st.error("Event Master Excel non trovato.")
            st.stop()

    try:
        fogli = excel_fogli(str(p), p.stat().st_mtime)
        foglio = st.selectbox("Foglio del master", fogli)
        master = excel_cache(str(p), p.stat().st_mtime, foglio)
    except Exception as exc:
        st.error(f"Errore nella lettura del master: {exc}")
        st.stop()

    if master.empty:
        st.info("Il foglio selezionato è vuoto.")
        st.stop()

    ruoli = {r: colonne_ruolo(master, r) for r in ALIAS}

    with st.expander("Colonne riconosciute e collegamenti"):
        st.caption(
            "Puoi aggiungere una colonna al ruolo corretto se il suo nome "
            "non è tra quelli riconosciuti. Non selezionare colonne di fonti, "
            "note, percentuali o pernottamenti già calcolati."
        )
        extra_role = st.selectbox(
            "Ruolo da integrare",
            ["Nessuno"] + list(ALIAS),
        )
        if extra_role != "Nessuno":
            extra_col = st.selectbox(
                "Colonna del master",
                ["Nessuna"] + list(master.columns),
            )
            if extra_col != "Nessuna":
                ruoli[extra_role] = [extra_col] + [
                    c for c in ruoli[extra_role] if c != extra_col
                ]

        st.dataframe(
            pd.DataFrame([
                {"Ruolo": r, "Colonne": " · ".join(c)}
                for r, c in ruoli.items()
            ]),
            hide_index=True,
            use_container_width=True,
        )

    eventi = prepara_master(master, ruoli)

    catalogo, controllo, simulatore, contesto, salvati = st.tabs([
        "Catalogo completo",
        "Controllo lettura",
        "Simulatore",
        "Contesto turistico",
        "Screening precedenti",
    ])

    with catalogo:
        st.caption(
            "Tutte le righe dell'Excel restano disponibili, "
            "indipendentemente dalla loro utilizzabilità nel simulatore."
        )
        a, b = st.columns(2)
        ricerca = a.text_input("Cerca evento")
        sport = b.selectbox(
            "Sport",
            ["Tutti"] + sorted(eventi["Sport"].unique()),
            key="catalogo_sport",
        )
        usa_regione = st.checkbox(
            "Applica il filtro regionale al catalogo",
            value=False,
        )

        vista = eventi.copy()
        if sport != "Tutti":
            vista = vista[vista["Sport"].eq(sport)]
        if usa_regione and regione != "Tutte":
            vista = vista[vista["NUTS2"].eq(regione)]
        if ricerca.strip():
            vista = vista[
                vista["Evento"].str.contains(
                    ricerca.strip(), case=False, regex=False, na=False
                )
            ]

        a, b, c = st.columns(3)
        a.metric("Righe del master", numero(len(master)))
        b.metric("Eventi identificati", numero(eventi["ID evento"].nunique()))
        c.metric("Righe mostrate", numero(len(vista)))

        originale = master.iloc[
            (vista["Riga master"] - 1).astype(int).tolist()
        ]
        tabella(originale, "catalogo_event_master.csv", 380)

    with controllo:
        buone = eventi[eventi["Calcolabile"]]
        sospese = eventi[~eventi["Calcolabile"]]

        a, b, c = st.columns(3)
        a.metric("Righe lette", numero(len(eventi)))
        b.metric("Calcolabili", numero(len(buone)))
        c.metric("Da verificare / escluse", numero(len(sospese)))

        st.info(
            "I valori delle singole righe sono trattati come già territorializzati. "
            "Non viene applicato un nuovo riparto uniforme tra regioni."
        )

        classi = (
            eventi.groupby(["Classe", "Categoria", "Calcolabile"], dropna=False)
            .size().reset_index(name="Righe")
        )
        tabella(classi, "classificazione_letta.csv")

        if not sospese.empty:
            st.subheader("Diagnostica delle righe non calcolabili")
            tabella(
                sospese[[
                    "Riga master", "Evento", "Sport", "NUTS2",
                    "Tipo originale", "Colonna usata",
                    "Valore originale", "Problema"
                ]],
                "diagnostica_eventi.csv",
                380,
            )

        with st.expander("Base completa usata nei calcoli"):
            tabella(eventi, "base_eventi_letta.csv")

        st.caption(
            "La lettura numerica può riconoscere un singolo valore accompagnato "
            "da testo. Intervalli e descrizioni con più numeri restano da verificare. "
            "Il testo originale rimane disponibile per controllare stime "
            "approssimative e limiti inferiori."
        )

    with simulatore:
        st.write(
            "Stima esplorativa dei pernottamenti associabili agli eventi. "
            "Persone e proxy vengono mantenuti separati."
        )
        st.caption(
            "La quota è applicata alla base selezionata dell'evento, come "
            "parametro complessivo dello scenario. Non viene aggiunto "
            "un secondo moltiplicatore automatico sulla provenienza."
        )

        with st.expander("Filtri", expanded=False):
            sport_scelti = filtro_tutti(
                "Sport", sorted(eventi["Sport"].unique()), "sim_sport"
            )

            selezione_ids = eventi[
                eventi["Sport"].isin(sport_scelti)
            ].drop_duplicates("ID evento")

            etichette = selezione_ids.set_index("ID evento")["Evento"].to_dict()

            ids = filtro_tutti(
                "Eventi", selezione_ids["ID evento"].tolist(), "sim_ids",
                formato=lambda k: f"{etichette[k]} · {k}",
            )

        scelti = eventi[
            eventi["Sport"].isin(sport_scelti)
            & eventi["ID evento"].isin(ids)
        ].copy()

        if regione != "Tutte":
            scelti = scelti[scelti["NUTS2"].eq(regione)]

        st.caption(
            "Territorio: "
            + ("Italia" if regione == "Tutte" else str(nomi_reg[regione]))
        )

        sospese = scelti[~scelti["Calcolabile"]]
        if not sospese.empty:
            st.warning(
                f"{len(sospese)} righe selezionate non entrano nel calcolo. "
                "Il motivo e il valore originale sono in Controllo lettura."
            )

        mesi = mesi_eventi(scelti)

        if mesi.empty:
            st.info("Nessuna riga calcolabile per questa selezione.")
        else:
            with st.expander("Periodo", expanded=False):
                anni = filtro_tutti(
                    "Anni", sorted(mesi["year"].unique()), "sim_anni"
                )
                mesi_scelti = filtro_tutti(
                    "Mesi", list(range(1, 13)), "sim_mesi"
                )

            mesi = mesi[
                mesi["year"].isin(anni)
                & mesi["month"].isin(mesi_scelti)
            ]

            if mesi.empty:
                st.info("Nessuna riga nel periodo selezionato.")
            else:
                sport_attivi = sorted(mesi["Sport"].unique())
                a, b = st.columns([2, 1])
                preset = a.selectbox(
                    "Preset comune",
                    [5, 15, 25, 40, 60],
                    index=2,
                    format_func=lambda q: f"{q}%",
                )
                if b.button("Applica agli sport"):
                    for s in sport_attivi:
                        st.session_state["q_" + s] = int(preset)

                parametri = {}

                for s in sport_attivi:
                    durata_max = int(
                        mesi.loc[mesi["Sport"].eq(s), "Durata"].max()
                    )
                    massimo = max(1, min(durata_max, 5))

                    with st.expander(
                        f"{s} · quota e notti",
                        expanded=len(sport_attivi) <= 2,
                    ):
                        c1, c2 = st.columns(2)

                        quota = c1.slider(
                            "Quota pernottanti (%)",
                            0, 100, 25,
                            key="q_" + s,
                        )

                        if massimo == 1:
                            notti = 1
                            c2.metric("Notti per pernottante", 1)
                        else:
                            chiave_notti = "n_" + s
                            if chiave_notti in st.session_state:
                                st.session_state[chiave_notti] = min(
                                    int(st.session_state[chiave_notti]),
                                    massimo,
                                )
                            notti = c2.slider(
                                "Notti per pernottante",
                                1, massimo, 1,
                                key=chiave_notti,
                            )

                        st.caption(
                            "0% = nessun pernottante; 100% = tutta la base "
                            "considerata. Le notti sono limitate anche alla "
                            "durata del singolo evento. Nessun minimo "
                            "automatico per tifosi ospiti o trasferte."
                        )

                        parametri[s] = {"quota": quota, "notti": notti}

                risultato = simula(mesi, parametri)
                mensile = riepilogo(risultato)

                a, b, c = st.columns(3)
                a.metric(
                    "Notti · base persone",
                    numero(risultato.loc[
                        risultato["Categoria"].eq("PERSONE"),
                        "Pernottamenti simulati"
                    ].sum()),
                )
                b.metric(
                    "Notti · base proxy",
                    numero(risultato.loc[
                        risultato["Categoria"].eq("PROXY"),
                        "Pernottamenti simulati"
                    ].sum()),
                )
                c.metric(
                    "Righe evento–regione–mese ≥ 2%",
                    numero(risultato["Rapporto %"].ge(2).sum()),
                )

                st.caption(
                    "Il conteggio sopra soglia riguarda le righe, non gli eventi "
                    "unici. I proxy non sono visitatori unici e non vengono "
                    "sommati alle persone in un unico totale."
                )

                mancanti = risultato[
                    risultato["guest_nights_total"].isna()
                    | risultato["guest_nights_total"].le(0)
                ]
                if not mancanti.empty:
                    st.info(
                        "Per alcune righe manca un denominatore Eurostat positivo: "
                        "i pernottamenti sono simulati, ma il rapporto resta n.d."
                    )

                andamento = (
                    mensile.groupby(["data", "Categoria"])[
                        "Pernottamenti_simulati"
                    ].sum().reset_index()
                )
                grafico(px.bar(
                    andamento,
                    x="data", y="Pernottamenti_simulati",
                    color="Categoria", barmode="group",
                    color_discrete_map={
                        "PERSONE": "#126982",
                        "PROXY": "#DC813F",
                    },
                    labels={
                        "data": "Mese",
                        "Pernottamenti_simulati": "Pernottamenti simulati",
                    },
                ))

                st.subheader("Confronto regione–mese")
                tabella(mensile, "simulazione_regioni_mesi.csv")

                with st.expander("Cinque scenari a confronto"):
                    confronti = []
                    for q in [5, 15, 25, 40, 60]:
                        scenario = simula(mesi, parametri, preset=q)
                        for categoria, gruppo in scenario.groupby("Categoria"):
                            confronti.append({
                                "Quota": f"{q}%",
                                "Categoria": categoria,
                                "Pernottamenti": gruppo[
                                    "Pernottamenti simulati"
                                ].sum(),
                                "Righe almeno 2%": gruppo[
                                    "Rapporto %"
                                ].ge(2).sum(),
                            })

                    confronto = pd.DataFrame(confronti)
                    st.caption(
                        "Variano solo le quote pernottanti; le notti per sport "
                        "restano quelle impostate."
                    )
                    tabella(confronto, "confronto_cinque_scenari.csv")

                with st.expander("Dettaglio evento–regione–mese"):
                    tabella(risultato, "simulazione_eventi_mesi.csv", 380)

    with contesto:
        if reg.empty:
            st.info("Dati regionali non disponibili.")
        else:
            codici = sorted(nomi_reg, key=lambda k: str(nomi_reg[k]))
            scelta = st.selectbox(
                "Regione del confronto",
                codici,
                index=codici.index(regione) if regione in codici else 0,
                format_func=lambda k: str(nomi_reg[k]),
            )
            serie = reg[reg["nuts2_code"].eq(scelta)].copy()
            precedente = serie[
                ["year", "month", "guest_nights_total"]
            ].copy()
            precedente["year"] += 1
            precedente = precedente.rename(
                columns={"guest_nights_total": "Stesso mese anno precedente"}
            )
            serie = serie.merge(
                precedente, on=["year", "month"],
                how="left", validate="one_to_one",
            )
            denominatore = serie["Stesso mese anno precedente"].where(
                serie["Stesso mese anno precedente"] > 0
            )
            serie["Variazione %"] = (
                100 * (serie["guest_nights_total"] / denominatore - 1)
            )

            grafico(px.line(
                serie.sort_values("data"),
                x="data", y="guest_nights_total",
                color_discrete_sequence=["#126982"],
                labels={"data": "Mese", "guest_nights_total": "Pernottamenti"},
            ))
            st.caption(
                "Confronto descrittivo. Non separa l'effetto degli eventi "
                "da stagionalità, festività e altri fattori."
            )
            tabella(serie, "contesto_regionale.csv")

    with salvati:
        disponibili = [n for n in SCREENING if percorso(n).exists()]

        if not disponibili:
            st.info("Nessun file di screening trovato.")
        else:
            scelta = st.selectbox("File già elaborato", disponibili)
            st.caption(
                "Questi sono gli output precedenti del progetto. "
                "Non vengono sovrascritti né reinterpretati come basi grezze. "
                "Il confronto con il nuovo calcolo va effettuato a parità "
                "di eventi, periodo, quote, notti e definizione della base."
            )
            try:
                tabella(csv(percorso(scelta)), "storico_" + scelta, 380)
            except Exception as exc:
                st.error(str(exc))


elif pagina == "Forecast 2026–2030":
    provinciale = provincia != "Tutte"
    chiave = "fp" if provinciale else "fr"
    richiedi(chiave)

    dati = filtra_territorio(DATI[chiave])
    prefisso = "forecast_" if provinciale else "gn_"
    codice = "nuts3_code" if provinciale else "nuts2_code"

    st.caption(
        "Scenari letti dai CSV esistenti. "
        "La simulazione degli eventi non modifica il forecast."
    )

    scelti = st.multiselect("Scenari", SCENARI, default=SCENARI)

    if not scelti or dati.empty:
        st.info("Seleziona uno scenario con dati disponibili.")
    else:
        righe = []

        for scenario in scelti:
            campo = prefisso + scenario.lower()
            annuale = aggrega(dati, ["year"], campo)
            annuale = annuale.rename(columns={campo: "Pernottamenti"})
            annuale["Scenario"] = scenario
            righe.append(annuale)

        lungo = pd.concat(righe, ignore_index=True)
        grafico(px.line(
            lungo.sort_values("year"),
            x="year", y="Pernottamenti", color="Scenario",
            markers=True,
            color_discrete_map=COLORI,
            labels={"year": "Anno"},
        ))

        a, b = st.columns(2)
        anno = a.selectbox(
            "Anno cartografico", sorted(dati["year"].unique(), reverse=True)
        )
        scenario = b.selectbox("Scenario cartografico", scelti)
        campo = prefisso + scenario.lower()

        d = aggrega(dati[dati["year"].eq(anno)], [codice], campo)
        mappa(
            d, 3 if provinciale else 2,
            codice, campo,
            f"Forecast {scenario.lower()} · {anno}", "forecast"
        )
        tabella(lungo, "forecast_annuale.csv")


else:
    lettura, metodo, forecast, qualita = st.tabs([
        "Fonti e interpretazione",
        "Metodo eventi",
        "Forecast",
        "Qualità",
    ])

    with lettura:
        st.markdown(
            """
            ### Dati osservati, simulazione e previsione

            **Dati osservati:** pernottamenti degli ospiti rilevati
            nell'ambito coperto dalla fonte Eurostat.

            **Simulazione:** pernottamenti teorici ottenuti applicando
            ipotesi alla base degli eventi.

            **Forecast:** tre scenari già elaborati e salvati nei CSV.
            Il simulatore non li modifica.

            ### Unità territoriali e temporali

            La base mensile del progetto è **NUTS2 × mese**,
            riferita a `tour_ce_omn12`.

            Le serie `tour_ce_oan3` sono utilizzate a livello
            **NUTS3 × anno**. Non vengono create osservazioni
            provinciali mensili artificiali.

            Un pernottamento degli ospiti non equivale a una prenotazione:
            tre ospiti per due notti generano sei pernottamenti.

            La copertura Eurostat non coincide automaticamente con
            tutto il turismo regionale o con il solo Airbnb.

            [Metadati Eurostat](https://ec.europa.eu/eurostat/cache/metadata/en/tour_ce_esms.htm)
            """
        )

    with metodo:
        st.markdown(
            """
            ### 1. Conservazione del master

            Il catalogo mantiene tutte le righe del foglio Excel.
            Le righe non calcolabili sono segnalate, non cancellate.

            I valori delle righe territoriali vengono mantenuti:
            non sono divisi nuovamente per il numero delle regioni.

            Una cella con più regioni non viene ripartita automaticamente.
            Eventuali duplicazioni evento–territorio vengono segnalate
            prima della somma.

            ### 2. Basi persone

            Persone, persone stimate e partecipanti vengono utilizzati
            come base persone, conservando la classe originale.

            La qualità della stima dipende anche dalla definizione
            e dalla precisione della fonte.

            ### 3. Basi proxy

            Ingressi cumulativi e biglietti non vengono chiamati
            visitatori unici.

            Per la simulazione viene costruita una base equivalente:
            **volume della riga / durata dell'evento**.

            Questa è una media giornaliera equivalente, non una stima
            del numero di individui distinti. Nel caso dei biglietti,
            non dimostra neppure la partecipazione effettiva.

            Una base equivalente già presente nel master non viene
            divisa una seconda volta.

            F1, MotoGP e Superbike mantengono il trattamento
            degli ingressi cumulativi concordato nel progetto.

            ### 4. Quote e notti

            La quota per sport varia da 0 a 100%, con valore iniziale 25%.
            I preset comuni sono 5%, 15%, 25%, 40% e 60%.

            Il parametro esprime la quota della base considerata
            associata a pernottamento nello scenario.

            Se la base contiene tutti i partecipanti, questa quota
            incorpora implicitamente anche l'ipotesi sulla provenienza.
            Se contiene già solo non residenti, è una quota condizionata
            a quella popolazione. Le due interpretazioni non vanno confuse.

            Non viene aggiunto un secondo fattore automatico sulla
            provenienza e non vengono imposti minimi per tifosi ospiti.

            Le notti partono da una e sono limitate, per ciascun evento,
            al minimo tra durata e cinque notti.
            Non sono inclusi soggiorni prima o dopo l'evento.

            ### 5. Ripartizione mensile

            Il peso mensile è dato dai giorni dell'evento nel mese
            divisi per la durata complessiva. È una convenzione uniforme,
            non una distribuzione osservata dei pernottamenti.
            """
        )

        st.latex(
            r"N_{e,r,m}=B_{e,r}\cdot q_e\cdot n_e"
            r"\cdot\frac{d_{e,m}}{d_e}"
        )

        st.markdown(
            """
            **B** è la base già territoriale, **q** la quota in forma
            decimale, **n** le notti, **d** i giorni.

            Il rapporto di screening divide il risultato per i
            pernottamenti Eurostat della stessa regione e dello stesso mese.
            """
        )

        st.latex(
            r"R_{e,r,m}=100\cdot\frac{N_{e,r,m}}{GN_{r,m}}"
        )

        st.markdown(
            """
            ### 6. Cosa significa il risultato

            Il rapporto è una misura di dimensione relativa,
            non una stima causale dell'incremento turistico.

            La soglia del 2% serve a individuare casi da approfondire:
            non è una soglia di significatività statistica.

            Persone e proxy restano separati. Le somme fra eventi
            possono comunque includere gli stessi individui o soggiorni.

            Il numeratore può riguardare alloggi non compresi nella fonte
            Eurostat: in quel caso il rapporto confronta perimetri diversi.

            Lo sci di fondo resta nel catalogo storico ma non entra
            nella simulazione.
            """
        )

    with forecast:
        st.markdown(
            """
            ### Lettura degli scenari

            L'app legge i forecast regionali e provinciali già prodotti.
            Non riaddestra i modelli e non ricostruisce i loro parametri
            dai risultati finali.

            Prudente, Intermedio ed Espansivo sono scenari,
            non intervalli probabilistici automaticamente calibrati.

            La documentazione del modello deve specificare periodo
            di training, trattamento del COVID, anno base, stagionalità,
            crescita e riparto territoriale.

            ### Validazione

            I modelli vanno confrontati sugli stessi periodi di test
            con una baseline stagionale semplice, senza usare
            informazioni future nell'addestramento.

            WAPE e bias sono utili per distinguere dimensione
            dell'errore e tendenza a sovrastimare o sottostimare.

            I risultati di test precedenti non vengono presentati
            come ricalcolati da questa dashboard.

            ### Province

            Un forecast provinciale costruito con quote territoriali
            costanti eredita la crescita regionale.
            Non rappresenta una previsione indipendente delle variazioni
            nella distribuzione della domanda tra province.
            """
        )

        st.latex(
            r"WAPE=100\cdot"
            r"\frac{\sum_t|y_t-\hat y_t|}{\sum_t|y_t|}"
        )
        st.latex(
            r"Bias=100\cdot"
            r"\frac{\sum_t(\hat y_t-y_t)}{\sum_t y_t}"
        )

    with qualita:
        controlli = []

        for k in SCHEMI:
            d = DATI[k]
            campi = [
                c for c in SCHEMI[k]
                if c.startswith(("guest_nights", "gn_", "forecast_"))
            ]

            controlli.append({
                "File": FILES[k].name,
                "Stato": ERRORI.get(k, "Chiavi valide e non duplicate"),
                "Righe": len(d),
                "Valori mancanti": (
                    int(d[campi].isna().sum().sum()) if not d.empty else None
                ),
                "Valori negativi": (
                    int(d[campi].lt(0).sum().sum()) if not d.empty else None
                ),
            })

        tabella(pd.DataFrame(controlli), "qualita_file.csv")

        if not reg.empty:
            copertura = (
                reg.groupby(["nuts2_code", "year"])
                .agg(
                    Mesi_presenti=("month", "nunique"),
                    Mesi_valorizzati=("guest_nights_total", "count"),
                )
                .reset_index()
            )
            with st.expander("Copertura regionale"):
                tabella(copertura, "copertura_regionale.csv")

        if not fr.empty and not fp.empty:
            confronti = []

            for s in SCENARI:
                a = aggrega(
                    fr, ["nuts2_code", "year"], "gn_" + s.lower()
                ).rename(columns={"gn_" + s.lower(): "Regione"})
                b = aggrega(
                    fp, ["nuts2_code", "year"], "forecast_" + s.lower()
                ).rename(
                    columns={"forecast_" + s.lower(): "Somma province"}
                )
                c = a.merge(
                    b, on=["nuts2_code", "year"],
                    how="outer", validate="one_to_one",
                )
                c["Scenario"] = s
                c["Scarto"] = c["Somma province"] - c["Regione"]
                confronti.append(c)

            with st.expander("Riconciliazione forecast"):
                tabella(
                    pd.concat(confronti, ignore_index=True),
                    "riconciliazione_forecast.csv",
                )

        st.markdown(
            """
            ### Limiti e controlli da mantenere

            - Valori mancanti non vengono trasformati in zero.
            - Gli anni parziali non sono confrontati come anni completi
              nel riepilogo nazionale.
            - I codici territoriali devono corrispondere alla versione
              NUTS della cartografia.
            - Un numero recuperato da una descrizione mantiene
              il testo originale, inclusi eventuali qualificatori.
            - Gli screening precedenti restano disponibili senza modifiche.
              La loro equivalenza con la nuova simulazione non è presunta.

            Per misurare causalmente l'effetto di un evento servirebbero
            un controfattuale credibile, territori comparabili,
            controlli per altri fattori e verifica delle tendenze precedenti.

            Un aumento nello stesso mese dell'evento, da solo,
            non dimostra che l'evento lo abbia causato.
            """
        )


st.divider()
st.caption(f"Airbnb&Sports · {VERSIONE} · [SBL Consultancy]({SITO})")
