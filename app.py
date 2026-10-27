# ============================================================
# AIRBNB&SPORTS — app.py
#
# Dipendenze:
# streamlit
# pandas
# plotly
# openpyxl
# keplergl
# ============================================================

from pathlib import Path
from urllib.request import Request, urlopen
import base64
import json
import math
import re
import tempfile

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
    initial_sidebar_state="expanded",
)

ROOT = Path(__file__).resolve().parent
VERSIONE = "v5.0 · Screening verificato"
SITO = "https://www.sblconsultancy.it/"

PAGINE = [
    "Panorama nazionale",
    "Esplora un territorio",
    "Eventi e simulatore",
    "Forecast 2026–2030",
    "Dati e metodo",
]

SCENARI = ["Prudente", "Intermedio", "Espansivo"]

COLORI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C",
}

COLORI_BASI = {
    "Persone": "#126982",
    "Equivalenti da presenze": "#DC813F",
    "Equivalenti da biglietti": "#C94F7C",
}


def percorso(nome):
    for cartella in (ROOT, ROOT / "output_eurostat"):
        p = cartella / nome
        if p.is_file():
            return p
    return ROOT / nome


FILES = {
    "reg": percorso("eurostat_master_final.csv"),
    "prov": percorso("eurostat_nuts3_annuale_2018_2025.csv"),
    "fr": percorso("forecast_nuts2_2026_2030_3_scenari.csv"),
    "fp": percorso("forecast_nuts3_2026_2030_3_scenari.csv"),
    "master": percorso(
        "AirbnbSports_Event_Master_Simone_Visitatori.xlsx"
    ),
    "screen": percorso("screening_eventi_completo.csv"),
}

ARCHIVIO = [
    "screening_eventi_completo.csv",
    "screening_eventi_equivalenti.csv",
    "screening_regioni_mesi_completo.csv",
    "screening_regioni_mesi_equivalenti.csv",
    "screening_pernottamenti_5_scenari_notti.csv",
    "screening_pernottamenti_5_scenari.csv",
]


# ============================================================
# 2. LETTURA E PRESENTAZIONE
# ============================================================

@st.cache_data(show_spinner=False)
def leggi_csv_cache(nome, modifica):
    df = pd.read_csv(nome, encoding="utf-8-sig")

    if len(df.columns) == 1 and ";" in str(df.columns[0]):
        df = pd.read_csv(
            nome, encoding="utf-8-sig", sep=";"
        )

    df.columns = [str(c).strip() for c in df.columns]
    return df


def leggi_csv(p):
    if not p.is_file():
        raise FileNotFoundError(f"File non trovato: {p.name}")

    return leggi_csv_cache(str(p), p.stat().st_mtime)


@st.cache_data(show_spinner=False)
def leggi_master_cache(nome, modifica):
    return pd.read_excel(nome, sheet_name="Eventi")


def numero(v, decimali=0):
    if pd.isna(v):
        return "n.d."

    return (
        f"{float(v):,.{decimali}f}"
        .replace(",", "§")
        .replace(".", ",")
        .replace("§", ".")
    )


def percentuale(v):
    return "n.d." if pd.isna(v) else f"{numero(v, 1)}%"


def somma_completa(s):
    if s.empty or s.isna().any():
        return float("nan")
    return s.sum()


def aggrega(df, gruppi, campo):
    return (
        df.groupby(gruppi, dropna=False)[campo]
        .agg(somma_completa)
        .reset_index()
    )


def data_singola(v):
    if pd.isna(v):
        return pd.NaT

    if isinstance(v, (int, float)):
        if 20000 <= float(v) <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(float(v), unit="D")
            ).normalize()
        return pd.NaT

    s = str(v).strip()
    iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", s))

    d = pd.to_datetime(
        s, errors="coerce", dayfirst=not iso
    )

    if pd.isna(d):
        return pd.NaT

    if d.tzinfo is not None:
        d = d.tz_localize(None)

    return d.normalize()


def tabella_pulita(df):
    d = df.copy()

    nomi_date = {
        "data", "inizio", "fine",
        "data inizio", "data fine",
        "data_inizio", "data_fine",
        "start_date", "end_date",
    }

    for c in d.columns:
        if pd.api.types.is_datetime64_any_dtype(d[c]):
            d[c] = d[c].dt.strftime("%d/%m/%Y")

        elif d[c].dtype == object:
            d[c] = d[c].map(
                lambda x: pd.NA
                if isinstance(x, str)
                and x.strip().lower() in (
                    "", "nan", "none", "nat", "null"
                )
                else x
            )

            if str(c).lower().strip() in nomi_date:
                date = d[c].map(data_singola)
                d[c] = d[c].where(
                    date.isna(),
                    date.dt.strftime("%d/%m/%Y"),
                )

    return d.dropna(axis=1, how="all")


def tabella(df, nome, altezza=340):
    if df.empty:
        st.info("Nessuna riga per questa selezione.")
        return

    d = tabella_pulita(df)

    st.dataframe(
        d,
        hide_index=True,
        use_container_width=True,
        height=min(altezza, 40 + 35 * len(d)),
    )

    st.download_button(
        "Scarica CSV",
        d.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=nome,
        mime="text/csv",
        key="download_" + nome,
    )


def grafico(fig, altezza=340):
    fig.update_layout(
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TESTO, size=12),
        margin=dict(l=12, r=12, t=35, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            x=0,
        ),
        hovermode="x unified",
    )

    fig.update_xaxes(showgrid=False, automargin=True)
    fig.update_yaxes(gridcolor=BORDO, automargin=True)

    st.plotly_chart(fig, use_container_width=True)


def filtro_tutti(label, opzioni, key, formato=None):
    tutti = st.checkbox(
        f"{label}: tutti",
        value=True,
        key=key + "_tutti",
    )

    if tutti:
        return list(opzioni)

    kwargs = {}
    if formato:
        kwargs["format_func"] = formato

    return st.multiselect(
        label,
        list(opzioni),
        default=[],
        key=key,
        **kwargs,
    )


# ============================================================
# 3. LOGO E NAVIGAZIONE
# ============================================================

def mostra_logo():
    formati = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }

    for cartella in (
        ROOT,
        ROOT / "assets",
        ROOT / "images",
        ROOT / "img",
        ROOT / "output_eurostat",
    ):
        if not cartella.exists():
            continue

        for p in sorted(cartella.iterdir()):
            if (
                p.is_file()
                and p.stem.lower() == "logo_sbl"
                and p.suffix.lower() in formati
            ):
                dati = base64.b64encode(
                    p.read_bytes()
                ).decode()

                st.markdown(
                    f"""
                    <a href="{SITO}" target="_blank"
                       rel="noopener noreferrer"
                       aria-label="SBL Consultancy">
                        <img
                          src="data:{formati[p.suffix.lower()]};base64,{dati}"
                          alt="SBL Consultancy"
                          style="display:block;width:108px;
                                 max-width:100%;height:auto;
                                 margin-bottom:12px;">
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
                return

    st.markdown(f"[**SBL CONSULTANCY**]({SITO})")


with st.sidebar:
    mostra_logo()
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


# ============================================================
# 4. STILE — RIMOZIONE BARRA BIANCA SUPERIORE
# ============================================================

st.markdown(
    f"""
    <style>
    .stApp {{
        background: {SFONDO};
        color: {TESTO};
    }}

    /* Elimina fondo e ingombro della barra superiore. */
    header[data-testid="stHeader"] {{
        background: transparent !important;
        box-shadow: none !important;
        border: none !important;
        height: 0 !important;
        min-height: 0 !important;
        overflow: visible !important;
    }}

    /* Nasconde strumenti e decorazione della barra. */
    [data-testid="stToolbar"],
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    .stAppDeployButton,
    #MainMenu {{
        display: none !important;
    }}

    /* Mantiene disponibile la riapertura della sidebar. */
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stExpandSidebarButton"] {{
        visibility: visible !important;
        opacity: 1 !important;
        pointer-events: auto !important;
        position: fixed !important;
        top: 8px !important;
        left: 8px !important;
        z-index: 100000 !important;
        color: {TESTO} !important;
    }}

    .block-container {{
        max-width: 1320px;
        padding: 1.5rem 1.6rem 2rem 1.6rem;
    }}

    [data-testid="stSidebar"] {{
        background: {SIDE};
        min-width: 255px;
        max-width: 280px;
    }}

    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] h4 {{
        color: #FFFFFF !important;
    }}

    [data-testid="stSidebar"] [data-testid="stCaptionContainer"] p {{
        color: #E2F0F5 !important;
    }}

    h1 {{font-size: 1.8rem !important;}}
    h2 {{font-size: 1.3rem !important;}}
    h3 {{font-size: 1.08rem !important;}}

    p, label {{font-size: 0.92rem;}}

    [data-testid="stMain"] p,
    [data-testid="stMain"] label,
    [data-testid="stMain"] h1,
    [data-testid="stMain"] h2,
    [data-testid="stMain"] h3 {{
        color: {TESTO};
    }}

    [data-testid="stCaptionContainer"] p {{
        color: {SECONDARIO} !important;
        font-size: 0.83rem !important;
    }}

    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stMultiSelect"] [data-baseweb="select"] > div {{
        background: #F7FAFC !important;
        border-color: #B5D8E1 !important;
        min-height: 36px !important;
        border-radius: 8px !important;
    }}

    [data-testid="stSelectbox"] [data-baseweb="select"] *,
    [data-testid="stMultiSelect"] [data-baseweb="select"] *,
    [data-baseweb="popover"] *,
    [role="listbox"] *,
    [role="option"] * {{
        color: #082B46 !important;
        -webkit-text-fill-color: #082B46 !important;
        font-size: 0.87rem !important;
    }}

    [role="listbox"], [role="option"] {{
        background: #F7FAFC !important;
    }}

    [data-baseweb="tag"] {{
        background: #D8EDF3 !important;
        border-radius: 5px !important;
    }}

    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {{
        background: #F7FAFC !important;
        color: #082B46 !important;
        -webkit-text-fill-color: #082B46 !important;
    }}

    [data-testid="stMetric"] {{
        background: {CARD};
        border: 1px solid {BORDO};
        border-radius: 12px;
        padding: 11px 13px;
    }}

    [data-testid="stMetricValue"] {{
        color: {TESTO};
        font-size: 1.55rem !important;
    }}

    [data-testid="stMetricLabel"] p {{
        font-size: 0.81rem !important;
    }}

    [data-testid="stVerticalBlock"] {{
        gap: 0.8rem;
    }}

    .stButton > button,
    .stDownloadButton > button {{
        background: #F7FAFC;
        color: #082B46 !important;
        border: 1px solid #B5D8E1;
        border-radius: 8px;
    }}

    .stButton > button p,
    .stDownloadButton > button p {{
        color: #082B46 !important;
        font-size: 0.84rem !important;
    }}

    @media (max-width: 768px) {{
        .block-container {{
            padding-top: 3rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 5. DATI TURISTICI E FORECAST
# ============================================================

SCHEMI = {
    "reg": [
        "year", "month", "nuts2_code",
        "nuts2_name", "guest_nights_total",
    ],
    "prov": [
        "year", "nuts3_code", "nuts3_name",
        "nuts2_code", "guest_nights_total",
    ],
    "fr": [
        "year", "month", "nuts2_code",
        "gn_prudente", "gn_intermedio", "gn_espansivo",
    ],
    "fp": [
        "year", "nuts3_code", "nuts2_code",
        "forecast_prudente",
        "forecast_intermedio",
        "forecast_espansivo",
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
        d = leggi_csv(FILES[chiave])

        mancanti = [
            c for c in richieste if c not in d.columns
        ]
        if mancanti:
            raise ValueError(
                "Colonne mancanti: " + ", ".join(mancanti)
            )

        for c in d:
            if c in ("nuts2_code", "nuts3_code"):
                d[c] = (
                    d[c].astype("string")
                    .str.strip().str.upper()
                )

            elif (
                c in ("year", "month")
                or c.startswith(
                    ("guest_nights", "gn_", "forecast_")
                )
            ):
                d[c] = pd.to_numeric(
                    d[c], errors="raise"
                ).replace(
                    [float("inf"), -float("inf")],
                    float("nan"),
                )

        if d[CHIAVI[chiave]].isna().any().any():
            raise ValueError(
                "Chiavi territoriali o temporali mancanti."
            )

        if d.duplicated(CHIAVI[chiave]).any():
            raise ValueError(
                "Chiavi duplicate: aggregazione sospesa."
            )

        if d["year"].mod(1).ne(0).any():
            raise ValueError("Anno non intero.")

        d["year"] = d["year"].astype(int)

        if "month" in d:
            if (
                ~d["month"].between(1, 12)
                | d["month"].mod(1).ne(0)
            ).any():
                raise ValueError("Mesi non validi.")

            d["month"] = d["month"].astype(int)
            d["data"] = pd.to_datetime(
                dict(
                    year=d["year"],
                    month=d["month"],
                    day=1,
                ),
                errors="coerce",
            )

            if d["data"].isna().any():
                raise ValueError("Date non valide.")

        DATI[chiave] = d

    except Exception as exc:
        DATI[chiave] = pd.DataFrame()
        ERRORI[chiave] = str(exc)


reg = DATI["reg"]
prov = DATI["prov"]
fr = DATI["fr"]
fp = DATI["fp"]

nomi_reg = (
    reg.drop_duplicates("nuts2_code")
    .set_index("nuts2_code")["nuts2_name"]
    .to_dict()
    if not reg.empty else {}
)

nomi_prov = (
    prov.drop_duplicates("nuts3_code")
    .set_index("nuts3_code")["nuts3_name"]
    .to_dict()
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
        ["Tutte"] + sorted(
            nomi_reg,
            key=lambda k: str(nomi_reg[k]),
        ),
        format_func=lambda k: (
            "Tutta Italia"
            if k == "Tutte" else str(nomi_reg[k])
        ),
    )

    disponibili = prov.copy()

    if not prov.empty and regione != "Tutte":
        disponibili = prov[
            prov["nuts2_code"].eq(regione)
        ]

    codici_prov = (
        sorted(
            disponibili["nuts3_code"].unique(),
            key=lambda k: str(nomi_prov.get(k, k)),
        )
        if not disponibili.empty else []
    )

    provincia = st.selectbox(
        "Provincia",
        ["Tutte"] + codici_prov,
        format_func=lambda k: (
            "Tutte le province"
            if k == "Tutte"
            else str(nomi_prov.get(k, k))
        ),
    )

    st.caption(
        "Gli eventi sono confrontati con regione e mese. "
        "Il filtro provinciale riguarda le serie annuali."
    )


def filtra_territorio(df):
    d = df.copy()

    if regione != "Tutte" and "nuts2_code" in d:
        d = d[d["nuts2_code"].eq(regione)]

    if provincia != "Tutte" and "nuts3_code" in d:
        d = d[d["nuts3_code"].eq(provincia)]

    return d


# ============================================================
# 6. SCREENING VERIFICATO
#
# Fonte del simulatore:
# screening_eventi_completo.csv
#
# NON ricostruisce la base dal master.
# NON divide nuovamente per durata.
# NON applica nuovi riparti territoriali.
# NON espande nuovamente gli eventi sui mesi.
# ============================================================

def prepara_screening(originale):
    d = originale.copy()

    richieste = [
        "id_riga_territoriale",
        "evento", "sport", "nuts2",
        "anno", "mese", "inizio", "fine",
        "durata_giorni", "giorni_nel_mese",
        "quota_mese",
        "pubblico_originale", "pubblico_equivalente",
        "classe_base", "metodo_base",
        "notti_applicate", "gn_eurostat_mese",
        "pct_medio", "pernottamenti_medio",
    ]

    mancanti = [
        c for c in richieste if c not in d.columns
    ]

    if mancanti:
        raise ValueError(
            "Lo screening non ha lo schema verificato. "
            "Colonne mancanti: " + ", ".join(mancanti)
        )

    chiavi = [
        "id_riga_territoriale", "nuts2", "anno", "mese"
    ]

    if d[chiavi].isna().any().any():
        raise ValueError("Chiavi mancanti nello screening.")

    for c in (
        "evento", "sport", "nuts2",
        "classe_base", "metodo_base"
    ):
        if d[c].isna().any():
            raise ValueError(f"Valori mancanti in {c}.")
        d[c] = d[c].astype(str).str.strip()
        if d[c].eq("").any():
            raise ValueError(f"Valori vuoti in {c}.")

    d["nuts2"] = d["nuts2"].str.upper()
    d["classe_base"] = d["classe_base"].str.upper()

    numeriche = [
        "anno", "mese", "durata_giorni",
        "giorni_nel_mese", "quota_mese",
        "pubblico_originale", "pubblico_equivalente",
        "notti_applicate", "gn_eurostat_mese",
        "pct_medio", "pernottamenti_medio",
    ]

    for c in numeriche:
        d[c] = pd.to_numeric(
            d[c], errors="raise"
        ).replace(
            [float("inf"), -float("inf")],
            float("nan"),
        )

    obbligatorie = [
        "anno", "mese", "durata_giorni",
        "giorni_nel_mese", "quota_mese",
        "pubblico_originale", "pubblico_equivalente",
        "notti_applicate", "pct_medio",
        "pernottamenti_medio",
    ]

    if d[obbligatorie].isna().any().any():
        raise ValueError(
            "Valori necessari al calcolo mancanti nello screening. "
            "Nessuna riga è stata esclusa automaticamente."
        )

    for c in (
        "anno", "mese", "durata_giorni",
        "giorni_nel_mese", "notti_applicate"
    ):
        if d[c].mod(1).ne(0).any():
            raise ValueError(f"Valori non interi in {c}.")

    d["anno"] = d["anno"].astype(int)
    d["mese"] = d["mese"].astype(int)

    if not d["mese"].between(1, 12).all():
        raise ValueError("Mesi fuori dall'intervallo 1–12.")

    if d.duplicated(chiavi).any():
        raise ValueError(
            "Chiavi evento-territorio-mese duplicate."
        )

    if not d["classe_base"].isin(
        ["PERSONE", "EQUIVALENTE"]
    ).all():
        raise ValueError(
            "Classi diverse da PERSONE/EQUIVALENTE "
            "nel file completo: verificarne la versione."
        )

    if (
        d["pubblico_equivalente"].lt(0).any()
        or d["pubblico_originale"].lt(0).any()
        or d["durata_giorni"].lt(1).any()
        or d["giorni_nel_mese"].lt(1).any()
        or not d["quota_mese"].between(0, 1).all()
        or not d["pct_medio"].between(0, 100).all()
    ):
        raise ValueError("Valori fuori intervallo nello screening.")

    d["inizio"] = d["inizio"].map(data_singola)
    d["fine"] = d["fine"].map(data_singola)

    if (
        d[["inizio", "fine"]].isna().any().any()
        or d["fine"].lt(d["inizio"]).any()
    ):
        raise ValueError("Date non valide nello screening.")

    durata_date = (d["fine"] - d["inizio"]).dt.days + 1
    if durata_date.ne(d["durata_giorni"]).any():
        raise ValueError(
            "La durata salvata non coincide con le date."
        )

    d["data"] = pd.to_datetime(
        dict(year=d["anno"], month=d["mese"], day=1),
        errors="coerce",
    )

    if d["data"].isna().any():
        raise ValueError("Anno/mese non valido.")

    mese_fine = d["data"] + pd.offsets.MonthEnd(0)

    inizio_locale = pd.concat(
        [d["inizio"], d["data"]], axis=1
    ).max(axis=1)

    fine_locale = pd.concat(
        [d["fine"], mese_fine], axis=1
    ).min(axis=1)

    giorni_attesi = (
        fine_locale - inizio_locale
    ).dt.days + 1

    if (
        giorni_attesi.lt(1).any()
        or giorni_attesi.ne(d["giorni_nel_mese"]).any()
    ):
        raise ValueError(
            "Giorni mensili non coerenti con le date."
        )

    quota_attesa = (
        d["giorni_nel_mese"] / d["durata_giorni"]
    )

    if (
        d["quota_mese"] - quota_attesa
    ).abs().gt(0.00001).any():
        raise ValueError(
            "Quota mensile non coerente con i giorni."
        )

    gruppo = ["id_riga_territoriale", "nuts2"]

    somme_quote = d.groupby(gruppo)["quota_mese"].sum()

    if somme_quote.sub(1).abs().gt(0.00001).any():
        raise ValueError(
            "Le quote mensili non sommano a 1 "
            "per alcune righe territoriali."
        )

    invarianti = [
        "evento", "sport", "inizio", "fine",
        "durata_giorni", "pubblico_equivalente",
        "classe_base", "metodo_base",
    ]

    for c in invarianti:
        if d.groupby(gruppo)[c].nunique(
            dropna=False
        ).gt(1).any():
            raise ValueError(
                f"{c} cambia tra i mesi della stessa riga territoriale."
            )

    denominatori = d.groupby(
        ["nuts2", "anno", "mese"]
    )["gn_eurostat_mese"].nunique(dropna=False)

    if denominatori.gt(1).any():
        raise ValueError(
            "Denominatori Eurostat diversi per la stessa regione-mese."
        )

    d["tipo_base_visuale"] = "Equivalenti da presenze"

    d.loc[
        d["classe_base"].eq("PERSONE"),
        "tipo_base_visuale",
    ] = "Persone"

    biglietti = (
        d["classe_base"].eq("EQUIVALENTE")
        & d["metodo_base"].str.contains(
            "bigliett|ticket",
            case=False,
            regex=True,
            na=False,
        )
    )

    d.loc[
        biglietti,
        "tipo_base_visuale",
    ] = "Equivalenti da biglietti"

    # Identificativo per la selezione degli eventi:
    # conserva la stessa selezione nelle diverse regioni.
    d["evento_key"] = (
        d["sport"] + " | " + d["evento"]
        + " | " + d["inizio"].dt.strftime("%Y-%m-%d")
        + " | " + d["fine"].dt.strftime("%Y-%m-%d")
    )

    d["Regione"] = d["nuts2"].map(nomi_reg).fillna(d["nuts2"])

    return d


def calcola(dati, quote, notti):
    d = dati.copy()

    d["quota_pct"] = d["sport"].map(quote).astype(float)

    d["notti_scenario"] = pd.concat(
        [
            d["sport"].map(notti).rename("richieste"),
            d["durata_giorni"].clip(upper=5),
        ],
        axis=1,
    ).min(axis=1)

    d["pernottamenti_simulati"] = (
        d["pubblico_equivalente"]
        * d["quota_pct"] / 100
        * d["notti_scenario"]
        * d["quota_mese"]
    )

    d["rapporto_pct"] = (
        100 * d["pernottamenti_simulati"]
        / d["gn_eurostat_mese"].where(
            d["gn_eurostat_mese"] > 0
        )
    )

    return d


def riassumi_simulazione(d):
    r = (
        d.groupby(
            [
                "nuts2", "Regione", "anno", "mese",
                "data", "tipo_base_visuale",
            ],
            dropna=False,
        )
        .agg(
            Pernottamenti_simulati=(
                "pernottamenti_simulati", "sum"
            ),
            Domanda_osservata=(
                "gn_eurostat_mese", "first"
            ),
            Eventi=("evento_key", "nunique"),
        )
        .reset_index()
    )

    r["Rapporto_percentuale"] = (
        100 * r["Pernottamenti_simulati"]
        / r["Domanda_osservata"].where(
            r["Domanda_osservata"] > 0
        )
    )

    return r


# ============================================================
# 7. MAPPE KEPLER
# ============================================================

@st.cache_data(show_spinner=False, ttl=86400)
def confini_nuts(livello, versione):
    nome = (
        f"NUTS_RG_20M_{versione}_4326_LEVL_{livello}.geojson"
    )
    p = percorso(nome)

    if p.is_file():
        contenuto = json.loads(
            p.read_text(encoding="utf-8")
        )
    else:
        url = (
            "https://gisco-services.ec.europa.eu/"
            f"distribution/v2/nuts/geojson/{nome}"
        )

        richiesta = Request(
            url,
            headers={"User-Agent": "AirbnbSports/5.0"},
        )

        with urlopen(richiesta, timeout=30) as risposta:
            contenuto = json.load(risposta)

    return [
        f for f in contenuto["features"]
        if str(
            f["properties"].get("NUTS_ID", "")
        ).startswith("IT")
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


def mappa(df, livello, codice, campo, titolo, key):
    st.subheader(titolo)

    if df.empty:
        st.info("Nessun dato cartografico.")
        return

    if KeplerGl is None:
        st.error(
            "Kepler.gl non è installato: "
            "aggiungi keplergl a requirements.txt."
        )
        return

    with st.expander("Impostazioni mappa"):
        versione = st.selectbox(
            "Versione NUTS",
            [2024, 2021, 2016],
            key="versione_" + key,
        )
        editor = st.checkbox(
            "Mostra strumenti Kepler",
            key="strumenti_" + key,
        )

    try:
        d = df.dropna(subset=[codice, campo])

        if d.duplicated(codice).any():
            raise ValueError(
                "Più valori per lo stesso territorio."
            )

        lookup = d.set_index(codice)[campo].to_dict()
        nomi = nomi_reg if livello == 2 else nomi_prov

        features = []
        punti = []
        abbinati = set()

        for f in confini_nuts(livello, versione):
            nuts = f["properties"]["NUTS_ID"]

            if nuts not in lookup:
                continue

            abbinati.add(nuts)

            features.append({
                "type": "Feature",
                "geometry": f["geometry"],
                "properties": {
                    "Territorio": str(nomi.get(nuts, nuts)),
                    "NUTS": nuts,
                    "Valore": float(lookup[nuts]),
                    "Pernottamenti": numero(lookup[nuts]),
                },
            })

            punti.extend(
                coordinate(f["geometry"]["coordinates"])
            )

        mancanti = set(lookup) - abbinati

        if mancanti:
            st.warning(
                "Codici senza confine corrispondente: "
                + ", ".join(sorted(mancanti))
            )

        if not features:
            raise ValueError(
                "Nessun confine abbinato ai dati."
            )

        xs, ys = zip(*punti)
        ampiezza = max(
            max(xs) - min(xs),
            1.4 * (max(ys) - min(ys)),
            0.1,
        )

        zoom = max(
            3, min(9, math.log2(360 / ampiezza) - 1.05)
        )

        stile = {
            "version": 8,
            "name": "SBL Physical",
            "sources": {
                "physical": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/"
                        "rest/services/World_Physical_Map/"
                        "MapServer/tile/{z}/{y}/{x}"
                    ],
                    "tileSize": 256,
                    "maxzoom": 8,
                    "attribution": (
                        "Tiles © Esri — "
                        "Source: US National Park Service"
                    ),
                }
            },
            "layers": [
                {
                    "id": "background",
                    "type": "background",
                    "paint": {
                        "background-color": "#FAF9F6"
                    },
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
                            "style": stile,
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
                            "columns": {
                                "geojson": "_geojson"
                            },
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
                            "colorField": {
                                "name": "Valore",
                                "type": "real",
                            },
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
                                    {
                                        "name": "Territorio",
                                        "format": None,
                                    },
                                    {
                                        "name": "NUTS",
                                        "format": None,
                                    },
                                    {
                                        "name": "Pernottamenti",
                                        "format": None,
                                    },
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

            k.save_to_html(
                file_name=str(p),
                read_only=not editor,
            )

            contenuto = p.read_text(encoding="utf-8")

        components.html(
            contenuto,
            height=430,
            scrolling=False,
        )

        st.caption(
            "Confini: Eurostat/GISCO. "
            "Sfondo: Esri, US National Park Service. "
            "La visualizzazione richiede WebGL e accesso "
            "ai servizi cartografici."
        )

    except Exception as exc:
        st.warning(f"Mappa non disponibile: {exc}")

    with st.expander("Valori cartografici"):
        tabella(df, "mappa_" + key + ".csv")


# ============================================================
# 8. PAGINE
# ============================================================

st.title(pagina)


# ------------------------------------------------------------
# PANORAMA NAZIONALE
# ------------------------------------------------------------

if pagina == "Panorama nazionale":
    richiedi("reg")

    st.caption(
        "Quadro nazionale. I filtri territoriali "
        "si applicano alle altre sezioni."
    )

    copertura = (
        reg.groupby(["year", "nuts2_code"])
        .agg(
            mesi=("month", "nunique"),
            validi=("guest_nights_total", "count"),
        )
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

    if anni_completi:
        annuale = aggrega(
            reg[reg["year"].isin(anni_completi)],
            ["year"],
            "guest_nights_total",
        )

        ultimo = max(anni_completi)
        valori = annuale.set_index(
            "year"
        )["guest_nights_total"]

        attuale = valori.loc[ultimo]
        precedente = valori.get(
            ultimo - 1, float("nan")
        )

        variazione = (
            100 * (attuale / precedente - 1)
            if pd.notna(precedente) and precedente > 0
            else float("nan")
        )

        a, b, c = st.columns(3)
        a.metric(
            f"Pernottamenti {ultimo}",
            numero(attuale),
        )
        b.metric(
            "Variazione annua",
            percentuale(variazione),
        )
        c.metric(
            "Territori NUTS2",
            numero(reg["nuts2_code"].nunique()),
        )

        grafico(px.bar(
            annuale,
            x="year",
            y="guest_nights_total",
            color_discrete_sequence=["#126982"],
            labels={
                "year": "Anno",
                "guest_nights_total": "Pernottamenti",
            },
        ))

    else:
        st.info(
            "Non risultano anni completi per tutto il panel."
        )

    anni = sorted(
        reg["year"].unique(), reverse=True
    )

    indice = (
        anni.index(max(anni_completi))
        if anni_completi else 0
    )

    anno = st.selectbox(
        "Anno della mappa",
        anni,
        index=indice,
    )

    st.caption(
        "Per gli anni parziali sono rappresentati "
        "i soli mesi disponibili."
    )

    d = aggrega(
        reg[reg["year"].eq(anno)],
        ["nuts2_code"],
        "guest_nights_total",
    )

    mappa(
        d, 2,
        "nuts2_code",
        "guest_nights_total",
        f"Pernottamenti regionali · {anno}",
        "nazionale",
    )


# ------------------------------------------------------------
# TERRITORIO
# ------------------------------------------------------------

elif pagina == "Esplora un territorio":
    if provincia != "Tutte":
        richiedi("prov")

        d = filtra_territorio(prov)
        asse = "year"

        st.subheader(
            str(nomi_prov.get(provincia, provincia))
        )
        st.caption(
            "Serie provinciale annuale. "
            "Nessuna mensilizzazione artificiale."
        )

    else:
        richiedi("reg")

        d = filtra_territorio(reg)
        asse = "data"

        st.subheader(
            "Italia" if regione == "Tutte"
            else str(nomi_reg[regione])
        )

    serie = aggrega(
        d, [asse], "guest_nights_total"
    )

    grafico(px.line(
        serie.sort_values(asse),
        x=asse,
        y="guest_nights_total",
        color_discrete_sequence=["#126982"],
        labels={
            asse: "Periodo",
            "guest_nights_total": "Pernottamenti",
        },
    ))

    if provincia == "Tutte":
        st.subheader("Stagionalità")

        stagioni = aggrega(
            d,
            ["year", "month"],
            "guest_nights_total",
        )
        stagioni["Anno"] = stagioni["year"].astype(str)

        grafico(px.line(
            stagioni.sort_values(["year", "month"]),
            x="month",
            y="guest_nights_total",
            color="Anno",
            labels={
                "month": "Mese",
                "guest_nights_total": "Pernottamenti",
            },
        ))

    tabella(d, "territorio.csv")


# ------------------------------------------------------------
# EVENTI E SIMULATORE
# ------------------------------------------------------------

elif pagina == "Eventi e simulatore":
    try:
        originale_screen = leggi_csv(FILES["screen"])
        screen = prepara_screening(originale_screen)

    except Exception as exc:
        st.error(
            "Impossibile usare lo screening completo: "
            + str(exc)
        )
        st.info(
            "Non viene applicato un riconoscimento alternativo "
            "dal master e nessuna riga viene eliminata "
            "per aggirare il problema."
        )
        st.stop()

    # Lo sci di fondo, se presente in futuri aggiornamenti,
    # resta consultabile nel catalogo ma fuori simulazione.
    escluso_perimetro = (
        screen["sport"].str.casefold().eq("sci di fondo")
    )

    base_simulatore = screen.loc[
        ~escluso_perimetro
    ].copy()

    a, b, c = st.columns(3)

    a.metric(
        "Righe territoriali",
        numero(screen["id_riga_territoriale"].nunique()),
    )
    b.metric(
        "Righe evento–regione–mese",
        numero(len(screen)),
    )
    c.metric(
        "Eventi identificati",
        numero(screen["evento_key"].nunique()),
    )

    st.caption(
        "Il simulatore usa lo screening completo. "
        "Basi, classificazioni e riparti sono quelli "
        "già elaborati nel progetto."
    )

    catalogo_tab, simulatore_tab, controllo_tab, contesto_tab, archivio_tab = st.tabs([
        "Catalogo",
        "Simulatore",
        "Dati utilizzati",
        "Contesto turistico",
        "Archivio",
    ])

    with catalogo_tab:
        try:
            p = FILES["master"]

            master = leggi_master_cache(
                str(p), p.stat().st_mtime
            )
            master.columns = [
                str(c).strip() for c in master.columns
            ]

            a, b = st.columns(2)

            ricerca = a.text_input(
                "Cerca evento",
                key="catalogo_ricerca",
            )

            sports = (
                sorted(master["Sport"].dropna().astype(str).unique())
                if "Sport" in master else []
            )

            sport = b.selectbox(
                "Sport",
                ["Tutti"] + sports,
                key="catalogo_sport",
            )

            applica_regione = st.checkbox(
                "Applica il filtro regionale al catalogo",
                value=False,
            )

            vista = master.copy()

            if sport != "Tutti":
                vista = vista[
                    vista["Sport"].astype(str).eq(sport)
                ]

            nome = (
                "Evento originale"
                if "Evento originale" in vista
                else "Evento"
            )

            if ricerca.strip() and nome in vista:
                vista = vista[
                    vista[nome].astype(str).str.contains(
                        ricerca.strip(),
                        case=False,
                        regex=False,
                        na=False,
                    )
                ]

            if (
                applica_regione
                and regione != "Tutte"
                and "Codici NUTS2" in vista
            ):
                vista = vista[
                    vista["Codici NUTS2"]
                    .astype(str)
                    .str.contains(
                        regione, regex=False, na=False
                    )
                ]

            st.caption(
                f"{len(vista)} righe mostrate su {len(master)}. "
                "Il catalogo conserva le colonne originali "
                "con fonti, note e qualità dei dati."
            )

            tabella(
                vista,
                "catalogo_event_master.csv",
                380,
            )

        except Exception as exc:
            st.warning(
                f"Catalogo Excel non disponibile: {exc}"
            )
            st.caption(
                "Il simulatore resta utilizzabile "
                "attraverso lo screening."
            )

    with simulatore_tab:
        st.write(
            "Modifica le ipotesi di pernottamento. "
            "Le tre basi rimangono distinte nei risultati."
        )

        st.caption(
            "La quota è applicata alla base dell'evento. "
            "Non viene aggiunto un secondo moltiplicatore "
            "automatico per la provenienza."
        )

        d = base_simulatore.copy()

        if regione != "Tutte":
            d = d[d["nuts2"].eq(regione)]

        st.caption(
            "Territorio: "
            + (
                "Italia"
                if regione == "Tutte"
                else str(nomi_reg[regione])
            )
            + ". Il filtro provinciale non modifica questa analisi."
        )

        with st.expander("Filtri", expanded=False):
            sport_scelti = filtro_tutti(
                "Sport",
                sorted(d["sport"].unique()),
                "sim_sport",
            )

            d = d[d["sport"].isin(sport_scelti)]

            anni_scelti = filtro_tutti(
                "Anni",
                sorted(d["anno"].unique()),
                "sim_anni",
            )

            mesi_scelti = filtro_tutti(
                "Mesi",
                list(range(1, 13)),
                "sim_mesi",
            )

            d = d[
                d["anno"].isin(anni_scelti)
                & d["mese"].isin(mesi_scelti)
            ]

            elenco = d.drop_duplicates(
                "evento_key"
            ).set_index("evento_key")

            etichette = {
                k: (
                    f"{r['evento']} · "
                    f"{r['inizio']:%d/%m/%Y}"
                )
                for k, r in elenco.iterrows()
            }

            eventi_scelti = filtro_tutti(
                "Eventi",
                elenco.index.tolist(),
                "sim_eventi",
                formato=lambda k: etichette[k],
            )

            d = d[d["evento_key"].isin(eventi_scelti)]

            includi_biglietti = st.checkbox(
                "Includi gli equivalenti da biglietti",
                value=True,
                help=(
                    "Sono scenari costruiti da biglietti pubblicati "
                    "o stimati, non persone uniche osservate."
                ),
            )

            if not includi_biglietti:
                d = d[
                    d["tipo_base_visuale"].ne(
                        "Equivalenti da biglietti"
                    )
                ]

        if d.empty:
            st.info("Nessun evento per i filtri selezionati.")

        else:
            sports = sorted(d["sport"].unique())

            a, b = st.columns([2, 1])

            preset = a.selectbox(
                "Preset comune",
                [5, 15, 25, 40, 60],
                index=2,
                format_func=lambda q: f"{q}%",
            )

            if b.button("Applica agli sport"):
                for s in sports:
                    st.session_state["quota_" + s] = int(preset)

            quote = {}
            notti = {}

            for s in sports:
                massimo = max(
                    1,
                    min(
                        5,
                        int(
                            d.loc[
                                d["sport"].eq(s),
                                "durata_giorni",
                            ].max()
                        ),
                    ),
                )

                with st.expander(
                    f"{s} · quota e notti",
                    expanded=len(sports) <= 2,
                ):
                    c1, c2 = st.columns(2)

                    quote[s] = c1.slider(
                        "Quota pernottanti (%)",
                        0, 100, 25,
                        key="quota_" + s,
                    )

                    if massimo == 1:
                        notti[s] = 1
                        c2.metric(
                            "Notti per pernottante", 1
                        )

                    else:
                        key_notti = "notti_" + s

                        if key_notti in st.session_state:
                            st.session_state[key_notti] = min(
                                massimo,
                                max(
                                    1,
                                    int(st.session_state[key_notti]),
                                ),
                            )

                        notti[s] = c2.slider(
                            "Notti per pernottante",
                            1, massimo, 1,
                            key=key_notti,
                        )

                    st.caption(
                        "0% = nessun pernottante; 100% = tutta "
                        "la base considerata. Le notti sono limitate "
                        "alla durata del singolo evento e a un massimo di 5."
                    )

            risultato = calcola(d, quote, notti)
            mensile = riassumi_simulazione(risultato)

            metriche = st.columns(3)

            for col, tipo in zip(
                metriche, COLORI_BASI.keys()
            ):
                totale = risultato.loc[
                    risultato["tipo_base_visuale"].eq(tipo),
                    "pernottamenti_simulati",
                ].sum()

                col.metric(
                    tipo,
                    numero(totale),
                    help="Pernottamenti simulati",
                )

            st.caption(
                "Pernottamenti simulati. "
                "Gli equivalenti non sono persone uniche. "
                "Le tre basi non vengono presentate come un unico totale."
            )

            a, b, c = st.columns(3)

            a.metric(
                "Eventi selezionati",
                numero(risultato["evento_key"].nunique()),
            )

            b.metric(
                "Righe mensili selezionate",
                numero(len(risultato)),
            )

            c.metric(
                "Righe con rapporto ≥ 2%",
                numero(
                    risultato["rapporto_pct"].ge(2).sum()
                ),
            )

            st.caption(
                "Il conteggio sopra soglia riguarda righe "
                "evento–regione–mese, non eventi unici."
            )

            mancanti = risultato[
                risultato["gn_eurostat_mese"].isna()
                | risultato["gn_eurostat_mese"].le(0)
            ]

            if not mancanti.empty:
                st.info(
                    f"{len(mancanti)} righe senza denominatore "
                    "Eurostat positivo: il rapporto è n.d.; "
                    "la simulazione resta disponibile."
                )

            andamento = (
                risultato.groupby(
                    ["data", "tipo_base_visuale"]
                )["pernottamenti_simulati"]
                .sum()
                .reset_index()
            )

            grafico(px.bar(
                andamento,
                x="data",
                y="pernottamenti_simulati",
                color="tipo_base_visuale",
                barmode="group",
                color_discrete_map=COLORI_BASI,
                labels={
                    "data": "Mese",
                    "pernottamenti_simulati": "Pernottamenti simulati",
                    "tipo_base_visuale": "Base",
                },
            ))

            st.subheader("Confronto con la domanda regionale")

            tabella(
                mensile,
                "simulazione_regioni_mesi.csv",
            )

            with st.expander("Cinque scenari a confronto"):
                confronti = []

                for q in [5, 15, 25, 40, 60]:
                    scenario = calcola(
                        d,
                        {s: q for s in sports},
                        notti,
                    )

                    for tipo, gruppo in scenario.groupby(
                        "tipo_base_visuale"
                    ):
                        confronti.append({
                            "Quota pernottanti": f"{q}%",
                            "Base": tipo,
                            "Pernottamenti simulati": gruppo[
                                "pernottamenti_simulati"
                            ].sum(),
                            "Righe con rapporto almeno 2%": gruppo[
                                "rapporto_pct"
                            ].ge(2).sum(),
                        })

                confronto = pd.DataFrame(confronti)

                st.caption(
                    "Variano soltanto le quote pernottanti. "
                    "Le notti restano quelle impostate per sport."
                )

                grafico(px.bar(
                    confronto,
                    x="Quota pernottanti",
                    y="Pernottamenti simulati",
                    color="Base",
                    barmode="group",
                    color_discrete_map=COLORI_BASI,
                ))

                tabella(
                    confronto,
                    "simulazione_cinque_scenari.csv",
                )

            with st.expander("Dettaglio e basi utilizzate"):
                campi = [
                    "id_riga_territoriale",
                    "evento", "sport", "Regione",
                    "anno", "mese",
                    "tipo_base_visuale",
                    "classe_base", "metodo_base",
                    "pubblico_originale",
                    "pubblico_equivalente",
                    "durata_giorni", "quota_mese",
                    "quota_pct", "notti_scenario",
                    "pernottamenti_simulati",
                    "gn_eurostat_mese", "rapporto_pct",
                ]

                tabella(
                    risultato[campi],
                    "simulazione_dettaglio.csv",
                    380,
                )

            st.caption(
                "I ricalcoli usano le basi e le quote mensili salvate "
                "nel CSV. Possono differire di pochi centesimi dagli "
                "output precedenti, calcolati prima dell'arrotondamento."
            )

    with controllo_tab:
        st.success(
            "Schema, chiavi, date, durata e ripartizione mensile "
            "dello screening hanno superato i controlli dell'app."
        )

        riepilogo = (
            screen.groupby(
                ["tipo_base_visuale", "classe_base"],
                dropna=False,
            )
            .agg(
                Righe_mensili=("evento", "size"),
                Righe_territoriali=(
                    "id_riga_territoriale", "nunique"
                ),
                Notti_scenario_medio_salvate=(
                    "pernottamenti_medio", "sum"
                ),
            )
            .reset_index()
        )

        tabella(
            riepilogo,
            "riepilogo_screening.csv",
        )

        st.caption(
            "I risultati salvati sono quelli presenti nel file "
            "e non cambiano con i cursori del simulatore."
        )

        biglietti = screen[
            screen["tipo_base_visuale"].eq(
                "Equivalenti da biglietti"
            )
        ]

        with st.expander(
            "Eventi con base equivalente da biglietti"
        ):
            st.write(
                "Queste righe sono mantenute, con un'etichetta "
                "specifica. Il rapporto biglietti/durata è una "
                "convenzione di scenario: non identifica "
                "visitatori unici."
            )

            tabella(
                biglietti[[
                    "evento", "sport", "Regione",
                    "anno", "mese",
                    "pubblico_originale",
                    "pubblico_equivalente",
                    "durata_giorni",
                    "metodo_base",
                    "pernottamenti_medio",
                ]],
                "basi_da_biglietti.csv",
            )

        if not reg.empty:
            confronto = screen[
                ["nuts2", "anno", "mese", "gn_eurostat_mese"]
            ].drop_duplicates()

            confronto = confronto.merge(
                reg[[
                    "nuts2_code", "year", "month",
                    "guest_nights_total",
                ]],
                left_on=["nuts2", "anno", "mese"],
                right_on=["nuts2_code", "year", "month"],
                how="left",
                validate="one_to_one",
                indicator=True,
            )

            confronto["Scarto denominatore"] = (
                confronto["gn_eurostat_mese"]
                - confronto["guest_nights_total"]
            )

            diversa_disponibilita = (
                confronto["gn_eurostat_mese"].isna()
                != confronto["guest_nights_total"].isna()
            )

            differenze = confronto[
                confronto["_merge"].ne("both")
                | diversa_disponibilita
                | confronto["Scarto denominatore"].abs().gt(0.01)
            ]

            with st.expander(
                "Confronto denominatori con il master Eurostat"
            ):
                st.caption(
                    "Il simulatore mantiene il denominatore dello "
                    "screening per riprodurne il perimetro. "
                    "Eventuali revisioni di Eurostat vengono segnalate."
                )

                if differenze.empty:
                    st.success(
                        "Nessuna differenza rilevata "
                        "oltre la tolleranza di 0,01."
                    )
                else:
                    tabella(
                        differenze,
                        "differenze_denominatori.csv",
                    )

        with st.expander("Screening completo utilizzato"):
            tabella(
                screen,
                "screening_utilizzato.csv",
                380,
            )

    with contesto_tab:
        if reg.empty:
            st.info("Dati regionali non disponibili.")

        else:
            codici = sorted(
                nomi_reg,
                key=lambda k: str(nomi_reg[k]),
            )

            scelta = st.selectbox(
                "Regione del confronto",
                codici,
                index=(
                    codici.index(regione)
                    if regione in codici else 0
                ),
                format_func=lambda k: str(nomi_reg[k]),
            )

            serie = reg[
                reg["nuts2_code"].eq(scelta)
            ].copy()

            precedente = serie[[
                "year", "month", "guest_nights_total"
            ]].copy()

            precedente["year"] += 1

            precedente = precedente.rename(columns={
                "guest_nights_total": "Stesso mese anno precedente"
            })

            serie = serie.merge(
                precedente,
                on=["year", "month"],
                how="left",
                validate="one_to_one",
            )

            denominatore = serie[
                "Stesso mese anno precedente"
            ].where(
                serie["Stesso mese anno precedente"] > 0
            )

            serie["Variazione percentuale"] = (
                100 * (
                    serie["guest_nights_total"]
                    / denominatore - 1
                )
            )

            grafico(px.line(
                serie.sort_values("data"),
                x="data",
                y="guest_nights_total",
                color_discrete_sequence=["#126982"],
                labels={
                    "data": "Mese",
                    "guest_nights_total": "Pernottamenti",
                },
            ))

            st.caption(
                "Confronto descrittivo: non separa l'effetto "
                "degli eventi da stagionalità, festività, prezzi "
                "e altri fattori."
            )

            tabella(
                serie,
                "contesto_regionale.csv",
            )

    with archivio_tab:
        disponibili = [
            n for n in ARCHIVIO
            if percorso(n).is_file()
        ]

        if disponibili:
            scelta = st.selectbox(
                "Elaborazione salvata",
                disponibili,
            )

            st.info(
                "I file storici non sono necessariamente "
                "intercambiabili: alcuni usano ingressi cumulativi "
                "direttamente, altri basi equivalenti. "
                "Il simulatore usa soltanto screening_eventi_completo.csv."
            )

            try:
                tabella(
                    leggi_csv(percorso(scelta)),
                    "archivio_" + scelta,
                    380,
                )
            except Exception as exc:
                st.error(str(exc))


# ------------------------------------------------------------
# FORECAST
# ------------------------------------------------------------

elif pagina == "Forecast 2026–2030":
    provinciale = provincia != "Tutte"
    chiave = "fp" if provinciale else "fr"

    richiedi(chiave)

    d = filtra_territorio(DATI[chiave])

    prefisso = (
        "forecast_" if provinciale else "gn_"
    )
    codice = (
        "nuts3_code" if provinciale else "nuts2_code"
    )

    st.caption(
        "Scenari letti dai CSV del progetto. "
        "La simulazione degli eventi non modifica il forecast."
    )

    scelti = st.multiselect(
        "Scenari",
        SCENARI,
        default=SCENARI,
    )

    if not scelti or d.empty:
        st.info(
            "Seleziona uno scenario con dati disponibili."
        )

    else:
        righe = []

        for scenario in scelti:
            campo = prefisso + scenario.lower()

            annuale = aggrega(
                d, ["year"], campo
            ).rename(columns={
                campo: "Pernottamenti"
            })

            annuale["Scenario"] = scenario
            righe.append(annuale)

        lungo = pd.concat(righe, ignore_index=True)

        grafico(px.line(
            lungo.sort_values("year"),
            x="year",
            y="Pernottamenti",
            color="Scenario",
            markers=True,
            color_discrete_map=COLORI,
            labels={"year": "Anno"},
        ))

        a, b = st.columns(2)

        anno = a.selectbox(
            "Anno della mappa",
            sorted(d["year"].unique(), reverse=True),
        )
        scenario = b.selectbox(
            "Scenario della mappa",
            scelti,
        )

        campo = prefisso + scenario.lower()

        geografici = aggrega(
            d[d["year"].eq(anno)],
            [codice],
            campo,
        )

        mappa(
            geografici,
            3 if provinciale else 2,
            codice,
            campo,
            f"Forecast {scenario.lower()} · {anno}",
            "forecast",
        )

        tabella(
            lungo,
            "forecast_annuale.csv",
        )


# ------------------------------------------------------------
# DATI E METODO
# ------------------------------------------------------------

else:
    fonti_tab, eventi_tab, forecast_tab, qualita_tab = st.tabs([
        "Fonti e lettura",
        "Metodo eventi",
        "Forecast",
        "Qualità e limiti",
    ])

    with fonti_tab:
        st.markdown(
            """
            ### Tre livelli distinti

            **Osservato:** pernottamenti rilevati nell'ambito
            coperto dalla fonte Eurostat.

            **Simulato:** pernottamenti teorici ottenuti applicando
            ipotesi alle basi degli eventi.

            **Previsto:** scenari di forecast già elaborati.
            I risultati del simulatore non vengono aggiunti
            automaticamente alle previsioni.

            ### Fonti

            `tour_ce_omn12` è il riferimento mensile NUTS2
            del progetto.

            `tour_ce_oan3` è il riferimento annuale NUTS3.

            Tre ospiti per due notti generano sei pernottamenti.
            Questa misura è diversa da prenotazioni, persone uniche
            e notti di occupazione dell'alloggio.

            La copertura della fonte non coincide automaticamente
            con tutto il turismo regionale o con il solo Airbnb.

            ### Ruolo dei file

            Il **master Excel** alimenta il catalogo:
            eventi, fonti, note e qualità.

            **`screening_eventi_completo.csv`** alimenta il simulatore:
            basi già elaborate, territori, mesi e denominatori.

            La dashboard non interpreta nuovamente le colonne
            del master per ricostruire le basi.

            [Metadati Eurostat](https://ec.europa.eu/eurostat/cache/metadata/en/tour_ce_esms.htm)
            """
        )

    with eventi_tab:
        st.markdown(
            """
            ### Basi di calcolo

            **Persone:** la base disponibile è trattata come persone,
            mantenendo il metodo documentato nel file.

            **Equivalenti da presenze:** la base proxy è già stata
            elaborata nello screening. Non viene divisa nuovamente
            per la durata.

            **Equivalenti da biglietti:** scenario basato su biglietti
            pubblicati o stimati. Non identifica visitatori unici
            e non dimostra l'effettiva partecipazione.

            I tre gruppi sono mostrati separatamente.

            ### Territorio e mesi

            I valori sono già territorializzati.
            La dashboard non applica nuovi riparti tra regioni.

            Gli eventi sono già distribuiti sulle righe mensili.
            La quota del mese viene applicata una sola volta.

            La ripartizione in proporzione ai giorni è una
            convenzione uniforme, non una distribuzione osservata
            dei pernottamenti.

            ### Quote e notti

            La quota pernottanti varia da 0 a 100%, con default 25%.
            I preset sono 5%, 15%, 25%, 40% e 60%.

            Il parametro si applica alla base disponibile.
            Se la base include tutti i partecipanti, incorpora
            implicitamente anche l'ipotesi sulla provenienza.
            Non è automaticamente una quota riferita ai soli
            non residenti.

            Non sono imposti minimi per tifosi ospiti o trasferte.

            Le notti partono da una e sono limitate al minimo
            tra la durata del singolo evento e cinque.
            Non includono soggiorni prima o dopo l'evento.
            """
        )

        st.latex(
            r"N_{e,r,m}=B_{e,r}\cdot q_e\cdot n_e\cdot w_{e,m}"
        )

        st.markdown(
            """
            **B** è `pubblico_equivalente`, già elaborato.
            **q** è la quota pernottanti divisa per 100.
            **n** sono le notti applicate.
            **w** è `quota_mese`.

            Per le righe classificate PERSONE, la colonna
            `pubblico_equivalente` contiene la base persone:
            il nome tecnico della colonna non cambia la classe.

            ### Rapporto con la domanda regionale
            """
        )

        st.latex(
            r"R_{e,r,m}=100\cdot\frac{N_{e,r,m}}{GN_{r,m}}"
        )

        st.markdown(
            """
            Il denominatore è quello salvato nello screening,
            relativo alla stessa regione e allo stesso mese.

            Se manca o non è positivo, il rapporto resta
            non disponibile.

            Il 2% è una soglia esplorativa, non una soglia
            di significatività statistica.

            Il rapporto indica una dimensione relativa:
            non misura il turismo causato dall'evento.

            Il numeratore può comprendere pernottamenti in alloggi
            fuori dal perimetro Eurostat. In quel caso il confronto
            riguarda perimetri diversi.

            Le somme tra eventi possono comprendere persone
            o soggiorni ripetuti: non sono un conteggio
            di turisti unici aggiuntivi.
            """
        )

    with forecast_tab:
        st.markdown(
            """
            ### Scenari letti dai file del progetto

            L'app legge i forecast regionali e provinciali.
            Non riaddestra i modelli e non ricostruisce
            i loro parametri dai risultati finali.

            Prudente, Intermedio ed Espansivo sono scenari.
            Non sono automaticamente intervalli di confidenza.

            La documentazione del modello a monte deve indicare
            periodo di stima, trattamento del COVID, stagionalità,
            anno base, crescita e riparto territoriale.

            ### Validazione

            Confrontare i modelli sugli stessi periodi di test,
            con una baseline stagionale semplice e senza usare
            informazioni future nella stima.

            WAPE e bias misurano rispettivamente la dimensione
            relativa dell'errore e la tendenza a sovrastimare
            o sottostimare.

            I risultati delle verifiche precedenti non vengono
            presentati come ricalcolati da questa dashboard.

            ### Province

            Se le previsioni provinciali derivano da quote costanti
            della regione, non prevedono autonomamente
            cambiamenti nella distribuzione territoriale.
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

    with qualita_tab:
        controlli = []

        for k in SCHEMI:
            d = DATI[k]

            campi = [
                c for c in SCHEMI[k]
                if c.startswith(
                    ("guest_nights", "gn_", "forecast_")
                )
            ]

            controlli.append({
                "File": FILES[k].name,
                "Stato": ERRORI.get(
                    k, "Chiavi valide e non duplicate"
                ),
                "Righe": len(d),
                "Valori mancanti": (
                    int(d[campi].isna().sum().sum())
                    if not d.empty else None
                ),
                "Valori negativi": (
                    int(d[campi].lt(0).sum().sum())
                    if not d.empty else None
                ),
            })

        tabella(
            pd.DataFrame(controlli),
            "controlli_file.csv",
        )

        if not reg.empty:
            copertura = (
                reg.groupby(["nuts2_code", "year"])
                .agg(
                    Mesi_presenti=("month", "nunique"),
                    Mesi_valorizzati=(
                        "guest_nights_total", "count"
                    ),
                )
                .reset_index()
            )

            with st.expander("Copertura regionale"):
                tabella(
                    copertura,
                    "copertura_regionale.csv",
                )

        if not fr.empty and not fp.empty:
            confronti = []

            for s in SCENARI:
                a = aggrega(
                    fr,
                    ["nuts2_code", "year"],
                    "gn_" + s.lower(),
                ).rename(columns={
                    "gn_" + s.lower(): "Regione"
                })

                b = aggrega(
                    fp,
                    ["nuts2_code", "year"],
                    "forecast_" + s.lower(),
                ).rename(columns={
                    "forecast_" + s.lower(): "Somma province"
                })

                c = a.merge(
                    b,
                    on=["nuts2_code", "year"],
                    how="outer",
                    validate="one_to_one",
                )

                c["Scenario"] = s
                c["Scarto"] = (
                    c["Somma province"] - c["Regione"]
                )
                confronti.append(c)

            with st.expander("Riconciliazione forecast"):
                tabella(
                    pd.concat(
                        confronti, ignore_index=True
                    ),
                    "riconciliazione_forecast.csv",
                )

        st.markdown(
            """
            ### Controlli e limiti

            I valori mancanti non vengono trasformati in zero.

            Gli anni parziali non sono presentati come anni
            completi nel riepilogo nazionale.

            Le basi dello screening non vengono riclassificate
            automaticamente dall'app.

            Un controllo aritmetico superato non certifica
            la validità causale o la precisione delle fonti.

            Stagionalità, festività, prezzi, trasporti e altri eventi
            possono influenzare la domanda turistica.

            Un evento può anche spostare domanda nel tempo
            o sostituire altri visitatori.

            Per stimare un effetto causale servono un controfattuale
            credibile, territori comparabili, controlli per altri
            fattori e verifica delle tendenze precedenti.
            """
        )


st.divider()
st.caption(
    f"Airbnb&Sports · {VERSIONE} · "
    f"[SBL Consultancy]({SITO})"
)
