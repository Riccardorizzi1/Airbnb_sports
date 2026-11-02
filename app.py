# AIRBNB&SPORTS — app.py
# Avvio locale: python -m streamlit run app.py

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
VERSIONE = "v5.1 · Screening verificato"
SITO = "https://www.sblconsultancy.it/"

PAGINE = [
    "Panorama nazionale",
    "Esplora un territorio",
    "Eventi e simulatore",
    "Forecast 2026–2030",
    "Dati e metodo",
]

SCENARI = ["Prudente", "Intermedio", "Espansivo"]

COLORI_SCENARI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C",
}

COLORI_BASI = {
    "Persone": "#6DDDEB",
    "Equivalenti da presenze": "#DC813F",
    "Equivalenti da biglietti": "#C94F7C",
}

TESTO = "#F3FAFC"
SECONDARIO = "#D3E5EC"
BORDO = "#315462"
CARD = "#102D3B"
ACCENTO = "#6DDDEB"

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
        "forecast_prudente", "forecast_intermedio", "forecast_espansivo",
    ],
}

CHIAVI = {
    "reg": ["nuts2_code", "year", "month"],
    "prov": ["nuts3_code", "year"],
    "fr": ["nuts2_code", "year", "month"],
    "fp": ["nuts3_code", "year"],
}


def percorso(nome):
    for cartella in (ROOT, ROOT / "output_eurostat"):
        candidato = cartella / nome
        if candidato.is_file():
            return candidato
    return ROOT / nome


FILES = {
    "reg": percorso("eurostat_master_final.csv"),
    "prov": percorso("eurostat_nuts3_annuale_2018_2025.csv"),
    "fr": percorso("forecast_nuts2_2026_2030_3_scenari.csv"),
    "fp": percorso("forecast_nuts3_2026_2030_3_scenari.csv"),
    "master": percorso("AirbnbSports_Event_Master_Simone_Visitatori.xlsx"),
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
# 2. TEMA SCURO SBL
# ============================================================

st.markdown(
    """
    <style>
    :root { color-scheme: dark; }

    .stApp {
        background: linear-gradient(
            100deg, #071822 0%, #0C2B39 47%, #10475B 100%
        );
        color: #F3FAFC;
    }

    header[data-testid="stHeader"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
        height: 0 !important;
        min-height: 0 !important;
        overflow: visible !important;
    }

    [data-testid="stDecoration"], [data-testid="stToolbar"],
    [data-testid="stStatusWidget"], .stAppDeployButton, #MainMenu {
        display: none !important;
    }

    [data-testid="stSidebarCollapsedControl"],
    [data-testid="collapsedControl"],
    [data-testid="stExpandSidebarButton"] {
        visibility: visible !important;
        opacity: 1 !important;
        pointer-events: auto !important;
        position: fixed !important;
        left: 8px !important;
        top: 8px !important;
        z-index: 100000 !important;
        color: #F3FAFC !important;
        background: #102D3B !important;
    }

    .block-container {
        max-width: 1320px;
        padding: 1.5rem 1.6rem 2rem;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(
            180deg, #071F31 0%, #0E4058 52%, #126982 100%
        );
        min-width: 255px;
        max-width: 280px;
    }

    [data-testid="stMain"] :is(h1,h2,h3,h4,p,label,li),
    [data-testid="stSidebar"] :is(h1,h2,h3,h4,p,label,li) {
        color: #F3FAFC;
    }

    h1 { font-size: 1.8rem !important; }
    h2 { font-size: 1.3rem !important; }
    h3 { font-size: 1.08rem !important; }
    p, label { font-size: 0.92rem; }

    [data-testid="stCaptionContainer"] p {
        color: #D3E5EC !important;
        font-size: 0.83rem !important;
    }

    [data-testid="stMain"] a, [data-testid="stSidebar"] a {
        color: #6DDDEB;
    }

    [data-testid="stSelectbox"] [data-baseweb="select"] > div,
    [data-testid="stMultiSelect"] [data-baseweb="select"] > div,
    [data-baseweb="input"], [data-baseweb="base-input"] {
        background: #102D3B !important;
        border-color: #315462 !important;
        min-height: 36px !important;
        border-radius: 8px !important;
    }

    [data-baseweb="select"] *, [data-baseweb="input"] *,
    [data-baseweb="popover"] *, [role="listbox"] *, [role="option"] * {
        color: #F3FAFC !important;
        -webkit-text-fill-color: #F3FAFC !important;
    }

    [data-baseweb="popover"] [data-baseweb="menu"],
    [role="listbox"], [role="option"] {
        background: #102D3B !important;
        color: #F3FAFC !important;
    }

    [role="option"]:hover, [role="option"][aria-selected="true"] {
        background: #17627E !important;
    }

    [data-baseweb="tag"] {
        background: #17627E !important;
        border-radius: 5px !important;
    }

    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input {
        background: #102D3B !important;
        color: #F3FAFC !important;
        -webkit-text-fill-color: #F3FAFC !important;
    }

    [data-testid="stMetric"] {
        background: #102D3B;
        border: 1px solid #315462;
        border-radius: 12px;
        padding: 11px 13px;
    }

    [data-testid="stMetricValue"] {
        color: #F3FAFC !important;
        font-size: 1.55rem !important;
    }

    [data-testid="stMetricLabel"] p { font-size: 0.81rem !important; }

    [data-testid="stVerticalBlock"] { gap: 0.8rem; }

    [data-testid="stExpander"] details,
    [data-testid="stDataFrame"] {
        border-color: #315462 !important;
    }

    [data-baseweb="tab"] { color: #D3E5EC !important; }
    [data-baseweb="tab"][aria-selected="true"] { color: #6DDDEB !important; }
    [data-baseweb="tab-highlight"] { background: #6DDDEB !important; }

    .stButton > button, .stDownloadButton > button {
        background: #102D3B !important;
        color: #F3FAFC !important;
        border: 1px solid #315462 !important;
        border-radius: 8px;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        border-color: #6DDDEB !important;
    }

    .stButton > button p, .stDownloadButton > button p {
        color: #F3FAFC !important;
        font-size: 0.84rem !important;
    }

    @media (max-width: 768px) {
        .block-container {
            padding: 3rem 1rem 2rem;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 3. LETTURA E PRESENTAZIONE
# ============================================================

@st.cache_data(show_spinner=False)
def leggi_csv_cache(nome, modifica):
    dati = pd.read_csv(nome, encoding="utf-8-sig")
    if len(dati.columns) == 1 and ";" in str(dati.columns[0]):
        dati = pd.read_csv(nome, encoding="utf-8-sig", sep=";")
    dati.columns = [str(c).strip() for c in dati.columns]
    return dati


def leggi_csv(file):
    if not file.is_file():
        raise FileNotFoundError(f"File non trovato: {file.name}")
    return leggi_csv_cache(str(file), file.stat().st_mtime_ns)


@st.cache_data(show_spinner=False)
def leggi_master(nome, modifica):
    dati = pd.read_excel(nome, sheet_name="Eventi")
    dati.columns = [str(c).strip() for c in dati.columns]
    return dati


def numero(valore, decimali=0):
    if pd.isna(valore):
        return "n.d."
    return (
        f"{float(valore):,.{decimali}f}"
        .replace(",", "§").replace(".", ",").replace("§", ".")
    )


def percentuale(valore):
    return "n.d." if pd.isna(valore) else f"{numero(valore, 1)}%"


def data_singola(valore):
    if pd.isna(valore):
        return pd.NaT
    if isinstance(valore, (int, float)):
        if 20000 <= valore <= 80000:
            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(valore, unit="D")
            ).normalize()
        return pd.NaT

    testo = str(valore).strip()
    iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", testo))
    giorno = pd.to_datetime(testo, errors="coerce", dayfirst=not iso)
    if pd.isna(giorno):
        return pd.NaT
    if giorno.tzinfo is not None:
        giorno = giorno.tz_localize(None)
    return giorno.normalize()


def somma_completa(serie):
    if serie.empty or serie.isna().any():
        return float("nan")
    return serie.sum()


def aggrega(dati, gruppi, campo):
    return (
        dati.groupby(gruppi, dropna=False)[campo]
        .agg(somma_completa)
        .reset_index()
    )


def tabella_pulita(dati):
    tab = dati.copy()
    nomi_date = {
        "data", "inizio", "fine", "data inizio", "data fine",
        "data_inizio", "data_fine", "start_date", "end_date",
    }

    for col in tab:
        if pd.api.types.is_datetime64_any_dtype(tab[col]):
            tab[col] = tab[col].dt.strftime("%d/%m/%Y")
        elif tab[col].dtype == object:
            tab[col] = tab[col].map(
                lambda v: pd.NA
                if isinstance(v, str)
                and v.strip().lower() in {"", "nan", "none", "nat", "null"}
                else v
            )
            if str(col).lower().strip() in nomi_date:
                date = pd.to_datetime(tab[col].map(data_singola))
                tab[col] = tab[col].where(
                    date.isna(), date.dt.strftime("%d/%m/%Y")
                )

    return tab.dropna(axis=1, how="all")


def tabella(dati, nome, altezza=340):
    if dati.empty:
        st.info("Nessuna riga per questa selezione.")
        return

    tab = tabella_pulita(dati)
    st.dataframe(
        tab, hide_index=True, use_container_width=True,
        height=min(altezza, 40 + 35 * len(tab)),
    )
    st.download_button(
        "Scarica CSV",
        tab.to_csv(index=False, sep=";").encode("utf-8-sig"),
        file_name=nome,
        mime="text/csv",
        key="download_" + nome,
    )


def grafico(fig, altezza=340):
    fig.update_layout(
        template="plotly_dark",
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=TESTO, size=12),
        hoverlabel=dict(bgcolor=CARD, font_color=TESTO),
        margin=dict(l=12, r=12, t=35, b=20),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=False, automargin=True, color=SECONDARIO)
    fig.update_yaxes(gridcolor=BORDO, automargin=True, color=SECONDARIO)
    st.plotly_chart(fig, use_container_width=True, theme=None)


def filtro_tutti(label, opzioni, key, formato=None):
    opzioni = list(opzioni)
    if st.checkbox(f"{label}: tutti", value=True, key=key + "_tutti"):
        return opzioni
    kwargs = {"format_func": formato} if formato else {}
    return st.multiselect(label, opzioni, default=[], key=key, **kwargs)


def mostra_logo():
    formati = {
        ".png": "image/png", ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg", ".webp": "image/webp",
        ".svg": "image/svg+xml",
    }

    for cartella in (
        ROOT, ROOT / "assets", ROOT / "images",
        ROOT / "img", ROOT / "output_eurostat",
    ):
        if not cartella.is_dir():
            continue
        for file in sorted(cartella.iterdir()):
            if (
                file.is_file()
                and file.stem.lower() == "logo_sbl"
                and file.suffix.lower() in formati
            ):
                dati = base64.b64encode(file.read_bytes()).decode()
                mime = formati[file.suffix.lower()]
                st.markdown(
                    f'<a href="{SITO}" target="_blank" rel="noopener noreferrer">'
                    f'<img src="data:{mime};base64,{dati}" alt="SBL Consultancy" '
                    'style="display:block;width:108px;max-width:100%;height:auto;'
                    'margin-bottom:12px;"></a>',
                    unsafe_allow_html=True,
                )
                return

    st.markdown(f"[**SBL CONSULTANCY**]({SITO})")


with st.sidebar:
    mostra_logo()
    st.markdown("### Airbnb&Sports")
    st.caption("Tourism Demand · Italia")
    pagina = st.radio("Sezione", PAGINE)
    st.caption(VERSIONE)


# ============================================================
# 4. DATI TERRITORIALI
# ============================================================

def prepara_territorio(chiave):
    dati = leggi_csv(FILES[chiave])
    mancanti = [c for c in SCHEMI[chiave] if c not in dati]
    if mancanti:
        raise ValueError("Colonne mancanti: " + ", ".join(mancanti))

    for col in dati:
        if col in ("nuts2_code", "nuts3_code"):
            dati[col] = dati[col].astype("string").str.strip().str.upper()
        elif (
            col in ("year", "month")
            or col.startswith(("guest_nights", "gn_", "forecast_"))
        ):
            dati[col] = pd.to_numeric(dati[col], errors="raise").replace(
                [float("inf"), -float("inf")], float("nan")
            )

    chiavi = CHIAVI[chiave]
    if dati[chiavi].isna().any().any():
        raise ValueError("Chiavi territoriali o temporali mancanti.")
    if dati.duplicated(chiavi).any():
        raise ValueError("Chiavi territoriali o temporali duplicate.")
    if dati["year"].mod(1).ne(0).any():
        raise ValueError("Anno non intero.")

    dati["year"] = dati["year"].astype(int)

    if "month" in dati:
        if (
            ~dati["month"].between(1, 12)
            | dati["month"].mod(1).ne(0)
        ).any():
            raise ValueError("Mesi non validi.")

        dati["month"] = dati["month"].astype(int)
        dati["data"] = pd.to_datetime(
            dict(year=dati["year"], month=dati["month"], day=1),
            errors="raise",
        )

    return dati


DATI = {}
ERRORI = {}

for chiave in SCHEMI:
    try:
        DATI[chiave] = prepara_territorio(chiave)
    except Exception as exc:
        DATI[chiave] = pd.DataFrame()
        ERRORI[chiave] = str(exc)

reg, prov, fr, fp = (DATI[k] for k in ("reg", "prov", "fr", "fp"))

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

    disponibili = prov
    if regione != "Tutte" and not prov.empty:
        disponibili = prov[prov["nuts2_code"].eq(regione)]

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
            "Tutte le province" if k == "Tutte"
            else str(nomi_prov.get(k, k))
        ),
    )

    st.caption(
        "Gli eventi sono confrontati con regione e mese. "
        "Il filtro provinciale riguarda le serie annuali."
    )


def filtra_territorio(dati):
    vista = dati.copy()
    if regione != "Tutte" and "nuts2_code" in vista:
        vista = vista[vista["nuts2_code"].eq(regione)]
    if provincia != "Tutte" and "nuts3_code" in vista:
        vista = vista[vista["nuts3_code"].eq(provincia)]
    return vista


# ============================================================
# 5. SIMULATORE BASATO SULLO SCREENING VERIFICATO
# ============================================================

def prepara_screening(originale):
    dati = originale.copy()

    chiavi = ["id_riga_territoriale", "nuts2", "anno", "mese"]
    testi = ["evento", "sport", "nuts2", "classe_base", "metodo_base"]
    numeriche = [
        "anno", "mese", "durata_giorni", "giorni_nel_mese", "quota_mese",
        "pubblico_originale", "pubblico_equivalente", "notti_applicate",
        "gn_eurostat_mese", "pct_medio", "pernottamenti_medio",
    ]
    richieste = set(chiavi + testi + numeriche + ["inizio", "fine"])
    mancanti = sorted(richieste - set(dati.columns))

    if mancanti:
        raise ValueError("Colonne dello screening mancanti: " + ", ".join(mancanti))
    if dati.empty:
        raise ValueError("Lo screening completo è vuoto.")
    if dati[chiavi].isna().any().any():
        raise ValueError("Chiavi mancanti nello screening.")

    for col in testi:
        if dati[col].isna().any():
            raise ValueError(f"Valori mancanti in {col}.")
        dati[col] = dati[col].astype(str).str.strip()
        if dati[col].eq("").any():
            raise ValueError(f"Valori vuoti in {col}.")

    dati["nuts2"] = dati["nuts2"].str.upper()
    dati["classe_base"] = dati["classe_base"].str.upper()

    for col in numeriche:
        dati[col] = pd.to_numeric(dati[col], errors="raise").replace(
            [float("inf"), -float("inf")], float("nan")
        )

    necessarie = [c for c in numeriche if c != "gn_eurostat_mese"]
    if dati[necessarie].isna().any().any():
        raise ValueError("Valori necessari al calcolo mancanti nello screening.")

    for col in ["anno", "mese", "durata_giorni", "giorni_nel_mese", "notti_applicate"]:
        if dati[col].mod(1).ne(0).any():
            raise ValueError(f"Valori non interi in {col}.")
        dati[col] = dati[col].astype(int)

    if dati.duplicated(chiavi).any():
        raise ValueError("Chiavi evento–territorio–mese duplicate.")
    if not dati["classe_base"].isin(["PERSONE", "EQUIVALENTE"]).all():
        raise ValueError("Classe non prevista nello screening completo.")

    if (
        not dati["mese"].between(1, 12).all()
        or not dati["quota_mese"].between(0, 1).all()
        or not dati["pct_medio"].between(0, 100).all()
        or dati["durata_giorni"].lt(1).any()
        or dati["giorni_nel_mese"].lt(1).any()
        or dati["pubblico_originale"].lt(0).any()
        or dati["pubblico_equivalente"].lt(0).any()
    ):
        raise ValueError("Valori fuori intervallo nello screening.")

    dati["inizio"] = pd.to_datetime(dati["inizio"].map(data_singola))
    dati["fine"] = pd.to_datetime(dati["fine"].map(data_singola))

    if dati[["inizio", "fine"]].isna().any().any():
        raise ValueError("Date evento mancanti o non valide.")

    durata = (dati["fine"] - dati["inizio"]).dt.days + 1
    if durata.ne(dati["durata_giorni"]).any():
        raise ValueError("Durata incoerente con le date dell'evento.")

    dati["data"] = pd.to_datetime(
        dict(year=dati["anno"], month=dati["mese"], day=1)
    )
    fine_mese = dati["data"] + pd.offsets.MonthEnd(0)
    inizio_locale = pd.concat([dati["inizio"], dati["data"]], axis=1).max(axis=1)
    fine_locale = pd.concat([dati["fine"], fine_mese], axis=1).min(axis=1)
    giorni = (fine_locale - inizio_locale).dt.days + 1

    if giorni.ne(dati["giorni_nel_mese"]).any():
        raise ValueError("Giorni nel mese incoerenti con le date.")

    quota_attesa = dati["giorni_nel_mese"] / dati["durata_giorni"]
    if dati["quota_mese"].sub(quota_attesa).abs().gt(0.00001).any():
        raise ValueError("Quota mensile incoerente con i giorni.")

    gruppo = dati.groupby(["id_riga_territoriale", "nuts2"], dropna=False)
    if gruppo["quota_mese"].sum().sub(1).abs().gt(0.00001).any():
        raise ValueError("Le quote mensili di una riga territoriale non sommano a 1.")

    invarianti = [
        "evento", "sport", "inizio", "fine", "durata_giorni",
        "pubblico_equivalente", "classe_base", "metodo_base",
    ]
    if gruppo[invarianti].nunique(dropna=False).gt(1).any().any():
        raise ValueError("La stessa riga territoriale cambia definizione tra mesi.")

    denominatori = dati.groupby(
        ["nuts2", "anno", "mese"]
    )["gn_eurostat_mese"].nunique(dropna=False)

    if denominatori.gt(1).any():
        raise ValueError("Denominatori diversi per la stessa regione e lo stesso mese.")

    dati["tipo_base_visuale"] = "Equivalenti da presenze"
    dati.loc[
        dati["classe_base"].eq("PERSONE"), "tipo_base_visuale"
    ] = "Persone"

    biglietti = (
        dati["classe_base"].eq("EQUIVALENTE")
        & dati["metodo_base"].str.contains("bigliett|ticket", case=False, na=False)
    )
    dati.loc[biglietti, "tipo_base_visuale"] = "Equivalenti da biglietti"

    dati["evento_key"] = (
        dati["sport"] + " | " + dati["evento"]
        + " | " + dati["inizio"].dt.strftime("%Y-%m-%d")
        + " | " + dati["fine"].dt.strftime("%Y-%m-%d")
    )
    dati["Regione"] = dati["nuts2"].map(nomi_reg).fillna(dati["nuts2"])
    return dati


def calcola(dati, quote, notti):
    risultato = dati.copy()

    risultato["quota_pct"] = risultato["sport"].map(quote).astype(float)
    richieste = risultato["sport"].map(notti)
    limite_evento = risultato["durata_giorni"].clip(upper=5)

    risultato["notti_scenario"] = pd.concat(
        [richieste, limite_evento], axis=1
    ).min(axis=1)

    risultato["pernottamenti_simulati"] = (
        risultato["pubblico_equivalente"]
        * risultato["quota_pct"] / 100
        * risultato["notti_scenario"]
        * risultato["quota_mese"]
    )

    denominatore = risultato["gn_eurostat_mese"].where(
        risultato["gn_eurostat_mese"] > 0
    )
    risultato["rapporto_pct"] = 100 * risultato["pernottamenti_simulati"] / denominatore
    return risultato


def riassumi_simulazione(dati):
    mensile = (
        dati.groupby(
            ["nuts2", "Regione", "anno", "mese", "data", "tipo_base_visuale"],
            dropna=False,
        )
        .agg(
            Pernottamenti_simulati=("pernottamenti_simulati", "sum"),
            Domanda_osservata=("gn_eurostat_mese", "first"),
            Eventi=("evento_key", "nunique"),
        )
        .reset_index()
    )
    mensile["Rapporto_percentuale"] = (
        100 * mensile["Pernottamenti_simulati"]
        / mensile["Domanda_osservata"].where(mensile["Domanda_osservata"] > 0)
    )
    return mensile


# ============================================================
# 6. MAPPA KEPLER CON CONFINI NUTS
# ============================================================

@st.cache_data(show_spinner=False, ttl=86400)
def confini_nuts(livello, versione):
    nome = f"NUTS_RG_20M_{versione}_4326_LEVL_{livello}.geojson"
    file = percorso(nome)

    if file.is_file():
        contenuto = json.loads(file.read_text(encoding="utf-8"))
    else:
        url = (
            "https://gisco-services.ec.europa.eu/"
            f"distribution/v2/nuts/geojson/{nome}"
        )
        richiesta = Request(url, headers={"User-Agent": "AirbnbSports/5.1"})
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


def mappa(dati, livello, codice, campo, titolo, key):
    st.subheader(titolo)

    if dati.empty:
        st.info("Nessun dato cartografico per questa selezione.")
        return
    if KeplerGl is None:
        st.error("Per visualizzare la mappa, aggiungi keplergl a requirements.txt.")
        return

    with st.expander("Impostazioni mappa"):
        versione = st.selectbox(
            "Versione NUTS", [2024, 2021, 2016], key="nuts_" + key
        )
        editor = st.checkbox("Mostra strumenti Kepler", key="strumenti_" + key)

    try:
        validi = dati.dropna(subset=[codice, campo])
        if validi.duplicated(codice).any():
            raise ValueError("Più valori per lo stesso territorio.")

        valori = validi.set_index(codice)[campo].to_dict()
        nomi = nomi_reg if livello == 2 else nomi_prov
        features = []
        punti = []
        abbinati = set()

        for feature in confini_nuts(livello, versione):
            nuts = feature["properties"]["NUTS_ID"]
            if nuts not in valori:
                continue
            abbinati.add(nuts)
            features.append({
                "type": "Feature",
                "geometry": feature["geometry"],
                "properties": {
                    "Territorio": str(nomi.get(nuts, nuts)),
                    "NUTS": nuts,
                    "Valore": float(valori[nuts]),
                    "Pernottamenti": numero(valori[nuts]),
                },
            })
            punti.extend(coordinate(feature["geometry"]["coordinates"]))

        mancanti = set(valori) - abbinati
        if mancanti:
            st.warning("Codici senza confine corrispondente: " + ", ".join(sorted(mancanti)))
        if not features:
            raise ValueError("Nessun confine abbinato ai dati.")

        xs, ys = zip(*punti)
        ampiezza = max(max(xs) - min(xs), 1.4 * (max(ys) - min(ys)), 0.1)
        zoom = max(3, min(9, math.log2(360 / ampiezza) - 1.05))

        stile_fisico = {
            "version": 8,
            "name": "SBL Physical",
            "sources": {
                "physical": {
                    "type": "raster",
                    "tiles": [
                        "https://server.arcgisonline.com/ArcGIS/rest/"
                        "services/World_Physical_Map/MapServer/tile/{z}/{y}/{x}"
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
                    "paint": {"raster-opacity": 0.48, "raster-saturation": -0.55},
                },
            ],
        }

        layer = {
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
                        "colors": ["#D1EFF2", "#B4DEE6", "#8CC6D2", "#599CAC", "#276579"],
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
        }

        config = {
            "version": "v1",
            "config": {
                "mapState": {
                    "latitude": (min(ys) + max(ys)) / 2,
                    "longitude": (min(xs) + max(xs)) / 2,
                    "zoom": zoom, "pitch": 0, "bearing": 0, "dragRotate": False,
                },
                "mapStyle": {
                    "styleType": "sbl_physical",
                    "mapStyles": {
                        "sbl_physical": {
                            "id": "sbl_physical", "label": "Sfondo fisico",
                            "custom": True, "style": stile_fisico,
                        }
                    },
                },
                "visState": {
                    "filters": [], "layers": [layer],
                    "interactionConfig": {
                        "tooltip": {
                            "enabled": True,
                            "fieldsToShow": {
                                "territori": [
                                    {"name": c, "format": None}
                                    for c in ["Territorio", "NUTS", "Pernottamenti"]
                                ]
                            },
                        }
                    },
                },
            },
        }

        kepler = KeplerGl(
            height=420, show_docs=False,
            data={"territori": {"type": "FeatureCollection", "features": features}},
            config=config,
        )

        with tempfile.TemporaryDirectory() as cartella:
            file = Path(cartella) / "mappa.html"
            kepler.save_to_html(file_name=str(file), read_only=not editor)
            contenuto = file.read_text(encoding="utf-8")

        components.html(contenuto, height=430, scrolling=False)
        st.caption(
            "Confini Eurostat/GISCO. Sfondo Esri, US National Park Service. "
            "La mappa richiede WebGL e accesso ai servizi cartografici."
        )

    except Exception as exc:
        st.warning(f"Mappa non disponibile: {exc}")

    with st.expander("Valori cartografici"):
        tabella(dati, "mappa_" + key + ".csv")


# ============================================================
# 7. PANORAMA E TERRITORIO
# ============================================================

st.title(pagina)

if pagina == "Panorama nazionale":
    richiedi("reg")
    st.caption(
        "Quadro nazionale. I filtri territoriali si applicano alle altre sezioni."
    )

    copertura = (
        reg.groupby(["year", "nuts2_code"])
        .agg(mesi=("month", "nunique"), validi=("guest_nights_total", "count"))
    )

    anni_completi = []
    for anno in sorted(reg["year"].unique()):
        anno_panel = copertura.loc[anno]
        if (
            len(anno_panel) == reg["nuts2_code"].nunique()
            and anno_panel["mesi"].eq(12).all()
            and anno_panel["validi"].eq(12).all()
        ):
            anni_completi.append(anno)

    if anni_completi:
        annuale = aggrega(
            reg[reg["year"].isin(anni_completi)], ["year"], "guest_nights_total"
        )
        ultimo = max(anni_completi)
        valori = annuale.set_index("year")["guest_nights_total"]
        precedente = valori.get(ultimo - 1, float("nan"))
        variazione = (
            100 * (valori.loc[ultimo] / precedente - 1)
            if pd.notna(precedente) and precedente > 0 else float("nan")
        )

        a, b, c = st.columns(3)
        a.metric(f"Pernottamenti {ultimo}", numero(valori.loc[ultimo]))
        b.metric("Variazione annua", percentuale(variazione))
        c.metric("Territori NUTS2", numero(reg["nuts2_code"].nunique()))

        grafico(px.bar(
            annuale, x="year", y="guest_nights_total",
            color_discrete_sequence=[ACCENTO],
            labels={"year": "Anno", "guest_nights_total": "Pernottamenti"},
        ))
    else:
        st.info("Non risultano anni completi per tutto il panel.")

    anni = sorted(reg["year"].unique(), reverse=True)
    indice = anni.index(max(anni_completi)) if anni_completi else 0
    anno = st.selectbox("Anno della mappa", anni, index=indice)
    st.caption("Per gli anni parziali sono rappresentati i soli mesi disponibili.")

    geografici = aggrega(
        reg[reg["year"].eq(anno)], ["nuts2_code"], "guest_nights_total"
    )
    mappa(
        geografici, 2, "nuts2_code", "guest_nights_total",
        f"Pernottamenti regionali · {anno}", "nazionale",
    )

elif pagina == "Esplora un territorio":
    if provincia != "Tutte":
        richiedi("prov")
        vista = filtra_territorio(prov)
        asse = "year"
        st.subheader(str(nomi_prov.get(provincia, provincia)))
        st.caption("Serie provinciale annuale.")
    else:
        richiedi("reg")
        vista = filtra_territorio(reg)
        asse = "data"
        st.subheader("Italia" if regione == "Tutte" else str(nomi_reg[regione]))

    serie = aggrega(vista, [asse], "guest_nights_total")
    grafico(px.line(
        serie.sort_values(asse), x=asse, y="guest_nights_total",
        color_discrete_sequence=[ACCENTO],
        labels={asse: "Periodo", "guest_nights_total": "Pernottamenti"},
    ))

    if provincia == "Tutte":
        st.subheader("Stagionalità")
        stagioni = aggrega(vista, ["year", "month"], "guest_nights_total")
        stagioni["Anno"] = stagioni["year"].astype(str)
        grafico(px.line(
            stagioni.sort_values(["year", "month"]),
            x="month", y="guest_nights_total", color="Anno",
            labels={"month": "Mese", "guest_nights_total": "Pernottamenti"},
        ))

    tabella(vista, "territorio.csv")


# ============================================================
# 8. EVENTI E SIMULATORE
# ============================================================

elif pagina == "Eventi e simulatore":
    try:
        screen = prepara_screening(leggi_csv(FILES["screen"]))
    except Exception as exc:
        st.error(f"Impossibile usare screening_eventi_completo.csv: {exc}")
        st.stop()

    base_simulatore = screen[
        screen["sport"].str.casefold().ne("sci di fondo")
    ].copy()

    a, b, c = st.columns(3)
    a.metric("Righe territoriali", numero(screen["id_riga_territoriale"].nunique()))
    b.metric("Righe evento–regione–mese", numero(len(screen)))
    c.metric("Eventi identificati", numero(screen["evento_key"].nunique()))

    st.caption(
        "Basi, classificazioni e riparti provengono dallo screening completo."
    )

    catalogo, simulatore, controlli, contesto, archivio = st.tabs([
        "Catalogo", "Simulatore", "Dati utilizzati", "Contesto turistico", "Archivio"
    ])

    with catalogo:
        try:
            file = FILES["master"]
            master = leggi_master(str(file), file.stat().st_mtime_ns)
            a, b = st.columns(2)
            ricerca = a.text_input("Cerca evento", key="catalogo_ricerca")
            sports = (
                sorted(master["Sport"].dropna().astype(str).unique())
                if "Sport" in master else []
            )
            sport = b.selectbox("Sport", ["Tutti"] + sports, key="catalogo_sport")
            applica_regione = st.checkbox("Applica il filtro regionale al catalogo")

            vista = master.copy()
            if sport != "Tutti":
                vista = vista[vista["Sport"].astype(str).eq(sport)]

            nome = "Evento originale" if "Evento originale" in vista else "Evento"
            if ricerca.strip() and nome in vista:
                vista = vista[vista[nome].astype(str).str.contains(
                    ricerca.strip(), case=False, regex=False, na=False
                )]
            if applica_regione and regione != "Tutte" and "Codici NUTS2" in vista:
                vista = vista[vista["Codici NUTS2"].astype(str).str.contains(
                    regione, regex=False, na=False
                )]

            st.caption(f"{len(vista)} righe mostrate su {len(master)} nel master.")
            tabella(vista, "catalogo_event_master.csv", 380)

        except Exception as exc:
            st.warning(f"Catalogo Excel non disponibile: {exc}")

    with simulatore:
        st.write(
            "Modifica la quota pernottanti e le notti per sport. "
            "I risultati distinguono persone, equivalenti da presenze "
            "ed equivalenti da biglietti."
        )

        vista = base_simulatore.copy()
        if regione != "Tutte":
            vista = vista[vista["nuts2"].eq(regione)]

        territorio = "Italia" if regione == "Tutte" else str(nomi_reg[regione])
        st.caption(
            f"Territorio: {territorio}. Il confronto turistico è regionale e mensile."
        )

        with st.expander("Filtri"):
            sports = filtro_tutti("Sport", sorted(vista["sport"].unique()), "sim_sport")
            vista = vista[vista["sport"].isin(sports)]

            anni = filtro_tutti("Anni", sorted(vista["anno"].unique()), "sim_anni")
            mesi = filtro_tutti("Mesi", range(1, 13), "sim_mesi")
            vista = vista[vista["anno"].isin(anni) & vista["mese"].isin(mesi)]

            elenco = vista.drop_duplicates("evento_key").set_index("evento_key")
            etichette = {
                k: f"{r['evento']} · {r['inizio']:%d/%m/%Y}"
                for k, r in elenco.iterrows()
            }
            scelti = filtro_tutti(
                "Eventi", elenco.index, "sim_eventi", formato=lambda k: etichette[k]
            )
            vista = vista[vista["evento_key"].isin(scelti)]

            includi_biglietti = st.checkbox(
                "Includi gli equivalenti da biglietti", value=True,
                help="Sono basi costruite da biglietti pubblicati o stimati.",
            )
            if not includi_biglietti:
                vista = vista[
                    vista["tipo_base_visuale"].ne("Equivalenti da biglietti")
                ]

        if vista.empty:
            st.info("Nessun evento per i filtri selezionati.")
        else:
            sports = sorted(vista["sport"].unique())
            a, b = st.columns([2, 1])
            preset = a.selectbox(
                "Preset comune", [5, 15, 25, 40, 60], index=2,
                format_func=lambda q: f"{q}%",
            )
            if b.button("Applica agli sport"):
                for sport in sports:
                    st.session_state["quota_" + sport] = int(preset)

            quote = {}
            notti = {}

            for sport in sports:
                durata_massima = int(
                    vista.loc[vista["sport"].eq(sport), "durata_giorni"].max()
                )
                massimo = min(5, durata_massima)

                with st.expander(f"{sport} · quota e notti", expanded=len(sports) <= 2):
                    c1, c2 = st.columns(2)
                    quote[sport] = c1.slider(
                        "Quota pernottanti (%)", 0, 100, 25, key="quota_" + sport
                    )

                    if massimo == 1:
                        notti[sport] = 1
                        c2.metric("Notti per pernottante", 1)
                    else:
                        key_notti = "notti_" + sport
                        if key_notti in st.session_state:
                            st.session_state[key_notti] = max(
                                1, min(massimo, int(st.session_state[key_notti]))
                            )
                        notti[sport] = c2.slider(
                            "Notti per pernottante",
                            1, massimo, 1, key=key_notti,
                        )

                    st.caption(
                        "0% = nessuno della base considerata. "
                        "100% = tutta la base. Le notti sono limitate "
                        "anche alla durata del singolo evento."
                    )

            risultato = calcola(vista, quote, notti)
            mensile = riassumi_simulazione(risultato)

            for col, tipo in zip(st.columns(3), COLORI_BASI):
                totale = risultato.loc[
                    risultato["tipo_base_visuale"].eq(tipo), "pernottamenti_simulati"
                ].sum()
                col.metric(tipo, numero(totale), help="Pernottamenti simulati")

            a, b, c = st.columns(3)
            a.metric("Eventi selezionati", numero(risultato["evento_key"].nunique()))
            b.metric("Righe mensili selezionate", numero(len(risultato)))
            c.metric("Righe con rapporto ≥ 2%", numero(risultato["rapporto_pct"].ge(2).sum()))

            st.caption(
                "La soglia si applica alle righe evento–regione–mese. "
                "Gli equivalenti sono una misura di scenario."
            )

            senza_denominatore = risultato["rapporto_pct"].isna().sum()
            if senza_denominatore:
                st.info(
                    f"{senza_denominatore} righe senza denominatore Eurostat "
                    "positivo. Il rapporto è n.d."
                )

            andamento = (
                risultato.groupby(["data", "tipo_base_visuale"])["pernottamenti_simulati"]
                .sum().reset_index()
            )
            grafico(px.bar(
                andamento, x="data", y="pernottamenti_simulati",
                color="tipo_base_visuale", barmode="group",
                color_discrete_map=COLORI_BASI,
                labels={
                    "data": "Mese", "tipo_base_visuale": "Base",
                    "pernottamenti_simulati": "Pernottamenti simulati",
                },
            ))

            st.subheader("Confronto con la domanda regionale")
            tabella(mensile, "simulazione_regioni_mesi.csv")

            with st.expander("Cinque scenari a confronto"):
                righe = []
                for q in [5, 15, 25, 40, 60]:
                    scenario = calcola(vista, {s: q for s in sports}, notti)
                    for tipo, gruppo in scenario.groupby("tipo_base_visuale"):
                        righe.append({
                            "Quota pernottanti": f"{q}%",
                            "Base": tipo,
                            "Pernottamenti simulati": gruppo["pernottamenti_simulati"].sum(),
                            "Righe con rapporto almeno 2%": gruppo["rapporto_pct"].ge(2).sum(),
                        })
                confronto = pd.DataFrame(righe)
                st.caption("Variano le quote pernottanti. Le notti sono quelle scelte per sport.")
                grafico(px.bar(
                    confronto, x="Quota pernottanti", y="Pernottamenti simulati",
                    color="Base", barmode="group", color_discrete_map=COLORI_BASI,
                ))
                tabella(confronto, "simulazione_cinque_scenari.csv")

            with st.expander("Dettaglio e basi utilizzate"):
                campi = [
                    "id_riga_territoriale", "evento", "sport", "Regione", "anno", "mese",
                    "tipo_base_visuale", "classe_base", "metodo_base",
                    "pubblico_originale", "pubblico_equivalente", "durata_giorni",
                    "quota_mese", "quota_pct", "notti_scenario",
                    "pernottamenti_simulati", "gn_eurostat_mese", "rapporto_pct",
                ]
                tabella(risultato[campi], "simulazione_dettaglio.csv", 380)

            st.caption(
                "Il ricalcolo usa le basi e le quote esportate nel CSV. "
                "L'arrotondamento può produrre differenze di pochi centesimi "
                "rispetto ai risultati salvati."
            )

    with controlli:
        st.success("Superati i controlli su chiavi, date, durata e ripartizione mensile.")
        riepilogo = (
            screen.groupby(["tipo_base_visuale", "classe_base"])
            .agg(
                Righe_mensili=("evento", "size"),
                Righe_territoriali=("id_riga_territoriale", "nunique"),
                Notti_scenario_medio_salvate=("pernottamenti_medio", "sum"),
            ).reset_index()
        )
        tabella(riepilogo, "riepilogo_screening.csv")
        st.caption("I totali salvati provengono dal CSV e sono indipendenti dai cursori.")

        with st.expander("Eventi con base equivalente da biglietti"):
            st.write(
                "Per queste righe il file completo usa una base convenzionale "
                "da biglietti pubblicati o stimati."
            )
            biglietti = screen[screen["tipo_base_visuale"].eq("Equivalenti da biglietti")]
            tabella(biglietti[[
                "evento", "sport", "Regione", "anno", "mese", "pubblico_originale",
                "pubblico_equivalente", "durata_giorni", "metodo_base", "pernottamenti_medio",
            ]], "basi_da_biglietti.csv")

        if not reg.empty:
            with st.expander("Confronto denominatori con il master Eurostat"):
                denominatori = screen[
                    ["nuts2", "anno", "mese", "gn_eurostat_mese"]
                ].drop_duplicates()

                confronto = denominatori.merge(
                    reg[["nuts2_code", "year", "month", "guest_nights_total"]],
                    left_on=["nuts2", "anno", "mese"],
                    right_on=["nuts2_code", "year", "month"],
                    how="left", validate="one_to_one", indicator=True,
                )
                confronto["Scarto denominatore"] = (
                    confronto["gn_eurostat_mese"] - confronto["guest_nights_total"]
                )
                diverse = confronto[
                    confronto["_merge"].ne("both")
                    | confronto["gn_eurostat_mese"].isna().ne(
                        confronto["guest_nights_total"].isna()
                    )
                    | confronto["Scarto denominatore"].abs().gt(0.01)
                ]
                st.caption("Il simulatore usa il denominatore salvato nello screening.")
                if diverse.empty:
                    st.success("Nessuna differenza rilevata oltre la tolleranza di 0,01.")
                else:
                    tabella(diverse, "differenze_denominatori.csv")

        with st.expander("Screening completo utilizzato"):
            tabella(screen, "screening_utilizzato.csv", 380)

    with contesto:
        if reg.empty:
            st.info("Dati regionali non disponibili.")
        else:
            codici = sorted(nomi_reg, key=lambda k: str(nomi_reg[k]))
            scelta = st.selectbox(
                "Regione del confronto", codici,
                index=codici.index(regione) if regione in codici else 0,
                format_func=lambda k: str(nomi_reg[k]),
            )
            serie = reg[reg["nuts2_code"].eq(scelta)].copy()
            precedente = serie[["year", "month", "guest_nights_total"]].copy()
            precedente["year"] += 1
            precedente = precedente.rename(
                columns={"guest_nights_total": "Stesso mese anno precedente"}
            )
            serie = serie.merge(
                precedente, on=["year", "month"], how="left", validate="one_to_one"
            )
            riferimento = serie["Stesso mese anno precedente"]
            serie["Variazione percentuale"] = 100 * (
                serie["guest_nights_total"] / riferimento.where(riferimento > 0) - 1
            )
            grafico(px.line(
                serie.sort_values("data"), x="data", y="guest_nights_total",
                color_discrete_sequence=[ACCENTO],
                labels={"data": "Mese", "guest_nights_total": "Pernottamenti"},
            ))
            st.caption(
                "Il confronto con lo stesso mese dell'anno precedente è descrittivo. "
                "Stagionalità, festività, prezzi e altri eventi possono influenzarlo."
            )
            tabella(serie, "contesto_regionale.csv")

    with archivio:
        disponibili = [n for n in ARCHIVIO if percorso(n).is_file()]
        if disponibili:
            scelta = st.selectbox("Elaborazione salvata", disponibili)
            st.info(
                "Alcuni screening storici utilizzano gli ingressi cumulativi "
                "direttamente. Il simulatore usa screening_eventi_completo.csv "
                "con le basi equivalenti già elaborate."
            )
            try:
                tabella(leggi_csv(percorso(scelta)), "archivio_" + scelta, 380)
            except Exception as exc:
                st.error(str(exc))


# ============================================================
# 9. FORECAST
# ============================================================

elif pagina == "Forecast 2026–2030":
    provinciale = provincia != "Tutte"
    chiave = "fp" if provinciale else "fr"
    richiedi(chiave)

    vista = filtra_territorio(DATI[chiave])
    prefisso = "forecast_" if provinciale else "gn_"
    codice = "nuts3_code" if provinciale else "nuts2_code"

    st.caption(
        "Scenari letti dai CSV del progetto. "
        "I cursori del simulatore eventi agiscono soltanto sulla simulazione."
    )

    scelti = st.multiselect("Scenari", SCENARI, default=SCENARI)
    if not scelti or vista.empty:
        st.info("Seleziona uno scenario con dati disponibili.")
    else:
        righe = []
        for scenario in scelti:
            campo = prefisso + scenario.lower()
            annuale = aggrega(vista, ["year"], campo).rename(
                columns={campo: "Pernottamenti"}
            )
            annuale["Scenario"] = scenario
            righe.append(annuale)

        lungo = pd.concat(righe, ignore_index=True)
        grafico(px.line(
            lungo.sort_values("year"), x="year", y="Pernottamenti",
            color="Scenario", markers=True, color_discrete_map=COLORI_SCENARI,
            labels={"year": "Anno"},
        ))

        a, b = st.columns(2)
        anno = a.selectbox("Anno della mappa", sorted(vista["year"].unique(), reverse=True))
        scenario = b.selectbox("Scenario della mappa", scelti)
        campo = prefisso + scenario.lower()

        geografici = aggrega(vista[vista["year"].eq(anno)], [codice], campo)
        mappa(
            geografici, 3 if provinciale else 2, codice, campo,
            f"Forecast {scenario.lower()} · {anno}", "forecast",
        )
        tabella(lungo, "forecast_annuale.csv")


# ============================================================
# 10. DATI E METODO
# ============================================================

else:
    fonti, metodo, previsione, qualita = st.tabs([
        "Fonti e lettura", "Metodo eventi", "Forecast", "Qualità e limiti",
    ])

    with fonti:
        st.markdown(
            """
            ### Domanda di ricerca

            Quanto cambia la domanda di alloggi di breve durata in una
            regione italiana nei mesi interessati da grandi eventi sportivi?
            Per rispondere occorre distinguere tre quantità.

            **La domanda osservata** è il numero di pernottamenti rilevati
            da Eurostat nel territorio e nel periodo considerati.

            **La domanda simulata** si ottiene applicando ipotesi sul
            pernottamento alla base di pubblico di un evento. Serve a
            valutare la dimensione che l'evento potrebbe assumere rispetto
            alla domanda regionale.

            **La domanda prevista** è la traiettoria riportata nei file
            di forecast. Per attribuire all'evento una variazione rispetto
            a quella traiettoria occorrerebbe stimare separatamente
            cosa sarebbe successo in sua assenza.

            ### Unità territoriale e temporale

            La base mensile è Eurostat **`tour_ce_omn12`**. L'unità centrale
            è **NUTS2 × mese**. Per la provincia si utilizza la serie annuale
            **`tour_ce_oan3`**. Un totale provinciale annuale non permette
            di conoscere direttamente la domanda del mese di una gara.

            I confini cartografici sono quelli ufficiali NUTS. Il codice
            territoriale e la versione della geometria devono corrispondere.
            Una regione senza geometria abbinata può comunque avere dati.

            ### Che cosa conta un pernottamento

            Tre ospiti che soggiornano due notti generano **sei guest nights**.
            Un singolo soggiorno può quindi generare più pernottamenti ospite.
            Le notti di occupazione dell'alloggio misurano ancora un'altra
            quantità. Per questo *guest nights*, *stays* e *occupied nights*
            non sono intercambiabili.

            La fonte copre gli alloggi prenotati attraverso le piattaforme
            comprese nella rilevazione. Il dato non rappresenta automaticamente
            tutta la ricettività regionale o il solo mercato Airbnb.
            La distinzione domestico/estero riguarda la residenza secondo
            la fonte e non identifica da sola i visitatori provenienti
            da fuori regione.

            Un indice dei prezzi dei servizi di alloggio, come l'HICP usato
            nel progetto, fornisce un controllo macroeconomico. Non permette
            di ricavare il prezzo medio di un annuncio o l'ADR di Airbnb.

            ### Come leggere i file

            Il **master Excel** conserva eventi, sedi, fonti, note e qualità.
            Lo **screening completo** contiene basi già classificate e ripartite
            per territorio e mese. I **CSV di forecast** contengono le
            previsioni già elaborate a monte.

            Nel dataset verificato, 369 righe territoriali generano 377 righe
            mensili. Un evento con più sedi o a cavallo di due mesi può
            comparire più volte senza essere un nuovo evento.

            [Metadati ufficiali Eurostat](https://ec.europa.eu/eurostat/cache/metadata/en/tour_ce_esms.htm)
            """
        )

    with metodo:
        st.markdown(
            """
            ### Dalla fonte alla base di pubblico

            La prima domanda è che cosa abbia effettivamente contato la fonte.
            Un partecipante, un ingresso giornaliero e un biglietto acquistato
            non rappresentano necessariamente lo stesso individuo.

            **Persone.** La base è trattata come persone secondo il metodo
            documentato nella riga. Può derivare da visitatori unici,
            partecipanti o stime di persone. Il campo `metodo_base` consente
            di distinguere la provenienza del valore.

            **Equivalenti da presenze.** Per le manifestazioni con ingressi
            cumulativi, lo screening può usare una media giornaliera equivalente.
            Se una gara registra 300.000 ingressi in tre giorni, la media è
            100.000 ingressi al giorno. La fonte non dice quanti spettatori
            siano tornati in più giornate.

            **Equivalenti da biglietti.** Il file completo applica una
            convenzione analoga a biglietti pubblicati o stimati. Una prevendita
            minima, una soglia dichiarata o un biglietto emesso mantengono
            i limiti della fonte originale. Un biglietto non dimostra
            né l'ingresso effettivo né il pernottamento.

            Nel confronto verificato, dieci righe mensili relative a cinque
            eventi avevano un valore equivalente nel file completo ed erano
            considerate non convertibili nel precedente file degli equivalenti.
            Il simulatore le identifica come **equivalenti da biglietti**.
            Il filtro dedicato permette di vedere come cambiano i risultati
            includendole o escludendole.

            ### Territorio e mese

            La base `pubblico_equivalente` è già attribuita alla riga
            territoriale. Una manifestazione con quattro sedi regionali può
            avere quattro valori diversi. Dividere nuovamente ciascun valore
            per quattro ne ridurrebbe artificialmente il peso.

            La quota `quota_mese` distribuisce la riga territoriale nel tempo.
            Un evento di tre giorni con due giorni in agosto e uno in settembre
            ha pesi di 2/3 e 1/3. I pesi della stessa riga territoriale devono
            sommare a uno.

            Questa distribuzione assume un'esposizione uniforme per giorno.
            Il pubblico può concentrarsi nella finale o nel fine settimana.
            Anche i soggiorni possono attraversare il confine del mese.
            I dati mensili non risolvono questi dettagli giornalieri.

            ### Quota pernottanti e notti

            Il cursore per sport va da 0 a 100%, con valore iniziale 25%.
            I preset 5%, 15%, 25%, 40% e 60% servono a confrontare ipotesi.
            Non sono frequenze osservate per ciascuno sport.

            La quota si applica alla base effettivamente disponibile.
            Se questa comprende tutto il pubblico, il parametro incorpora
            implicitamente la provenienza e la propensione al pernottamento.
            Se la base comprende soltanto non residenti, la quota si riferisce
            a quel gruppo. Non viene applicato un secondo moltiplicatore
            automatico, né un minimo basato sui tifosi ospiti.

            Le notti partono da una. Per ogni riga si applica il minimo tra
            notti scelte, durata dell'evento e cinque. Il vincolo rende
            esplicito il perimetro della simulazione. Soggiorni prima o dopo
            l'evento richiederebbero un'ipotesi aggiuntiva.
            """
        )

        st.latex(r"N_{e,r,m}=B_{e,r}\cdot q_e\cdot n_e\cdot w_{e,m}")

        st.markdown(
            """
            **B** è la base territoriale già elaborata. **q** è la quota
            pernottanti divisa per 100. **n** sono le notti applicate
            al singolo evento e **w** è `quota_mese`.

            Una base di 10.000 persone, una quota del 25%, due notti
            e un peso mensile del 50% producono **2.500 pernottamenti
            simulati** nel mese. La formula usa la quota mensile una volta sola.

            Il nome tecnico `pubblico_equivalente` è comune a tutte le righe.
            La classificazione `classe_base` indica se il valore appartiene
            alla base persone o alla base equivalente.
            """
        )

        st.latex(r"R_{e,r,m}=100\cdot\frac{N_{e,r,m}}{GN_{r,m}}")

        st.markdown(
            """
            **GN** è la domanda Eurostat della stessa regione e dello stesso
            mese, salvata nello screening. Con 2.500 pernottamenti simulati
            e 100.000 osservati, il rapporto è **2,5%**. Misura la dimensione
            dello scenario rispetto al riferimento regionale.

            Il numeratore potrebbe comprendere soggiorni in strutture fuori
            dal perimetro Eurostat. Inoltre, una parte degli spettatori avrebbe
            potuto visitare il territorio anche in assenza dell'evento.
            Il rapporto non identifica la quota di turismo causata dall'evento.

            ### Soglia e aggregazioni

            Il 2% è una soglia esplorativa per selezionare casi da approfondire.
            Il conteggio sopra soglia riguarda righe evento–regione–mese.
            Un evento può quindi comparire più volte.

            Nel riepilogo regionale mensile si sommano i pernottamenti simulati
            per tipo di base e si usa una sola volta il denominatore del mese.
            Le somme tra eventi possono includere persone o soggiorni ripetuti.
            I risultati delle tre basi rimangono distinti.

            Per riprodurre un confronto storico devono coincidere eventi,
            basi, quote, notti e periodo. Un vecchio scenario calcolato sugli
            ingressi cumulativi può produrre un totale diverso anche a parità
            di eventi. Le basi e le quote esportate nel CSV possono inoltre
            introdurre piccoli scarti di arrotondamento.
            """
        )

    with previsione:
        st.markdown(
            """
            ### Che cosa rappresentano le previsioni

            La dashboard legge i CSV di forecast prodotti nel progetto.
            Prudente, Intermedio ed Espansivo descrivono traiettorie costruite
            con ipotesi diverse. Non hanno automaticamente il significato
            di percentili o intervalli di confidenza.

            Il modello va documentato indicando periodo di addestramento,
            periodo di test, trattamento del COVID, anno base, stagionalità,
            ipotesi di crescita e riparto territoriale. Il 2020–2021 comprende
            uno shock strutturale. Trattarlo come un periodo ordinario
            cambierebbe il riferimento utilizzato per la domanda attesa.

            ### Validazione temporale

            Un confronto informativo usa gli stessi territori, mesi e orizzonti
            per tutti i modelli. La previsione deve essere costruita usando
            solo le informazioni disponibili prima del periodo di test.
            Una baseline stagionale, come il valore dello stesso mese
            dell'anno precedente, offre un riferimento concreto.

            I risultati vanno letti anche per territorio e periodo.
            Una buona prestazione aggregata può nascondere difficoltà
            in singole regioni o durante particolari mesi.
            """
        )

        st.latex(
            r"WAPE=100\cdot\frac{\sum_t|y_t-\hat y_t|}{\sum_t|y_t|}"
        )
        st.markdown(
            """
            Il **WAPE** rapporta gli errori assoluti al volume osservato.
            Sullo stesso campione, un valore inferiore indica un errore
            aggregato più contenuto. I territori con maggiore volume pesano
            di più nel risultato complessivo.
            """
        )

        st.latex(
            r"Bias=100\cdot\frac{\sum_t(\hat y_t-y_t)}{\sum_t y_t}"
        )
        st.markdown(
            """
            Con questa convenzione, un **bias positivo** indica sovrastima
            complessiva e uno negativo indica sottostima. Errori di segno
            opposto possono compensarsi. Il bias va quindi letto insieme
            a una misura dell'errore assoluto. Le metriche richiedono
            un denominatore diverso da zero.

            ### Province e relazione con gli eventi

            Un forecast provinciale costruito con quote costanti eredita
            la dinamica regionale. La riconciliazione tra somma delle province
            e totale regionale verifica la coerenza contabile, mentre
            l'accuratezza richiede un confronto con osservazioni successive.

            I pernottamenti simulati non vengono aggiunti al forecast.
            Eventi ricorrenti possono essere già incorporati nella domanda
            storica. Aggiungerli integralmente rischierebbe di contare
            due volte parte della stessa domanda.

            La previsione dell'incremento dovuto a un futuro evento richiede
            prima una stima credibile dell'effetto storico rispetto a un
            controfattuale. La dashboard attuale mantiene distinti il forecast
            di base e l'esercizio di simulazione.
            """
        )

    with qualita:
        controlli_file = []
        for chiave, richieste in SCHEMI.items():
            dati = DATI[chiave]
            campi = [
                c for c in richieste
                if c.startswith(("guest_nights", "gn_", "forecast_"))
            ]
            controlli_file.append({
                "File": FILES[chiave].name,
                "Stato": ERRORI.get(chiave, "Chiavi valide e non duplicate"),
                "Righe": len(dati),
                "Valori mancanti": (
                    int(dati[campi].isna().sum().sum()) if not dati.empty else None
                ),
                "Valori negativi": (
                    int(dati[campi].lt(0).sum().sum()) if not dati.empty else None
                ),
            })
        tabella(pd.DataFrame(controlli_file), "controlli_file.csv")

        if not reg.empty:
            with st.expander("Copertura regionale"):
                copertura = (
                    reg.groupby(["nuts2_code", "year"])
                    .agg(
                        Mesi_presenti=("month", "nunique"),
                        Mesi_valorizzati=("guest_nights_total", "count"),
                    ).reset_index()
                )
                tabella(copertura, "copertura_regionale.csv")

        if not fr.empty and not fp.empty:
            with st.expander("Riconciliazione forecast"):
                confronti = []
                for scenario in SCENARI:
                    col_reg = "gn_" + scenario.lower()
                    col_prov = "forecast_" + scenario.lower()
                    totale_reg = aggrega(fr, ["nuts2_code", "year"], col_reg).rename(
                        columns={col_reg: "Regione"}
                    )
                    totale_prov = aggrega(fp, ["nuts2_code", "year"], col_prov).rename(
                        columns={col_prov: "Somma province"}
                    )
                    confronto = totale_reg.merge(
                        totale_prov, on=["nuts2_code", "year"],
                        how="outer", validate="one_to_one",
                    )
                    confronto["Scenario"] = scenario
                    confronto["Scarto"] = confronto["Somma province"] - confronto["Regione"]
                    confronti.append(confronto)

                tabella(
                    pd.concat(confronti, ignore_index=True),
                    "riconciliazione_forecast.csv",
                )

        st.markdown(
            """
            ### Copertura e qualità

            I controlli verificano struttura, chiavi, duplicazioni,
            valori mancanti e copertura temporale. Nello screening si
            verificano anche la durata, i giorni attribuiti ai mesi,
            la somma delle quote e la stabilità della base territoriale.

            Un dato mancante non equivale a zero. Un anno con meno di
            dodici mesi non è un totale annuale completo. La copertura
            del panel va verificata anche rispetto all'insieme atteso
            dei territori, perché il conteggio dei mesi delle regioni
            presenti non rivela da solo una regione interamente assente.

            Le fonti degli eventi possono contenere conteggi pubblicati,
            stime, approssimazioni o limiti inferiori. Qualità e precisione
            possono cambiare tra sport ed eventi. Il catalogo conserva
            fonti e note per permettere di risalire al singolo dato.

            ### Interpretazione dell'andamento osservato

            Il turismo regionale può variare per stagionalità, festività,
            prezzi, trasporti e manifestazioni contemporanee. Un grande evento
            può attirare visitatori, spostare viaggi da un altro periodo o
            sostituire turisti che evitano il territorio durante la manifestazione.

            Una partita di un solo giorno può essere documentata nel catalogo,
            ma il dato regionale mensile può risultare troppo aggregato
            per distinguerne l'effetto da queste altre variazioni.

            Per stimare un incremento attribuibile all'evento occorre un termine
            di confronto credibile. Bisogna considerare gli stessi mesi in anni
            precedenti, le tendenze prima dell'evento e territori comparabili
            non esposti nello stesso periodo. Conta anche la possibilità
            che i visitatori soggiornino in una regione vicina.

            Il confronto con l'anno precedente e lo screening del 2% aiutano
            a scegliere i casi da approfondire. Un'analisi causale richiederebbe
            un disegno ulteriore, per esempio uno studio di evento o una
            difference-in-differences con ipotesi verificate sui dati.
            """
        )


st.divider()
st.caption(f"Airbnb&Sports · {VERSIONE} · [SBL Consultancy]({SITO})")
