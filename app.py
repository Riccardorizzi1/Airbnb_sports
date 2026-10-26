
from pathlib import Path
from urllib.request import Request, urlopen
import json
import re

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import math
import tempfile
import streamlit.components.v1 as components
from keplergl import KeplerGl

VERSIONE = "v3.0 · 25/10/2026"

st.set_page_config(
    page_title="Airbnb&Sports",
    page_icon="🏟️",
    layout="wide"
)

ROOT = Path(__file__).resolve().parent

FILES = {
    "reg": ROOT / "eurostat_master_final.csv",
"prov": ROOT / "eurostat_nuts3_annuale_2018_2025.csv",
"fp": ROOT / "forecast_nuts3_2026_2030_3_scenari.csv",
    "fp": ROOT / "output_eurostat/forecast_nuts3_2026_2030_3_scenari.csv",
    "eventi": ROOT / "AirbnbSports_Event_Master_Simone_Visitatori.xlsx",
    "screening": ROOT / "screening_eventi_completo.csv",
    "mesi_eventi": ROOT / "screening_regioni_mesi_completo.csv",
}

SCENARI = ["Prudente", "Intermedio", "Espansivo"]

COLORI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C"
}

PAGINE = [
    "Panorama nazionale",
    "Analisi storica Eurostat",
    "Analisi storica sportiva",
    "Forecast 2026–2030",
    "Metodologia"
]


# ============================================================
# 1. LETTURA DATI
# ============================================================

@st.cache_data(show_spinner=False)
def csv(path):
    return pd.read_csv(
        path,
        encoding="utf-8-sig",
        low_memory=False
    )


@st.cache_data(show_spinner=False)
def excel(path):
    return pd.read_excel(path, sheet_name=0)


@st.cache_data(ttl=86400, show_spinner=False)
def confini(livello):

    url = (
        "https://gisco-services.ec.europa.eu/"
        "distribution/v2/nuts/geojson/"
        f"NUTS_RG_20M_2024_4326_LEVL_{livello}.geojson"
    )

    request = Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"}
    )

    with urlopen(request, timeout=30) as response:
        geo = json.load(response)

    features = []

    for feature in geo["features"]:

        code = str(
            feature.get("properties", {}).get(
                "NUTS_ID", ""
            )
        )

        if (
            code.startswith("IT")
            and len(code) == livello + 2
        ):
            feature["id"] = code
            features.append(feature)

    geo["features"] = features

    return geo


# ============================================================
# 2. FUNZIONI GENERALI
# ============================================================

def numero(x, decimali=0):

    if pd.isna(x):
        return "—"

    return (
        f"{float(x):,.{decimali}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def pct(x):

    if pd.isna(x):
        return "—"

    return f"{x:+.2f}%".replace(".", ",")


def colonna(df, *nomi):

    trad = {
        re.sub(
            "[^a-z0-9]", "", str(c).lower()
        ): c
        for c in df.columns
    }

    for nome in nomi:

        key = re.sub(
            "[^a-z0-9]", "", nome.lower()
        )

        if key in trad:
            return trad[key]

    return None


def grafico(fig, height=400):

    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=TESTO,
        margin=dict(l=5, r=5, t=40, b=18),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            y=1.14
        )
    )

    fig.update_yaxes(
        gridcolor="rgba(122,157,169,.25)",
        zeroline=False,
        tickformat=",~s"
    )

    fig.update_xaxes(
        gridcolor="rgba(122,157,169,.12)",
        zeroline=False
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


def tabella(df, filename=None):

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if filename:

        st.download_button(
            "Scarica CSV",
            df.to_csv(
                index=False,
                sep=";"
            ).encode("utf-8-sig"),
            file_name=filename,
            mime="text/csv"
        )



def mappa(df, livello, codice, valore, titolo):

    st.subheader(titolo)

    dati = df[
        [codice, "territorio", valore]
    ].copy()

    dati[codice] = (
        dati[codice].astype(str).str.strip()
    )

    dati[valore] = pd.to_numeric(
        dati[valore],
        errors="coerce"
    )

    dati = dati.dropna(subset=[valore])

    if dati.empty:
        st.warning("Nessun dato disponibile per la mappa.")
        return

    if dati[codice].duplicated().any():
        st.error("Codici NUTS duplicati.")
        return

    mostra = st.checkbox(
        "Visualizza mappa Kepler.gl",
        value=True,
        key="kepler_" + str(livello) + "_" + titolo
    )

    if mostra:

        strumenti = st.toggle(
            "Mostra controlli Kepler",
            value=False,
            key="tools_" + str(livello) + "_" + titolo
        )

        try:

            # Geometrie ufficiali Eurostat/GISCO
            territorio = confini(livello)

            valori = (
                dati.set_index(codice)[
                    ["territorio", valore]
                ].to_dict("index")
            )

            features = []
            longitudini = []
            latitudini = []

            # Coordinate ricavate dai confini ufficiali
            def coord_bounds(x):

                if isinstance(x, (list, tuple)):

                    if (
                        len(x) >= 2
                        and all(
                            isinstance(a, (int, float))
                            for a in x[:2]
                        )
                    ):

                        longitudini.append(float(x[0]))
                        latitudini.append(float(x[1]))

                    else:

                        for y in x:
                            coord_bounds(y)

            # Unione dei dati e delle geometrie
            for f in territorio.get("features", []):

                ident = str(
                    f.get("properties", {}).get(
                        "NUTS_ID",
                        f.get("id", "")
                    )
                )

                if ident not in valori:
                    continue

                v = float(valori[ident][valore])

                features.append({
                    "type": "Feature",
                    "geometry": f["geometry"],
                    "properties": {
                        "nuts_code": ident,
                        "territorio": str(
                            valori[ident]["territorio"]
                        ),
                        "guest_nights": v,
                        "guest_nights_it": (
                            f"{v:,.0f}".replace(",", ".")
                        )
                    }
                })

                coord_bounds(
                    f["geometry"]["coordinates"]
                )

            if not features:
                raise ValueError(
                    "I codici NUTS non corrispondono "
                    "ai confini GISCO."
                )

            # Centro e zoom calcolati dalle geometrie
            west = min(longitudini)
            east = max(longitudini)
            south = min(latitudini)
            north = max(latitudini)

            lon = (west + east) / 2
            lat = (south + north) / 2

            ampiezza = max(
                east - west,
                (north - south) * 1.4,
                0.1
            )

            zoom = max(
                3.0,
                min(
                    11.0,
                    math.log2(360 / ampiezza) - 0.9
                )
            )

            geojson = {
                "type": "FeatureCollection",
                "features": features
            }

            dataset = "AirbnbSports_NUTS"

            palette = [
                "#C2F0F6",
                "#88D6E7",
                "#4BAFC9",
                "#126982",
                "#084A68"
            ]

            # Configurazione Kepler.gl
            config = {
                "version": "v1",
                "config": {

                    "mapState": {
                        "latitude": lat,
                        "longitude": lon,
                        "zoom": zoom,
                        "pitch": 0,
                        "bearing": 0
                    },

                    "mapStyle": {
                        "styleType": (
                            "dark" if dark else "light"
                        )
                    },

                    "visState": {

                        "filters": [],

                        "layers": [{
                            "id": "airbnbsports_geojson",
                            "type": "geojson",

                            "config": {
                                "dataId": dataset,
                                "label": (
                                    f"Guest nights NUTS{livello}"
                                ),
                                "color": [40, 165, 199],

                                "columns": {
                                    "geojson": "_geojson"
                                },

                                "isVisible": True,

                                "visConfig": {
                                    "opacity": 0.85,
                                    "strokeOpacity": 0.8,
                                    "thickness": 0.6,
                                    "stroked": True,
                                    "filled": True,

                                    "colorRange": {
                                        "name": "AirbnbSports blu",
                                        "type": "sequential",
                                        "category": "Custom",
                                        "colors": palette
                                    }
                                }
                            },

                            "visualChannels": {
                                "colorField": {
                                    "name": "guest_nights",
                                    "type": "real"
                                },
                                "colorScale": "quantile"
                            }
                        }],

                        "interactionConfig": {
                            "tooltip": {
                                "fieldsToShow": {
                                    dataset: [
                                        {
                                            "name": "territorio",
                                            "format": None
                                        },
                                        {
                                            "name": "nuts_code",
                                            "format": None
                                        },
                                        {
                                            "name": "guest_nights_it",
                                            "format": None
                                        }
                                    ]
                                },
                                "enabled": True
                            }
                        }
                    }
                }
            }

            # Generazione Kepler
            kepler = KeplerGl(
                height=720,
                data={dataset: geojson},
                config=config
            )

            # File HTML temporaneo
            with tempfile.TemporaryDirectory(
                prefix="airbnbsports_kepler_"
            ) as tmp:

                file_html = Path(tmp) / "mappa.html"

                kepler.save_to_html(
                    file_name=str(file_html),
                    read_only=not strumenti
                )

                html = file_html.read_text(
                    encoding="utf-8"
                )

            components.html(
                html,
                height=740,
                scrolling=False
            )

            st.caption(
                f"Kepler.gl · {len(features)} territori "
                f"NUTS{livello} · Confini GISCO 2024"
            )

        except Exception as exc:

            st.error(
                f"Errore nella mappa Kepler.gl: {exc}"
            )

    # Tabella con dati originali
    tabella(
        dati.sort_values(
            valore,
            ascending=False
        )
    )



# ============================================================
# 3. SIDEBAR E TEMA
# ============================================================

with st.sidebar:

    st.markdown("## ◈ Airbnb&Sports")
    st.caption("Tourism Demand · Italia")
    st.caption("VERSIONE " + VERSIONE)

    st.divider()

    pagina = st.radio(
        "Sezione",
        PAGINE
    )

    st.divider()

    dark = st.toggle(
        "Modalità scura",
        value=False
    )

    st.divider()


if dark:

    SIDE = (
        "#071F31",
        "#0E4058",
        "#126982"
    )

    BG = (
        "#071822",
        "#0C2B39",
        "#10475B"
    )

    CARD = "#102D3B"
    TESTO = "#F3FAFC"
    BORDO = "#315462"
    OMBRA = "rgba(0,0,0,.28)"
    ACCENTO = "#6DDDEB"

    SCALA = [
        "#16465C",
        "#267C9B",
        "#4CBCD1",
        "#A4F1F4"
    ]

else:

    SIDE = (
        "#153F57",
        "#17627E",
        "#259BBA"
    )

    BG = (
        "#C7F1F6",
        "#9FE5F0",
        "#5BC6E2"
    )

    CARD = "#FDFEFF"
    TESTO = "#082B46"
    BORDO = "#D4E7EC"
    OMBRA = "rgba(8,43,70,.10)"
    ACCENTO = "#0E718F"

    SCALA = [
        "#D9F6FA",
        "#A0E0EB",
        "#4DBAD4",
        "#126889",
        "#082B46"
    ]


# ============================================================
# 4. CSS - DESIGN DI RIFERIMENTO
# ============================================================

st.markdown(f"""
<style>

header[data-testid="stHeader"] {{
    background: transparent !important;
}}

section[data-testid="stSidebar"] {{
    min-width: 340px !important;
    width: 340px !important;

    background: linear-gradient(
        180deg,
        {SIDE[0]},
        {SIDE[1]} 52%,
        {SIDE[2]}
    ) !important;
}}

section[data-testid="stSidebar"] * {{
    color: white !important;
}}

.stApp,
[data-testid="stAppViewContainer"],
[data-testid="stMain"] {{

    background: linear-gradient(
        100deg,
        {BG[0]},
        {BG[1]} 47%,
        {BG[2]}
    ) !important;

    color: {TESTO} !important;
}}

.block-container {{
    max-width: 1550px !important;
    padding: 2rem 2.5rem 3.4rem !important;
}}

[data-testid="stMain"] h1,
[data-testid="stMain"] h2,
[data-testid="stMain"] h3,
[data-testid="stMain"] p,
[data-testid="stMain"] label {{
    color: {TESTO} !important;
}}

[data-testid="stMetric"] {{
    background: {CARD} !important;
    border: 1px solid {BORDO} !important;
    border-radius: 16px !important;
    padding: 1rem 1.2rem !important;
    box-shadow: 0 5px 14px {OMBRA} !important;
    min-height: 115px;
}}

[data-testid="stMetric"] * {{
    color: {TESTO} !important;
}}

[data-testid="stDataFrame"],
[data-testid="stExpander"] {{
    border: 1px solid {BORDO} !important;
    border-radius: 14px !important;
    overflow: hidden;
}}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 5. CARICAMENTO MASTER E FORECAST
# ============================================================

obbligatori = ("reg", "prov", "fr", "fp")

mancanti = [
    str(FILES[k].relative_to(ROOT))
    for k in obbligatori
    if not FILES[k].exists()
]

if mancanti:

    st.error(
        "Mancano alcuni CSV del progetto:"
    )

    st.code("\n".join(mancanti))
    st.stop()


try:

    reg = csv(str(FILES["reg"]))
    prov = csv(str(FILES["prov"]))
    fr = csv(str(FILES["fr"]))
    fp = csv(str(FILES["fp"]))

    richieste = [
        (
            reg,
            [
                "year", "month",
                "nuts2_code", "nuts2_name",
                "guest_nights_total"
            ]
        ),
        (
            prov,
            [
                "year", "nuts3_code",
                "nuts3_name", "nuts2_code",
                "guest_nights_total"
            ]
        ),
        (
            fr,
            [
                "year", "month", "nuts2_code"
            ] + [
                f"gn_{s.lower()}"
                for s in SCENARI
            ]
        ),
        (
            fp,
            [
                "year", "nuts3_code", "nuts2_code"
            ] + [
                f"forecast_{s.lower()}"
                for s in SCENARI
            ]
        )
    ]

    for df, colonne in richieste:

        assenti = set(colonne) - set(df.columns)

        if assenti:
            raise ValueError(
                f"Colonne mancanti: {sorted(assenti)}"
            )

        df["year"] = df["year"].astype(int)

    for df in (reg, fr):

        df["month"] = df["month"].astype(int)

        df["data"] = pd.to_datetime(
            dict(
                year=df.year,
                month=df.month,
                day=1
            )
        )

except Exception as exc:

    st.error(
        f"Errore lettura dati: {exc}"
    )

    st.stop()


names2 = (
    reg.drop_duplicates("nuts2_code")
    .set_index("nuts2_code")["nuts2_name"]
    .to_dict()
)

names3 = (
    prov.drop_duplicates("nuts3_code")
    .set_index("nuts3_code")["nuts3_name"]
    .to_dict()
)


# ============================================================
# 6. FILTRI TERRITORIALI
# ============================================================

with st.sidebar:

    st.subheader("Selezione territoriale")

    regione = st.selectbox(
        "Regione NUTS2",
        ["Italia"] + sorted(names2),
        format_func=lambda c: (
            c if c == "Italia"
            else f"{names2[c]} · {c}"
        )
    )

    pr = (
        prov if regione == "Italia"
        else prov[prov.nuts2_code == regione]
    )

    provincia = st.selectbox(
        "Provincia NUTS3",
        ["Tutte"] + sorted(
            pr.nuts3_code.unique()
        ),
        format_func=lambda c: (
            c if c == "Tutte"
            else f"{names3[c]} · {c}"
        )
    )

    st.divider()

    scenario = st.selectbox(
        "Scenario predefinito",
        SCENARI,
        index=1
    )


# ============================================================
# 7. HEADER
# ============================================================

st.title("Airbnb&Sports")

st.caption(
    "Domanda di short-term accommodation "
    "tramite piattaforme online · Italia"
)


# ============================================================
# 8. PANORAMA NAZIONALE
# ============================================================

if pagina == "Panorama nazionale":

    st.header(
        "Panorama nazionale · riferimento 2022–2025"
    )

    agg = (
        reg[reg.year.between(2022, 2025)]
        .groupby("year", as_index=False)[
            "guest_nights_total"
        ]
        .sum()
    )

    annual = agg.set_index("year")[
        "guest_nights_total"
    ]

    a24 = annual.loc[2024]
    a25 = annual.loc[2025]

    a30 = fr.loc[
        fr.year == 2030,
        f"gn_{scenario.lower()}"
    ].sum()

    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "Guest nights 2025",
        numero(a25)
    )

    c2.metric(
        "Crescita 2025/24",
        pct(100 * (a25 / a24 - 1))
    )

    c3.metric(
        f"Previsione 2030 · {scenario}",
        numero(a30)
    )

    c4.metric(
        "Copertura",
        "21 NUTS2 · 107 NUTS3"
    )

    fig = px.bar(
        agg,
        x="year",
        y="guest_nights_total",
        color_discrete_sequence=[ACCENTO]
    )

    fig.update_xaxes(dtick=1)

    grafico(fig, 360)

    st.caption(
        "2020–2021 esclusi dal periodo principale "
        "per lo shock Covid; restano nel master."
    )

    regions = (
        reg[reg.year == 2025]
        .groupby(
            "nuts2_code",
            as_index=False
        )["guest_nights_total"]
        .sum()
    )

    regions["territorio"] = (
        regions.nuts2_code.map(names2)
    )

    mappa(
        regions,
        2,
        "nuts2_code",
        "guest_nights_total",
        "Domanda regionale osservata 2025"
    )


# ============================================================
# 9. ANALISI STORICA EUROSTAT
# ============================================================

elif pagina == "Analisi storica Eurostat":

    st.header(
        "Analisi storica · 2022–2025 "
        "e mesi disponibili 2026"
    )

    livello = st.radio(
        "Frequenza",
        [
            "Mensile NUTS2",
            "Annuale NUTS3"
        ],
        horizontal=True
    )

    indicatori = {
        "Guest nights totali": "guest_nights_total",
        "Domestici": "guest_nights_domestic",
        "Stranieri": "guest_nights_foreign",
        "Stays": "stays_total",
        "Notti occupate": "occupied_nights_total"
    }

    base = (
        reg if livello == "Mensile NUTS2"
        else prov
    )

    disponibili = {
        k: v
        for k, v in indicatori.items()
        if v in base.columns
    }

    misura = st.selectbox(
        "Indicatore Eurostat",
        list(disponibili)
    )

    field = disponibili[misura]

    df = base[base.year >= 2022]

    if regione != "Italia":
        df = df[df.nuts2_code == regione]

    if (
        livello == "Annuale NUTS3"
        and provincia != "Tutte"
    ):
        df = df[df.nuts3_code == provincia]

    if df.empty:

        st.warning(
            "Nessun dato per il territorio."
        )
        st.stop()

    if livello == "Mensile NUTS2":

        serie = (
            df.groupby(
                "data",
                as_index=False
            )[field]
            .sum()
            .sort_values("data")
        )

        c1, c2 = st.columns(2)

        c1.metric(
            "Ultimo mese osservato",
            numero(serie[field].iloc[-1])
        )

        c2.metric(
            "Data ultima osservazione",
            serie.data.iloc[-1].strftime("%m/%Y")
        )

        fig = px.line(
            serie,
            x="data",
            y=field,
            color_discrete_sequence=[ACCENTO]
        )

        grafico(fig)

        st.subheader(
            "Stagionalità: confronta "
            "lo stesso mese tra anni"
        )

        stag = serie.copy()
        stag["anno"] = stag.data.dt.year.astype(str)
        stag["mese"] = stag.data.dt.month

        fig = px.line(
            stag,
            x="mese",
            y=field,
            color="anno",
            markers=True,
            color_discrete_sequence=[
                "#126889",
                "#249784",
                "#559EB8",
                "#DC813F",
                "#C94F7C"
            ]
        )

        fig.update_xaxes(dtick=1)
        grafico(fig)

        mesi26 = stag.loc[
            stag.anno == "2026",
            "mese"
        ].tolist()

        if mesi26:

            x26 = stag.loc[
                (stag.anno == "2026")
                & stag.mese.isin(mesi26),
                field
            ].sum()

            x25 = stag.loc[
                (stag.anno == "2025")
                & stag.mese.isin(mesi26),
                field
            ].sum()

            if x25 > 0:

                st.info(
                    "Confronto degli stessi mesi "
                    "2026/2025: "
                    + pct(100 * (x26 / x25 - 1))
                )

    else:

        serie = (
            df.groupby(
                "year",
                as_index=False
            )[field]
            .sum()
        )

        val = serie.set_index("year")[field]

        c1, c2 = st.columns(2)

        c1.metric(
            "Totale 2025",
            numero(
                val.get(2025, float("nan"))
            )
        )

        crescita = (
            pct(100 * (val[2025] / val[2024] - 1))
            if val.get(2024, 0) > 0
            else "—"
        )

        c2.metric(
            "Variazione 2025/2024",
            crescita
        )

        fig = px.bar(
            serie,
            x="year",
            y=field,
            color_discrete_sequence=[ACCENTO]
        )

        fig.update_xaxes(dtick=1)
        grafico(fig)

        st.caption(
            "NUTS3 annuale: non vengono generati "
            "dati mensili provinciali."
        )

    tabella(
        serie,
        "storico_eurostat.csv"
    )


# ============================================================
# 10. ANALISI STORICA SPORTIVA
# ============================================================

elif pagina == "Analisi storica sportiva":

    st.header(
        "Analisi storica degli eventi sportivi"
    )

    st.info(
        "Analisi descrittiva della dimensione "
        "degli eventi rispetto alla domanda regionale. "
        "Non è una stima causale e non si somma al forecast."
    )

    t1, t2, t3 = st.tabs([
        "Event Master",
        "Screening del 2%",
        "Eventi e domanda osservata"
    ])

    # --------------------------------------------------------
    # 10.A EVENT MASTER
    # --------------------------------------------------------

    with t1:

        if FILES["eventi"].exists():

            try:

                ev = excel(
                    str(FILES["eventi"])
                ).copy()

                sport = colonna(
                    ev, "Sport", "sport"
                )

                nome = colonna(
                    ev,
                    "Evento originale",
                    "Evento",
                    "event_name"
                )

                start = colonna(
                    ev,
                    "Data inizio",
                    "data_inizio",
                    "start_date"
                )

                area = colonna(
                    ev,
                    "Codici NUTS2",
                    "nuts2_code",
                    "nuts2"
                )

                eventid = colonna(
                    ev,
                    "event_id",
                    "id_evento"
                )

                if start:

                    ev["_data"] = pd.to_datetime(
                        ev[start],
                        errors="coerce",
                        dayfirst=True
                    )

                    ev = ev[
                        ev["_data"].dt.year >= 2022
                    ]

                if area and regione != "Italia":

                    ev = ev[
                        ev[area]
                        .astype(str)
                        .str.contains(
                            regione,
                            regex=False,
                            na=False
                        )
                    ]

                if sport:

                    opzioni = sorted(
                        ev[sport]
                        .dropna()
                        .astype(str)
                        .unique()
                    )

                    scelte = st.multiselect(
                        "Sport",
                        opzioni,
                        default=opzioni
                    )

                    ev = ev[
                        ev[sport].isin(scelte)
                    ]

                if eventid:

                    distinti = ev[eventid].nunique()

                elif nome and start:

                    distinti = len(
                        ev.drop_duplicates(
                            [nome, start]
                        )
                    )

                else:

                    distinti = None

                a, b = st.columns(2)

                a.metric(
                    "Eventi distinti",
                    numero(distinti)
                    if distinti is not None
                    else "N/D"
                )

                b.metric(
                    "Righe territoriali",
                    numero(len(ev))
                )

                if sport and not ev.empty:

                    counts = (
                        ev.groupby(
                            sport,
                            as_index=False
                        )
                        .size()
                    )

                    fig = px.bar(
                        counts,
                        x=sport,
                        y="size",
                        color_discrete_sequence=[ACCENTO]
                    )

                    grafico(fig, 320)

                tabella(
                    ev.drop(
                        columns=["_data"],
                        errors="ignore"
                    ),
                    "event_master_filtrato.csv"
                )

            except Exception as exc:

                st.error(
                    f"Problema lettura Event Master: {exc}"
                )

        else:

            st.warning(
                "Event Master non trovato: "
                + FILES["eventi"].name
            )

    # --------------------------------------------------------
    # 10.B SCREENING DEL 2%
    # --------------------------------------------------------

    with t2:

        st.write(
            "Lo screening confronta "
            "**pernottamenti teorici** e guest nights "
            "NUTS2 nello stesso mese. "
            "La soglia **2%** è esplorativa, "
            "non statistica."
        )

        st.caption(
            "Preset considerati nell'analisi: "
            "5%, 15%, 25%, 40%, 60% di pernottanti. "
            "Le notti sono limitate dalla durata "
            "dell'evento. Nessun effetto causale implicito."
        )

        if FILES["screening"].exists():

            screen = csv(
                str(FILES["screening"])
            ).copy()

            loc = colonna(
                screen,
                "nuts2_code",
                "nuts2",
                "codici_nuts2"
            )

            if loc and regione != "Italia":

                screen = screen[
                    screen[loc]
                    .astype(str)
                    .str.contains(
                        regione,
                        regex=False,
                        na=False
                    )
                ]

            st.metric(
                "Righe disponibili nello screening",
                numero(len(screen))
            )

            suggest = [
                c for c in screen
                if any(
                    q in c.lower()
                    for q in (
                        "incidenza",
                        "rapporto",
                        "pct",
                        "percent",
                        "peso"
                    )
                )
            ]

            chosen = st.selectbox(
                "Colonna percentuale del file "
                "(facoltativo)",
                ["Non selezionare"] + suggest
            )

            if chosen != "Non selezionare":

                scala = st.radio(
                    "Unità della colonna",
                    [
                        "Percentuale: 2,5",
                        "Frazione: 0,025"
                    ],
                    horizontal=True
                )

                vals = pd.to_numeric(
                    screen[chosen],
                    errors="coerce"
                )

                if scala.startswith("Frazione"):
                    vals = vals * 100

                screen = screen.assign(
                    incidenza_pct=vals
                )

                st.metric(
                    "Righe con incidenza teorica ≥ 2%",
                    numero((vals >= 2).sum())
                )

                if st.checkbox(
                    "Solo righe sopra soglia"
                ):

                    screen = screen[
                        screen.incidenza_pct >= 2
                    ]

            st.caption(
                "Selezionare la colonna che misura "
                "il rapporto evento/domanda. "
                "Una percentuale diversa non "
                "rappresenta lo screening del 2%."
            )

            tabella(
                screen,
                "screening_eventi_filtrato.csv"
            )

        else:

            st.warning(
                "screening_eventi_completo.csv "
                "non trovato. Nessun risultato inventato."
            )

    # --------------------------------------------------------
    # 10.C EVENTI E DOMANDA OSSERVATA
    # --------------------------------------------------------

    with t3:

        sel = (
            regione
            if regione != "Italia"
            else st.selectbox(
                "Regione NUTS2",
                sorted(names2),
                format_func=lambda c: names2[c]
            )
        )

        anno = st.selectbox(
            "Anno storico",
            [2022, 2023, 2024, 2025],
            index=3
        )

        mese = st.selectbox(
            "Mese",
            list(range(1, 13)),
            index=4
        )

        selreg = reg[
            (reg.nuts2_code == sel)
            & (reg.month == mese)
            & reg.year.between(2022, 2025)
        ]

        ita = reg[
            (reg.month == mese)
            & reg.year.between(2022, 2025)
        ]

        regionale = (
            selreg.set_index("year")[
                "guest_nights_total"
            ]
        )

        nazionale = (
            ita.groupby("year")[
                "guest_nights_total"
            ].sum()
        )

        confronti = pd.DataFrame({
            "Regione": regionale,
            "Resto Italia": nazionale - regionale
        }).sort_index()

        yoy = confronti.pct_change() * 100

        fig = go.Figure()

        for campo, colore in [
            ("Regione", ACCENTO),
            ("Resto Italia", COLORI["Intermedio"])
        ]:

            fig.add_trace(
                go.Scatter(
                    x=yoy.index,
                    y=yoy[campo],
                    name=campo + " YoY %",
                    mode="lines+markers",
                    line=dict(
                        color=colore,
                        width=3
                    )
                )
            )

        grafico(fig, 350)

        if anno in confronti.index:

            st.metric(
                f"{names2[sel]}, "
                f"{mese:02}/{anno} · guest nights",
                numero(
                    confronti.loc[anno, "Regione"]
                )
            )

        tabella(
            confronti.reset_index(
                names="year"
            ),
            "domanda_mese_territorio.csv"
        )

        if FILES["mesi_eventi"].exists():

            em = csv(
                str(FILES["mesi_eventi"])
            )

            rc = colonna(
                em,
                "nuts2_code",
                "nuts2"
            )

            yc = colonna(
                em,
                "year",
                "anno"
            )

            mc = colonna(
                em,
                "month",
                "mese"
            )

            if rc and yc and mc:

                mask = (
                    em[rc]
                    .astype(str)
                    .str.contains(
                        sel,
                        regex=False,
                        na=False
                    )
                    & pd.to_numeric(
                        em[yc],
                        errors="coerce"
                    ).eq(anno)
                    & pd.to_numeric(
                        em[mc],
                        errors="coerce"
                    ).eq(mese)
                )

                st.subheader(
                    "Eventi associati al mese selezionato"
                )

                tabella(
                    em.loc[mask]
                )

        st.caption(
            "La differenza tra regione e resto d'Italia "
            "è descrittiva: non identifica da sola "
            "un effetto causale degli eventi."
        )


# ============================================================
# 11. FORECAST 2026-2030
# ============================================================

elif pagina == "Forecast 2026–2030":

    st.header(
        "Forecast 2026–2030 · dal 2022"
    )

    attivi = st.multiselect(
        "Scenari nel grafico principale",
        SCENARI,
        default=SCENARI
    )

    fig = go.Figure()

    if provincia != "Tutte":

        # Serie provinciale annuale
        storico = prov[
            (prov.nuts3_code == provincia)
            & (prov.year >= 2022)
        ].sort_values("year")

        futuro = fp[
            fp.nuts3_code == provincia
        ].sort_values("year")

        xx = storico.year.tolist()
        yy = storico.guest_nights_total.tolist()

        st.subheader(
            f"{names3[provincia]} · NUTS3 annuale"
        )

        for s in attivi:

            fig.add_trace(
                go.Scatter(
                    x=[xx[-1]] + futuro.year.tolist(),
                    y=[yy[-1]]
                    + futuro[
                        f"forecast_{s.lower()}"
                    ].tolist(),
                    name=s,
                    mode="lines+markers",
                    line=dict(
                        color=COLORI[s],
                        dash="dash",
                        width=3
                    )
                )
            )

    else:

        # Serie regionale / nazionale mensile
        base = (
            reg
            if regione == "Italia"
            else reg[reg.nuts2_code == regione]
        )

        storico = (
            base[base.year >= 2022]
            .groupby(
                "data",
                as_index=False
            )["guest_nights_total"]
            .sum()
            .sort_values("data")
        )

        future_base = (
            fr
            if regione == "Italia"
            else fr[fr.nuts2_code == regione]
        )

        future_base = future_base[
            future_base.data > storico.data.max()
        ]

        cols = [
            f"gn_{s.lower()}"
            for s in SCENARI
        ]

        futuro = (
            future_base.groupby(
                "data",
                as_index=False
            )[cols]
            .sum()
            .sort_values("data")
        )

        xx = storico.data.tolist()
        yy = storico.guest_nights_total.tolist()

        titolo = (
            "Italia"
            if regione == "Italia"
            else names2[regione]
        )

        st.subheader(
            titolo + " · NUTS2 mensile"
        )

        for s in attivi:

            fig.add_trace(
                go.Scatter(
                    x=[xx[-1]] + futuro.data.tolist(),
                    y=[yy[-1]]
                    + futuro[
                        f"gn_{s.lower()}"
                    ].tolist(),
                    name=s,
                    mode="lines",
                    line=dict(
                        color=COLORI[s],
                        dash="dash",
                        width=3
                    )
                )
            )

    # Dati osservati
    fig.add_trace(
        go.Scatter(
            x=xx,
            y=yy,
            name="Osservato",
            mode=(
                "lines+markers"
                if provincia != "Tutte"
                else "lines"
            ),
            line=dict(
                color=ACCENTO,
                width=3
            )
        )
    )

    fig.update_yaxes(
        title="Guest nights"
    )

    grafico(fig, 480)

    st.caption(
        "Serie osservata dal 2022. "
        "NUTS2 2026: gennaio–marzo osservati, "
        "restanti mesi previsti. "
        "Per le province sono disponibili "
        "solo dati annuali."
    )

    # KPI e confronto annuale
    cards = st.columns(3)

    yearly = pd.DataFrame({
        "Anno": list(range(2026, 2031))
    })

    for card, s in zip(cards, SCENARI):

        if provincia != "Tutte":
            data = fp[
                fp.nuts3_code == provincia
            ]
            campo = f"forecast_{s.lower()}"

        else:
            data = (
                fr if regione == "Italia"
                else fr[
                    fr.nuts2_code == regione
                ]
            )
            campo = f"gn_{s.lower()}"

        annual = (
            data.groupby("year")[campo]
            .sum()
        )

        yearly[s] = yearly.Anno.map(annual)

        card.metric(
            "2030 · " + s,
            numero(annual.loc[2030])
        )

    st.subheader(
        "Confronto annuale dei tre scenari"
    )

    fig = go.Figure()

    for s in SCENARI:

        fig.add_trace(
            go.Scatter(
                x=yearly.Anno,
                y=yearly[s],
                name=s,
                mode="lines+markers",
                line=dict(
                    color=COLORI[s],
                    width=3
                )
            )
        )

    grafico(fig, 330)

    tabella(
        yearly,
        "forecast_tre_scenari.csv"
    )

    # Mappa forecast
    st.divider()

    anno_m = st.select_slider(
        "Anno mappa",
        options=[
            2026, 2027, 2028, 2029, 2030
        ],
        value=2030
    )

    livello_m = st.radio(
        "Geografia mappa",
        ["NUTS2", "NUTS3"],
        horizontal=True
    )

    scenario_m = st.selectbox(
        "Scenario mappa",
        SCENARI,
        index=SCENARI.index(scenario)
    )

    if livello_m == "NUTS2":

        campo = f"gn_{scenario_m.lower()}"

        df = (
            fr[fr.year == anno_m]
            .groupby(
                "nuts2_code",
                as_index=False
            )[campo]
            .sum()
        )

        df["territorio"] = (
            df.nuts2_code.map(names2)
        )

        mappa(
            df,
            2,
            "nuts2_code",
            campo,
            f"Previsione {anno_m} · "
            f"{scenario_m} · NUTS2"
        )

    else:

        campo = (
            f"forecast_{scenario_m.lower()}"
        )

        df = fp.loc[
            fp.year == anno_m,
            ["nuts3_code", campo]
        ].copy()

        df["territorio"] = (
            df.nuts3_code.map(names3)
        )

        mappa(
            df,
            3,
            "nuts3_code",
            campo,
            f"Previsione {anno_m} · "
            f"{scenario_m} · NUTS3"
        )


# ============================================================
# 12. METODOLOGIA E CONTROLLI
# ============================================================

else:

    st.header(
        "Metodologia e controlli"
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "Regioni NUTS2",
        reg.nuts2_code.nunique()
    )

    b.metric(
        "Province NUTS3",
        prov.nuts3_code.nunique()
    )

    c.metric(
        "Mesi NUTS2",
        reg[
            ["year", "month"]
        ].drop_duplicates().shape[0]
    )

    d.metric(
        "Anni NUTS3",
        prov.year.nunique()
    )

    checks = [
        (
            "21 NUTS2",
            reg.nuts2_code.nunique() == 21
        ),
        (
            "107 NUTS3",
            prov.nuts3_code.nunique() == 107
        ),
        (
            "1260 righe forecast NUTS2",
            len(fr) == 1260
        ),
        (
            "535 righe forecast NUTS3",
            len(fp) == 535
        )
    ]

    for s in SCENARI:

        a = (
            fr.groupby(
                ["nuts2_code", "year"]
            )[f"gn_{s.lower()}"]
            .sum()
            .round()
        )

        b = (
            fp.groupby(
                ["nuts2_code", "year"]
            )[f"forecast_{s.lower()}"]
            .sum()
        )

        errore = (
            a.sub(b, fill_value=0)
            .abs()
            .max()
        )

        checks.append(
            (
                f"Riconciliazione NUTS3/NUTS2 {s}",
                errore <= 1
            )
        )

    tabella(
        pd.DataFrame([
            {
                "Controllo": nome,
                "Esito": (
                    "OK"
                    if valido
                    else "DA VERIFICARE"
                )
            }
            for nome, valido in checks
        ])
    )

    st.subheader(
        "Decisioni del progetto"
    )

    st.write(
        "**Periodo 2022–2025:** riferimento "
        "post-Covid. I mesi osservati del 2026 "
        "sono incorporati nel forecast regionale."
    )

    st.write(
        "**Eurostat:** `tour_ce_omn12` "
        "per NUTS2 mensile; `tour_ce_oan3` "
        "per NUTS3 annuale. "
        "Non misura i prezzi Airbnb."
    )

    st.write(
        "**Forecast:** stagionalità 2025, "
        "tasso iniziale 50% nazionale "
        "e 50% regionale, tre ipotesi "
        "di rallentamento della crescita. "
        "NUTS3 ripartito secondo quote 2025."
    )

    st.write(
        "**Sport:** catalogo e screening "
        "teorico del 2%, senza imputare "
        "effetti causali né sommarli al forecast."
    )

    st.subheader(
        "Backtest documentati"
    )

    st.caption(
        "Risultati precedenti riportati come "
        "riferimento, non ricalcolati all'avvio."
    )

    backtest = pd.DataFrame(
        [
            (
                "Baseline stagionale 2024",
                "11,11%",
                "Test storico"
            ),
            (
                "Media 2023–2024",
                "17,13%",
                "Test storico"
            ),
            (
                "Scenario prudente",
                "7,48%",
                "Test 2025"
            ),
            (
                "Scenario intermedio",
                "8,39%",
                "Test 2025"
            ),
            (
                "Scenario espansivo",
                "9,89%",
                "Test 2025"
            )
        ],
        columns=[
            "Metodo",
            "WAPE",
            "Campione"
        ]
    )

    tabella(backtest)


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Airbnb&Sports · " + VERSIONE
    + " · Eurostat, screening eventi "
    + "e tre scenari distinti"
)
