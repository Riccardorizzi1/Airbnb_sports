
# AIRBNB&SPORTS - STREAMLIT + KEPLER.GL
# File principale per GitHub e Streamlit Community Cloud: app.py

from pathlib import Path
from urllib.request import Request, urlopen
import json
import math
import re
import tempfile

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components
try:
    from keplergl import KeplerGl
except ImportError:
    KeplerGl = None

st.set_page_config(
    page_title="Airbnb&Sports",
    page_icon="🏟️",
    layout="wide"
)

VERSIONE = "v3.3 · Revisione dashboard"
ROOT = Path(__file__).resolve().parent


# ============================================================
# 1. PERCORSI E FILE DEL PROGETTO
# ============================================================

def file_dati(nome):
    """
    Cerca prima nella cartella principale del repository
    e poi nella sottocartella output_eurostat.
    """
    diretto = ROOT / nome
    annidato = ROOT / "output_eurostat" / nome

    return diretto if diretto.exists() else annidato


FILES = {
    "reg": file_dati(
        "eurostat_master_final.csv"
    ),
    "prov": file_dati(
        "eurostat_nuts3_annuale_2018_2025.csv"
    ),
    "fr": file_dati(
        "forecast_nuts2_2026_2030_3_scenari.csv"
    ),
    "fp": file_dati(
        "forecast_nuts3_2026_2030_3_scenari.csv"
    ),
    "eventi": ROOT / (
        "AirbnbSports_Event_Master_Simone_Visitatori.xlsx"
    ),
    "screening": ROOT / "screening_eventi_completo.csv",
    "mesi_eventi": ROOT / "screening_regioni_mesi_completo.csv",
}

SCENARI = [
    "Prudente",
    "Intermedio",
    "Espansivo"
]

COLORI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C"
}

PAGINE = [
    "Panorama nazionale",
    "Esplora un territorio",
    "Eventi e domanda turistica",
    "Forecast 2026–2030",
    "Dati e metodo"
]


# ============================================================
# 2. LETTURA DEI DATI
# ============================================================

@st.cache_data(show_spinner=False)
def leggi_csv(path):
    return pd.read_csv(
        path,
        encoding="utf-8-sig",
        low_memory=False
    )


@st.cache_data(show_spinner=False)
def leggi_excel(path):
    return pd.read_excel(
        path,
        sheet_name=0
    )


@st.cache_data(ttl=86400, show_spinner=False)
def confini_nuts(livello):

    url = (
        "https://gisco-services.ec.europa.eu/"
        "distribution/v2/nuts/geojson/"
        f"NUTS_RG_20M_2024_4326_LEVL_{livello}.geojson"
    )

    request = Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"}
    )

    with urlopen(request, timeout=35) as response:
        geo = json.load(response)

    features = []

    for feature in geo.get("features", []):

        codice = str(
            feature.get("properties", {}).get(
                "NUTS_ID", ""
            )
        )

        if (
            codice.startswith("IT")
            and len(codice) == livello + 2
        ):
            features.append(feature)

    if not features:
        raise ValueError(
            "Confini NUTS italiani non disponibili"
        )

    return features


# ============================================================
# 3. FUNZIONI GENERALI
# ============================================================

def numero(valore, decimali=0):

    if pd.isna(valore):
        return "—"

    return (
        f"{float(valore):,.{decimali}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def percentuale(valore):

    if pd.isna(valore):
        return "—"

    return f"{valore:+.2f}%".replace(".", ",")


def trova_colonna(df, *nomi):

    disponibili = {
        re.sub(
            r"[^a-z0-9]",
            "",
            str(c).lower()
        ): c
        for c in df.columns
    }

    for nome in nomi:

        key = re.sub(
            r"[^a-z0-9]",
            "",
            nome.lower()
        )

        col = disponibili.get(key)

        if col is not None:
            return col

    return None


def date_eventi(serie):
    """Date ISO, italiane, datetime Excel; valori non interpretabili diventano NaT."""
    def converti(valore):
        if pd.isna(valore):
            return pd.NaT
        if isinstance(valore, (int, float)):
            return pd.to_datetime(valore, unit="D", origin="1899-12-30", errors="coerce") if 20000 < valore < 80000 else pd.NaT
        testo = str(valore).strip()
        iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", testo))
        return pd.to_datetime(testo, errors="coerce", dayfirst=not iso)
    return pd.to_datetime(serie.map(converti), errors="coerce")


def grafico(fig, altezza=410):

    fig.update_layout(
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=TESTO,
        hovermode="x unified",
        margin=dict(l=10, r=10, t=40, b=15),
        legend=dict(
            orientation="h",
            y=1.14
        )
    )

    etichette = {
        "data": "Mese", "year": "Anno", "month": "Mese", "mese": "Mese",
        "guest_nights_total": "Pernottamenti degli ospiti",
        "guest_nights_domestic": "Pernottamenti · residenti in Italia",
        "guest_nights_foreign": "Pernottamenti · residenti all’estero",
        "occupied_nights_total": "Notti di alloggio occupate",
        "stays_total": "Soggiorni", "size": "Righe territoriali"
    }
    for asse in (fig.layout.xaxis, fig.layout.yaxis):
        if asse.title.text in etichette:
            asse.title.text = etichette[asse.title.text]
    fig.update_xaxes(
        gridcolor="rgba(120,150,170,0.12)"
    )

    fig.update_yaxes(
        gridcolor="rgba(120,150,170,0.20)",
        tickformat=",~s"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )


def tabella(df, nome=None):

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )

    if nome is not None:

        st.download_button(
            "Scarica CSV",
            df.to_csv(
                index=False,
                sep=";"
            ).encode("utf-8-sig"),
            file_name=nome,
            mime="text/csv"
        )


# ============================================================
# 4. MAPPE KEPLER.GL
# ============================================================

def limiti(geometria, x, y):

    if not isinstance(
        geometria,
        (list, tuple)
    ):
        return

    if (
        len(geometria) >= 2
        and all(
            isinstance(v, (int, float))
            for v in geometria[:2]
        )
    ):

        x.append(float(geometria[0]))
        y.append(float(geometria[1]))

    else:

        for elemento in geometria:
            limiti(elemento, x, y)


def mappa_kepler(df, livello, codice, campo, titolo):

    st.subheader(titolo)

    dati = df[
        [codice, "territorio", campo]
    ].copy()

    dati[codice] = dati[codice].astype(str)

    dati[campo] = pd.to_numeric(
        dati[campo],
        errors="coerce"
    )

    dati = dati.dropna(
        subset=[campo]
    )

    if (
        dati.empty
        or dati[codice].duplicated().any()
    ):

        st.warning(
            "Mappa non disponibile: dati mancanti "
            "o codici NUTS duplicati."
        )
        return

    mostra = st.checkbox(
        "Mostra mappa Kepler.gl",
        value=True,
        key=f"mostra_{titolo}"
    )

    motore = st.radio(
        "Visualizzazione della mappa", ["Kepler.gl", "Mappa semplice (senza WebGL)"],
        index=0 if KeplerGl is not None else 1, horizontal=True, key=f"motore_{titolo}"
    ) if mostra else None
    if mostra and motore == "Mappa semplice (senza WebGL)":
        try:
            geo = {"type": "FeatureCollection", "features": confini_nuts(livello)}
            fig = px.choropleth(
                dati, geojson=geo, locations=codice,
                featureidkey="properties.NUTS_ID", color=campo,
                hover_name="territorio", color_continuous_scale=["#C7F1F6", "#5BC6E2", "#0E718F"],
                labels={campo: "Pernottamenti degli ospiti"}
            )
            fig.update_geos(fitbounds="locations", visible=False, bgcolor="rgba(0,0,0,0)")
            grafico(fig, 600)
            st.caption("Geometrie ufficiali GISCO NUTS 2024; nessuna coordinata artificiale.")
        except Exception as exc:
            st.warning(f"Mappa non disponibile: {exc}. Consulta la tabella sottostante.")
    if mostra and motore == "Kepler.gl" and KeplerGl is None:
        st.info("Kepler.gl non è installato. Seleziona la mappa semplice.")
    if mostra and motore == "Kepler.gl" and KeplerGl is not None:
        st.caption("Se il riquadro è vuoto o segnala WebGL, scegli la mappa semplice.")
        editor = st.toggle(
            "Mostra strumenti Kepler",
            value=False,
            key=f"editor_{titolo}"
        )

        try:

            valori = (
                dati.set_index(codice)
                .to_dict("index")
            )

            features = []
            xs = []
            ys = []

            for feature in confini_nuts(livello):

                nuts = feature[
                    "properties"
                ]["NUTS_ID"]

                if nuts not in valori:
                    continue

                record = valori[nuts]

                features.append({
                    "type": "Feature",
                    "geometry": feature["geometry"],
                    "properties": {
                        "nuts_code": nuts,
                        "territorio": str(
                            record["territorio"]
                        ),
                        "guest_nights": float(
                            record[campo]
                        ),
                        "guest_nights_it": numero(
                            record[campo]
                        ),
                    }
                })

                limiti(
                    feature["geometry"]["coordinates"],
                    xs,
                    ys
                )

            if not features:

                raise ValueError(
                    "Nessun codice NUTS dei dati "
                    "corrisponde ai confini GISCO."
                )

            ampiezza = max(
                max(xs) - min(xs),
                1.4 * (max(ys) - min(ys)),
                0.1
            )

            zoom = max(
                3.0,
                min(
                    11.0,
                    math.log2(360 / ampiezza) - 0.9
                )
            )

            data_id = "AirbnbSports_NUTS"

            config = {
                "version": "v1",
                "config": {

                    "mapState": {
                        "longitude": (
                            min(xs) + max(xs)
                        ) / 2,

                        "latitude": (
                            min(ys) + max(ys)
                        ) / 2,

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
                            "id": "nuts_polygons",
                            "type": "geojson",

                            "config": {
                                "dataId": data_id,

                                "label": (
                                    f"Guest nights NUTS{livello}"
                                ),

                                "color": [
                                    40, 165, 199
                                ],

                                "columns": {
                                    "geojson": "_geojson"
                                },

                                "isVisible": True,

                                "visConfig": {
                                    "opacity": 0.85,
                                    "strokeOpacity": 0.80,
                                    "thickness": 0.6,

                                    "colorRange": {
                                        "name": "AirbnbSports Blu",
                                        "type": "sequential",
                                        "category": "Custom",

                                        "colors": [
                                            "#C2F0F6",
                                            "#88D6E7",
                                            "#4BAFC9",
                                            "#126982",
                                            "#084A68"
                                        ]
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
                                    data_id: [
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

            kepler = KeplerGl(
                height=680,
                data={
                    data_id: {
                        "type": "FeatureCollection",
                        "features": features
                    }
                },
                config=config
            )

            with tempfile.TemporaryDirectory(
                prefix="airbnbsports_"
            ) as tmp:

                pagina = Path(tmp) / "mappa.html"

                kepler.save_to_html(
                    file_name=str(pagina),
                    read_only=not editor
                )

                html = pagina.read_text(
                    encoding="utf-8"
                )

            components.html(
                html,
                height=700,
                scrolling=False
            )

            st.caption(
                f"Kepler.gl · {len(features)} territori · "
                "confini ufficiali GISCO NUTS 2024"
            )

        except Exception as exc:

            st.warning(
                f"Mappa Kepler non disponibile: {exc}"
            )

    tabella(
        dati.sort_values(
            campo,
            ascending=False
        )
    )


# ============================================================
# 5. SIDEBAR E TEMA
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


# ============================================================
# 6. CSS - IDENTITA GRAFICA
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

    box-shadow:
        0 5px 14px {OMBRA} !important;

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

section[data-testid="stSidebar"] [data-baseweb="select"] * {{
    color: #082B46 !important;
}}
section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background-color: #FDFEFF !important;
}}
</style>
""", unsafe_allow_html=True)


# ============================================================
# 7. CARICAMENTO MASTER E FORECAST
# ============================================================

# Tutte e quattro le chiavi sono definite in FILES.
obbligatori = (
    "reg",
    "prov",
    "fr",
    "fp"
)

mancanti = [
    str(FILES[k])
    for k in obbligatori
    if not FILES[k].is_file()
]

if mancanti:

    st.error(
        "File CSV mancanti nel repository:"
    )

    st.code("\n".join(mancanti))
    st.stop()


try:

    reg, prov, fr, fp = [
        leggi_csv(str(FILES[k]))
        for k in obbligatori
    ]

    controlli = [
        (
            reg,
            [
                "year", "month",
                "nuts2_code",
                "nuts2_name",
                "guest_nights_total"
            ]
        ),
        (
            prov,
            [
                "year",
                "nuts3_code",
                "nuts3_name",
                "nuts2_code",
                "guest_nights_total"
            ]
        ),
        (
            fr,
            [
                "year",
                "month",
                "nuts2_code"
            ] + [
                f"gn_{s.lower()}"
                for s in SCENARI
            ]
        ),
        (
            fp,
            [
                "year",
                "nuts3_code",
                "nuts2_code"
            ] + [
                f"forecast_{s.lower()}"
                for s in SCENARI
            ]
        )
    ]

    for df, colonne in controlli:

        assenti = sorted(
            set(colonne) - set(df.columns)
        )

        if assenti:
            raise ValueError(
                f"Colonne mancanti: {assenti}"
            )

        df["year"] = pd.to_numeric(
            df["year"],
            errors="raise"
        ).astype(int)

    for df in (reg, fr):

        df["month"] = pd.to_numeric(
            df["month"],
            errors="raise"
        ).astype(int)

        df["data"] = pd.to_datetime(
            dict(
                year=df.year,
                month=df.month,
                day=1
            )
        )

except Exception as exc:

    st.error(
        f"Errore nella lettura dei CSV: {exc}"
    )
    st.stop()


nomi_reg = (
    reg.drop_duplicates("nuts2_code")
    .set_index("nuts2_code")["nuts2_name"]
    .to_dict()
)

nomi_prov = (
    prov.drop_duplicates("nuts3_code")
    .set_index("nuts3_code")["nuts3_name"]
    .to_dict()
)


# ============================================================
# 8. FILTRI TERRITORIALI
# ============================================================

with st.sidebar:

    st.subheader("Selezione territoriale")

    regione = st.selectbox(
        "Regione NUTS2",
        ["Italia"] + sorted(nomi_reg),

        format_func=lambda x: (
            "Italia — tutte le regioni" if x == "Italia"
            else f"{nomi_reg[x]} · {x}"
        )
    )

    province_ammissibili = (
        prov if regione == "Italia"
        else prov[
            prov.nuts2_code == regione
        ]
    )

    provincia = st.selectbox(
        "Provincia NUTS3",
        ["Tutte"] + sorted(
            province_ammissibili.nuts3_code.unique()
        ),

        format_func=lambda x: (
            "Tutte le province" if x == "Tutte"
            else f"{nomi_prov[x]} · {x}"
        )
    )

    st.divider()

    scenario = "Intermedio"
    if pagina in ("Panorama nazionale", "Forecast 2026–2030"):
        scenario = st.selectbox("Scenario predefinito", SCENARI, index=1)
    if pagina in ("Panorama nazionale", "Dati e metodo"):
        st.caption("Questa pagina mostra il quadro nazionale; i filtri territoriali valgono nelle altre sezioni.")


# ============================================================
# 9. HEADER
# ============================================================

st.title("Airbnb&Sports")

st.caption(
    "Domanda di short-term accommodation "
    "tramite piattaforme online · Italia"
)


# ============================================================
# 10. PANORAMA NAZIONALE
# ============================================================

if pagina == "Panorama nazionale":

    st.header(
        "Panorama nazionale · periodo 2022–2025"
    )

    annual = (
        reg[reg.year.between(2022, 2025)]
        .groupby("year", as_index=False)
        .guest_nights_total.sum()
    )

    valori = (
        annual.set_index("year")
        .guest_nights_total
    )

    v24 = valori[2024]
    v25 = valori[2025]

    v30 = fr.loc[
        fr.year == 2030,
        f"gn_{scenario.lower()}"
    ].sum()

    a, b, c, d = st.columns(4)

    a.metric(
        "Guest nights 2025",
        numero(v25)
    )

    b.metric(
        "Crescita 2025/2024",
        percentuale(
            100 * (v25 / v24 - 1)
        )
    )

    c.metric(
        f"Forecast 2030 · {scenario}",
        numero(v30)
    )

    d.metric(
        "Territori",
        "21 NUTS2 · 107 NUTS3"
    )

    fig = px.bar(
        annual,
        x="year",
        y="guest_nights_total",
        color_discrete_sequence=[ACCENTO]
    )

    fig.update_xaxes(dtick=1)

    grafico(fig, 360)

    st.caption(
        "2020–2021 esclusi dalla visualizzazione ordinaria "
        "a causa dello shock Covid."
    )

    geo = (
        reg[reg.year == 2025]
        .groupby(
            "nuts2_code",
            as_index=False
        )
        .guest_nights_total.sum()
    )

    geo["territorio"] = (
        geo.nuts2_code.map(nomi_reg)
    )

    mappa_kepler(
        geo,
        2,
        "nuts2_code",
        "guest_nights_total",
        "Domanda regionale osservata · 2025"
    )


# ============================================================
# 11. ANALISI STORICA EUROSTAT
# ============================================================

elif pagina == "Esplora un territorio":

    st.header(
        "Analisi storica Eurostat · dal 2022"
    )

    livello = st.radio(
        "Frequenza",
        [
            "Mensile NUTS2",
            "Annuale NUTS3"
        ],
        horizontal=True
    )

    territorio = "Italia" if regione == "Italia" else nomi_reg[regione]
    if livello == "Annuale NUTS3" and provincia != "Tutte":
        territorio = nomi_prov[provincia]
    st.subheader(territorio)
    if livello == "Mensile NUTS2" and provincia != "Tutte":
        st.caption("La vista mensile usa la regione selezionata (Italia se non selezionata). Per la provincia scegli Annuale NUTS3.")
    st.caption("Guest nights = pernottamenti degli ospiti; stays = soggiorni; notti occupate = notti di alloggio. Eurostat non misura il prezzo Airbnb.")
    indicatori = {
        "Guest nights totali": "guest_nights_total",
        "Pernottamenti domestici": "guest_nights_domestic",
        "Pernottamenti stranieri": "guest_nights_foreign",
        "Stays": "stays_total",
        "Notti occupate": "occupied_nights_total"
    }

    base = (
        reg if livello == "Mensile NUTS2"
        else prov
    )

    opzioni = {
        nome: col
        for nome, col in indicatori.items()
        if col in base.columns
    }

    misura = st.selectbox(
        "Indicatore Eurostat",
        list(opzioni)
    )

    campo = opzioni[misura]

    dati = base[
        base.year >= 2022
    ]

    if regione != "Italia":

        dati = dati[
            dati.nuts2_code == regione
        ]

    if (
        livello == "Annuale NUTS3"
        and provincia != "Tutte"
    ):

        dati = dati[
            dati.nuts3_code == provincia
        ]

    if dati.empty:

        st.info(
            "Nessun dato per la selezione corrente."
        )

    elif livello == "Mensile NUTS2":

        serie = (
            dati.groupby(
                "data",
                as_index=False
            )[campo]
            .sum()
            .sort_values("data")
        )

        a, b = st.columns(2)

        a.metric(
            "Ultimo mese osservato",
            numero(serie[campo].iloc[-1])
        )

        b.metric(
            "Mese",
            serie.data.iloc[-1].strftime("%m/%Y")
        )

        fig = px.line(
            serie,
            x="data",
            y=campo,
            color_discrete_sequence=[ACCENTO]
        )

        grafico(fig)

        st.subheader(
            "Stagionalità · stessi mesi tra anni"
        )

        stag = serie.copy()

        stag["Anno"] = (
            stag.data.dt.year.astype(str)
        )

        stag["Mese"] = (
            stag.data.dt.month
        )

        fig = px.line(
            stag,
            x="Mese",
            y=campo,
            color="Anno",
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

        mesi26 = sorted(
            stag.loc[
                stag.Anno == "2026",
                "Mese"
            ].unique()
        )

        if len(mesi26):

            g26 = stag[
                (stag.Anno == "2026")
                & stag.Mese.isin(mesi26)
            ][campo].sum()

            g25 = stag[
                (stag.Anno == "2025")
                & stag.Mese.isin(mesi26)
            ][campo].sum()

            if g25 > 0:

                st.info(
                    "Variazione 2026/2025 sugli stessi mesi: "
                    + percentuale(
                        100 * (g26 / g25 - 1)
                    )
                )

        tabella(
            serie,
            "storico_nuts2.csv"
        )

    else:

        serie = (
            dati.groupby(
                "year",
                as_index=False
            )[campo]
            .sum()
            .sort_values("year")
        )

        valori = (
            serie.set_index("year")[campo]
        )

        a, b = st.columns(2)

        a.metric(
            "Totale 2025",
            numero(
                valori.get(
                    2025,
                    float("nan")
                )
            )
        )

        yoy = (
            100 * (
                valori[2025] / valori[2024] - 1
            )
            if valori.get(2024, 0) > 0
            else float("nan")
        )

        b.metric(
            "2025/2024",
            percentuale(yoy)
        )

        fig = px.bar(
            serie,
            x="year",
            y=campo,
            color_discrete_sequence=[ACCENTO]
        )

        fig.update_xaxes(dtick=1)

        grafico(fig)

        st.caption(
            "Le osservazioni NUTS3 sono annuali, "
            "non mensili."
        )

        tabella(
            serie,
            "storico_nuts3.csv"
        )


# ============================================================
# 12. ANALISI STORICA SPORTIVA
# ============================================================

elif pagina == "Eventi e domanda turistica":

    st.header(
        "Analisi storica degli eventi sportivi"
    )

    st.info(
        "Analisi descrittiva: lo screening del 2% "
        "non dimostra un effetto causale e non "
        "viene sommato al forecast."
    )

    tab1, tab2, tab3 = st.tabs([
        "Event Master",
        "Screening del 2%",
        "Eventi e domanda osservata"
    ])

    # --------------------------------------------------------
    # EVENT MASTER
    # --------------------------------------------------------

    with tab1:

        if FILES["eventi"].is_file():

            eventi = leggi_excel(
                str(FILES["eventi"])
            ).copy()

            sport = trova_colonna(
                eventi,
                "Sport",
                "sport"
            )

            nome = trova_colonna(
                eventi,
                "Evento originale",
                "Evento",
                "event_name"
            )

            data = trova_colonna(
                eventi,
                "Data inizio",
                "data_inizio",
                "start_date"
            )

            area = trova_colonna(
                eventi,
                "Codici NUTS2",
                "nuts2_code",
                "nuts2"
            )

            identificativo = trova_colonna(
                eventi,
                "event_id",
                "id_evento"
            )

            if data:

                eventi["_data"] = date_eventi(eventi[data])
                senza_data = eventi["_data"].isna()
                if senza_data.any():
                    st.warning(f"{senza_data.sum()} righe hanno una data iniziale mancante o non leggibile: restano consultabili, ma non sono collocabili nel calendario.")
                fine = trova_colonna(eventi, "Data fine", "data_fine", "end_date")
                termine = date_eventi(eventi[fine]) if fine else eventi["_data"]
                termine = termine.fillna(eventi["_data"])
                eventi = eventi[senza_data | (termine >= pd.Timestamp("2022-01-01"))].copy()

            if area and regione != "Italia":

                eventi = eventi[
                    eventi[area]
                    .astype(str)
                    .str.contains(
                        regione,
                        regex=False,
                        na=False
                    )
                ]

            if sport:

                sport_disponibili = sorted(
                    eventi[sport]
                    .dropna()
                    .astype(str)
                    .unique()
                )

                scelti = st.multiselect(
                    "Sport",
                    sport_disponibili,
                    default=sport_disponibili
                )

                eventi = eventi[
                    eventi[sport].isin(scelti)
                ]

            if nome:
                ricerca = st.text_input("Cerca un evento", key="cerca_evento")
                if ricerca.strip():
                    eventi = eventi[eventi[nome].astype(str).str.contains(ricerca.strip(), case=False, regex=False, na=False)]
            if data and not eventi.empty:
                anni = sorted(eventi["_data"].dropna().dt.year.unique().tolist())
                selezione_anni = st.multiselect("Anno di inizio evento", anni, default=anni)
                eventi = eventi[eventi["_data"].isna() | eventi["_data"].dt.year.isin(selezione_anni)]
            if sport:
                # Esclusione già concordata nel progetto.
                eventi = eventi[~eventi[sport].astype(str).str.casefold().eq("sci di fondo")]

            if identificativo:

                unici = (
                    eventi[identificativo].nunique()
                )

            elif nome and data:

                unici = len(
                    eventi.drop_duplicates(
                        [nome, data]
                    )
                )

            else:

                unici = None

            a, b = st.columns(2)

            a.metric(
                "Eventi distinti",
                numero(unici)
                if unici is not None
                else "N/D"
            )

            b.metric(
                "Righe territoriali",
                numero(len(eventi))
            )

            if sport and not eventi.empty:

                counts = (
                    eventi.groupby(
                        sport,
                        as_index=False
                    ).size()
                )

                fig = px.bar(
                    counts,
                    x=sport,
                    y="size",
                    color_discrete_sequence=[ACCENTO]
                )

                grafico(fig, 330)

            st.caption("Il grafico conta righe territoriali: un evento multisede può comparire in più regioni.")
            if nome and not eventi.empty:
                st.subheader("Scheda evento")
                # Raggruppamento esplicito; niente fusione automatica tra edizioni.
                chiavi = eventi[nome].fillna("Evento senza nome").astype(str)
                if data:
                    chiavi = chiavi + " · " + eventi["_data"].dt.strftime("%d/%m/%Y").fillna("data non disponibile")
                if identificativo:
                    chiavi = chiavi + " · ID " + eventi[identificativo].astype("string").fillna("non disponibile")
                scelta_evento = st.selectbox("Evento e relativa edizione", sorted(chiavi.unique()))
                dettaglio = eventi.loc[chiavi.eq(scelta_evento)].drop(columns=["_data"], errors="ignore")
                st.dataframe(dettaglio, use_container_width=True, hide_index=True)
                with st.expander("Come leggere attendance e qualità del dato"):
                    st.write("**Persone / visitatori unici:** persone distinte, se documentate dalla fonte. **Ingressi cumulativi:** accessi che possono includere la stessa persona più volte. **Biglietti:** non equivalgono automaticamente a persone distinte.")
                    st.write("**EQUIVALENTE:** proxy eventualmente già calcolata come ingressi cumulativi divisi per giorni; non è una stima di visitatori unici. Questa app conserva i valori del master e non applica nuove conversioni.")
                    st.write("Per eventi multisede consulta le singole righe territoriali, la fonte e il metodo di ripartizione. Non sommare indicatori con unità diverse.")
            with st.expander("Catalogo completo filtrato", expanded=True):
                tabella(
                    eventi.drop(columns=["_data"], errors="ignore"),
                    "event_master_filtrato.csv"
                )

        else:

            st.warning(
                "File Excel Event Master non disponibile "
                "nel repository."
            )

    # --------------------------------------------------------
    # SCREENING DEL 2%
    # --------------------------------------------------------

    with tab2:

        st.write(
            "Il rapporto fra pernottamenti teorici "
            "di un evento e guest nights mensili NUTS2 "
            "è uno screening esplorativo."
        )

        st.latex(r"\text{Incidenza teorica (\%)} = 100 \times \frac{\text{pernottamenti teorici attribuiti al mese e alla regione}}{\text{guest nights osservate nello stesso mese e regione}}")
        st.info("Superare il 2% indica un evento da approfondire, non un effetto dimostrato né significatività statistica. I risultati sono letti dai file esistenti, senza ricalcolare le ipotesi.")
        st.caption(
            "Preset discussi: 5%, 15%, 25%, 40%, 60%. "
            "Non equivalgono a prenotazioni aggiuntive osservate."
        )

        if FILES["screening"].is_file():

            screen = leggi_csv(
                str(FILES["screening"])
            ).copy()

            area = trova_colonna(
                screen,
                "nuts2_code",
                "nuts2",
                "codici_nuts2"
            )

            if area and regione != "Italia":

                screen = screen[
                    screen[area]
                    .astype(str)
                    .str.contains(
                        regione,
                        regex=False,
                        na=False
                    )
                ]

            st.metric(
                "Righe dello screening",
                numero(len(screen))
            )

            possibili = [
                c for c in screen.columns
                if any(
                    t in c.lower()
                    for t in (
                        "incidenza",
                        "rapporto",
                        "pct",
                        "percent",
                        "peso"
                    )
                )
            ]

            scelta = st.selectbox(
                "Colonna del rapporto evento/domanda "
                "(facoltativa)",
                ["Nessuna"] + possibili
            )

            if scelta != "Nessuna":

                scala = st.radio(
                    "Unità",
                    [
                        "Percentuale (2,5)",
                        "Frazione (0,025)"
                    ],
                    horizontal=True
                )

                incidenza = pd.to_numeric(
                    screen[scelta],
                    errors="coerce"
                )

                if scala.startswith("Frazione"):
                    incidenza = 100 * incidenza

                screen["incidenza_pct"] = incidenza

                if incidenza.isna().any():
                    st.warning(f"{incidenza.isna().sum()} rapporti mancanti o non numerici esclusi dal conteggio sopra soglia.")
                st.metric(
                    "Righe ≥ 2%",
                    numero(
                        (incidenza >= 2).sum()
                    )
                )

                if st.checkbox(
                    "Mostra solo sopra soglia"
                ):

                    screen = screen[
                        screen.incidenza_pct >= 2
                    ]

            st.caption(
                "Selezionare la colonna che rappresenta "
                "il rapporto evento/domanda, non "
                "una percentuale diversa."
            )

            tabella(
                screen,
                "screening_filtrato.csv"
            )

        else:

            st.warning(
                "File di screening non disponibile."
            )

    # --------------------------------------------------------
    # EVENTI E DOMANDA OSSERVATA
    # --------------------------------------------------------

    with tab3:

        scelta_reg = (
            regione if regione != "Italia"
            else st.selectbox(
                "Regione NUTS2",
                sorted(nomi_reg),
                format_func=lambda x: nomi_reg[x]
            )
        )

        anno = st.selectbox(
            "Anno",
            [2022, 2023, 2024, 2025],
            index=3
        )

        mese = st.selectbox(
            "Mese",
            list(range(1, 13)),
            index=4
        )

        base = reg[
            reg.year.between(2022, 2025)
            & (reg.month == mese)
        ]

        nazionale = (
            base.groupby("year")
            .guest_nights_total.sum()
        )

        locale = (
            base[
                base.nuts2_code == scelta_reg
            ]
            .set_index("year")
            .guest_nights_total
        )

        confronto = pd.DataFrame({
            "Regione": locale,
            "Resto Italia": nazionale - locale
        }).sort_index()

        yoy = confronto.pct_change() * 100

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

        grafico(fig, 360)

        if anno in confronto.index:

            st.metric(
                f"{nomi_reg[scelta_reg]} · {mese:02}/{anno}",
                numero(
                    confronto.loc[
                        anno,
                        "Regione"
                    ]
                )
            )

        tabella(
            confronto.reset_index(
                names="year"
            ),
            "confronto_regionale.csv"
        )

        if FILES["eventi"].is_file():
            catalogo = leggi_excel(str(FILES["eventi"])).copy()
            ci = trova_colonna(catalogo, "Data inizio", "data_inizio", "start_date")
            cf = trova_colonna(catalogo, "Data fine", "data_fine", "end_date")
            cr = trova_colonna(catalogo, "Codici NUTS2", "nuts2_code", "nuts2")
            if ci and cr:
                inizio = date_eventi(catalogo[ci])
                fine = date_eventi(catalogo[cf]).fillna(inizio) if cf else inizio
                primo = pd.Timestamp(year=anno, month=mese, day=1)
                ultimo = primo + pd.offsets.MonthEnd(0)
                maschera = (catalogo[cr].astype(str).str.contains(scelta_reg, regex=False, na=False)
                            & inizio.le(ultimo) & fine.ge(primo) & fine.ge(inizio))
                selezionati = catalogo.loc[maschera].copy()
                if not selezionati.empty:
                    da = inizio.loc[maschera].clip(lower=primo)
                    a = fine.loc[maschera].clip(upper=ultimo)
                    selezionati["Giorni evento nel mese"] = (a - da).dt.days + 1
                st.subheader("Eventi del catalogo presenti nel mese")
                st.caption("Inclusi anche gli eventi iniziati nel mese precedente. I giorni misurano esposizione, non pernottamenti aggiuntivi.")
                tabella(selezionati)

        if FILES["mesi_eventi"].is_file():

            em = leggi_csv(
                str(FILES["mesi_eventi"])
            )

            rc = trova_colonna(
                em,
                "nuts2_code",
                "nuts2"
            )

            yc = trova_colonna(
                em,
                "year",
                "anno"
            )

            mc = trova_colonna(
                em,
                "month",
                "mese"
            )

            if rc and yc and mc:

                mask = (
                    em[rc]
                    .astype(str)
                    .str.contains(
                        scelta_reg,
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
                    "Eventi del territorio e mese selezionati"
                )

                tabella(
                    em.loc[mask]
                )

        st.caption(
            "Confronto descrittivo: differenze fra regioni "
            "non identificano automaticamente l'effetto sportivo."
        )


# ============================================================
# 13. FORECAST 2026-2030
# ============================================================

elif pagina == "Forecast 2026–2030":

    st.header(
        "Forecast 2026–2030 · storico dal 2022"
    )

    st.info("I tre scenari rappresentano ipotesi alternative, non intervalli di confidenza. Non includono incrementi sportivi aggiunti automaticamente.")
    with st.expander("Ipotesi e lettura delle previsioni"):
        st.write("Stagionalità 2025; tasso iniziale composto al 50% dalla componente nazionale e al 50% da quella regionale; tre ipotesi di rallentamento. I valori sono caricati dai CSV già prodotti.")
        st.write("Le previsioni provinciali ripartiscono quelle regionali con quote 2025 costanti: non sono modelli indipendenti per provincia. Lo storico provinciale resta annuale.")
    attivi = st.multiselect(
        "Scenari nel grafico principale",
        SCENARI,
        default=SCENARI
    )

    fig = go.Figure()

    # --------------------------------------------------------
    # PREVISIONI PROVINCIALI
    # --------------------------------------------------------

    if provincia != "Tutte":

        storico = prov[
            (prov.nuts3_code == provincia)
            & (prov.year >= 2022)
        ].sort_values("year")

        futuro = fp[
            fp.nuts3_code == provincia
        ].sort_values("year")

        x = storico.year.tolist()
        y = storico.guest_nights_total.tolist()

        st.subheader(
            nomi_prov[provincia] + " · NUTS3 annuale"
        )

        if not x:
            st.warning("Nessun dato storico per il territorio selezionato.")
            st.stop()
        for s in attivi:

            fig.add_trace(
                go.Scatter(
                    x=[x[-1]] + futuro.year.tolist(),

                    y=[y[-1]] + futuro[
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

    # --------------------------------------------------------
    # PREVISIONI REGIONALI E NAZIONALI
    # --------------------------------------------------------

    else:

        base = (
            reg if regione == "Italia"
            else reg[
                reg.nuts2_code == regione
            ]
        )

        storico = (
            base[
                base.year >= 2022
            ]
            .groupby(
                "data",
                as_index=False
            )
            .guest_nights_total.sum()
            .sort_values("data")
        )

        future_base = (
            fr if regione == "Italia"
            else fr[
                fr.nuts2_code == regione
            ]
        )

        future_base = future_base[
            future_base.data > storico.data.max()
        ]

        campi = [
            f"gn_{s.lower()}"
            for s in SCENARI
        ]

        futuro = (
            future_base.groupby(
                "data",
                as_index=False
            )[campi]
            .sum()
            .sort_values("data")
        )

        x = storico.data.tolist()
        y = storico.guest_nights_total.tolist()

        st.subheader(
            (
                "Italia" if regione == "Italia"
                else nomi_reg[regione]
            ) + " · NUTS2 mensile"
        )

        if not x:
            st.warning("Nessun dato storico per il territorio selezionato.")
            st.stop()
        for s in attivi:

            fig.add_trace(
                go.Scatter(
                    x=[x[-1]] + futuro.data.tolist(),

                    y=[y[-1]] + futuro[
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

    # --------------------------------------------------------
    # SERIE OSSERVATA
    # --------------------------------------------------------

    if not x:
        st.warning("Storico non disponibile per la selezione corrente.")
        st.stop()
    fig.add_shape(type="line", x0=x[-1], x1=x[-1], y0=0, y1=1, yref="paper",
                  line=dict(color=ACCENTO, width=1, dash="dot"))
    st.caption("La linea verticale segna l’ultima osservazione. Le linee tratteggiate rappresentano le previsioni.")
    fig.add_trace(
        go.Scatter(
            x=x,
            y=y,
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
        "NUTS2 2026: i mesi già pubblicati sono osservati; "
        "NUTS3 2026 è una previsione annuale."
    )

    # --------------------------------------------------------
    # KPI E CONFRONTO ANNUALE
    # --------------------------------------------------------

    kpi = st.columns(3)

    confronto = pd.DataFrame({
        "Anno": list(range(2026, 2031))
    })

    for box, s in zip(kpi, SCENARI):

        if provincia != "Tutte":

            dati = fp[
                fp.nuts3_code == provincia
            ]

            campo = f"forecast_{s.lower()}"

        else:

            dati = (
                fr if regione == "Italia"
                else fr[
                    fr.nuts2_code == regione
                ]
            )

            campo = f"gn_{s.lower()}"

        annuo = (
            dati.groupby("year")[campo].sum()
        )

        confronto[s] = (
            confronto.Anno.map(annuo)
        )

        box.metric(
            "2030 · " + s,
            numero(annuo.loc[2030])
        )

    st.subheader(
        "Confronto annuale dei tre scenari"
    )

    fig = go.Figure()

    for s in SCENARI:

        fig.add_trace(
            go.Scatter(
                x=confronto.Anno,
                y=confronto[s],
                name=s,
                mode="lines+markers",

                line=dict(
                    color=COLORI[s],
                    width=3
                )
            )
        )

    grafico(fig, 340)

    tabella(
        confronto,
        "forecast_scenari.csv"
    )

    # --------------------------------------------------------
    # MAPPA DEL FORECAST
    # --------------------------------------------------------

    st.divider()

    anno_mappa = st.select_slider(
        "Anno mappa",
        options=[
            2026, 2027, 2028, 2029, 2030
        ],
        value=2030
    )

    geo_mappa = st.radio(
        "Geografia",
        ["NUTS2", "NUTS3"],
        horizontal=True
    )

    scen_mappa = st.selectbox(
        "Scenario mappa",
        SCENARI,
        index=SCENARI.index(scenario)
    )

    if geo_mappa == "NUTS2":

        campo = f"gn_{scen_mappa.lower()}"

        dati = (
            fr[
                fr.year == anno_mappa
            ]
            .groupby(
                "nuts2_code",
                as_index=False
            )[campo]
            .sum()
        )

        dati["territorio"] = (
            dati.nuts2_code.map(nomi_reg)
        )

        mappa_kepler(
            dati,
            2,
            "nuts2_code",
            campo,
            f"Previsione NUTS2 {anno_mappa} · {scen_mappa}"
        )

    else:

        campo = (
            f"forecast_{scen_mappa.lower()}"
        )

        dati = fp.loc[
            fp.year == anno_mappa,
            ["nuts3_code", campo]
        ].copy()

        dati["territorio"] = (
            dati.nuts3_code.map(nomi_prov)
        )

        mappa_kepler(
            dati,
            3,
            "nuts3_code",
            campo,
            f"Previsione NUTS3 {anno_mappa} · {scen_mappa}"
        )


# ============================================================
# 14. METODOLOGIA E CONTROLLI
# ============================================================

else:

    st.header(
        "Metodologia e controlli"
    )

    a, b, c, d = st.columns(4)

    a.metric(
        "NUTS2",
        reg.nuts2_code.nunique()
    )

    b.metric(
        "NUTS3",
        prov.nuts3_code.nunique()
    )

    c.metric(
        "Mesi osservati NUTS2",
        reg[
            ["year", "month"]
        ].drop_duplicates().shape[0]
    )

    d.metric(
        "Anni osservati NUTS3",
        prov.year.nunique()
    )

    risultati = [
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

    for nome_df, dataset, chiave in [
        ("Master NUTS2", reg, ["nuts2_code", "year", "month"]),
        ("Master NUTS3", prov, ["nuts3_code", "year"]),
        ("Forecast NUTS2", fr, ["nuts2_code", "year", "month"]),
        ("Forecast NUTS3", fp, ["nuts3_code", "year"]),
    ]:
        risultati.append((f"{nome_df}: chiavi univoche", not dataset.duplicated(chiave).any()))
    st.caption(f"Ultimo mese osservato regionale: {reg['data'].max():%m/%Y}. Ultimo anno provinciale: {prov['year'].max()}.")
    for s in SCENARI:

        r = (
            fr.groupby(
                ["nuts2_code", "year"]
            )[f"gn_{s.lower()}"]
            .sum()
            .round()
        )

        p = (
            fp.groupby(
                ["nuts2_code", "year"]
            )[f"forecast_{s.lower()}"]
            .sum()
        )

        risultati.append((
            f"Riconciliazione NUTS3/NUTS2 · {s}",

            r.subtract(
                p,
                fill_value=0
            ).abs().max() <= 1
        ))

    tabella(
        pd.DataFrame([
            {
                "Controllo": nome,
                "Esito": (
                    "OK" if esito
                    else "DA VERIFICARE"
                )
            }
            for nome, esito in risultati
        ])
    )

    st.subheader(
        "Impostazione metodologica"
    )

    st.write(
        "**Periodo principale:** 2022–2025, "
        "più i mesi osservati del 2026."
    )

    st.write(
        "**Fonte:** Eurostat `tour_ce_omn12` "
        "(NUTS2 mensile) e `tour_ce_oan3` "
        "(NUTS3 annuale)."
    )

    st.write(
        "**Forecast:** tassi iniziali 50% nazionali "
        "e 50% regionali, tre scenari, quote "
        "provinciali 2025 mantenute costanti."
    )

    st.write(
        "**Eventi sportivi:** screening del 2% "
        "descrittivo, nessun incremento causale "
        "sommato automaticamente alla previsione."
    )

    st.caption(
        "Backtest storici documentati, "
        "non ricalcolati nell'app. I campioni indicati non sono tutti identici: "
        "non costituiscono una classifica omogenea dei modelli. "
        "L’errore storico non garantisce la precisione delle proiezioni al 2030."
    )

    tabella(
        pd.DataFrame(
            [
                (
                    "Baseline stagionale 2024",
                    "11,11%",
                    "test storico"
                ),
                (
                    "Media 2023–2024",
                    "17,13%",
                    "test storico"
                ),
                (
                    "Prudente",
                    "7,48%",
                    "test 2025"
                ),
                (
                    "Intermedio",
                    "8,39%",
                    "test 2025"
                ),
                (
                    "Espansivo",
                    "9,89%",
                    "test 2025"
                )
            ],
            columns=[
                "Metodo",
                "WAPE",
                "Campione"
            ]
        )
    )


# ============================================================
# 15. FOOTER
# ============================================================

st.divider()

st.caption(
    "Airbnb&Sports · "
    + VERSIONE
    + " · Eurostat, eventi e forecast separati"
)
