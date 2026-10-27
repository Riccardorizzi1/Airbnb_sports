# ============================================================
# AIRBNB&SPORTS — SBL CONSULTANCY
# File principale: app.py
# ============================================================

from pathlib import Path
from urllib.request import Request, urlopen
import base64
import html
import json
import re
import unicodedata

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components


# ============================================================
# 1. CONFIGURAZIONE
# ============================================================

st.set_page_config(
    page_title="Airbnb&Sports | SBL Consultancy",
    page_icon="🏟️",
    layout="wide",
)

ROOT = Path(__file__).resolve().parent
VERSIONE = "v4.0 · Simulatore eventi"
SITO_AZIENDA = "https://www.sblconsultancy.it/"

SCENARI = ["Prudente", "Intermedio", "Espansivo"]

COLORI = {
    "Prudente": "#249784",
    "Intermedio": "#DC813F",
    "Espansivo": "#C94F7C",
}

PRESET = {
    "Molto basso · 5%": 5,
    "Basso · 15%": 15,
    "Medio · 25%": 25,
    "Alto · 40%": 40,
    "Molto alto · 60%": 60,
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
        if percorso.is_file():
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
    "screening": file_dati("screening_eventi_completo.csv"),
    "mesi_eventi": file_dati("screening_regioni_mesi_completo.csv"),
}


# ============================================================
# 2. FUNZIONI DI SUPPORTO
# ============================================================

def normalizza(testo):
    testo = unicodedata.normalize("NFKD", str(testo))
    testo = "".join(
        carattere
        for carattere in testo
        if not unicodedata.combining(carattere)
    )
    return re.sub(r"[^a-z0-9]", "", testo.lower())


def trova_colonna(df, *nomi):
    disponibili = {
        normalizza(colonna): colonna
        for colonna in df.columns
    }

    for nome in nomi:
        risultato = disponibili.get(normalizza(nome))
        if risultato is not None:
            return risultato

    return None


def numero(valore, decimali=0):
    if pd.isna(valore):
        return "—"

    return (
        f"{float(valore):,.{decimali}f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


def percentuale(valore, segno=True):
    if pd.isna(valore):
        return "—"

    formato = (
        f"{float(valore):+.2f}%"
        if segno
        else f"{float(valore):.2f}%"
    )

    return formato.replace(".", ",")


def numero_dato(valore):
    """
    Legge numeri già numerici o stringhe semplici.
    Intervalli, 'oltre', '~' e testo non vengono convertiti
    automaticamente in valori puntuali.
    """
    if pd.isna(valore):
        return float("nan")

    if isinstance(valore, (int, float)):
        return float(valore)

    testo = str(valore).strip()
    testo = testo.replace("\u00a0", "").replace(" ", "")

    if not testo:
        return float("nan")

    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+(,\d+)?", testo):
        testo = testo.replace(".", "").replace(",", ".")

    elif re.fullmatch(r"-?\d{1,3}(,\d{3})+(\.\d+)?", testo):
        testo = testo.replace(",", "")

    elif "," in testo and "." not in testo:
        testo = testo.replace(",", ".")

    return pd.to_numeric(testo, errors="coerce")


def data_singola(valore):
    if pd.isna(valore):
        return pd.NaT

    if isinstance(valore, (int, float)):
        if 20000 < valore < 80000:
            return pd.to_datetime(
                valore,
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            )
        return pd.NaT

    testo = str(valore).strip()
    iso = bool(re.match(r"^\d{4}-\d{2}-\d{2}", testo))

    risultato = pd.to_datetime(
        testo,
        dayfirst=not iso,
        errors="coerce",
    )

    if pd.isna(risultato):
        return pd.NaT

    return pd.Timestamp(risultato).tz_localize(None).normalize()


def date_eventi(serie):
    return pd.to_datetime(
        serie.map(data_singola),
        errors="coerce",
    )


@st.cache_data(show_spinner=False)
def leggi_csv(percorso):
    df = pd.read_csv(
        percorso,
        encoding="utf-8-sig",
        low_memory=False,
    )

    if len(df.columns) == 1 and ";" in str(df.columns[0]):
        df = pd.read_csv(
            percorso,
            sep=";",
            encoding="utf-8-sig",
            low_memory=False,
        )

    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def leggi_excel(percorso):
    df = pd.read_excel(percorso, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def prepara_tabella(df):
    """
    Modifica soltanto la copia visualizzata/esportata.
    Non modifica i dati usati nei calcoli.
    """
    risultato = df.copy()

    for colonna in risultato.columns:
        serie = risultato[colonna]

        if pd.api.types.is_object_dtype(serie):
            risultato[colonna] = serie.map(
                lambda x: (
                    pd.NA
                    if isinstance(x, str)
                    and x.strip().lower() in {
                        "", "nan", "none", "nat", "null", "<na>"
                    }
                    else x
                )
            )

    risultato = risultato.dropna(axis=1, how="all")

    for colonna in risultato.columns:
        serie = risultato[colonna]
        nome = normalizza(colonna)

        nome_data = nome in {
            "data",
            "datainizio",
            "datafine",
            "startdate",
            "enddate",
            "date",
            "inizio",
            "fine",
            "meseosservazione",
        }

        if pd.api.types.is_datetime64_any_dtype(serie):
            risultato[colonna] = (
                serie.dt.strftime("%d/%m/%Y")
                .fillna("")
            )

        elif nome_data and not pd.api.types.is_numeric_dtype(serie):
            convertita = date_eventi(serie)
            maschera = convertita.notna()

            if maschera.any():
                visualizzata = serie.astype("string").fillna("")
                visualizzata.loc[maschera] = (
                    convertita.loc[maschera]
                    .dt.strftime("%d/%m/%Y")
                )
                risultato[colonna] = visualizzata

    return risultato


def tabella(df, nome=None):
    visualizzata = prepara_tabella(df)

    if visualizzata.empty:
        st.info("Nessuna riga disponibile per questa selezione.")
        return

    st.dataframe(
        visualizzata,
        hide_index=True,
        use_container_width=True,
    )

    if nome:
        st.download_button(
            "Scarica CSV",
            visualizzata.to_csv(
                index=False,
                sep=";",
            ).encode("utf-8-sig"),
            file_name=nome,
            mime="text/csv",
            key=f"download_{nome}",
        )


def grafico(fig, altezza=420):
    fig.update_layout(
        height=altezza,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=TESTO,
        hovermode="x unified",
        margin=dict(l=15, r=15, t=45, b=20),
        legend=dict(
            orientation="h",
            y=1.14,
        ),
    )

    fig.update_xaxes(
        gridcolor="rgba(120,150,170,0.12)"
    )

    fig.update_yaxes(
        gridcolor="rgba(120,150,170,0.20)",
        tickformat=",~s",
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
    )


# ============================================================
# 3. LOGO AZIENDALE CLICCABILE
# ============================================================

def trova_logo():
    estensioni = {
        ".png",
        ".svg",
        ".jpg",
        ".jpeg",
        ".webp",
    }

    cartelle = [
        ROOT,
        ROOT / "assets",
        ROOT / "images",
        ROOT / "img",
        ROOT / "output_eurostat",
    ]

    for cartella in cartelle:
        if not cartella.is_dir():
            continue

        for percorso in cartella.iterdir():
            if (
                percorso.is_file()
                and percorso.stem.lower() == "logo_sbl"
                and percorso.suffix.lower() in estensioni
            ):
                return percorso

    return None


def mostra_logo():
    percorso = trova_logo()

    if percorso is None:
        st.markdown(
            f"### [SBL Consultancy]({SITO_AZIENDA})"
        )
        st.caption(
            "Logo non trovato: il file deve chiamarsi "
            "logo_sbl con estensione PNG, SVG, JPG, JPEG o WEBP."
        )
        return

    mime = {
        ".png": "image/png",
        ".svg": "image/svg+xml",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
    }[percorso.suffix.lower()]

    contenuto = base64.b64encode(
        percorso.read_bytes()
    ).decode("ascii")

    st.markdown(
        f"""
        <a href="{SITO_AZIENDA}"
           target="_blank"
           rel="noopener noreferrer"
           aria-label="Visita il sito di SBL Consultancy"
           style="display:block; margin-bottom:18px;">
            <img
                src="data:{mime};base64,{contenuto}"
                alt="SBL Consultancy"
                style="
                    display:block;
                    max-width:260px;
                    width:100%;
                    height:auto;
                    object-fit:contain;
                "
            />
        </a>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 4. MAPPA CON SFONDO GEOGRAFICO
# ============================================================

@st.cache_data(ttl=86400, show_spinner=False)
def confini_nuts(livello):
    url = (
        "https://gisco-services.ec.europa.eu/"
        "distribution/v2/nuts/geojson/"
        f"NUTS_RG_20M_2024_4326_LEVL_{livello}.geojson"
    )

    richiesta = Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
    )

    with urlopen(richiesta, timeout=35) as risposta:
        geo = json.load(risposta)

    features = [
        f
        for f in geo.get("features", [])
        if str(
            f.get("properties", {}).get("NUTS_ID", "")
        ).startswith("IT")
        and len(
            str(f.get("properties", {}).get("NUTS_ID", ""))
        ) == livello + 2
    ]

    if not features:
        raise ValueError(
            "Confini territoriali italiani non disponibili."
        )

    return {
        "type": "FeatureCollection",
        "features": features,
    }


def mappa(df, livello, codice, campo, titolo):
    st.subheader(titolo)

    dati = df[
        [codice, "territorio", campo]
    ].copy()

    dati[codice] = dati[codice].astype(str)
    dati[campo] = pd.to_numeric(
        dati[campo],
        errors="coerce",
    )

    dati = dati.dropna(subset=[campo])

    if dati.empty:
        st.info("Nessun dato disponibile per la mappa.")
        return

    if dati[codice].duplicated().any():
        st.warning(
            "Codici territoriali duplicati: "
            "la mappa non viene costruita."
        )
        tabella(dati)
        return

    mostra = st.checkbox(
        "Mostra mappa",
        value=True,
        key=f"map_{titolo}",
    )

    if mostra:
        try:
            geo = confini_nuts(livello)
            valori = dati.set_index(codice).to_dict("index")

            features = []

            for feature in geo["features"]:
                nuts = feature["properties"]["NUTS_ID"]

                if nuts not in valori:
                    continue

                record = valori[nuts]

                features.append({
                    "type": "Feature",
                    "geometry": feature["geometry"],
                    "properties": {
                        "codice": nuts,
                        "territorio": str(record["territorio"]),
                        "valore": float(record[campo]),
                        "valore_testo": numero(record[campo]),
                    },
                })

            if not features:
                raise ValueError(
                    "Nessuna corrispondenza fra codici "
                    "dei dati e geometrie GISCO."
                )

            dati_geo = {
                "type": "FeatureCollection",
                "features": features,
            }

            contenuto_geo = json.dumps(
                dati_geo,
                ensure_ascii=False,
            ).replace("</", "<\\/")

            sfondo = "#071822" if dark else "#C7F1F6"

            documento = """
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<link
  rel="stylesheet"
  href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"
/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<style>
html, body {
    margin: 0;
    width: 100%;
    height: 100%;
    font-family: Arial, sans-serif;
}
#map {
    width: 100%;
    height: 580px;
    background: __SFONDO__;
    border-radius: 14px;
}
.legend {
    background: white;
    color: #082B46;
    padding: 10px 12px;
    border-radius: 8px;
    box-shadow: 0 2px 8px rgba(0,0,0,.15);
    font-size: 12px;
    line-height: 1.5;
}
.legend-bar {
    width: 150px;
    height: 12px;
    background: linear-gradient(
        90deg, #C2F0F6, #88D6E7, #4BAFC9, #126982, #084A68
    );
    margin: 6px 0;
}
</style>
</head>
<body>
<div id="map"></div>
<script>
const geo = __GEO__;
const map = L.map("map", {
    scrollWheelZoom: false
});

L.tileLayer(
    "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
    {
        maxZoom: 18,
        attribution: '&copy; OpenStreetMap contributors'
    }
).addTo(map);

const values = geo.features.map(f => f.properties.valore);
const minValue = Math.min(...values);
const maxValue = Math.max(...values);
const colors = ["#C2F0F6", "#88D6E7", "#4BAFC9", "#126982", "#084A68"];

function colorFor(value) {
    if (maxValue === minValue) return colors[2];
    const scaled = (value - minValue) / (maxValue - minValue);
    return colors[Math.min(4, Math.floor(scaled * 5))];
}

function safeText(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;");
}

const polygons = L.geoJSON(geo, {
    style: feature => ({
        color: "#084A68",
        weight: 1,
        fillColor: colorFor(feature.properties.valore),
        fillOpacity: 0.72
    }),
    onEachFeature: (feature, layer) => {
        const p = feature.properties;
        const text =
            "<b>" + safeText(p.territorio) + "</b><br>" +
            safeText(p.codice) + "<br>" +
            "Pernottamenti: " + safeText(p.valore_testo);

        layer.bindTooltip(text);
        layer.bindPopup(text);

        layer.on("mouseover", () => {
            layer.setStyle({weight: 3, fillOpacity: 0.88});
        });

        layer.on("mouseout", () => {
            polygons.resetStyle(layer);
        });
    }
}).addTo(map);

map.fitBounds(polygons.getBounds(), {padding: [15, 15]});

const legend = L.control({position: "bottomright"});
legend.onAdd = function() {
    const div = L.DomUtil.create("div", "legend");
    div.innerHTML =
        "<b>Pernottamenti degli ospiti</b>" +
        '<div class="legend-bar"></div>' +
        minValue.toLocaleString("it-IT", {maximumFractionDigits: 0}) +
        " — " +
        maxValue.toLocaleString("it-IT", {maximumFractionDigits: 0});
    return div;
};
legend.addTo(map);
</script>
</body>
</html>
"""
            documento = documento.replace(
                "__SFONDO__",
                sfondo,
            ).replace(
                "__GEO__",
                contenuto_geo,
            )

            components.html(
                documento,
                height=600,
                scrolling=False,
            )

            st.caption(
                "Confini ufficiali GISCO NUTS 2024. "
                "Sfondo OpenStreetMap. "
                "La mappa non richiede WebGL."
            )

        except Exception as exc:
            st.warning(
                f"Mappa non disponibile: {exc}. "
                "I dati rimangono consultabili nella tabella."
            )

    tabella(
        dati.sort_values(campo, ascending=False)
    )


# ============================================================
# 5. SIDEBAR E COLORI AZIENDALI
# ============================================================

with st.sidebar:
    mostra_logo()

    st.markdown("## Airbnb&Sports")
    st.caption("Tourism Demand · Italia")
    st.caption(VERSIONE)

    st.divider()

    pagina = st.radio(
        "Sezione",
        PAGINE,
    )

    dark = st.toggle(
        "Modalità scura",
        value=False,
    )

    st.divider()


if dark:
    SIDE = ("#071F31", "#0E4058", "#126982")
    BG = ("#071822", "#0C2B39", "#10475B")
    CARD = "#102D3B"
    TESTO = "#F3FAFC"
    BORDO = "#315462"
    OMBRA = "rgba(0,0,0,.28)"
    ACCENTO = "#6DDDEB"

else:
    SIDE = ("#153F57", "#17627E", "#259BBA")
    BG = ("#C7F1F6", "#9FE5F0", "#5BC6E2")
    CARD = "#FDFEFF"
    TESTO = "#082B46"
    BORDO = "#D4E7EC"
    OMBRA = "rgba(8,43,70,.10)"
    ACCENTO = "#0E718F"


st.markdown(
    f"""
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

section[data-testid="stSidebar"] [data-baseweb="select"] * {{
    color: #082B46 !important;
}}

section[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background: #FDFEFF !important;
}}
</style>
""",
    unsafe_allow_html=True,
)


# ============================================================
# 6. CARICAMENTO DATI PRINCIPALI
# ============================================================

obbligatori = ["reg", "prov", "fr", "fp"]

mancanti = [
    FILES[k].name
    for k in obbligatori
    if not FILES[k].is_file()
]

if mancanti:
    st.error("Mancano questi file del progetto:")
    st.code("\n".join(mancanti))
    st.stop()


try:
    reg, prov, fr, fp = [
        leggi_csv(str(FILES[k])).copy()
        for k in obbligatori
    ]

    specifiche = [
        (
            reg,
            ["year", "month", "nuts2_code", "nuts2_name"],
            ["guest_nights_total"],
        ),
        (
            prov,
            ["year", "nuts3_code", "nuts3_name", "nuts2_code"],
            ["guest_nights_total"],
        ),
        (
            fr,
            ["year", "month", "nuts2_code"],
            [f"gn_{s.lower()}" for s in SCENARI],
        ),
        (
            fp,
            ["year", "nuts3_code", "nuts2_code"],
            [f"forecast_{s.lower()}" for s in SCENARI],
        ),
    ]

    for df, chiavi, misure in specifiche:
        assenti = sorted(
            set(chiavi + misure) - set(df.columns)
        )

        if assenti:
            raise ValueError(
                "Colonne mancanti: " + ", ".join(assenti)
            )

        df["year"] = pd.to_numeric(
            df["year"],
            errors="raise",
        ).astype(int)

        for c in df.columns:
            if (
                c.startswith("guest_nights_")
                or c.startswith("occupied_nights_")
                or c.startswith("stays_")
                or c in misure
            ):
                df[c] = pd.to_numeric(
                    df[c],
                    errors="coerce",
                )

        for c in ["nuts2_code", "nuts3_code"]:
            if c in df:
                df[c] = (
                    df[c].astype("string")
                    .str.strip()
                    .str.upper()
                )

    for df in [reg, fr]:
        df["month"] = pd.to_numeric(
            df["month"],
            errors="raise",
        ).astype(int)

        df["data"] = pd.to_datetime({
            "year": df["year"],
            "month": df["month"],
            "day": 1,
        })

    for df, chiavi in [
        (reg, ["nuts2_code", "year", "month"]),
        (prov, ["nuts3_code", "year"]),
        (fr, ["nuts2_code", "year", "month"]),
        (fp, ["nuts3_code", "year"]),
    ]:
        if df.duplicated(chiavi).any():
            raise ValueError(
                "Chiavi duplicate nel dataset: "
                + ", ".join(chiavi)
            )

except Exception as exc:
    st.error(f"Errore nella lettura dei dati: {exc}")
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


with st.sidebar:
    st.subheader("Territorio")

    regione = st.selectbox(
        "Regione",
        ["Italia"] + sorted(nomi_reg),
        format_func=lambda x: (
            "Italia — tutte le regioni"
            if x == "Italia"
            else f"{nomi_reg[x]} · {x}"
        ),
    )

    province_ammissibili = (
        prov
        if regione == "Italia"
        else prov[prov["nuts2_code"].eq(regione)]
    )

    provincia = st.selectbox(
        "Provincia",
        ["Tutte"] + sorted(
            province_ammissibili["nuts3_code"]
            .dropna()
            .unique()
        ),
        format_func=lambda x: (
            "Tutte le province"
            if x == "Tutte"
            else f"{nomi_prov[x]} · {x}"
        ),
    )

    scenario = "Intermedio"

    if pagina in ["Panorama nazionale", "Forecast 2026–2030"]:
        scenario = st.selectbox(
            "Scenario predefinito",
            SCENARI,
            index=1,
        )

    if pagina in ["Panorama nazionale", "Dati e metodo"]:
        st.caption(
            "Questa sezione presenta il quadro nazionale."
        )

    if pagina == "Eventi e simulatore":
        st.caption(
            "La simulazione usa la regione e il mese. "
            "Il filtro provinciale non disaggrega la domanda mensile."
        )


st.title("Airbnb&Sports")
st.caption(
    "SBL Consultancy · Domanda di alloggi di breve durata "
    "prenotati tramite piattaforme online"
)


# ============================================================
# 7. FUNZIONI EVENT MASTER E SIMULAZIONE
# ============================================================

def selettore_colonna(
    df,
    etichetta,
    candidati,
    chiave,
):
    trovata = trova_colonna(df, *candidati)
    opzioni = ["— Non disponibile —"] + list(df.columns)

    indice = (
        opzioni.index(trovata)
        if trovata is not None
        else 0
    )

    scelta = st.selectbox(
        etichetta,
        opzioni,
        index=indice,
        key=chiave,
    )

    return None if scelta == opzioni[0] else scelta


def mappa_colonne_eventi(df):
    """
    Riconoscimento dei nomi più comuni.
    Le associazioni restano visibili e correggibili.
    """
    definizioni = {
        "nome": (
            "Nome evento",
            ["Evento originale", "Evento", "event_name", "nome_evento"],
        ),
        "id": (
            "Identificativo evento",
            ["event_id", "id_evento", "ID evento"],
        ),
        "sport": (
            "Sport",
            ["Sport", "sport"],
        ),
        "inizio": (
            "Data di inizio",
            ["Data inizio", "data_inizio", "start_date", "inizio"],
        ),
        "fine": (
            "Data di fine",
            ["Data fine", "data_fine", "end_date", "fine"],
        ),
        "nuts": (
            "Codice NUTS2",
            ["Codici NUTS2", "nuts2_code", "NUTS2", "codice_nuts2"],
        ),
        "persone": (
            "Persone / visitatori unici documentati o stimati",
            [
                "Visitatori unici",
                "visitatori_unici",
                "Visitatori unici stimati",
                "Visitatori unici/stimati",
                "Visitatori unici / stimati",
                "visitatori_unici_stimati",
                "unique_visitors",
            ],
        ),
        "valore": (
            "Attendance / presenze originali",
            [
                "Presenze",
                "Presenze totali",
                "presenze_totali",
                "attendance",
                "Attendance",
                "Presenze stimate",
                "presenze_stimate",
                "Valore presenze",
                "valore_presenze",
            ],
        ),
        "tipo": (
            "Tipo dell’indicatore attendance",
            [
                "Tipo indicatore",
                "tipo_indicatore",
                "tipo_indicatore_presenze",
                "tipo_base",
                "Tipo di indicatore",
                "attendance_type",
            ],
        ),
        "equivalente": (
            "Eventuale base equivalente già calcolata",
            [
                "base_equivalente",
                "presenze_equivalenti",
                "visitatori_equivalenti",
                "equivalenti",
                "attendance_equivalente",
            ],
        ),
    }

    mapping = {}

    with st.expander(
        "Colonne utilizzate: riconoscimento automatico e verifica",
        expanded=False,
    ):
        st.caption(
            "Questa associazione non cambia i file. "
            "Se un nome non viene riconosciuto, scegli la colonna "
            "corretta. Non associare ingressi cumulativi alla "
            "colonna dei visitatori unici."
        )

        for chiave, (etichetta, candidati) in definizioni.items():
            mapping[chiave] = selettore_colonna(
                df,
                etichetta,
                candidati,
                f"mapping_{chiave}",
            )

    return mapping


def costruisci_eventi(df, mapping):
    """
    Una riga = evento × territorio già presente nel master.
    Non ripartisce arbitrariamente righe con più NUTS2.
    """
    richieste = ["nome", "sport", "inizio", "nuts"]

    assenti = [
        x
        for x in richieste
        if mapping.get(x) is None
    ]

    if assenti:
        return pd.DataFrame(), (
            "Associa le colonne obbligatorie: "
            + ", ".join(assenti)
        )

    righe = []

    for posizione, (_, record) in enumerate(df.iterrows()):
        nome = str(record[mapping["nome"]])
        sport = str(record[mapping["sport"]]).strip()

        if normalizza(sport) == "scidifondo":
            continue

        inizio = data_singola(record[mapping["inizio"]])

        fine = (
            data_singola(record[mapping["fine"]])
            if mapping.get("fine")
            else inizio
        )

        if pd.isna(fine):
            fine = inizio

        codici = re.findall(
            r"\bIT[A-Z0-9]{2}\b",
            str(record[mapping["nuts"]]).upper(),
        )
        codici = sorted(set(codici))

        nuts = codici[0] if len(codici) == 1 else None

        durata = (
            (fine - inizio).days + 1
            if pd.notna(inizio)
            and pd.notna(fine)
            and fine >= inizio
            else None
        )

        persone = (
            numero_dato(record[mapping["persone"]])
            if mapping.get("persone")
            else float("nan")
        )

        valore = (
            numero_dato(record[mapping["valore"]])
            if mapping.get("valore")
            else float("nan")
        )

        equivalente = (
            numero_dato(record[mapping["equivalente"]])
            if mapping.get("equivalente")
            else float("nan")
        )

        tipo_originale = (
            str(record[mapping["tipo"]]).strip()
            if mapping.get("tipo")
            else ""
        )
        tipo = normalizza(tipo_originale)

        base = float("nan")
        categoria = "NON UTILIZZABILE"
        metodo = "Base non disponibile o unità non identificata"

        if pd.notna(persone) and persone >= 0:
            base = persone
            categoria = "PERSONE"
            metodo = (
                "Colonna visitatori unici/documentati/stimati"
            )

        elif pd.notna(equivalente) and equivalente >= 0:
            base = equivalente
            categoria = "EQUIVALENTE"
            metodo = "Base equivalente già presente nel file"

        elif pd.notna(valore) and valore >= 0:
            if (
                "bigliett" in tipo
                or "ticket" in tipo
            ):
                metodo = (
                    "Biglietti: nessuna conversione automatica "
                    "in persone"
                )

            elif (
                "equivalent" in tipo
            ):
                base = valore
                categoria = "EQUIVALENTE"
                metodo = "Indicatore già classificato equivalente"

            elif (
                "cumulativ" in tipo
                or "ingress" in tipo
            ):
                if durata and durata > 0:
                    base = valore / durata
                    categoria = "EQUIVALENTE"
                    metodo = (
                        "Proxy: ingressi / giorni dell’evento; "
                        "non visitatori unici"
                    )

            elif (
                "person" in tipo
                or "partecipant" in tipo
                or "visitator" in tipo
                or "unique" in tipo
            ):
                base = valore
                categoria = "PERSONE"
                metodo = (
                    "Attendance classificata come "
                    "persone/partecipanti"
                )

        problemi = []

        if durata is None:
            problemi.append("Date mancanti o incoerenti")

        if nuts is None:
            problemi.append(
                "NUTS2 mancante o più NUTS2 nella stessa riga"
            )
        elif nuts not in nomi_reg:
            problemi.append("NUTS2 non presente nel master Eurostat")

        if pd.isna(base):
            problemi.append(metodo)

        if mapping.get("id"):
            id_originale = record[mapping["id"]]
        else:
            id_originale = None

        if pd.isna(id_originale) or id_originale is None:
            etichetta_data = (
                inizio.strftime("%Y-%m-%d")
                if pd.notna(inizio)
                else "data_non_disponibile"
            )
            event_id = f"{nome}|{etichetta_data}"
        else:
            event_id = str(id_originale)

        righe.append({
            "riga_master": posizione,
            "event_id": event_id,
            "Evento": nome,
            "Sport": sport,
            "Data inizio": inizio,
            "Data fine": fine,
            "Durata giorni": durata,
            "nuts2_code": nuts,
            "Regione": nomi_reg.get(nuts, ""),
            "Attendance originale": valore,
            "Tipo indicatore originale": tipo_originale,
            "Base simulazione": base,
            "Categoria base": categoria,
            "Metodo base": metodo,
            "Utilizzabile": len(problemi) == 0,
            "Note verifica": "; ".join(problemi),
        })

    return pd.DataFrame(righe), None


def espandi_eventi_mese(eventi):
    """
    Ripartizione uniforme per giorni di esposizione.
    Non ricostruisce domanda turistica giornaliera.
    """
    righe = []

    for _, evento in eventi.iterrows():
        if not evento["Utilizzabile"]:
            continue

        inizio = evento["Data inizio"]
        fine = evento["Data fine"]
        durata = int(evento["Durata giorni"])

        mesi = pd.period_range(
            inizio.to_period("M"),
            fine.to_period("M"),
            freq="M",
        )

        for periodo in mesi:
            primo = periodo.start_time.normalize()
            ultimo = periodo.end_time.normalize()

            giorni = (
                min(fine, ultimo) - max(inizio, primo)
            ).days + 1

            riga = evento.to_dict()
            riga["year"] = periodo.year
            riga["month"] = periodo.month
            riga["data"] = primo
            riga["Giorni evento nel mese"] = giorni
            riga["Quota temporale"] = giorni / durata
            righe.append(riga)

    if not righe:
        return pd.DataFrame()

    risultato = pd.DataFrame(righe)

    denominatore = reg[
        [
            "nuts2_code",
            "year",
            "month",
            "guest_nights_total",
        ]
    ].copy()

    risultato = risultato.merge(
        denominatore,
        on=["nuts2_code", "year", "month"],
        how="left",
        validate="many_to_one",
    )

    return risultato


def calcola_scenario(base, quote, notti):
    risultato = base.copy()

    risultato["Quota pernottanti (%)"] = (
        risultato["Sport"].map(quote).astype(float)
    )

    risultato["Notti richieste"] = (
        risultato["Sport"].map(notti).astype(float)
    )

    risultato["Notti applicate"] = risultato[
        ["Notti richieste", "Durata giorni"]
    ].min(axis=1).clip(lower=1, upper=5)

    risultato["Pernottanti teorici evento"] = (
        risultato["Base simulazione"]
        * risultato["Quota pernottanti (%)"]
        / 100
    )

    risultato["Notti teoriche nel mese"] = (
        risultato["Pernottanti teorici evento"]
        * risultato["Notti applicate"]
        * risultato["Quota temporale"]
    )

    denominatore = risultato["guest_nights_total"].where(
        risultato["guest_nights_total"] > 0
    )

    risultato["Incidenza teorica (%)"] = (
        100
        * risultato["Notti teoriche nel mese"]
        / denominatore
    )

    risultato["Sopra soglia 2%"] = (
        risultato["Incidenza teorica (%)"] >= 2
    )

    return risultato


def riepilogo_mensile(risultato):
    gruppo = [
        "nuts2_code",
        "Regione",
        "year",
        "month",
        "data",
        "Categoria base",
    ]

    aggregato = (
        risultato.groupby(gruppo, as_index=False)
        .agg({
            "Notti teoriche nel mese": "sum",
            "guest_nights_total": "first",
            "event_id": "nunique",
        })
        .rename(columns={
            "event_id": "Eventi distinti",
        })
    )

    denominatore = aggregato["guest_nights_total"].where(
        aggregato["guest_nights_total"] > 0
    )

    aggregato["Incidenza teorica aggregata (%)"] = (
        100
        * aggregato["Notti teoriche nel mese"]
        / denominatore
    )

    return aggregato


# ============================================================
# 8. PANORAMA NAZIONALE
# ============================================================

if pagina == "Panorama nazionale":
    st.header("Panorama nazionale · 2022–2025")

    storico = (
        reg[reg["year"].between(2022, 2025)]
        .groupby("year")["guest_nights_total"]
        .sum(min_count=1)
    )

    v24 = storico.get(2024, float("nan"))
    v25 = storico.get(2025, float("nan"))

    crescita = (
        100 * (v25 / v24 - 1)
        if pd.notna(v24) and v24 > 0
        else float("nan")
    )

    futuro = fr.loc[
        fr["year"].eq(2030),
        f"gn_{scenario.lower()}",
    ].sum(min_count=1)

    a, b, c, d = st.columns(4)

    a.metric("Pernottamenti 2025", numero(v25))
    b.metric("Crescita 2025/2024", percentuale(crescita))
    c.metric(f"2030 · {scenario}", numero(futuro))
    d.metric(
        "Territori",
        f"{reg.nuts2_code.nunique()} NUTS2 · "
        f"{prov.nuts3_code.nunique()} NUTS3",
    )

    annuale = storico.rename(
        "Pernottamenti"
    ).reset_index()

    fig = px.bar(
        annuale,
        x="year",
        y="Pernottamenti",
        labels={"year": "Anno"},
        color_discrete_sequence=[ACCENTO],
    )
    fig.update_xaxes(dtick=1)
    grafico(fig, 360)

    st.caption(
        "2020–2021 esclusi dalla visualizzazione ordinaria. "
        "I dati descrivono la domanda sulle piattaforme "
        "incluse nella rilevazione Eurostat."
    )

    geo = (
        reg[reg["year"].eq(2025)]
        .groupby("nuts2_code", as_index=False)["guest_nights_total"]
        .sum(min_count=1)
    )

    geo["territorio"] = geo["nuts2_code"].map(nomi_reg)

    mappa(
        geo,
        2,
        "nuts2_code",
        "guest_nights_total",
        "Domanda regionale osservata · 2025",
    )


# ============================================================
# 9. ESPLORA UN TERRITORIO
# ============================================================

elif pagina == "Esplora un territorio":
    st.header("Domanda turistica osservata")

    livello = st.radio(
        "Frequenza",
        ["Mensile NUTS2", "Annuale NUTS3"],
        horizontal=True,
    )

    base = reg if livello == "Mensile NUTS2" else prov

    indicatori = {
        "Pernottamenti degli ospiti": "guest_nights_total",
        "Pernottamenti domestici": "guest_nights_domestic",
        "Pernottamenti stranieri": "guest_nights_foreign",
        "Soggiorni": "stays_total",
        "Notti di alloggio occupate": "occupied_nights_total",
    }

    disponibili = {
        nome: campo
        for nome, campo in indicatori.items()
        if campo in base.columns
    }

    misura = st.selectbox(
        "Indicatore",
        list(disponibili),
    )
    campo = disponibili[misura]

    dati = base[base["year"] >= 2022].copy()

    if regione != "Italia":
        dati = dati[dati["nuts2_code"].eq(regione)]

    territorio = (
        "Italia"
        if regione == "Italia"
        else nomi_reg[regione]
    )

    if livello == "Annuale NUTS3" and provincia != "Tutte":
        dati = dati[dati["nuts3_code"].eq(provincia)]
        territorio = nomi_prov[provincia]

    elif livello == "Mensile NUTS2" and provincia != "Tutte":
        st.caption(
            "La provincia non viene disaggregata mensilmente. "
            "Questa vista usa il territorio regionale selezionato."
        )

    st.subheader(territorio)

    if dati.empty:
        st.info("Nessun dato disponibile.")
        st.stop()

    if livello == "Mensile NUTS2":
        serie = (
            dati.groupby("data", as_index=False)[campo]
            .sum(min_count=1)
            .sort_values("data")
        )

        a, b = st.columns(2)
        a.metric(misura, numero(serie[campo].iloc[-1]))
        b.metric(
            "Ultimo mese",
            serie["data"].iloc[-1].strftime("%m/%Y"),
        )

        fig = px.line(
            serie,
            x="data",
            y=campo,
            labels={"data": "Mese", campo: misura},
            color_discrete_sequence=[ACCENTO],
        )
        grafico(fig)

        stag = serie.copy()
        stag["Anno"] = stag["data"].dt.year.astype(str)
        stag["Mese"] = stag["data"].dt.month

        st.subheader("Stagionalità · confronto degli stessi mesi")

        fig = px.line(
            stag,
            x="Mese",
            y=campo,
            color="Anno",
            markers=True,
            labels={campo: misura},
            color_discrete_sequence=[
                "#126889",
                "#249784",
                "#559EB8",
                "#DC813F",
                "#C94F7C",
            ],
        )
        fig.update_xaxes(dtick=1)
        grafico(fig)

        ultimo_anno = int(serie["data"].dt.year.max())
        mesi = serie.loc[
            serie["data"].dt.year.eq(ultimo_anno),
            "data",
        ].dt.month.tolist()

        attuale = serie[
            serie["data"].dt.year.eq(ultimo_anno)
            & serie["data"].dt.month.isin(mesi)
        ][campo].sum(min_count=1)

        precedente = serie[
            serie["data"].dt.year.eq(ultimo_anno - 1)
            & serie["data"].dt.month.isin(mesi)
        ][campo].sum(min_count=1)

        if pd.notna(precedente) and precedente > 0:
            st.info(
                f"Variazione {ultimo_anno}/{ultimo_anno - 1} "
                "sugli stessi mesi: "
                + percentuale(100 * (attuale / precedente - 1))
            )

        tabella(serie, "storico_regionale.csv")

    else:
        serie = (
            dati.groupby("year", as_index=False)[campo]
            .sum(min_count=1)
            .sort_values("year")
        )

        fig = px.bar(
            serie,
            x="year",
            y=campo,
            labels={"year": "Anno", campo: misura},
            color_discrete_sequence=[ACCENTO],
        )
        fig.update_xaxes(dtick=1)
        grafico(fig)

        st.caption(
            "Il dato provinciale è annuale: "
            "non vengono creati dati mensili artificiali."
        )

        tabella(serie, "storico_provinciale.csv")


# ============================================================
# 10. EVENTI E SIMULATORE
# ============================================================

elif pagina == "Eventi e simulatore":
    st.header("Eventi sportivi e screening esplorativo")

    st.info(
        "Il simulatore produce pernottamenti teorici "
        "condizionati alle ipotesi selezionate. "
        "Non misura un effetto causale e non modifica il forecast."
    )

    if not FILES["eventi"].is_file():
        st.error(
            "Event Master non trovato: "
            + FILES["eventi"].name
        )
        st.stop()

    try:
        master = leggi_excel(
            str(FILES["eventi"])
        ).copy()
    except Exception as exc:
        st.error(f"Impossibile leggere l’Event Master: {exc}")
        st.stop()

    mapping = mappa_colonne_eventi(master)
    eventi, errore = costruisci_eventi(master, mapping)

    if errore:
        st.warning(errore)
        tabella(master)
        st.stop()

    if eventi.empty:
        st.info("Il master non contiene eventi utilizzabili.")
        st.stop()

    if regione != "Italia":
        eventi = eventi[
            eventi["nuts2_code"].eq(regione)
        ].copy()

    catalogo_tab, simulatore_tab, storico_tab, file_tab = st.tabs([
        "Catalogo e scheda evento",
        "Simulatore · soglia 2%",
        "Eventi e domanda osservata",
        "Screening già calcolati",
    ])

    # --------------------------------------------------------
    # CATALOGO
    # --------------------------------------------------------

    with catalogo_tab:
        ricerca = st.text_input(
            "Cerca evento",
            key="cerca_catalogo",
        )

        sport_catalogo = sorted(
            eventi["Sport"].dropna().unique()
        )

        scelti_catalogo = st.multiselect(
            "Sport nel catalogo",
            sport_catalogo,
            default=sport_catalogo,
        )

        filtrati = eventi[
            eventi["Sport"].isin(scelti_catalogo)
        ].copy()

        if ricerca.strip():
            filtrati = filtrati[
                filtrati["Evento"].str.contains(
                    ricerca.strip(),
                    case=False,
                    regex=False,
                    na=False,
                )
            ]

        a, b, c = st.columns(3)
        a.metric("Eventi distinti", filtrati["event_id"].nunique())
        b.metric("Righe territoriali", len(filtrati))
        c.metric(
            "Righe con base utilizzabile",
            int(filtrati["Utilizzabile"].sum()),
        )

        if not filtrati.empty:
            chiavi = (
                filtrati["Evento"]
                + " · "
                + filtrati["Data inizio"]
                .dt.strftime("%d/%m/%Y")
                .fillna("data non disponibile")
                + " · "
                + filtrati["event_id"]
            )

            selezione = st.selectbox(
                "Scheda evento",
                sorted(chiavi.unique()),
            )

            dettaglio = filtrati.loc[
                chiavi.eq(selezione)
            ]

            st.subheader("Dati utilizzati dal simulatore")
            tabella(
                dettaglio.drop(
                    columns=["riga_master"],
                    errors="ignore",
                )
            )

            st.subheader("Informazioni originali e fonti")
            posizioni = dettaglio["riga_master"].astype(int).tolist()
            tabella(master.iloc[posizioni])

        with st.expander(
            "Catalogo filtrato",
            expanded=False,
        ):
            tabella(
                filtrati.drop(
                    columns=["riga_master"],
                    errors="ignore",
                ),
                "catalogo_eventi_filtrato.csv",
            )

        st.caption(
            "Persone ed equivalenti restano distinti. "
            "Le date sono visualizzate senza orario. "
            "Le colonne completamente vuote sono nascoste."
        )

    # --------------------------------------------------------
    # SIMULATORE
    # --------------------------------------------------------

    with simulatore_tab:
        st.subheader("Imposta le ipotesi per ciascuno sport")

        st.write(
            "Lo slider indica la quota della **base selezionata "
            "che ipotizziamo pernotti**. "
            "0% significa nessun pernottante; 100% significa "
            "che tutta quella base pernotta."
        )

        st.caption(
            "La base non coincide automaticamente con i visitatori "
            "fuori regione. Domestic/foreign e tifosi ospiti non "
            "vengono usati come sostituti automatici di tale quota."
        )

        utilizzabili = eventi[
            eventi["Utilizzabile"]
        ].copy()

        esclusi = eventi[
            ~eventi["Utilizzabile"]
        ].copy()

        if not esclusi.empty:
            st.warning(
                f"{len(esclusi)} righe non entrano nel simulatore "
                "per date, territorio, base o unità non utilizzabili."
            )

            with st.expander("Mostra righe escluse e motivi"):
                tabella(esclusi)

        if utilizzabili.empty:
            st.info(
                "Nessuna base simulabile. Controlla le associazioni "
                "delle colonne nell’espansore in alto."
            )

        else:
            mensile = espandi_eventi_mese(utilizzabili)

            anni = sorted(
                mensile.loc[
                    mensile["year"] >= 2022,
                    "year",
                ].unique()
            )

            if not anni:
                st.info("Nessun evento dal 2022 in avanti.")

            else:
                c1, c2 = st.columns(2)

                with c1:
                    anni_scelti = st.multiselect(
                        "Anni della simulazione",
                        anni,
                        default=anni,
                    )

                with c2:
                    mesi_scelti = st.multiselect(
                        "Mesi della simulazione",
                        list(range(1, 13)),
                        default=list(range(1, 13)),
                    )

                sport_disponibili = sorted(
                    mensile["Sport"].unique()
                )

                sport_scelti = st.multiselect(
                    "Sport inclusi",
                    sport_disponibili,
                    default=sport_disponibili,
                )

                base = mensile[
                    mensile["year"].isin(anni_scelti)
                    & mensile["month"].isin(mesi_scelti)
                    & mensile["Sport"].isin(sport_scelti)
                ].copy()

                if base.empty:
                    st.info("Nessun evento per i filtri selezionati.")

                else:
                    etichette_evento = (
                        base[
                            ["event_id", "Evento", "Data inizio"]
                        ]
                        .drop_duplicates("event_id")
                        .set_index("event_id")
                    )

                    ids = etichette_evento.index.tolist()

                    eventi_scelti = st.multiselect(
                        "Eventi inclusi",
                        ids,
                        default=ids,
                        format_func=lambda x: (
                            str(etichette_evento.loc[x, "Evento"])
                            + " · "
                            + str(
                                etichette_evento.loc[
                                    x, "Data inizio"
                                ].strftime("%d/%m/%Y")
                            )
                        ),
                    )

                    base = base[
                        base["event_id"].isin(eventi_scelti)
                    ].copy()

                    if base.empty:
                        st.info("Seleziona almeno un evento.")

                    else:
                        sport_attivi = sorted(
                            base["Sport"].unique()
                        )

                        preset = st.selectbox(
                            "Preset globale",
                            list(PRESET),
                            index=2,
                        )

                        # Inizializzazione prima della creazione
                        # dei widget.
                        for sport in sport_attivi:
                            key_q = f"quota_sport_{sport}"
                            key_n = f"notti_sport_{sport}"

                            if key_q not in st.session_state:
                                st.session_state[key_q] = 25

                            if key_n not in st.session_state:
                                st.session_state[key_n] = 1

                        if st.button(
                            "Applica il preset a tutti gli sport",
                            key="applica_preset",
                        ):
                            for sport in sport_attivi:
                                st.session_state[
                                    f"quota_sport_{sport}"
                                ] = PRESET[preset]

                        if st.button(
                            "Ripristina 25% e 1 notte",
                            key="reset_simulatore",
                        ):
                            for sport in sport_attivi:
                                st.session_state[
                                    f"quota_sport_{sport}"
                                ] = 25
                                st.session_state[
                                    f"notti_sport_{sport}"
                                ] = 1

                        quote = {}
                        notti = {}

                        for sport in sport_attivi:
                            st.markdown(f"**{html.escape(sport)}**")

                            max_notti = int(
                                min(
                                    5,
                                    base.loc[
                                        base["Sport"].eq(sport),
                                        "Durata giorni",
                                    ].max(),
                                )
                            )
                            max_notti = max(1, max_notti)

                            key_q = f"quota_sport_{sport}"
                            key_n = f"notti_sport_{sport}"

                            if st.session_state[key_n] > max_notti:
                                st.session_state[key_n] = max_notti

                            sinistra, destra = st.columns(2)

                            with sinistra:
                                quote[sport] = st.slider(
                                    f"Quota pernottanti · {sport}",
                                    min_value=0,
                                    max_value=100,
                                    step=1,
                                    key=key_q,
                                    format="%d%%",
                                )

                            with destra:
                                if max_notti == 1:
                                    notti[sport] = 1
                                    st.caption(
                                        "Notti per pernottante: 1. "
                                        "Gli eventi selezionati durano un giorno."
                                    )
                                else:
                                    notti[sport] = st.slider(
                                        f"Notti per pernottante · {sport}",
                                        min_value=1,
                                        max_value=max_notti,
                                        step=1,
                                        key=key_n,
                                    )

                        st.caption(
                            "Per ogni evento le notti applicate sono "
                            "limitate alla sua durata e a un massimo "
                            "di 5. Non sono imposti minimi di "
                            "pernottanti basati sui tifosi ospiti."
                        )

                        risultato = calcola_scenario(
                            base,
                            quote,
                            notti,
                        )

                        persone = risultato[
                            risultato["Categoria base"].eq("PERSONE")
                        ]

                        proxy = risultato[
                            risultato["Categoria base"].eq("EQUIVALENTE")
                        ]

                        a, b, c, d = st.columns(4)

                        a.metric(
                            "Notti teoriche · base persone",
                            numero(
                                persone["Notti teoriche nel mese"].sum()
                            ),
                        )

                        b.metric(
                            "Notti teoriche · base equivalente",
                            numero(
                                proxy["Notti teoriche nel mese"].sum()
                            ),
                        )

                        c.metric(
                            "Righe evento-regione-mese ≥ 2%",
                            int(risultato["Sopra soglia 2%"].sum()),
                        )

                        d.metric(
                            "Eventi selezionati",
                            risultato["event_id"].nunique(),
                        )

                        senza_domanda = (
                            risultato["guest_nights_total"].isna()
                            | risultato["guest_nights_total"].le(0)
                        )

                        if senza_domanda.any():
                            st.warning(
                                f"{int(senza_domanda.sum())} righe "
                                "non hanno domanda Eurostat positiva "
                                "disponibile: le notti sono simulate, "
                                "ma l’incidenza non viene calcolata."
                            )

                        st.info(
                            "Le due basi sono mostrate separatamente: "
                            "le notti da equivalenti non sono "
                            "pernottamenti misurati su visitatori unici."
                        )

                        st.subheader(
                            "Confronto dei cinque preset"
                        )

                        confronto_preset = []

                        for nome_preset, quota in PRESET.items():
                            alternativo = calcola_scenario(
                                base,
                                {s: quota for s in sport_attivi},
                                notti,
                            )

                            for categoria in [
                                "PERSONE",
                                "EQUIVALENTE",
                            ]:
                                subset = alternativo[
                                    alternativo["Categoria base"].eq(
                                        categoria
                                    )
                                ]

                                confronto_preset.append({
                                    "Preset": nome_preset,
                                    "Base": categoria,
                                    "Notti teoriche": subset[
                                        "Notti teoriche nel mese"
                                    ].sum(),
                                    "Righe ≥ 2%": int(
                                        subset["Sopra soglia 2%"].sum()
                                    ),
                                })

                        confronto_preset = pd.DataFrame(
                            confronto_preset
                        )

                        fig = px.bar(
                            confronto_preset,
                            x="Preset",
                            y="Notti teoriche",
                            color="Base",
                            barmode="group",
                            color_discrete_map={
                                "PERSONE": ACCENTO,
                                "EQUIVALENTE": COLORI["Intermedio"],
                            },
                        )
                        grafico(fig)

                        st.caption(
                            "Nel confronto varia la quota globale; "
                            "restano ferme le notti scelte per sport "
                            "e gli eventi selezionati."
                        )

                        tabella(
                            confronto_preset,
                            "simulatore_confronto_preset.csv",
                        )

                        st.subheader(
                            "Incidenza per regione e mese"
                        )

                        aggregato = riepilogo_mensile(risultato)

                        st.caption(
                            "Il denominatore regionale mensile viene "
                            "conteggiato una sola volta per gruppo. "
                            "L’aggregazione somma esposizioni teoriche: "
                            "non elimina eventuali visitatori comuni "
                            "a eventi diversi."
                        )

                        tabella(
                            aggregato,
                            "simulatore_regioni_mesi.csv",
                        )

                        st.subheader(
                            "Dettaglio evento × regione × mese"
                        )

                        solo_soglia = st.checkbox(
                            "Mostra soltanto le righe almeno al 2%",
                            key="soglia_live",
                        )

                        dettaglio = (
                            risultato[
                                risultato["Sopra soglia 2%"]
                            ]
                            if solo_soglia
                            else risultato
                        )

                        colonne = [
                            "Evento",
                            "Sport",
                            "Regione",
                            "year",
                            "month",
                            "Data inizio",
                            "Data fine",
                            "Categoria base",
                            "Base simulazione",
                            "Metodo base",
                            "Quota pernottanti (%)",
                            "Notti applicate",
                            "Giorni evento nel mese",
                            "Quota temporale",
                            "Notti teoriche nel mese",
                            "guest_nights_total",
                            "Incidenza teorica (%)",
                            "Sopra soglia 2%",
                        ]

                        tabella(
                            dettaglio[colonne],
                            "simulatore_dettaglio.csv",
                        )

                        with st.expander(
                            "Formula e interpretazione dei risultati"
                        ):
                            st.latex(
                                r"T_{e,r,m}="
                                r"B_{e,r}\times q_s"
                                r"\times n_e\times w_{e,m}"
                            )

                            st.write(
                                "**B:** base persone oppure equivalente. "
                                "**q:** quota di pernottanti. "
                                "**n:** notti applicate, limitate "
                                "alla durata dell’evento e a 5. "
                                "**w:** quota dei giorni dell’evento "
                                "che ricade nel mese."
                            )

                            st.latex(
                                r"I_{e,r,m}=100\times"
                                r"\frac{T_{e,r,m}}{Y_{r,m}}"
                            )

                            st.write(
                                "**Y:** guest nights Eurostat osservate "
                                "nella stessa regione e nello stesso mese. "
                                "Il 2% è una soglia esplorativa, "
                                "non significatività statistica."
                            )

                            st.write(
                                "La ripartizione temporale uniforme "
                                "è un’ipotesi operativa: non ricostruisce "
                                "i soggiorni effettivi prima, durante "
                                "o dopo l’evento."
                            )

    # --------------------------------------------------------
    # EVENTI E DOMANDA OSSERVATA
    # --------------------------------------------------------

    with storico_tab:
        scelta_reg = (
            regione
            if regione != "Italia"
            else st.selectbox(
                "Regione da confrontare",
                sorted(nomi_reg),
                format_func=lambda x: nomi_reg[x],
            )
        )

        anni_storici = sorted(
            reg.loc[reg["year"] >= 2022, "year"].unique()
        )

        anno = st.selectbox(
            "Anno osservato",
            anni_storici,
            index=len(anni_storici) - 1,
        )

        mese = st.selectbox(
            "Mese osservato",
            list(range(1, 13)),
        )

        base_storica = reg[
            reg["year"].ge(2022)
            & reg["month"].eq(mese)
        ]

        nazionale = (
            base_storica.groupby("year")["guest_nights_total"]
            .sum(min_count=1)
        )

        locale = (
            base_storica[
                base_storica["nuts2_code"].eq(scelta_reg)
            ]
            .set_index("year")["guest_nights_total"]
        )

        confronto = pd.DataFrame({
            "Regione": locale,
            "Resto Italia": nazionale - locale,
        }).sort_index()

        yoy = confronto.pct_change(fill_method=None) * 100

        fig = go.Figure()

        for nome, colore in [
            ("Regione", ACCENTO),
            ("Resto Italia", COLORI["Intermedio"]),
        ]:
            fig.add_trace(go.Scatter(
                x=yoy.index,
                y=yoy[nome],
                name=nome,
                mode="lines+markers",
                line=dict(color=colore, width=3),
            ))

        fig.update_yaxes(title="Variazione annua (%)")
        fig.update_xaxes(title="Anno", dtick=1)
        grafico(fig)

        primo = pd.Timestamp(
            year=int(anno),
            month=int(mese),
            day=1,
        )
        ultimo = primo + pd.offsets.MonthEnd(0)

        selezione = eventi[
            eventi["nuts2_code"].eq(scelta_reg)
            & eventi["Data inizio"].le(ultimo)
            & eventi["Data fine"].ge(primo)
        ]

        st.subheader("Eventi presenti nel mese selezionato")

        tabella(
            selezione.drop(
                columns=["riga_master"],
                errors="ignore",
            )
        )

        st.caption(
            "Il resto d’Italia non è automaticamente "
            "un gruppo di controllo non esposto: "
            "può ospitare altri eventi. "
            "Il confronto è descrittivo."
        )

    # --------------------------------------------------------
    # RISULTATI STORICI GIÀ SALVATI
    # --------------------------------------------------------

    with file_tab:
        st.write(
            "Questi sono gli screening già presenti nei file. "
            "Gli slider del simulatore non li sovrascrivono."
        )

        for chiave, titolo, export in [
            (
                "screening",
                "Screening per evento",
                "screening_storico_eventi.csv",
            ),
            (
                "mesi_eventi",
                "Screening per regione e mese",
                "screening_storico_regioni_mesi.csv",
            ),
        ]:
            with st.expander(titolo):
                if FILES[chiave].is_file():
                    storico_file = leggi_csv(
                        str(FILES[chiave])
                    ).copy()

                    col_nuts = trova_colonna(
                        storico_file,
                        "nuts2_code",
                        "NUTS2",
                        "Codici NUTS2",
                    )

                    if col_nuts and regione != "Italia":
                        storico_file = storico_file[
                            storico_file[col_nuts]
                            .astype(str)
                            .str.contains(
                                regione,
                                regex=False,
                                na=False,
                            )
                        ]

                    tabella(storico_file, export)

                else:
                    st.info("File non disponibile.")


# ============================================================
# 11. FORECAST 2026–2030
# ============================================================

elif pagina == "Forecast 2026–2030":
    st.header("Forecast 2026–2030")

    st.info(
        "Gli scenari rappresentano ipotesi alternative, "
        "non intervalli di confidenza. "
        "Il simulatore sportivo non viene sommato al forecast."
    )

    attivi = st.multiselect(
        "Scenari da visualizzare",
        SCENARI,
        default=SCENARI,
    )

    fig = go.Figure()

    if provincia != "Tutte":
        storico = prov[
            prov["nuts3_code"].eq(provincia)
            & prov["year"].ge(2022)
        ].sort_values("year")

        futuro = fp[
            fp["nuts3_code"].eq(provincia)
        ].sort_values("year")

        st.subheader(
            nomi_prov[provincia] + " · NUTS3 annuale"
        )

        x = storico["year"].tolist()
        y = storico["guest_nights_total"].tolist()

        if not x:
            st.warning("Storico non disponibile.")
            st.stop()

        for s in attivi:
            fig.add_trace(go.Scatter(
                x=[x[-1]] + futuro["year"].tolist(),
                y=[y[-1]] + futuro[
                    f"forecast_{s.lower()}"
                ].tolist(),
                name=s,
                mode="lines+markers",
                line=dict(
                    color=COLORI[s],
                    dash="dash",
                    width=3,
                ),
            ))

        dati_annuali = futuro
        prefisso = "forecast_"

        st.caption(
            "Previsione provinciale ottenuta ripartendo "
            "quella regionale con quote 2025 costanti."
        )

    else:
        base = (
            reg
            if regione == "Italia"
            else reg[reg["nuts2_code"].eq(regione)]
        )

        storico = (
            base[base["year"].ge(2022)]
            .groupby("data", as_index=False)["guest_nights_total"]
            .sum(min_count=1)
            .sort_values("data")
        )

        dati_annuali = (
            fr
            if regione == "Italia"
            else fr[fr["nuts2_code"].eq(regione)]
        )

        campi = [
            f"gn_{s.lower()}"
            for s in SCENARI
        ]

        futuro = (
            dati_annuali[
                dati_annuali["data"] > storico["data"].max()
            ]
            .groupby("data", as_index=False)[campi]
            .sum(min_count=1)
            .sort_values("data")
        )

        x = storico["data"].tolist()
        y = storico["guest_nights_total"].tolist()

        if not x:
            st.warning("Storico non disponibile.")
            st.stop()

        st.subheader(
            (
                "Italia"
                if regione == "Italia"
                else nomi_reg[regione]
            ) + " · NUTS2 mensile"
        )

        for s in attivi:
            fig.add_trace(go.Scatter(
                x=[x[-1]] + futuro["data"].tolist(),
                y=[y[-1]] + futuro[
                    f"gn_{s.lower()}"
                ].tolist(),
                name=s,
                mode="lines",
                line=dict(
                    color=COLORI[s],
                    dash="dash",
                    width=3,
                ),
            ))

        prefisso = "gn_"

    fig.add_trace(go.Scatter(
        x=x,
        y=y,
        name="Osservato",
        mode="lines",
        line=dict(color=ACCENTO, width=3),
    ))

    fig.add_shape(
        type="line",
        x0=x[-1],
        x1=x[-1],
        y0=0,
        y1=1,
        yref="paper",
        line=dict(
            color=ACCENTO,
            dash="dot",
            width=1,
        ),
    )

    fig.update_yaxes(title="Pernottamenti degli ospiti")
    grafico(fig, 480)

    st.caption(
        "Linea continua: osservato. Linee tratteggiate: scenari. "
        "La linea verticale segna l’ultima osservazione."
    )

    confronto = pd.DataFrame({
        "Anno": list(range(2026, 2031))
    })

    cards = st.columns(3)

    for card, s in zip(cards, SCENARI):
        annuale = (
            dati_annuali.groupby("year")[
                prefisso + s.lower()
            ]
            .sum(min_count=1)
        )

        confronto[s] = confronto["Anno"].map(annuale)

        card.metric(
            "2030 · " + s,
            numero(annuale.get(2030, float("nan"))),
        )

    fig = go.Figure()

    for s in attivi:
        fig.add_trace(go.Scatter(
            x=confronto["Anno"],
            y=confronto[s],
            name=s,
            mode="lines+markers",
            line=dict(color=COLORI[s], width=3),
        ))

    fig.update_xaxes(title="Anno", dtick=1)
    fig.update_yaxes(title="Pernottamenti degli ospiti")
    grafico(fig, 350)

    tabella(confronto, "forecast_scenari.csv")

    st.caption(
        "I totali annuali sono letti dai CSV previsionali. "
        "Non vengono ricalcolati dall’app."
    )

    st.divider()

    anno_mappa = st.select_slider(
        "Anno della mappa",
        options=list(range(2026, 2031)),
        value=2030,
    )

    livello_mappa = st.radio(
        "Geografia della mappa",
        ["NUTS2", "NUTS3"],
        horizontal=True,
    )

    scenario_mappa = st.selectbox(
        "Scenario della mappa",
        SCENARI,
        index=SCENARI.index(scenario),
    )

    if livello_mappa == "NUTS2":
        campo = f"gn_{scenario_mappa.lower()}"

        geo = (
            fr[fr["year"].eq(anno_mappa)]
            .groupby("nuts2_code", as_index=False)[campo]
            .sum(min_count=1)
        )

        geo["territorio"] = geo["nuts2_code"].map(nomi_reg)

        mappa(
            geo,
            2,
            "nuts2_code",
            campo,
            f"NUTS2 · {anno_mappa} · {scenario_mappa}",
        )

    else:
        campo = f"forecast_{scenario_mappa.lower()}"

        geo = fp.loc[
            fp["year"].eq(anno_mappa),
            ["nuts3_code", campo],
        ].copy()

        geo["territorio"] = geo["nuts3_code"].map(nomi_prov)

        mappa(
            geo,
            3,
            "nuts3_code",
            campo,
            f"NUTS3 · {anno_mappa} · {scenario_mappa}",
        )


# ============================================================
# 12. DATI E METODO
# ============================================================

else:
    st.header("Dati e metodo")

    st.write(
        "**Domanda di ricerca:** quanto cambia la domanda "
        "di alloggi di breve durata nei territori italiani "
        "durante i periodi interessati da grandi eventi sportivi, "
        "e quali scenari possiamo costruire per il futuro?"
    )

    st.info(
        "La dashboard distingue osservazioni, simulazioni teoriche "
        "e forecast. Non ha ancora identificato un effetto "
        "causale degli eventi sportivi."
    )

    a, b, c, d = st.columns(4)

    a.metric("NUTS2", reg["nuts2_code"].nunique())
    b.metric("NUTS3", prov["nuts3_code"].nunique())
    c.metric(
        "Ultimo mese osservato",
        reg["data"].max().strftime("%m/%Y"),
    )
    d.metric(
        "Ultimo anno provinciale",
        int(prov["year"].max()),
    )

    sintesi, fonti, eventi_metodo, forecast_metodo, qualita, limiti = (
        st.tabs([
            "Interpretazione",
            "Fonti e indicatori",
            "Simulazione eventi",
            "Forecast",
            "Qualità e validazione",
            "Limiti e sviluppi",
        ])
    )

    with sintesi:
        st.subheader("Tre livelli di lettura")

        tabella(pd.DataFrame([
            [
                "Osservato",
                "Domanda registrata nei file Eurostat.",
                "Descrive ciò che è accaduto, non perché.",
            ],
            [
                "Simulazione",
                "Pernottamenti teorici secondo quote e notti scelte.",
                "È condizionata alle ipotesi, non un dato osservato.",
            ],
            [
                "Forecast",
                "Traiettorie future dei CSV previsionali.",
                "Non è una garanzia né una stima causale sportiva.",
            ],
        ], columns=[
            "Risultato",
            "Significato",
            "Interpretazione",
        ]))

        st.markdown(
            "1. **Eurostat Master:** dati turistici e controlli "
            "di copertura.\n"
            "2. **Event Master:** eventi, date, territori e fonti.\n"
            "3. **Incrocio storico:** collegamento regione–mese "
            "e confronto descrittivo.\n"
            "4. **Scenari e forecast:** simulazioni teoriche degli "
            "eventi e previsioni della domanda, mantenute separate."
        )

        st.write(
            "**Esempio:** una crescita del 12% nel mese di un evento "
            "non dimostra che l’evento abbia causato il 12%. "
            "Il risultato osservato può dipendere anche da "
            "stagionalità, trend, festività, offerta e altri fattori."
        )

    with fonti:
        st.subheader("Perimetro della rilevazione")

        st.write(
            "Eurostat considera gli alloggi di breve durata "
            "prenotati attraverso le piattaforme partecipanti. "
            "Il dato non coincide con il solo Airbnb né con "
            "l’intero mercato degli affitti brevi."
        )

        tabella(pd.DataFrame([
            [
                "tour_ce_omn12",
                "NUTS2 × mese",
                "Domanda regionale; riferimento per il simulatore.",
            ],
            [
                "tour_ce_oan3",
                "NUTS3 × anno",
                "Dettaglio provinciale annuale.",
            ],
            [
                "Event Master",
                "Evento × territorio",
                "Caratteristiche, attendance, fonti e qualità.",
            ],
            [
                "GISCO NUTS 2024",
                "Geometrie ufficiali",
                "Confini delle mappe.",
            ],
        ], columns=["Fonte", "Granularità", "Utilizzo"]))

        st.subheader("Indicatori")

        tabella(pd.DataFrame([
            [
                "Guest nights",
                "Pernottamenti delle persone ospitate.",
            ],
            [
                "Stays",
                "Soggiorni / prenotazioni, non numero di ospiti.",
            ],
            [
                "Occupied nights",
                "Notti di alloggio occupate, senza moltiplicare "
                "per il numero di persone.",
            ],
            [
                "Domestic / foreign",
                "Residenza in Italia / all’estero; "
                "non provenienza da dentro o fuori regione.",
            ],
        ], columns=["Indicatore", "Definizione"]))

        st.write(
            "**Esempio:** tre persone per due notti nello stesso "
            "alloggio corrispondono a un soggiorno, "
            "due notti occupate e sei guest nights."
        )

        st.markdown(
            "[Metadati ufficiali Eurostat]"
            "(https://ec.europa.eu/eurostat/cache/"
            "metadata/en/tour_ce_esms.htm)"
        )

        st.subheader("Periodo")

        st.write(
            "Il riferimento principale parte dal 2022, "
            "con i mesi 2026 effettivamente disponibili. "
            "Il 2020–2021 non è trattato come periodo ordinario. "
            "Il 2022 è una scelta operativa post-Covid, "
            "non la garanzia di completa normalizzazione "
            "in ogni territorio."
        )

        st.write(
            "Gli anni incompleti vanno confrontati sugli stessi "
            "mesi. Non vengono creati dati mensili provinciali "
            "partendo da osservazioni annuali."
        )

        st.write(
            "**Prezzi:** un eventuale HICP dei servizi di alloggio "
            "è un controllo macro, non un prezzo Airbnb. "
            "Questa app non stima ADR, ricavi o tassi di occupazione."
        )

    with eventi_metodo:
        st.subheader("Base e ipotesi della simulazione")

        st.write(
            "La simulazione parte dalle righe territoriali "
            "dell’Event Master. Non assegna arbitrariamente "
            "a una regione le righe contenenti più NUTS2. "
            "Le righe non interpretabili vengono segnalate."
        )

        st.write(
            "**Persone:** visitatori unici esplicitamente "
            "documentati/stimati oppure attendance classificata "
            "come persone o partecipanti.\n\n"
            "**Equivalenti:** proxy già presente nel file "
            "oppure ingressi cumulativi divisi per la durata. "
            "Non sono visitatori unici.\n\n"
            "**Biglietti:** non vengono convertiti automaticamente "
            "in persone. In assenza di una base persone o equivalente "
            "documentata, la riga resta fuori dal calcolo."
        )

        st.subheader("Quota di pernottanti")

        st.write(
            "Ogni sport ha uno slider 0–100%, inizialmente al 25%. "
            "I preset globali sono 5%, 15%, 25%, 40% e 60%. "
            "Non vengono imposti minimi legati ai tifosi ospiti."
        )

        st.write(
            "La percentuale si applica alla base mostrata. "
            "Solo se quella base rappresenta effettivamente "
            "visitatori fuori regione può essere interpretata "
            "come quota di tali visitatori che pernotta."
        )

        st.subheader("Notti e durata")

        st.write(
            "Il valore iniziale è una notte. "
            "La scelta è indipendente per sport. "
            "Per ogni evento si applica il minore fra "
            "notti richieste, durata dell’evento e cinque notti."
        )

        st.latex(
            r"n_e=\min(n_s,D_e,5)"
        )

        st.subheader("Eventi su più mesi")

        st.latex(
            r"d_{e,m}=\max\left("
            r"0,\min(f_e,f_m)-\max(i_e,i_m)+1"
            r"\right)"
        )

        st.latex(
            r"w_{e,m}=\frac{d_{e,m}}{D_e}"
        )

        st.write(
            "Un evento dal 25 gennaio all’8 febbraio ha "
            "sette giorni a gennaio e otto a febbraio. "
            "Le notti teoriche vengono distribuite "
            "proporzionalmente ai giorni di esposizione."
        )

        st.caption(
            "È una ripartizione uniforme dichiarata, "
            "non una ricostruzione di soggiorni giornalieri. "
            "Non identifica arrivi anticipati o partenze posticipate."
        )

        st.subheader("Formula")

        st.latex(
            r"T_{e,r,m}=B_{e,r}\times q_s\times n_e\times w_{e,m}"
        )

        st.write(
            "B è la base persone o equivalente; q è la quota "
            "di pernottanti; n sono le notti applicate; "
            "w è la quota temporale attribuita al mese."
        )

        st.latex(
            r"I_{e,r,m}=100\times\frac{T_{e,r,m}}{Y_{r,m}}"
        )

        st.write(
            "Y è la domanda Eurostat osservata nella stessa "
            "regione e nello stesso mese. "
            "Se Y manca o non è positivo, l’incidenza "
            "non viene calcolata."
        )

        st.info(
            "Il 2% è una soglia esplorativa, non un test "
            "di significatività. Superarla suggerisce "
            "un approfondimento, non dimostra un effetto."
        )

        st.write(
            "**Perimetro:** il simulatore non osserva il canale "
            "di prenotazione degli spettatori. "
            "Se le notti teoriche includono hotel, ospitalità "
            "privata o altri canali, il rapporto con Eurostat "
            "è un confronto di scala, non una quota verificata "
            "del mercato delle piattaforme."
        )

        st.write(
            "Persone ed equivalenti rimangono separati. "
            "La somma fra eventi non deduplica persone "
            "che partecipano a più manifestazioni."
        )

    with forecast_metodo:
        st.subheader("Previsioni caricate dai CSV")

        st.write(
            "La dashboard legge gli output previsionali; "
            "non addestra un modello durante la consultazione. "
            "L’impostazione dichiarata usa stagionalità 2025, "
            "componente iniziale 50% nazionale e 50% regionale "
            "e tre ipotesi di rallentamento."
        )

        st.latex(
            r"g_{r,0}=0.5g_{IT,0}+0.5g^{regionale}_{r,0}"
        )

        st.write(
            "Il codice che genera i CSV resta la fonte "
            "per finestra di stima, tassi e parametri esatti. "
            "Non vengono dedotti dal nome dello scenario."
        )

        tabella(pd.DataFrame([
            [
                "Prudente",
                "Crescita più contenuta fra le ipotesi adottate.",
            ],
            [
                "Intermedio",
                "Traiettoria intermedia, non probabilità maggiore.",
            ],
            [
                "Espansivo",
                "Crescita più elevata fra le ipotesi adottate.",
            ],
        ], columns=["Scenario", "Lettura"]))

        st.write(
            "Gli scenari non sono intervalli di confidenza. "
            "La distanza fra le curve rappresenta sensibilità "
            "alle ipotesi, non tutta l’incertezza."
        )

        st.subheader("Quote provinciali")

        st.latex(
            r"w_{p|r,2025}="
            r"\frac{Y_{p,2025}}{\sum_{j\in r}Y_{j,2025}}"
        )

        st.latex(
            r"\widehat Y_{p,t,s}="
            r"w_{p|r,2025}\widehat Y_{r,t,s}"
        )

        st.write(
            "Le quote provinciali del 2025 restano costanti. "
            "Non si tratta di modelli provinciali indipendenti. "
            "La somma provinciale deve riconciliarsi "
            "con il forecast regionale."
        )

        st.subheader("Rapporto con gli eventi")

        st.write(
            "La simulazione sportiva non viene sommata al forecast. "
            "Una baseline costruita sullo storico può già "
            "incorporare la normale presenza di eventi ricorrenti: "
            "non è automaticamente un controfattuale senza sport."
        )

        st.caption(
            "Nei grafici lo storico prevale sui mesi già osservati. "
            "I totali annuali restano quelli dei CSV: la loro "
            "coerenza con le osservazioni sovrapposte va "
            "verificata nello script di generazione."
        )

    with qualita:
        st.subheader("Controlli eseguiti sui dati")

        controlli_output = []

        for nome, df, chiavi, misure in [
            (
                "Master NUTS2",
                reg,
                ["nuts2_code", "year", "month"],
                ["guest_nights_total"],
            ),
            (
                "Master NUTS3",
                prov,
                ["nuts3_code", "year"],
                ["guest_nights_total"],
            ),
            (
                "Forecast NUTS2",
                fr,
                ["nuts2_code", "year", "month"],
                [f"gn_{s.lower()}" for s in SCENARI],
            ),
            (
                "Forecast NUTS3",
                fp,
                ["nuts3_code", "year"],
                [f"forecast_{s.lower()}" for s in SCENARI],
            ),
        ]:
            valori = df[misure].apply(
                pd.to_numeric,
                errors="coerce",
            )

            errori = {
                "Chiavi mancanti": int(
                    df[chiavi].isna().any(axis=1).sum()
                ),
                "Chiavi duplicate": int(
                    df.duplicated(chiavi).sum()
                ),
                "Valori mancanti, negativi o infiniti": int(
                    (
                        valori.isna()
                        | valori.lt(0)
                        | valori.isin([
                            float("inf"),
                            -float("inf"),
                        ])
                    ).sum().sum()
                ),
            }

            for controllo, n in errori.items():
                controlli_output.append({
                    "Dataset": nome,
                    "Controllo": controllo,
                    "Esito": "OK" if n == 0 else "DA VERIFICARE",
                    "Anomalie": n,
                })

        for s in SCENARI:
            regionale = (
                fr.groupby(["nuts2_code", "year"])[
                    f"gn_{s.lower()}"
                ]
                .sum(min_count=1)
            )

            provinciale = (
                fp.groupby(["nuts2_code", "year"])[
                    f"forecast_{s.lower()}"
                ]
                .sum(min_count=1)
            )

            verifica = pd.concat([
                regionale.rename("Regione"),
                provinciale.rename("Province"),
            ], axis=1)

            scarto = (
                verifica["Regione"] - verifica["Province"]
            ).abs()

            anomalie = int(
                (
                    verifica.isna().any(axis=1)
                    | scarto.gt(1)
                ).sum()
            )

            controlli_output.append({
                "Dataset": "Forecast",
                "Controllo": f"Riconciliazione NUTS3/NUTS2 · {s}",
                "Esito": (
                    "OK"
                    if anomalie == 0 and not verifica.empty
                    else "DA VERIFICARE"
                ),
                "Anomalie": anomalie,
            })

        tabella(
            pd.DataFrame(controlli_output),
            "controlli_dati.csv",
        )

        st.caption(
            "Tolleranza di riconciliazione: una guest night. "
            "Un esito OK verifica il requisito indicato, "
            "non certifica l’intera metodologia."
        )

        st.subheader("WAPE e bias")

        st.latex(
            r"WAPE=100\times"
            r"\frac{\sum_i|Y_i-\widehat Y_i|}{\sum_iY_i}"
        )

        st.write(
            "Il WAPE misura l’errore assoluto complessivo "
            "rispetto al volume osservato. "
            "Può nascondere errori importanti nei territori "
            "piccoli, perché pesa maggiormente quelli grandi."
        )

        st.latex(
            r"Bias=100\times"
            r"\frac{\sum_i(\widehat Y_i-Y_i)}{\sum_iY_i}"
        )

        st.caption(
            "Bias negativo: sottostima. "
            "Bias positivo: sovrastima. "
            "Le metriche richiedono un denominatore positivo."
        )

        st.subheader("Risultati documentati nel progetto")

        tabella(pd.DataFrame([
            ["Stagionale 2024", "11,11%", "Test storico"],
            ["Media 2023–2024", "17,13%", "Test storico"],
            ["Prudente", "7,48%", "Test 2025 dichiarato"],
            ["Intermedio", "8,39%", "Test 2025 dichiarato"],
            ["Espansivo", "9,89%", "Test 2025 dichiarato"],
        ], columns=["Metodo", "WAPE", "Campione dichiarato"]))

        st.warning(
            "Questi backtest sono riportati, non ricalcolati "
            "dalla dashboard. Campioni, finestre e parametri "
            "devono essere allineati prima di confrontarli "
            "come una classifica omogenea."
        )

        st.write(
            "La validazione deve usare solo informazioni "
            "disponibili alla data della previsione, separare "
            "cronologicamente training e test e confrontare "
            "i modelli sullo stesso campione. "
            "Quote e parametri del test non devono usare "
            "informazioni future."
        )

    with limiti:
        st.subheader("Limiti interpretativi")

        st.markdown(
            "- **Aggregazione:** un evento locale può essere "
            "diluito nel dato regionale mensile.\n"
            "- **Concomitanza:** altri eventi, festività, prezzi, "
            "trasporti e offerta possono influenzare la domanda.\n"
            "- **Sostituzione:** gli ospiti sportivi possono "
            "rimpiazzare altri visitatori.\n"
            "- **Spostamenti:** notti trasferite tra mesi o regioni "
            "non equivalgono sempre a un aumento nazionale.\n"
            "- **Attendance:** accessi e biglietti non sono "
            "automaticamente persone uniche.\n"
            "- **Copertura:** il catalogo non è un censimento "
            "completo di tutte le manifestazioni italiane."
        )

        st.subheader("Sviluppo successivo")

        st.write(
            "Per stimare un effetto incrementale occorre "
            "costruire un controfattuale credibile. "
            "Panel, event study e difference-in-differences "
            "sono possibili sviluppi da concordare, "
            "non analisi già svolte da questa app."
        )

        st.write(
            "Prima di formulare conclusioni causali servono "
            "verifiche sui trend precedenti, anticipazioni, "
            "territori di controllo, eventi concomitanti, "
            "incertezza e robustezza delle stime."
        )

        st.success(
            "Prima dati e ipotesi trasparenti e riproducibili; "
            "poi maggiore complessità solo se necessaria "
            "alla domanda di ricerca."
        )


# ============================================================
# 13. FOOTER
# ============================================================

st.divider()

st.caption(
    "Airbnb&Sports · "
    + VERSIONE
    + " · Dati osservati, simulazioni e forecast distinti"
)

st.markdown(
    f"[SBL Consultancy · sito aziendale]({SITO_AZIENDA})"
)

# FINE DEL FILE APP.PY
