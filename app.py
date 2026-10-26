
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
