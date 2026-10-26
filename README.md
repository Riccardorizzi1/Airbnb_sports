# Airbnb&Sports

Dashboard Streamlit per l'analisi della domanda
di short-term accommodation in Italia.

## Contenuto

- Eurostat NUTS2 mensile
- Eurostat NUTS3 annuale
- Forecast 2026–2030 con tre scenari
- Event Master sportivo
- Screening storico degli eventi
- Mappe territoriali Kepler.gl, se previste da app.py

## Pubblicazione

1. Estrarre lo ZIP.
2. Caricare tutti i file su GitHub.
3. Mantenere le sottocartelle originali.
4. Aprire https://share.streamlit.io/
5. Selezionare il repository GitHub.
6. Main file path: app.py
7. Avviare il deployment.

## Esecuzione locale

python -m pip install -r requirements.txt
python -m streamlit run app.py

## Nota metodologica

L'analisi principale utilizza il periodo dal 2022.

Lo screening sportivo non rappresenta una stima
causale dell'impatto sulle guest nights.

I tre scenari di forecast sono ipotesi alternative.
