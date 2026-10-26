            ("Master NUTS2", reg, ["nuts2_code", "year", "month"], ["guest_nights_total"]),
            ("Master NUTS3", prov, ["nuts3_code", "year"], ["guest_nights_total"]),
            ("Forecast NUTS2", fr, ["nuts2_code", "year", "month"], [f"gn_{s.lower()}" for s in SCENARI]),
            ("Forecast NUTS3", fp, ["nuts3_code", "year"], [f"forecast_{s.lower()}" for s in SCENARI]),
        ]:
            null_chiavi = int(df[chiave].isna().any(axis=1).sum())
            duplicati = int(df.duplicated(chiave).sum())
            valori = df[misure].apply(pd.to_numeric, errors="coerce")
            non_validi = int((valori.isna() | valori.isin([float("inf"), -float("inf")]) | valori.lt(0)).sum().sum())
            for controllo, n in [("Chiavi mancanti", null_chiavi), ("Chiavi duplicate", duplicati), ("Valori mancanti, non numerici, infiniti o negativi", non_validi)]:
                esiti.append([etichetta, controllo, "OK" if n == 0 else "DA VERIFICARE", str(n)])
        for etichetta, df in [("Master NUTS2", reg), ("Master NUTS3", prov)]:
            for base in ["guest_nights", "stays", "occupied_nights"]:
                colonne = [f"{base}_{x}" for x in ["total", "domestic", "foreign"]]
                if all(c in df for c in colonne):
                    v = df[colonne].apply(pd.to_numeric, errors="coerce")
                    scarto = (v.iloc[:, 0] - v.iloc[:, 1] - v.iloc[:, 2]).abs()
                    valido = v.notna().all(axis=1)
                    n = int((~valido | scarto.gt(1)).sum())
                    esiti.append([etichetta, f"Totale = domestici + stranieri · {base}", "OK" if n == 0 else "DA VERIFICARE", f"{n} righe da verificare; tolleranza 1"])
        for scenario_check in SCENARI:
            r = fr.groupby(["nuts2_code", "year"])[f"gn_{scenario_check.lower()}"].sum(min_count=1)
            p = fp.groupby(["nuts2_code", "year"])[f"forecast_{scenario_check.lower()}"].sum(min_count=1)
            confronto_check = pd.concat([r.rename("Regione"), p.rename("Somma province")], axis=1)
            scarto = confronto_check.diff(axis=1).iloc[:, 1].abs()
            ok = (not confronto_check.empty and confronto_check.notna().all().all() and scarto.le(1).all())
            esiti.append(["Forecast", f"Riconciliazione NUTS3/NUTS2 · {scenario_check}", "OK" if ok else "DA VERIFICARE", f"Scarto massimo: {numero(scarto.max(), 2)}; tolleranza 1"])
        tabella(pd.DataFrame(esiti, columns=["Dataset", "Controllo", "Esito", "Dettaglio"]), "controlli_metodologici.csv")
        st.write(
            "**Ulteriori verifiche necessarie a monte:** continuità del panel per ogni territorio, "
            "flag e revisioni Eurostat, corrispondenza dei codici con le geometrie, tracciabilità "
            "delle fonti evento, date e ripartizioni multisede, parametri e versione del forecast. "
            "I controlli sopra non sostituiscono queste verifiche."
        )
        st.subheader("Come valutare la capacità predittiva")
        st.latex(r"\mathrm{WAPE}=100\times\frac{\sum_i|Y_i-\widehat Y_i|}{\sum_iY_i}")
        st.write(
            "Il WAPE misura l’errore assoluto complessivo rispetto al volume osservato; valori minori "
            "indicano errori aggregati inferiori. Non è calcolabile se il denominatore è nullo. "
            "Essendo ponderato dai volumi, può nascondere errori elevati nei territori piccoli."
        )
        st.latex(r"\mathrm{Bias}=100\times\frac{\sum_i(\widehat Y_i-Y_i)}{\sum_iY_i}")
        st.caption("Con questa convenzione: bias negativo = sottostima; positivo = sovrastima. WAPE e bias vanno letti insieme.")
        st.write(
            "Una validazione convincente usa solo informazioni disponibili alla data di previsione, "
            "separa cronologicamente addestramento e test, e confronta i modelli sullo stesso campione "
            "con una baseline stagionale semplice. Quote, tassi e altri parametri del test "
            "non devono essere calcolati usando dati futuri. Sono utili verifiche per regione, "
            "stagione e orizzonte e, successivamente, più origini temporali di previsione."
        )
        st.subheader("Risultati storici riportati nel progetto")
        tabella(pd.DataFrame([
            ["Stagionale 2024", "11,11%", "Test storico; periodo esatto da verificare nello script"],
            ["Media 2023–2024", "17,13%", "Test storico; periodo esatto da verificare nello script"],
            ["Prudente", "7,48%", "Test 2025 dichiarato"],
            ["Intermedio", "8,39%", "Test 2025 dichiarato"],
            ["Espansivo", "9,89%", "Test 2025 dichiarato"]
        ], columns=["Metodo", "WAPE riportato", "Campione dichiarato"]))
        st.warning(
            "Valori documentati, non ricalcolati né verificati da questa app. "
            "Senza allineare campioni, parametri e date di addestramento non costituiscono "
            "una classifica omogenea. Il backtest non dimostra da solo affidabilità fino al 2030."
        )

    with limiti_tab:
        st.subheader("Che cosa possiamo concludere oggi")
        st.write(
            "Possiamo descrivere la domanda, confrontare periodi coerenti, individuare territori "
            "e mesi interessati dagli eventi e valutare scenari. Non possiamo attribuire "
            "automaticamente le variazioni allo sport, stimare prezzi Airbnb o trasferire "
            "l’effetto regionale a una città, una provincia o un singolo alloggio."
        )
        st.subheader("Perché un picco non basta")
        st.markdown(
            "- **Aggregazione:** un effetto locale o di pochi giorni può essere diluito nel totale regionale mensile.\n"
            "- **Fattori concomitanti:** festività, meteo, trasporti, offerta, prezzi e altri eventi possono cambiare insieme.\n"
            "- **Spostamenti:** pernottamenti possono trasferirsi da un mese o territorio all’altro senza aumento netto nazionale.\n"
            "- **Sostituzione:** i visitatori sportivi possono rimpiazzare altri ospiti; affluenza lorda e domanda aggiuntiva differiscono.\n"
            "- **Selezione:** gli eventi non sono assegnati casualmente ai territori e il catalogo può avere copertura diseguale."
        )
        st.subheader("Passaggio successivo da concordare: un controfattuale credibile")
        st.write(
            "Per stimare l’effetto incrementale occorre definire l’esposizione, selezionare confronti "
            "adeguati e stimare quale domanda ci si sarebbe aspettati senza l’evento. "
            "Un modello panel, un’event study o una difference-in-differences possono essere valutati "
            "soltanto se dati e ipotesi lo consentono; non sono procedure già eseguite da questa dashboard."
        )
        st.write(
            "Prima di interpretarli causalmente servono controlli sui trend precedenti, "
            "sull’assenza di anticipazioni e contaminazioni nei territori di confronto, "
            "sugli eventi concomitanti, sull’incertezza delle stime e sulla robustezza "
            "alla scelta del periodo. Per eventi ricorrenti o sovrapposti il confronto richiede "
            "particolare attenzione. Anche risultati piccoli o non distinguibili da zero sono informativi."
        )
        st.success(
            "Principio di lavoro: prima rendere riproducibili dati, ipotesi e validazione; "
            "poi aggiungere complessità solo se migliora la risposta alla domanda di ricerca."
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
