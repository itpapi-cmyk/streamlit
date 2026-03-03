import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from io import BytesIO
from datetime import datetime
from docx import Document


def euro(val):
    return f"EUR {val:,.0f}".replace(",", ".")


def format_number_it(val, decimals=2):
    if pd.isna(val):
        return ""
    s = f"{val:,.{decimals}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def parse_number_it(val):
    if pd.isna(val):
        return np.nan
    s = str(val).strip()
    if not s:
        return np.nan
    s = s.replace("\u00a0", " ").replace("EUR", "").replace("eur", "").replace(" ", "").replace("'", "")
    s = s.replace("€", "")
    neg_parentheses = s.startswith("(") and s.endswith(")")
    if neg_parentheses:
        s = s[1:-1].strip()

    has_comma = "," in s
    has_dot = "." in s
    if has_comma and has_dot:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        s = s.replace(".", "").replace(",", ".")
    elif has_dot:
        parts = s.split(".")
        if len(parts) > 2 and all(p.isdigit() for p in parts):
            s = "".join(parts)

    if not re.match(r"^-?\d+(\.\d+)?$", s):
        return np.nan
    n = float(s)
    return -n if neg_parentheses else n


st.set_page_config(page_title="Audit Sampling - Key Items & Items", layout="wide")
st.title("Selezione campione - Key Items e Items")

with st.sidebar.expander("Metodologia e funzionamento del programma"):
    st.markdown(
        """
**DESCRIZIONE DEL PROGRAMMA - AUDIT SAMPLING**

**1. Finalita del programma**
Il programma Audit Sampling - Key Items & Items e uno strumento di supporto alle attivita di revisione contabile, sviluppato per la selezione strutturata dei campioni secondo i principi degli ISA (International Standards on Auditing).

Consente di:
- identificare automaticamente i Key Items
- selezionare un campione di elementi residui
- applicare criteri coerenti con la materialita e il livello di rischio
- garantire riproducibilita e tracciabilita delle scelte

**2. Riferimenti normativi (ISA)**
Il programma e coerente con:
- ISA 300 - Pianificazione della revisione
- ISA 315 - Identificazione e valutazione dei rischi
- ISA 330 - Risposte ai rischi identificati
- ISA 500 - Elementi probativi
- ISA 530 - Campionamento di revisione

**3. Logica generale di funzionamento**
Il processo di selezione si articola in tre fasi:
- Individuazione dei Key Items: elementi di maggiore rilevanza monetaria rispetto al totale.
- Determinazione del residuo: popolazione residua al netto dei Key Items.
- Selezione del campione: mediante uno dei metodi disponibili (MUS, Intervallo, Casuale).

**4. Metodo MUS (Monetary Unit Sampling)**
Il metodo MUS:
- assegna maggiore probabilita agli importi piu elevati
- e coerente con test di esistenza e accuratezza
- e particolarmente adatto per crediti, ricavi, immobilizzazioni
- utilizza un seed fisso (42) per garantire riproducibilita

Il programma calcola automaticamente:
- intervallo di campionamento
- starting point
- numero di elementi da testare

**5. Parametri di selezione**
L'utente definisce:
- Soglia Key Items (%)
- Materialita (EUR)
- Confidence level (%)
- Metodo di selezione

Tutti i parametri sono visibili e tracciabili.

**6. Output del programma**
Il software produce:
- Riepilogo generale
- Elenco Key Items
- Campione selezionato
- Export Excel
- Export Word

Gli importi sono:
- espressi in euro
- senza decimali
- coerenti con la documentazione di revisione

**7. Riproducibilita e controlli**
- Seed fisso per evitare variazioni casuali
- Nessuna modifica ai dati originali
- Risultati coerenti a parita di input
- Processo pienamente tracciabile

**8. Finalita di revisione**
Il programma supporta:
- la pianificazione delle verifiche
- la definizione del campione
- la documentazione delle scelte
- la difendibilita del lavoro svolto
- l'allineamento agli ISA

**9. Valutazione degli errori e proiezione sull'universo**

Nel caso in cui, a seguito delle verifiche effettuate sugli elementi selezionati, vengano riscontrate differenze tra saldo contabile e saldo verificato, il programma consente la determinazione automatica della stima dell'errore sull'intera popolazione oggetto di campionamento.

**9.1 Determinazione del tainting**
Per ciascun elemento con errore viene calcolato il tainting, definito come:
- Errore / Valore contabile dell'elemento

Il tainting rappresenta la percentuale di errore rispetto al valore originario dell'item selezionato.

Gli elementi privi di differenze assumono tainting pari a zero.

**9.2 Most Likely Error (MLE)**
Nel caso di utilizzo del metodo MUS, l'errore proiettato sull'universo (Most Likely Error - MLE) e determinato secondo la seguente logica:
- Intervallo di campionamento = Universo residuo / Numero di selezioni
- MLE = Intervallo x Somma dei tainting riscontrati

In forma equivalente:
- MLE = (Somma dei tainting / Numero di selezioni) x Universo residuo

Il MLE rappresenta la stima puntuale dell'errore presente nella popolazione oggetto di campionamento.

**9.3 Valore calcolato secondo il metodo selezionato**
Il programma puo inoltre calcolare un valore ulteriore derivante dal metodo di campionamento prescelto, determinato sulla base dei parametri impostati dall'utente (universo, numerosita del campione, livello di confidenza).

Tale valore e riportato a fini informativi e documentali, in coerenza con il metodo applicato.

**9.4 Interpretazione dei risultati**
Il Most Likely Error rappresenta la stima centrale dell'errore.

Il confronto tra i valori stimati e la materialita definita in fase di pianificazione consente al revisore di:
- valutare la significativita delle differenze riscontrate
- determinare l'eventuale necessita di estendere le verifiche
- supportare il giudizio professionale finale

**9.5 Responsabilita professionale**
La stima dell'errore sull'universo costituisce uno strumento di supporto all'analisi, ma non sostituisce:
- la valutazione qualitativa delle differenze
- l'analisi delle cause degli errori
- il giudizio professionale del revisore

Eventuali decisioni in merito all'estensione delle procedure o alla richiesta di rettifiche restano di esclusiva competenza del professionista.

**10. Avvertenze e limiti di utilizzo**
Il presente strumento:
- non sostituisce il giudizio professionale del revisore
- supporta, ma non determina autonomamente, le conclusioni di revisione
- deve essere utilizzato nel contesto di una corretta valutazione del rischio
- presuppone la correttezza dei dati caricati dall'utente

L'utente resta responsabile:
- della coerenza dei parametri inseriti
- della valutazione finale dei risultati
- dell'adeguatezza del campione selezionato

Il software rappresenta uno strumento di supporto operativo, conforme ai principi di revisione, ma non sostitutivo dell'attivita professionale.
"""
    )

uploaded_file = st.file_uploader("Carica file Excel o CSV", type=["xlsx", "csv"])
if uploaded_file is None:
    st.stop()

ext = os.path.splitext(uploaded_file.name)[1].lower()
if ext == ".xlsx":
    df = pd.read_excel(uploaded_file)
elif ext == ".csv":
    found = False
    for sep in [",", ";", "\t"]:
        uploaded_file.seek(0)
        try:
            df = pd.read_csv(uploaded_file, sep=sep)
            if len(df.columns) > 1:
                found = True
                break
        except Exception:
            continue
    if not found:
        st.error("CSV non valido. Usa separatore virgola, punto e virgola o tab.")
        st.stop()
else:
    st.error("Formato file non supportato.")
    st.stop()

if df.empty:
    st.error("File caricato vuoto.")
    st.stop()

st.sidebar.header("Informazioni revisione")
societa = st.sidebar.text_input("Societa")
revisione_al_str = st.sidebar.text_input("Revisione al (GG/MM/AA)", value=datetime.now().strftime("%d/%m/%y"))
try:
    datetime.strptime(revisione_al_str, "%d/%m/%y")
except ValueError:
    st.sidebar.warning("Formato data non valido. Uso data odierna.")
    revisione_al_str = datetime.now().strftime("%d/%m/%y")
preparato_da = st.sidebar.text_input("Preparato da")
data_ora = datetime.now().strftime("%d/%m/%Y %H:%M")

st.sidebar.header("Mappatura colonne")
colonne = df.columns.tolist()
if len(colonne) < 3:
    st.error("Il file deve contenere almeno 3 colonne.")
    st.stop()
codice_col = st.sidebar.selectbox("Colonna Codice", colonne, index=0)
descr_col = st.sidebar.selectbox("Colonna Descrizione", colonne, index=1)
valore_col = st.sidebar.selectbox("Colonna Valore", colonne, index=2)
if len({codice_col, descr_col, valore_col}) < 3:
    st.warning("Le colonne devono essere diverse.")
    st.stop()

st.sidebar.header("Parametri di campionamento")
soglia_tipo = st.sidebar.radio("Tipo soglia Key Items", ["Percentuale sul totale", "Soglia numerica"])
if soglia_tipo == "Percentuale sul totale":
    perc_key = st.sidebar.number_input("Soglia Key Items (%)", 0, 100, 30)
    soglia_key_num = None
else:
    soglia_key_num = st.sidebar.number_input("Soglia Key Items (EUR)", min_value=0.0, value=10000.0, step=100.0, format="%.2f")
    perc_key = None

materialita = st.sidebar.number_input("Materialita (EUR)", min_value=1.0, value=1_000_000.0)
confidence_level = st.sidebar.number_input("Confidence Level (%)", 1, 100, 80)

st.sidebar.header("Metodo selezione Items")
metodo = st.sidebar.radio("", ["MUS", "Intervallo", "Casuale"])

manual_starting_point = None
starting_point_mode = None
if metodo in ["MUS", "Intervallo"]:
    starting_point_mode = st.sidebar.radio("Selezione Starting Point", ["Automatica", "Manuale"])
    if starting_point_mode == "Manuale":
        manual_starting_point = st.sidebar.number_input(
            "Inserisci valore Starting Point", min_value=0.0, value=0.0, step=1.0, format="%.2f"
        )

df[valore_col] = df[valore_col].apply(parse_number_it)
valori_non_validi = int(df[valore_col].isna().sum())
df = df.dropna(subset=[valore_col]).reset_index(drop=True)
if valori_non_validi > 0:
    st.warning(f"Scartate {valori_non_validi} righe con valore non numerico.")

df_base = df.copy()
df_mus = df.sort_values(by=valore_col, ascending=False).reset_index(drop=True)

tot_items = len(df_base)
tot_valore = float(df_base[valore_col].sum())
top5_val = float(df_base.sort_values(by=valore_col, ascending=False).head(5)[valore_col].sum())
top5_perc = top5_val / tot_valore * 100 if tot_valore != 0 and tot_items >= 5 else 100.0

st.subheader("Universo completo")
st.dataframe(
    df_base.style.format({valore_col: lambda x: format_number_it(x, 2)}),
    width="stretch",
)
c1, c2, c3 = st.columns(3)
c1.metric("Totale items", tot_items)
c2.metric("Valore totale", format_number_it(tot_valore, 0))
c3.metric("Top 5 (% valore)", f"{top5_perc:.2f}%")

calcola_campione = st.button("Calcola selezione campione")
if calcola_campione or ("key_items" in st.session_state and "items_selezionati" in st.session_state):
    if calcola_campione:
        if "errori_editor" in st.session_state:
            del st.session_state["errori_editor"]

        df_sorted = df_mus.copy()
        if soglia_tipo == "Percentuale sul totale":
            soglia_key = tot_valore * perc_key / 100
            df_sorted["cumulativo"] = df_sorted[valore_col].cumsum()
            key_items = df_sorted[df_sorted["cumulativo"] <= soglia_key].copy()
            if key_items.empty:
                key_items = df_sorted.head(1)
        else:
            key_items = df_sorted[df_sorted[valore_col] > soglia_key_num].copy()
            if key_items.empty:
                key_items = df_sorted.head(1)

        residuo = df_sorted.drop(key_items.index).copy()
        residuo_tot = float(residuo[valore_col].sum())
        confidence_factor = 100 * (1 - ((100 - confidence_level) / 100) ** (1 / 100))
        num_items = int(round((residuo_tot / materialita) * confidence_factor))
        num_items = max(num_items, 1)

        selected_items = pd.DataFrame(columns=df_sorted.columns)
        starting_point = None

        if metodo == "MUS" and residuo_tot > 0:
            np.random.seed(42)
            residuo = df_mus.drop(key_items.index)
            intervallo = residuo_tot / num_items
            if starting_point_mode == "Manuale" and manual_starting_point is not None:
                starting_point = manual_starting_point
            else:
                starting_point = np.random.uniform(0, intervallo)

            residuo["cumulativo"] = residuo[valore_col].cumsum()
            soglie = [starting_point + i * intervallo for i in range(num_items)]
            idx = set()
            for s in soglie:
                r = residuo[residuo["cumulativo"] >= s].head(1)
                if not r.empty:
                    idx.add(r.index[0])
            selected_items = residuo.loc[list(idx)]
        elif metodo == "Intervallo":
            residuo = df_base.drop(key_items.index)
            step = max(1, len(residuo) // num_items)
            if starting_point_mode == "Manuale" and manual_starting_point is not None:
                start_idx = int(round(manual_starting_point))
                selected_items = residuo.iloc[start_idx::step].head(num_items)
                starting_point = start_idx
            else:
                np.random.seed(42)
                start_idx = np.random.randint(0, step)
                selected_items = residuo.iloc[start_idx::step].head(num_items)
                starting_point = start_idx
        else:
            residuo = df_base.drop(key_items.index)
            np.random.seed(42)
            selected_items = residuo.sample(n=min(num_items, len(residuo)), random_state=42)

        intervallo_utilizzato = residuo_tot / len(selected_items) if len(selected_items) > 0 else 0.0
        st.session_state["key_items"] = key_items
        st.session_state["items_selezionati"] = selected_items
        st.session_state["starting_point"] = starting_point
        st.session_state["intervallo_utilizzato"] = intervallo_utilizzato
    else:
        key_items = st.session_state["key_items"]
        selected_items = st.session_state["items_selezionati"]
        starting_point = st.session_state.get("starting_point")
        intervallo_utilizzato = st.session_state.get("intervallo_utilizzato", 0.0)

    riepilogo = pd.DataFrame(
        {
            "Categoria": ["Universo", "Key Items", "Items selezionati"],
            "Numero items": [tot_items, len(key_items), len(selected_items)],
            "Valore (EUR)": [euro(tot_valore), euro(key_items[valore_col].sum()), euro(selected_items[valore_col].sum())],
            "% su totale": [
                "100.00%",
                f"{key_items[valore_col].sum()/tot_valore*100:.2f}%" if tot_valore else "0.00%",
                f"{selected_items[valore_col].sum()/tot_valore*100:.2f}%" if tot_valore else "0.00%",
            ],
            "Starting point": ["-", "-", f"{starting_point:.2f}" if starting_point is not None else "-"],
        }
    )
    st.subheader("Riepilogo selezione")
    st.table(riepilogo)

    st.subheader("Key Items")
    st.dataframe(
        key_items[[codice_col, descr_col, valore_col]].style.format({valore_col: lambda x: format_number_it(x, 2)}),
        width="stretch",
    )
    st.subheader("Items selezionati")
    st.dataframe(
        selected_items[[codice_col, descr_col, valore_col]].style.format({valore_col: lambda x: format_number_it(x, 2)}),
        width="stretch",
    )

    st.subheader("Errori emersi e proiezione sull'universo")
    saldo_col = "Saldo corretto emerso (EUR)"
    errore_col = "Errore (EUR)"
    errore_perc_col = "Errore (%)"

    errori_df = selected_items[[codice_col, descr_col, valore_col]].copy()
    errori_df[saldo_col] = errori_df[valore_col]
    errori_editati = st.data_editor(
        errori_df,
        key="errori_editor",
        hide_index=True,
        use_container_width=True,
        column_config={
            valore_col: st.column_config.NumberColumn(label=valore_col, format="%.2f", disabled=True),
            saldo_col: st.column_config.NumberColumn(label=saldo_col, format="%.2f"),
        },
    )

    errori_editati[saldo_col] = pd.to_numeric(errori_editati[saldo_col], errors="coerce").fillna(0.0)
    valori_originali = pd.to_numeric(errori_editati[valore_col], errors="coerce").fillna(0.0)
    saldi_corretti = pd.to_numeric(errori_editati[saldo_col], errors="coerce").fillna(0.0)
    errore_importo = saldi_corretti - valori_originali
    errore_perc = np.where(valori_originali != 0, (errore_importo / valori_originali) * 100, 0.0)
    somma_perc_errori = float(np.nansum(errore_perc))
    mle = (somma_perc_errori / 100) * intervallo_utilizzato

    dettaglio_errori = errori_editati[[codice_col, descr_col, valore_col, saldo_col]].copy()
    dettaglio_errori[errore_col] = errore_importo
    dettaglio_errori[errore_perc_col] = errore_perc
    dettaglio_errori_con_errori = dettaglio_errori[np.abs(dettaglio_errori[errore_col]) > 1e-12]

    if not dettaglio_errori_con_errori.empty:
        st.write("Dettaglio items con errore")
        st.dataframe(
            dettaglio_errori_con_errori.style.format(
                {
                    valore_col: lambda x: format_number_it(x, 2),
                    saldo_col: lambda x: format_number_it(x, 2),
                    errore_col: lambda x: format_number_it(x, 2),
                    errore_perc_col: lambda x: format_number_it(x, 4),
                }
            ),
            width="stretch",
        )
    else:
        st.write("Nessun errore rilevato negli items selezionati.")

    universo_no_key_items = max(0.0, float(tot_valore - key_items[valore_col].sum()))
    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("Universo (no key items)", format_number_it(universo_no_key_items, 0))
    m2.metric("Items selezionati", len(errori_editati))
    m3.metric("Intervallo utilizzato", format_number_it(intervallo_utilizzato, 0))
    m4.metric("Somma % errori", f"{somma_perc_errori:.4f}%")
    m5.metric("Most likely error (MLE)", format_number_it(mle, 0))
    st.write("nella determinazione del MLE si sono considerati sia gli errori negativi che gli errori positivi")

    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        info_df = pd.DataFrame(
            {
                "Parametro": ["Societa", "Revisione al", "Preparato da", "Materialita", "Confidence level"],
                "Valore": [societa, revisione_al_str, preparato_da, euro(materialita), f"{confidence_level}%"],
            }
        )
        info_df.to_excel(writer, sheet_name="Riepilogo", index=False, startrow=0)
        riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False, startrow=info_df.shape[0] + 2)
        key_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Key Items", index=False)
        selected_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Items selezionati", index=False)

        mle_info_df = pd.DataFrame(
            {
                "Parametro": [
                    "Universo (no key items)",
                    "Items selezionati",
                    "Intervallo utilizzato",
                    "Somma % errori",
                    "Most likely error (MLE)",
                    "Nota",
                ],
                "Valore": [
                    format_number_it(universo_no_key_items, 0),
                    len(errori_editati),
                    intervallo_utilizzato,
                    somma_perc_errori,
                    mle,
                    "nella determinazione del MLE si sono considerati sia gli errori negativi che gli errori positivi",
                ],
            }
        )
        mle_info_df.to_excel(writer, sheet_name="Errori e MLE", index=False, startrow=0)
        dettaglio_errori_con_errori.to_excel(writer, sheet_name="Errori e MLE", index=False, startrow=mle_info_df.shape[0] + 2)
    excel_buffer.seek(0)

    st.download_button(
        "Export Excel",
        data=excel_buffer,
        file_name="audit_sampling.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    def export_word():
        doc = Document()
        doc.add_heading("MEMO DI REVISIONE - SELEZIONE CAMPIONE", 0)
        doc.add_paragraph(f"Societa: {societa}")
        doc.add_paragraph(f"Revisione al: {revisione_al_str}")
        doc.add_paragraph(f"File: {uploaded_file.name}")
        doc.add_paragraph(f"Data: {data_ora}")
        doc.add_paragraph(f"Preparato da: {preparato_da}")
        doc.add_paragraph(f"Metodo: {metodo}")
        if starting_point is not None:
            doc.add_paragraph(f"Starting point: {starting_point:.2f}")
        doc.add_paragraph(f"Materialita: {euro(materialita)}  |  Confidence level: {confidence_level}%")
        doc.add_paragraph("\nRiepilogo selezione")
        t = doc.add_table(rows=riepilogo.shape[0] + 1, cols=riepilogo.shape[1])
        for j, col in enumerate(riepilogo.columns):
            t.cell(0, j).text = str(col)
        for i in range(riepilogo.shape[0]):
            for j in range(riepilogo.shape[1]):
                t.cell(i + 1, j).text = str(riepilogo.iloc[i, j])

        doc.add_paragraph("\nMost likely error (MLE)")
        doc.add_paragraph(f"Universo (no key items): {format_number_it(universo_no_key_items, 0)}")
        doc.add_paragraph(f"Items selezionati: {len(errori_editati)}")
        doc.add_paragraph(f"Intervallo utilizzato: {format_number_it(intervallo_utilizzato, 0)}")
        doc.add_paragraph(f"Somma % errori: {somma_perc_errori:.4f}%")
        doc.add_paragraph(f"Most likely error (MLE): {format_number_it(mle, 0)}")
        doc.add_paragraph("nella determinazione del MLE si sono considerati sia gli errori negativi che gli errori positivi")

        doc.add_paragraph("\nDettaglio items con errore")
        headers = [codice_col, descr_col, valore_col, saldo_col, errore_col, errore_perc_col]
        te = doc.add_table(rows=len(dettaglio_errori_con_errori) + 1, cols=len(headers))
        for j, h in enumerate(headers):
            te.cell(0, j).text = h
        for i, row in enumerate(dettaglio_errori_con_errori[headers].itertuples(index=False)):
            for j, v in enumerate(row):
                col_name = headers[j]
                if col_name in [errore_col, errore_perc_col]:
                    num_v = pd.to_numeric(v, errors="coerce")
                    te.cell(i + 1, j).text = "" if pd.isna(num_v) else f"{num_v:.2f}"
                else:
                    te.cell(i + 1, j).text = str(v)

        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export Word",
        data=export_word(),
        file_name="audit_sampling.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
