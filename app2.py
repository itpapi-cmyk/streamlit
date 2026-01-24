import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from docx import Document




# =====================
# Funzione formattazione euro
# =====================
def euro(val):
    return f"€ {val:,.0f}".replace(",", ".")

# =====================
# Configurazione pagina
# =====================
st.set_page_config(
    page_title="Audit Sampling – Key Items & Items",
    layout="wide"
)

st.title("Selezione campione – Key Items e Items")

with st.sidebar.expander("📘 Metodologia e funzionamento del programma"):
    st.markdown("""
### DESCRIZIONE DEL PROGRAMMA – AUDIT SAMPLING

---

## 1. Finalità del programma
Il programma **Audit Sampling – Key Items & Items** è uno strumento di supporto alle attività di revisione contabile,
sviluppato per la selezione strutturata dei campioni secondo i principi degli **ISA (International Standards on Auditing)**.

Consente di:
- identificare automaticamente i **Key Items**
- selezionare un campione di elementi residui
- applicare criteri coerenti con la **materialità** e il **livello di rischio**
- garantire **riproducibilità e tracciabilità delle scelte**

---

## 2. Riferimenti normativi (ISA)

Il programma è coerente con:
- **ISA 300** – Pianificazione della revisione
- **ISA 315** – Identificazione e valutazione dei rischi
- **ISA 330** – Risposte ai rischi identificati
- **ISA 500** – Elementi probativi
- **ISA 530** – Campionamento di revisione

---

## 3. Logica generale di funzionamento

Il processo di selezione si articola in tre fasi:

1. **Individuazione dei Key Items**  
   Elementi di maggiore rilevanza monetaria rispetto al totale.

2. **Determinazione del residuo**  
   Popolazione residua al netto dei Key Items.

3. **Selezione del campione**  
   Mediante uno dei metodi disponibili (MUS, Intervallo, Casuale).

---

## 4. Metodo MUS (Monetary Unit Sampling)

Il metodo MUS:
- assegna maggiore probabilità agli importi più elevati
- è coerente con test di esistenza e accuratezza
- è particolarmente adatto per crediti, ricavi, immobilizzazioni
- utilizza un **seed fisso (42)** per garantire riproducibilità

Il programma calcola automaticamente:
- intervallo di campionamento
- starting point
- numero di elementi da testare

---

## 5. Parametri di selezione

L’utente definisce:
- **Soglia Key Items (%)**
- **Materialità (€)**
- **Confidence level (%)**
- **Metodo di selezione**

Tutti i parametri sono visibili e tracciabili.

---

## 6. Output del programma

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

---

## 7. Riproducibilità e controlli

- Seed fisso per evitare variazioni casuali
- Nessuna modifica ai dati originali
- Risultati coerenti a parità di input
- Processo pienamente tracciabile

---

## 8. Finalità di revisione

Il programma supporta:
- la pianificazione delle verifiche
- la definizione del campione
- la documentazione delle scelte
- la difendibilità del lavoro svolto
- l’allineamento agli ISA

---

## 9. Avvertenze e limiti di utilizzo

Il presente strumento:
- **non sostituisce il giudizio professionale del revisore**
- supporta, ma non determina autonomamente, le conclusioni di revisione
- deve essere utilizzato nel contesto di una corretta valutazione del rischio
- presuppone la correttezza dei dati caricati dall’utente

L’utente resta responsabile:
- della coerenza dei parametri inseriti
- della valutazione finale dei risultati
- dell’adeguatezza del campione selezionato

Il software rappresenta uno **strumento di supporto operativo**, conforme ai principi di revisione, ma non sostitutivo dell’attività professionale.
""")

# =====================
# Upload file
# =====================
uploaded_file = st.file_uploader("Carica file Excel", type=["xlsx"])
if uploaded_file is None:
    st.stop()

df = pd.read_excel(uploaded_file)
if df.empty:
    st.error("File Excel vuoto.")
    st.stop()

# =====================
# Sidebar – info revisione
# =====================
st.sidebar.header("Informazioni revisione")
preparato_da = st.sidebar.text_input("Preparato da")
data_ora = datetime.now().strftime("%d/%m/%Y %H:%M")

# =====================
# Sidebar – mappatura colonne
# =====================
st.sidebar.header("Mappatura colonne")
colonne = df.columns.tolist()

codice_col = st.sidebar.selectbox("Colonna Codice", colonne, index=0)
descr_col = st.sidebar.selectbox("Colonna Descrizione", colonne, index=1)
valore_col = st.sidebar.selectbox("Colonna Valore", colonne, index=2)

if len({codice_col, descr_col, valore_col}) < 3:
    st.warning("Le colonne devono essere diverse.")
    st.stop()

# =====================
# Sidebar – parametri
# =====================

st.sidebar.header("Parametri di campionamento")

# Opzione soglia Key Items: percentuale o numerica
soglia_tipo = st.sidebar.radio(
    "Tipo soglia Key Items",
    ["Percentuale sul totale", "Soglia numerica"],
    key="key_soglia_tipo"
)
if soglia_tipo == "Percentuale sul totale":
    perc_key = st.sidebar.number_input("Soglia Key Items (%)", 0, 100, 30)
    soglia_key_num = None
else:
    soglia_key_num = st.sidebar.number_input("Soglia Key Items (€)", min_value=0.0, value=10000.0, step=100.0, format="%.2f")
    perc_key = None
materialita = st.sidebar.number_input("Materialità (€)", min_value=1.0, value=1_000_000.0)
confidence_level = st.sidebar.number_input("Confidence Level (%)", 1, 100, 80)




# Sidebar – metodo selezione
metodo = st.sidebar.radio(
    "Metodo selezione Items",
    ["MUS", "Intervallo", "Casuale"],
    key="key_metodo"
)

# Opzione per starting point automatico o manuale (solo MUS e Intervallo)
st.sidebar.header("Opzioni Starting Point")
manual_starting_point = None
starting_point_mode = None
if metodo in ["MUS", "Intervallo"]:
    starting_point_mode = st.sidebar.radio(
        "Selezione Starting Point",
        ["Automatica", "Manuale"],
        index=0,
        key="key_starting_point_mode"
    )
    if starting_point_mode == "Manuale":
        manual_starting_point = st.sidebar.number_input(
            "Inserisci valore Starting Point",
            min_value=0.0,
            value=0.0,
            step=1.0,
            format="%.2f"
        )

metodo = st.sidebar.radio(
    "Metodo selezione Items",
    ["MUS", "Intervallo", "Casuale"]
)

# =====================
# Normalizzazione dati
# =====================
df[valore_col] = pd.to_numeric(df[valore_col], errors="coerce")
df = df.dropna(subset=[valore_col])

# Dataset base (ordine originale del file)
df_base = df.copy().reset_index(drop=True)

# Dataset ordinato SOLO per MUS
df_mus = df.sort_values(by=valore_col, ascending=False).reset_index(drop=True)


# =====================
# UNIVERSO
# =====================

tot_items = len(df_base)
tot_valore = df_base[valore_col].sum()
top5_val = df_base.sort_values(by=valore_col, ascending=False).head(5)[valore_col].sum()

top5_perc = top5_val / tot_valore * 100 if tot_items >= 5 else 100

st.subheader("Universo completo")
st.dataframe(df_base, width="stretch")

c1, c2, c3 = st.columns(3)
c1.metric("Totale items", tot_items)
c2.metric("Valore totale", euro(tot_valore))
c3.metric("Top 5 (% valore)", f"{top5_perc:.2f}%")

# =====================
# CALCOLO CAMPIONE
# =====================
if st.button("Calcola selezione campione"):
    # Define df_sorted based on the sorting logic used for MUS
    df_sorted = df_mus.copy()


    # -------- KEY ITEMS --------
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
    residuo_tot = residuo[valore_col].sum()

    # -------- CONFIDENCE FACTOR --------
    confidence_factor = 100 * (1 - ((100 - confidence_level) / 100) ** (1 / 100))

    # -------- NUMERO ITEMS --------
    num_items = int(round((residuo_tot / materialita) * confidence_factor))
    num_items = max(num_items, 1)

    # -------- SELEZIONE --------
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
            # Se manuale, starting_point è l'indice di partenza (arrotondato)
            start_idx = int(round(manual_starting_point))
            selected_items = residuo.iloc[start_idx::step].head(num_items)
            starting_point = start_idx
        else:
            selected_items = residuo.iloc[::step].head(num_items)
            starting_point = 1

    elif metodo == "Casuale":
        residuo = df_base.drop(key_items.index)
        np.random.seed(42)
        selected_items = residuo.sample(
            n=min(num_items, len(residuo)),
            random_state=42
        )

    # =====================
    # Salvataggio session state
    # =====================
    st.session_state["key_items"] = key_items
    st.session_state["items_selezionati"] = selected_items
    st.session_state["starting_point"] = starting_point

    # =====================
    # RIEPILOGO
    # =====================
    st.subheader("Riepilogo selezione")
    riepilogo = pd.DataFrame({
        "Categoria": ["Universo", "Key Items", "Items selezionati"],
        "Numero items": [tot_items, len(key_items), len(selected_items)],
        "Valore (€)": [
            euro(tot_valore),
            euro(key_items[valore_col].sum()),
            euro(selected_items[valore_col].sum())
        ],
        "% su totale": [
            "100.00%",
            f"{key_items[valore_col].sum()/tot_valore*100:.2f}%",
            f"{selected_items[valore_col].sum()/tot_valore*100:.2f}%"
        ],
        "Starting point": ["-", "-", f"{starting_point:.2f}" if starting_point is not None else "-"]
    })
    st.table(riepilogo)

    st.subheader("Key Items")
    st.dataframe(key_items[[codice_col, descr_col, valore_col]], width="stretch")

    st.subheader("Items selezionati")
    st.dataframe(selected_items[[codice_col, descr_col, valore_col]], width="stretch")

    # =====================
    # EXPORT EXCEL
    # =====================
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
        # valori numerici senza decimali
        key_items_export = key_items[[codice_col, descr_col, valore_col]].copy()
        key_items_export[valore_col] = key_items_export[valore_col].round(0).astype(int)
        key_items_export.to_excel(writer, sheet_name="Key Items", index=False)

        selected_items_export = selected_items[[codice_col, descr_col, valore_col]].copy()
        selected_items_export[valore_col] = selected_items_export[valore_col].round(0).astype(int)
        selected_items_export.to_excel(writer, sheet_name="Items selezionati", index=False)
    excel_buffer.seek(0)

    st.download_button(
        "Export Excel",
        data=excel_buffer,
        file_name="audit_sampling.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =====================
    # EXPORT WORD
    # =====================
    def export_word():
        doc = Document()
        doc.add_heading("MEMO DI REVISIONE – SELEZIONE CAMPIONE", 0)
        doc.add_paragraph(f"File: {uploaded_file.name}")
        doc.add_paragraph(f"Data: {data_ora}")
        doc.add_paragraph(f"Preparato da: {preparato_da}")
        doc.add_paragraph(f"Metodo: {metodo}")
        if starting_point is not None:
            doc.add_paragraph(f"Starting point: {starting_point:.2f}")
        doc.add_paragraph("\nRiepilogo selezione")
        table = doc.add_table(rows=riepilogo.shape[0]+1, cols=riepilogo.shape[1])
        for j, col in enumerate(riepilogo.columns):
            table.cell(0, j).text = col
        for i in range(riepilogo.shape[0]):
            for j in range(riepilogo.shape[1]):
                table.cell(i+1, j).text = str(riepilogo.iloc[i, j])
        doc.add_paragraph("\nKey Items")
        table_key = doc.add_table(rows=len(key_items_export)+1, cols=3)
        table_key.cell(0,0).text = codice_col
        table_key.cell(0,1).text = descr_col
        table_key.cell(0,2).text = "Valore (€)"
        for i, row in enumerate(key_items_export.itertuples(index=False)):
            table_key.cell(i+1,0).text = str(row[0])
            table_key.cell(i+1,1).text = str(row[1])
            table_key.cell(i+1,2).text = str(row[2])
        doc.add_paragraph("\nItems selezionati")
        table_sel = doc.add_table(rows=len(selected_items_export)+1, cols=3)
        table_sel.cell(0,0).text = codice_col
        table_sel.cell(0,1).text = descr_col
        table_sel.cell(0,2).text = "Valore (€)"
        for i, row in enumerate(selected_items_export.itertuples(index=False)):
            table_sel.cell(i+1,0).text = str(row[0])
            table_sel.cell(i+1,1).text = str(row[1])
            table_sel.cell(i+1,2).text = str(row[2])
        buffer = BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export Word",
        data=export_word(),
        file_name="audit_sampling.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )