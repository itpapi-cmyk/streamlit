import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from datetime import datetime
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT




# =====================
# Funzione formattazione euro
# =====================
def euro(val):
    return f"€ {val:,.0f}".replace(",", ".")

# Funzione formattazione numeri in stile italiano (1.234.567,89)
def format_number_it(val, decimals=2):
    if pd.isna(val):
        return ""
    formatted = f"{val:,.{decimals}f}"
    return formatted.replace(",", "X").replace(".", ",").replace("X", ".")

def parse_number_it(val):
    if pd.isna(val):
        return np.nan

    s = str(val).strip()
    if not s:
        return np.nan

    s = s.replace("\u00a0", " ").replace("€", "").replace("EUR", "").replace("eur", "")

    neg_parentheses = s.startswith("(") and s.endswith(")")
    if neg_parentheses:
        s = s[1:-1].strip()

    s = s.replace(" ", "").replace("'", "")

    # Normalizzazione robusta:
    # - supporta formato italiano (1.234,56)
    # - supporta anche decimali con punto (1234.56)
    has_comma = "," in s
    has_dot = "." in s

    if has_comma and has_dot:
        # Il separatore più a destra è considerato decimale.
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            decimal_sep = ","
            thousands_sep = "."
        else:
            decimal_sep = "."
            thousands_sep = ","
        normalized = s.replace(thousands_sep, "").replace(decimal_sep, ".")
    elif has_comma:
        # Solo virgole: trattate come decimali.
        normalized = s.replace(".", "").replace(",", ".")
    elif has_dot:
        dot_count = s.count(".")
        if dot_count == 1:
            left, right = s.split(".")
            # Se ci sono 1-2 decimali, interpreta come separatore decimale.
            # Con 3 cifre a destra e parte sinistra lunga (>3), interpreta come migliaia.
            if right.isdigit() and len(right) in (1, 2):
                normalized = s
            elif right.isdigit() and len(right) == 3 and len(left.replace("-", "")) > 3:
                normalized = s.replace(".", "")
            else:
                normalized = s
        else:
            # Più punti: validi solo come separatori migliaia.
            parts = s.split(".")
            if all(p.isdigit() for p in parts) and all(len(p) == 3 for p in parts[1:]):
                normalized = "".join(parts)
            else:
                return np.nan
    else:
        normalized = s

    if not re.match(r"^-?\d+(\.\d+)?$", normalized):
        return np.nan

    try:
        n = float(normalized)
    except ValueError:
        return np.nan

    return -n if neg_parentheses else n

# =====================
# Configurazione pagina
# =====================
st.set_page_config(
    page_title="Audit Sampling – Key Items & Items",
    layout="wide"
)


st.title("Selezione campione – Key Items e Items")
st.markdown("<span style='color:blue;'>Il file da caricare deve riportare nella prima riga la descrizione delle colonne (es. codice, descrizione, importo...)</span>", unsafe_allow_html=True)

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
# Upload file (Excel o CSV)
# =====================
uploaded_file = st.file_uploader("Carica file Excel o CSV", type=["xlsx", "csv"])
if uploaded_file is None:
    st.stop()

# Leggi il file in base all'estensione

import os
filename = uploaded_file.name
ext = os.path.splitext(filename)[1].lower()
if ext == ".xlsx":
    df = pd.read_excel(uploaded_file)
elif ext == ".csv":
    import pandas.errors
    # Prova con vari separatori: virgola, punto e virgola, tab
    found = False
    for sep in [',', ';', '\t']:
        uploaded_file.seek(0)
        try:
            df = pd.read_csv(uploaded_file, sep=sep)
            if len(df.columns) > 1:
                found = True
                break
        except pandas.errors.ParserError:
            continue
    if not found:
        st.error("Il file CSV sembra avere una sola colonna o non è ben formattato. Prova a salvare il file con separatore virgola, punto e virgola o tab.")
        st.stop()
else:
    st.error("Formato file non supportato. Carica un file .xlsx o .csv.")
    st.stop()

if df.empty:
    st.error("File caricato vuoto.")
    st.stop()

# =====================
# Sidebar – info revisione
# =====================
st.sidebar.header("Informazioni revisione")
societa = st.sidebar.text_input("Società")
revisione_al_str = st.sidebar.text_input("Revisione al (GG/MM/AA)", value=datetime.now().strftime("%d/%m/%y"))
try:
    datetime.strptime(revisione_al_str, "%d/%m/%y")
except ValueError:
    st.sidebar.warning("Formato data non valido. Usa GG/MM/AA.")
    revisione_al_str = datetime.now().strftime("%d/%m/%y")
preparato_da = st.sidebar.text_input("Preparato da")
data_ora = datetime.now().strftime("%d/%m/%Y %H:%M")

# =====================
# Sidebar – mappatura colonne
# =====================

st.sidebar.header("Mappatura colonne")
colonne = df.columns.tolist()
if len(colonne) < 3:
    st.error("Il file deve contenere almeno 3 colonne (Codice, Descrizione, Valore). Colonne trovate: " + ", ".join(colonne))
    st.stop()
codice_col = st.sidebar.selectbox("Colonna Codice", colonne, index=0)
descr_col = st.sidebar.selectbox("Colonna Descrizione", colonne, index=1 if len(colonne) > 1 else 0)
valore_col = st.sidebar.selectbox("Colonna Valore", colonne, index=2 if len(colonne) > 2 else 0)


if len({codice_col, descr_col, valore_col}) < 3:
    st.warning("Le colonne devono essere diverse.")
    st.stop()

# Mostra il DataFrame subito dopo il caricamento per debug
st.write("Anteprima dati caricati:")
st.dataframe(df.head(10))

# =====================
# Normalizzazione dati
# =====================
# Parsing importi con logica italiana (es. 1.234,56)
df[valore_col] = df[valore_col].apply(parse_number_it)
valori_non_validi = df[valore_col].isna().sum()
df = df.dropna(subset=[valore_col])
warning_placeholder = st.empty()
if "hide_non_numeric_warning" not in st.session_state:
    st.session_state["hide_non_numeric_warning"] = False
if valori_non_validi > 0 and not st.session_state["hide_non_numeric_warning"]:
    warning_placeholder.warning(f"Scartate {valori_non_validi} righe: valore non numerico in formato italiano.")

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







# Metodo selezione Items (uniformato a Parametri di campionamento)
st.sidebar.header("Metodo selezione Items")
metodo = st.sidebar.radio(
    "",
    ["MUS", "Intervallo", "Casuale"],
    key="key_metodo"
)

# Opzione per starting point automatico o manuale (solo MUS e Intervallo)
manual_starting_point = None
starting_point_mode = None
if metodo in ["MUS", "Intervallo"]:
    st.sidebar.markdown("Opzioni Starting Point")
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

# Dataset base (ordine originale del file)
df_base = df.copy().reset_index(drop=True)

# Dataset ordinato SOLO per MUS
df_mus = df.sort_values(by=valore_col, ascending=False).reset_index(drop=True)

# =====================
# Dataset per visualizzazione Universo
# =====================
df_universo_view = df_base.copy()
df_universo_view[valore_col] = pd.to_numeric(df_universo_view[valore_col], errors="coerce")
# Imposta l'indice a partire da 1
df_universo_view.index = range(1, len(df_universo_view) + 1)

# =====================
# UNIVERSO
# =====================

tot_items = len(df_base)
tot_valore = df_base[valore_col].sum()
top5_val = df_base.sort_values(by=valore_col, ascending=False).head(5)[valore_col].sum()
if tot_valore == 0:
    st.error("Il totale valore dell'universo è 0. Verifica la colonna importi nel file caricato.")
    st.stop()

top5_perc = top5_val / tot_valore * 100 if tot_items >= 5 else 100

st.subheader("Universo completo")
st.dataframe(
    df_universo_view.style
    .format({valore_col: lambda x: format_number_it(x, 2)}),
    width="stretch"
)



c1, c2, c3 = st.columns(3)
c1.metric("Totale items", tot_items)
c2.metric("Valore totale", euro(tot_valore))
c3.metric("Top 5 (% valore)", f"{top5_perc:.2f}%")

# =====================
# CALCOLO CAMPIONE
# =====================
if st.button("Calcola selezione campione"):
    st.session_state["hide_non_numeric_warning"] = True
    warning_placeholder.empty()
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
            np.random.seed(42)
            start_idx = np.random.randint(0, step)
            selected_items = residuo.iloc[start_idx::step].head(num_items)
            starting_point = start_idx

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
    st.write(f"Materialità: € {materialita:,.0f}  |  Confidence level: {confidence_level}%")
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

    key_items_view = key_items[[codice_col, descr_col, valore_col]].copy()
    key_items_view[valore_col] = pd.to_numeric(key_items_view[valore_col], errors="coerce")
    if key_items_view[valore_col].isnull().any():
        st.warning("Attenzione: alcuni valori nei Key Items non sono numerici e sono stati impostati a NaN.")
    key_items_view.index = range(1, len(key_items_view) + 1)
    st.dataframe(
        key_items_view.style
        .format({valore_col: lambda x: format_number_it(x, 2)}),
        width="stretch"
    )

    st.subheader("Items selezionati")

    selected_items_view = selected_items[[codice_col, descr_col, valore_col]].copy()
    selected_items_view[valore_col] = pd.to_numeric(selected_items_view[valore_col], errors="coerce")
    if selected_items_view[valore_col].isnull().any():
        st.warning("Attenzione: alcuni valori negli Items selezionati non sono numerici e sono stati impostati a NaN.")
    selected_items_view.index = range(1, len(selected_items_view) + 1)
    st.dataframe(
        selected_items_view.style
        .format({valore_col: lambda x: format_number_it(x, 2)}),
        width="stretch"
    )

    # =====================
    # EXPORT EXCEL
    # =====================
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        # Sheet Riepilogo: aggiungo info revisione e parametri sopra la tabella
        info_df = pd.DataFrame({
            "Parametro": ["Societ\u00e0", "Revisione al", "Preparato da", "Materialit\u00e0", "Confidence level"],
            "Valore": [societa, revisione_al_str, preparato_da, euro(materialita), f"{confidence_level}%"]
        })
        info_df.to_excel(writer, sheet_name="Riepilogo", index=False, startrow=0)
        riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False, startrow=info_df.shape[0]+2)

        # valori numerici senza decimali
        key_items_export = key_items[[codice_col, descr_col, valore_col]].copy()
        key_items_export[valore_col] = pd.to_numeric(key_items_export[valore_col], errors="coerce")
        key_items_export[valore_col] = key_items_export[valore_col].round(0).astype('Int64')
        key_items_export.to_excel(writer, sheet_name="Key Items", index=False)

        selected_items_export = selected_items[[codice_col, descr_col, valore_col]].copy()
        selected_items_export[valore_col] = pd.to_numeric(selected_items_export[valore_col], errors="coerce")
        selected_items_export[valore_col] = selected_items_export[valore_col].round(0).astype('Int64')
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
        def set_cell_text(cell, text, align=None):
            cell.text = str(text)
            if align is not None and cell.paragraphs:
                cell.paragraphs[0].alignment = align

        doc = Document()
        doc.add_paragraph(f"Società: {societa}")
        doc.add_paragraph(f"Revisione al: {revisione_al_str}")
        doc.add_paragraph(f"Preparato da: {preparato_da}")
        doc.add_paragraph("")
        doc.add_heading("MEMO DI REVISIONE – SELEZIONE CAMPIONE", 0)
        doc.add_paragraph(f"File: {uploaded_file.name}")
        doc.add_paragraph(f"Data: {data_ora}")
        doc.add_paragraph(f"Metodo: {metodo}")
        if starting_point is not None:
            doc.add_paragraph(f"Starting point: {starting_point:.2f}")
        doc.add_paragraph(f"Materialità: € {materialita:,.0f}  |  Confidence level: {confidence_level}%")
        doc.add_paragraph("\nRiepilogo selezione")
        table = doc.add_table(rows=riepilogo.shape[0]+1, cols=riepilogo.shape[1])
        right_col_indexes = {1, 2, 3, 4}
        for j, col in enumerate(riepilogo.columns):
            header_align = WD_PARAGRAPH_ALIGNMENT.RIGHT if j in right_col_indexes else None
            set_cell_text(table.cell(0, j), col, header_align)
        for i in range(riepilogo.shape[0]):
            for j in range(riepilogo.shape[1]):
                align = WD_PARAGRAPH_ALIGNMENT.RIGHT if j in right_col_indexes else None
                set_cell_text(table.cell(i+1, j), riepilogo.iloc[i, j], align)
        doc.add_paragraph("\nKey Items")
        table_key = doc.add_table(rows=len(key_items_export)+1, cols=3)
        table_key.cell(0,0).text = codice_col
        table_key.cell(0,1).text = descr_col
        set_cell_text(table_key.cell(0,2), "Valore (\u20ac)", WD_PARAGRAPH_ALIGNMENT.RIGHT)
        for i, row in enumerate(key_items_export.itertuples(index=False)):
            set_cell_text(table_key.cell(i+1,0), row[0])
            set_cell_text(table_key.cell(i+1,1), row[1])
            set_cell_text(table_key.cell(i+1,2), row[2], WD_PARAGRAPH_ALIGNMENT.RIGHT)
        doc.add_paragraph("\nItems selezionati")
        table_sel = doc.add_table(rows=len(selected_items_export)+1, cols=3)
        table_sel.cell(0,0).text = codice_col
        table_sel.cell(0,1).text = descr_col
        set_cell_text(table_sel.cell(0,2), "Valore (\u20ac)", WD_PARAGRAPH_ALIGNMENT.RIGHT)
        for i, row in enumerate(selected_items_export.itertuples(index=False)):
            set_cell_text(table_sel.cell(i+1,0), row[0])
            set_cell_text(table_sel.cell(i+1,1), row[1])
            set_cell_text(table_sel.cell(i+1,2), row[2], WD_PARAGRAPH_ALIGNMENT.RIGHT)
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


