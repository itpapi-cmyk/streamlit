import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# =====================
# Funzione formattazione euro
# =====================
def euro(val):
    return f"€ {val:,.0f}".replace(",", ".")

# =====================
# Configurazione pagina
# =====================
st.set_page_config(
    page_title="Audit Sampling – Key Items & Selezione Items",
    layout="wide"
)
st.title("Selezione campione – Key Items e Items (manuale)")

# =====================
# Sidebar – informazioni revisione
# =====================
st.sidebar.header("Informazioni revisione")
preparato_da = st.sidebar.text_input("Preparato da")
data_ora = datetime.now().strftime("%d/%m/%Y %H:%M")

# =====================
# Upload file
# =====================
uploaded_file = st.file_uploader("Carica file Excel", type=["xlsx"])
if uploaded_file is None:
    st.info("Carica un file Excel per iniziare.")
    st.stop()

# =====================
# Lettura dati
# =====================
df = pd.read_excel(uploaded_file)
if df.empty:
    st.error("Il file Excel non contiene dati.")
    st.stop()

# =====================
# Sidebar – mappatura colonne
# =====================
st.sidebar.header("Mappatura colonne")
colonne = df.columns.tolist()
if len(colonne) < 3:
    st.error("Il file deve contenere almeno 3 colonne.")
    st.stop()

codice_col = st.sidebar.selectbox("Colonna Codice", colonne, index=0)
descr_col = st.sidebar.selectbox("Colonna Descrizione", colonne, index=1 if len(colonne) > 1 else 0)
valore_col = st.sidebar.selectbox("Colonna Valore (saldo/importo)", colonne, index=2 if len(colonne) > 2 else 0)
if len({codice_col, descr_col, valore_col}) < 3:
    st.warning("Seleziona tre colonne diverse per Codice, Descrizione e Valore.")
    st.stop()

# =====================
# Sidebar – parametri selezione
# =====================
st.sidebar.header("Parametri selezione manuale")
perc_key = st.sidebar.number_input("Soglia Key Items (%)", min_value=0, max_value=100, value=30, step=1)
num_items = st.sidebar.number_input("Numero items", min_value=1, value=5, step=1)
metodo = st.sidebar.selectbox("Metodo selezione", ["MUS", "Intervallo fisso", "Casuale"])

# =====================
# Normalizzazione dati
# =====================
df[codice_col] = df[codice_col].astype(str)
df[descr_col] = df[descr_col].astype(str)
df[valore_col] = pd.to_numeric(df[valore_col], errors="coerce")
df = df.dropna(subset=[valore_col])
if df.empty:
    st.error("La colonna Valore non contiene dati numerici validi.")
    st.stop()

# =====================
# Ordinamento
# =====================
df_sorted = df.sort_values(by=valore_col, ascending=False).reset_index(drop=True)

# =====================
# Totale universo + Copertura Top5
# =====================
totale_universo_val = df_sorted[valore_col].sum()
totale_universo_items = len(df_sorted)
top5_ratio = df_sorted.head(5)[valore_col].sum() / totale_universo_val * 100 if totale_universo_items >= 5 else 100

# =====================
# Anteprima universo completo
# =====================
st.subheader("Universo completo (ordinabile)")
st.dataframe(df_sorted, width="stretch")

# =====================
# Riepilogo e selezione campione
# =====================
campione = pd.DataFrame(columns=df_sorted.columns)
key_items = pd.DataFrame(columns=df_sorted.columns)
starting_point = None

if st.button("Calcola Key Items e Items"):
    # -------- KEY ITEMS --------
    soglia_key = totale_universo_val * perc_key / 100
    df_sorted["cumulativo"] = df_sorted[valore_col].cumsum()
    key_items = df_sorted[df_sorted["cumulativo"] <= soglia_key].copy()
    if key_items.empty:
        key_items = df_sorted.head(1)
    elif key_items["cumulativo"].iloc[-1] < soglia_key:
        key_items = df_sorted.head(len(key_items) + 1)

    # -------- RESIDUO --------
    residuo = df_sorted.drop(key_items.index).copy()
    residuo_tot = residuo[valore_col].sum()

    # -------- ITEMS SELEZIONATI --------
    campione = pd.DataFrame(columns=df_sorted.columns)
    if metodo == "MUS" and not residuo.empty and residuo_tot > 0:
        intervallo = residuo_tot / num_items
        starting_point = np.random.uniform(0, intervallo)
        soglie = [starting_point + i * intervallo for i in range(num_items)]
        residuo["cumulativo"] = residuo[valore_col].cumsum()
        idx = set()
        for s in soglie:
            r = residuo[residuo["cumulativo"] >= s].head(1)
            if not r.empty:
                idx.add(r.index[0])
        campione = residuo.loc[list(idx)]
    elif metodo == "Intervallo fisso" and not residuo.empty:
        starting_point = np.random.randint(0, len(residuo))
        step = max(1, len(residuo)//num_items)
        idx = [(starting_point + i*step)%len(residuo) for i in range(num_items)]
        campione = residuo.iloc[idx]
    elif metodo == "Casuale" and not residuo.empty:
        np.random.seed(42)  # Per coerenza ISA 530
        selected_idx = np.random.choice(residuo.index, size=min(num_items, len(residuo)), replace=False)
        campione = residuo.loc[selected_idx]

    # =====================
    # Riepilogo dettagliato
    # =====================
    riepilogo = pd.DataFrame({
        "Parametro": [
            "Totale universo",
            "Key Items",
            "Items selezionati",
            "Metodo selezione",
            "Starting point (MUS)"
        ],
        "Numero items": [
            totale_universo_items,
            len(key_items),
            len(campione),
            "",
            ""
        ],
        "Valore (€)": [
            euro(totale_universo_val),
            euro(key_items[valore_col].sum()),
            euro(campione[valore_col].sum()) if not campione.empty else "N/A",
            "",
            starting_point if starting_point is not None else ""
        ],
        "% sul totale universo": [
            "",
            f"{key_items[valore_col].sum()/totale_universo_val*100:.1f}%",
            f"{campione[valore_col].sum()/totale_universo_val*100:.1f}%" if not campione.empty else "",
            "",
            ""
        ],
        "Criterio selezione": [
            "",
            "",
            "",
            metodo,
            ""
        ]
    })
    riepilogo["Numero items"] = riepilogo["Numero items"].astype(str)  # Fix PyArrow warning

    st.subheader("Riepilogo selezione")
    st.dataframe(riepilogo, width="stretch")

    st.subheader("Key Items selezionati")
    st.dataframe(key_items[[codice_col, descr_col, valore_col]], width="stretch")

    st.subheader("Items selezionati")
    st.dataframe(campione[[codice_col, descr_col, valore_col]], width="stretch")

    # =====================
    # Export Excel
    # =====================
    def export_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
            key_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Key Items", index=False)
            campione[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Items selezionati", index=False)
        output.seek(0)
        return output

    st.download_button(
        "Export Excel",
        data=export_excel(),
        file_name="riepilogo_key_items.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =====================
    # Export PDF
    # =====================
    def export_pdf():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("<b>MEMO DI REVISIONE – SELEZIONE CAMPIONE</b>", styles["Title"]))
        elements.append(Spacer(1,12))
        elements.append(Paragraph(f"<b>File:</b> {uploaded_file.name}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Data:</b> {data_ora}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Preparato da:</b> {preparato_da}", styles["Normal"]))
        elements.append(Spacer(1,12))

        # Riepilogo
        elements.append(Paragraph("<b>Riepilogo selezione</b>", styles["Heading2"]))
        data_riepilogo = [riepilogo.columns.tolist()] + riepilogo.values.tolist()
        elements.append(Table(data_riepilogo, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        # Key Items dettagliati
        elements.append(Paragraph("<b>Key Items selezionati</b>", styles["Heading2"]))
        data_key = [[codice_col, descr_col, "Valore (€)"]] + key_items[[codice_col, descr_col, valore_col]].values.tolist()
        elements.append(Table(data_key, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        # Items selezionati dettagliati
        elements.append(Paragraph("<b>Items selezionati</b>", styles["Heading2"]))
        data_items = [[codice_col, descr_col, "Valore (€)"]] + campione[[codice_col, descr_col, valore_col]].values.tolist()
        elements.append(Table(data_items, hAlign='LEFT'))

        doc.build(elements)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export Memo di Revisione (PDF)",
        data=export_pdf(),
        file_name="memo_revisione_key_items.pdf",
        mime="application/pdf"
    )
