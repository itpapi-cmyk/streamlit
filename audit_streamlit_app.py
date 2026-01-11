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

st.title("Selezione campione – Key Items e Items residui")

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
# Sidebar – informazioni revisione
# =====================
st.sidebar.header("Informazioni revisione")
preparato_da = st.sidebar.text_input("Preparato da")
data_ora = datetime.now().strftime("%d/%m/%Y %H:%M")

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
# Universo – riepilogo base
# =====================
totale_universo_val = df_sorted[valore_col].sum()
n_total_items = len(df_sorted)
top5_val = df_sorted.head(5)[valore_col].sum()
top5_perc = top5_val / totale_universo_val * 100 if n_total_items >= 5 else 100

st.subheader("Universo completo (ordinabile)")
st.dataframe(df_sorted, width="stretch")

st.subheader("Riepilogo universo")
c1, c2, c3 = st.columns(3)
c1.metric("Numero totale items", n_total_items)
c2.metric("Saldo totale universo", euro(totale_universo_val))
c3.metric("Primi 5 items (%)", f"{top5_perc:.1f} – {euro(top5_val)}")

# =====================
# Sidebar – parametri selezione
# =====================
st.sidebar.header("Parametri selezione manuale")
perc_key = st.sidebar.number_input("Soglia Key Items (%)", min_value=0, max_value=100, value=30, step=1)
materialita = st.sidebar.number_input("Materialità (€)", min_value=1.0, value=1000.0, step=100.0, format="%.2f")
confidence_level = st.sidebar.number_input("Confidence Level (%)", min_value=1.0, max_value=100.0, value=95.0, step=1.0, format="%.1f")
criterio = st.sidebar.selectbox("Criterio selezione items residui", ["MUS", "Intervallo Items", "Casuale"])

# =====================
# Bottone per calcolo
# =====================
if st.button("Seleziona Key Items e Items residui"):

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

    # -------- Confidence factor --------
    conf_factor = 100 * (1 - ((100 - confidence_level) / 100) ** (1 / 100))

    # -------- Numero items residui --------
    if residuo_tot > 0 and materialita > 0:
        num_items_calc = int(round(residuo_tot / (materialita * conf_factor)))
        if num_items_calc < 1:
            num_items_calc = 1
    else:
        num_items_calc = 0

    # -------- Selezione Items --------
    mus_items = pd.DataFrame(columns=df_sorted.columns)
    starting_point = None

    if criterio == "MUS" and not residuo.empty:
        intervallo = residuo_tot / num_items_calc
        starting_point = np.random.uniform(0, intervallo)
        soglie = [starting_point + i * intervallo for i in range(num_items_calc)]
        residuo["cumulativo"] = residuo[valore_col].cumsum()
        idx = set()
        for s in soglie:
            r = residuo[residuo["cumulativo"] >= s].head(1)
            if not r.empty:
                idx.add(r.index[0])
        mus_items = residuo.loc[list(idx)]
    elif criterio == "Intervallo Items" and not residuo.empty:
        intervallo_idx = max(1, len(residuo) // num_items_calc)
        mus_items = residuo.iloc[::intervallo_idx].head(num_items_calc)
        starting_point = 1  # primo indice come riferimento
    elif criterio == "Casuale" and not residuo.empty:
        mus_items = residuo.sample(n=min(num_items_calc, len(residuo)), random_state=42)
        starting_point = None

    # -------- Riepilogo finale --------
    residuo_val = residuo_tot - mus_items[valore_col].sum()
    riepilogo = pd.DataFrame({
        "Categoria": ["Totale Universo", "Key Items", "Items selezionati"],
        "Numero items": [n_total_items, len(key_items), len(mus_items)],
        "Valore (€)": [totale_universo_val, key_items[valore_col].sum(), mus_items[valore_col].sum()],
        "% sul totale": [100, key_items[valore_col].sum()/totale_universo_val*100, mus_items[valore_col].sum()/totale_universo_val*100],
        "Criterio selezione": ["-", "-", criterio],
        "Starting point": ["-", "-", euro(starting_point) if criterio=="MUS" else starting_point if criterio=="Intervallo Items" else "-"]
    })

    st.subheader("Riepilogo selezione")
    st.table(riepilogo)

    # -------- Tabelle dettagliate --------
    st.subheader("Key Items")
    st.dataframe(key_items[[codice_col, descr_col, valore_col]], width="stretch")
    st.subheader("Items selezionati")
    st.dataframe(mus_items[[codice_col, descr_col, valore_col]], width="stretch")

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

        elements.append(Paragraph("<b>Riepilogo selezione</b>", styles["Heading2"]))
        data_riepilogo = [riepilogo.columns.tolist()] + riepilogo.values.tolist()
        elements.append(Table(data_riepilogo, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        elements.append(Paragraph("<b>Key Items selezionati</b>", styles["Heading2"]))
        data_key = [[codice_col, descr_col, "Valore (€)"]] + key_items[[codice_col, descr_col, valore_col]].fillna(0).round(0).astype(str).values.tolist()
        elements.append(Table(data_key, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        elements.append(Paragraph("<b>Items selezionati</b>", styles["Heading2"]))
        data_mus = [[codice_col, descr_col, "Valore (€)"]] + mus_items[[codice_col, descr_col, valore_col]].fillna(0).round(0).astype(str).values.tolist()
        elements.append(Table(data_mus, hAlign='LEFT'))

        doc.build(elements)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export Memo di Revisione (PDF)",
        data=export_pdf(),
        file_name="memo_revisione.pdf",
        mime="application/pdf"
    )

    # =====================
    # Export Excel
    # =====================
    def export_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
            key_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Key Items", index=False)
            mus_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Items selezionati", index=False)
        output.seek(0)
        return output

    st.download_button(
        "Export Excel",
        data=export_excel(),
        file_name="riepilogo_selezione.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
