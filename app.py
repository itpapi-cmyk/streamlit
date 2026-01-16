import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4

# =====================
# Funzione euro
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

perc_key = st.sidebar.number_input("Soglia Key Items (%)", 0, 100, 30)
materialita = st.sidebar.number_input("Materialità (€)", min_value=1.0, value=1_000_000.0)
confidence_level = st.sidebar.number_input("Confidence Level (%)", 1, 100, 80)

metodo = st.sidebar.radio(
    "Metodo selezione Items",
    ["MUS", "Intervallo", "Casuale"]
)

# =====================
# Normalizzazione dati
# =====================
df[valore_col] = pd.to_numeric(df[valore_col], errors="coerce")
df = df.dropna(subset=[valore_col])

df_sorted = df.sort_values(by=valore_col, ascending=False).reset_index(drop=True)

# =====================
# UNIVERSO
# =====================
tot_items = len(df_sorted)
tot_valore = df_sorted[valore_col].sum()
top5_val = df_sorted.head(5)[valore_col].sum()
top5_perc = top5_val / tot_valore * 100 if tot_items >= 5 else 100

st.subheader("Universo completo")
st.dataframe(df_sorted, width="stretch")

c1, c2, c3 = st.columns(3)
c1.metric("Totale items", tot_items)
c2.metric("Valore totale", euro(tot_valore))
c3.metric("Top 5 (% valore)", f"{top5_perc:.2f}%")

# =====================
# CALCOLO CAMPIONE
# =====================
if st.button("Calcola selezione campione"):

    # -------- KEY ITEMS --------
    soglia_key = tot_valore * perc_key / 100
    df_sorted["cumulativo"] = df_sorted[valore_col].cumsum()

    key_items = df_sorted[df_sorted["cumulativo"] <= soglia_key].copy()
    if key_items.empty:
        key_items = df_sorted.head(1)

    residuo = df_sorted.drop(key_items.index).copy()
    residuo_tot = residuo[valore_col].sum()

    # -------- CONFIDENCE FACTOR (Excel compliant) --------
    confidence_factor = 100 * (1 - ((100 - confidence_level) / 100) ** (1 / 100))

    # -------- NUMERO ITEMS --------
    num_items = int(round((residuo_tot / materialita) * confidence_factor))
    num_items = max(num_items, 1)

    # -------- SELEZIONE --------
    selected_items = pd.DataFrame(columns=df_sorted.columns)
    starting_point = None

    if metodo == "MUS" and residuo_tot > 0:
        intervallo = residuo_tot / num_items
        starting_point = np.random.uniform(0, intervallo)

        soglie = [starting_point + i * intervallo for i in range(num_items)]
        residuo["cumulativo"] = residuo[valore_col].cumsum()

        idx = set()
        for s in soglie:
            r = residuo[residuo["cumulativo"] >= s].head(1)
            if not r.empty:
                idx.add(r.index[0])

        selected_items = residuo.loc[list(idx)]

    elif metodo == "Intervallo":
        step = max(1, len(residuo) // num_items)
        selected_items = residuo.iloc[::step].head(num_items)
        starting_point = 1

    elif metodo == "Casuale":
        selected_items = residuo.sample(
            n=min(num_items, len(residuo)),
            random_state=42
        )

    # =====================
    # RIEPILOGO
    # =====================
    st.subheader("Riepilogo selezione")

    riepilogo = pd.DataFrame({
        "Categoria": ["Universo", "Key Items", "Items selezionati"],
        "Numero items": [tot_items, len(key_items), len(selected_items)],
        "Valore (€)": [
            tot_valore,
            key_items[valore_col].sum(),
            selected_items[valore_col].sum()
        ],
        "% su totale": [
            100,
            key_items[valore_col].sum() / tot_valore * 100,
            selected_items[valore_col].sum() / tot_valore * 100
        ]
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
        key_items.to_excel(writer, sheet_name="Key Items", index=False)
        selected_items.to_excel(writer, sheet_name="Items selezionati", index=False)
    excel_buffer.seek(0)

    st.download_button(
        "Export Excel",
        data=excel_buffer,
        file_name="audit_sampling.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # =====================
    # EXPORT PDF
    # =====================
    def export_pdf():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        elements.append(Paragraph("<b>MEMO DI REVISIONE – SELEZIONE CAMPIONE</b>", styles["Title"]))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"Metodo: {metodo}", styles["Normal"]))
        if starting_point is not None:
            elements.append(Paragraph(f"Starting point: {starting_point:.2f}", styles["Normal"]))
        elements.append(Spacer(1, 12))

        data = [riepilogo.columns.tolist()] + riepilogo.values.tolist()
        elements.append(Table(data))

        doc.build(elements)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export PDF",
        data=export_pdf(),
        file_name="audit_sampling.pdf",
        mime="application/pdf"
    )
