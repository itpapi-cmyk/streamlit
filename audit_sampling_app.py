# Forcing rebuilt on Streamlit Cloud
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas
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
    page_title="Audit Sampling – Key Items & MUS",
    layout="wide"
)

st.title("Selezione campione – Key Items e MUS (manuale)")

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
# Totale universo + Copertura Top5
# =====================
totale_universo = df_sorted[valore_col].sum()
n_elementi = len(df_sorted)
top5_ratio = df_sorted.head(5)[valore_col].sum() / totale_universo * 100 if n_elementi >= 5 else 100

st.subheader("Riepilogo universo")
c1, c2, c3 = st.columns(3)
c1.metric("Numero totale items", n_elementi)
c2.metric("Valore totale universo", euro(totale_universo))
c3.metric("Copertura primi 5 items (%)", f"{top5_ratio:.1f}")

# =====================
# Sidebar – parametri manuali
# =====================
st.sidebar.header("Parametri selezione manuale")
perc_key = st.sidebar.number_input("Soglia Key Items (%)", min_value=0, max_value=100, value=60, step=1)
num_mus = st.sidebar.number_input("Numero MUS Items", min_value=1, value=5, step=1)

# =====================
# Tabella completa
# =====================
st.subheader("Universo completo (ordinabile)")
st.dataframe(df_sorted, width="stretch")

# =====================
# Calcolo Key Items e MUS
# =====================
if st.button("Calcola Key Items e MUS"):
    # -------- KEY ITEMS --------
    soglia_key = totale_universo * perc_key / 100
    df_sorted["cumulativo"] = df_sorted[valore_col].cumsum()
    key_items = df_sorted[df_sorted["cumulativo"] <= soglia_key].copy()
    if key_items.empty:
        key_items = df_sorted.head(1)
    elif key_items["cumulativo"].iloc[-1] < soglia_key:
        key_items = df_sorted.head(len(key_items) + 1)

    # -------- RESIDUO --------
    residuo = df_sorted.drop(key_items.index).copy()
    residuo_tot = residuo[valore_col].sum()

    # -------- MUS --------
    mus_items = pd.DataFrame(columns=df_sorted.columns)
    starting_point = None
    if not residuo.empty and residuo_tot > 0:
        intervallo = residuo_tot / num_mus
        starting_point = np.random.uniform(0, intervallo)
        soglie = [starting_point + i * intervallo for i in range(num_mus)]
        residuo["cumulativo"] = residuo[valore_col].cumsum()
        idx = set()
        for s in soglie:
            r = residuo[residuo["cumulativo"] >= s].head(1)
            if not r.empty:
                idx.add(r.index[0])
        mus_items = residuo.loc[list(idx)]

    # =====================
    # Mostra risultati
    # =====================
    st.subheader("Key Items selezionati")
    st.dataframe(key_items[[codice_col, descr_col, valore_col]], width="stretch")

    st.subheader("MUS Items selezionati")
    st.dataframe(mus_items[[codice_col, descr_col, valore_col]], width="stretch")

    # =====================
    # Riepilogo finale numerico
    # =====================
    st.subheader("Riepilogo finale")
    riepilogo = pd.DataFrame({
        "Categoria": ["Key Items", "MUS Items", "Totale Universo"],
        "Numero Items": [len(key_items), len(mus_items), n_elementi],
        "Valore Totale": [key_items[valore_col].sum(), mus_items[valore_col].sum(), totale_universo]
    })
    # Formatta valori in euro
    riepilogo["Valore Totale"] = riepilogo["Valore Totale"].apply(euro)
    st.table(riepilogo)

    # =====================
    # Grafico a torta (più piccolo in Streamlit)
    # =====================
    residuo_val = totale_universo - key_items[valore_col].sum() - mus_items[valore_col].sum()
    valori_pie = [key_items[valore_col].sum(), mus_items[valore_col].sum(), residuo_val]
    etichette = ["Key Items", "MUS Items", "Residuo"]
    colori = ["red", "blue", "lightgrey"]

    fig1, ax1 = plt.subplots(figsize=(4,4))  # dimensioni ridotte per Streamlit
    ax1.pie(valori_pie, labels=etichette, autopct='%1.1f%%', colors=colori)
    ax1.set_title("Distribuzione valore universo")
    st.pyplot(fig1)

    # =====================
    # Scatter plot (più piccolo in Streamlit)
    # =====================
    df_sorted["Categoria"] = "Residuo"
    df_sorted.loc[key_items.index, "Categoria"] = "Key Items"
    df_sorted.loc[mus_items.index, "Categoria"] = "MUS Items"
    color_map = {"Key Items": "red", "MUS Items": "blue", "Residuo": "lightgrey"}

    fig2, ax2 = plt.subplots(figsize=(6,3))  # dimensioni ridotte per Streamlit
    for cat in df_sorted["Categoria"].unique():
        subset = df_sorted[df_sorted["Categoria"] == cat]
        ax2.scatter(subset.index, subset[valore_col], c=color_map[cat], label=cat)
    ax2.set_xlabel("Item")
    ax2.set_ylabel("Valore")
    ax2.set_title("Distribuzione items – Key Items e MUS")
    ax2.legend()
    st.pyplot(fig2)

    # =====================
    # Export PDF completo
    # =====================
    def export_memo_pdf():
        buffer = BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        styles = getSampleStyleSheet()
        elements = []

        # Titolo e info
        elements.append(Paragraph("<b>MEMO DI REVISIONE – SELEZIONE CAMPIONE</b>", styles["Title"]))
        elements.append(Spacer(1,12))
        elements.append(Paragraph(f"<b>File:</b> {uploaded_file.name}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Data:</b> {data_ora}", styles["Normal"]))
        elements.append(Paragraph(f"<b>Preparato da:</b> {preparato_da}", styles["Normal"]))
        elements.append(Spacer(1,12))

        # Riepilogo
        elements.append(Paragraph("<b>1. Riepilogo generale</b>", styles["Heading2"]))
        data_riepilogo = [
            ["Categoria", "Numero Items", "Valore Totale"],
            ["Key Items", len(key_items), euro(key_items[valore_col].sum())],
            ["MUS Items", len(mus_items), euro(mus_items[valore_col].sum())],
            ["Totale Universo", n_elementi, euro(totale_universo)]
        ]
        elements.append(Table(data_riepilogo, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        # Starting point MUS
        elements.append(Paragraph(f"<b>Starting point MUS:</b> {starting_point:.2f}", styles["Normal"]))
        elements.append(Spacer(1,12))

        # Grafico a torta
        fig1.savefig("pie_chart.png")
        elements.append(Paragraph("<b>Grafico a torta</b>", styles["Heading2"]))
        elements.append(Image("pie_chart.png", width=400, height=300))
        elements.append(Spacer(1,12))

        # Scatter plot
        fig2.savefig("scatter_plot.png")
        elements.append(Paragraph("<b>Scatter plot items</b>", styles["Heading2"]))
        elements.append(Image("scatter_plot.png", width=400, height=300))
        elements.append(Spacer(1,12))

        # Key Items dettagliati
        elements.append(Paragraph("<b>2. Key Items selezionati</b>", styles["Heading2"]))
        data_key = [[codice_col, descr_col, "Valore"]] + key_items[[codice_col, descr_col, valore_col]].values.tolist()
        elements.append(Table(data_key, hAlign='LEFT'))
        elements.append(Spacer(1,12))

        # MUS Items dettagliati
        elements.append(Paragraph("<b>3. MUS Items selezionati</b>", styles["Heading2"]))
        data_mus = [[codice_col, descr_col, "Valore"]] + mus_items[[codice_col, descr_col, valore_col]].values.tolist()
        elements.append(Table(data_mus, hAlign='LEFT'))

        doc.build(elements)
        buffer.seek(0)
        return buffer

    st.download_button(
        "Export Memo di Revisione (PDF)",
        data=export_memo_pdf(),
        file_name="memo_revisione_key_mus.pdf",
        mime="application/pdf"
    )

    # =====================
    # Export Excel completo
    # =====================
    def export_excel():
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            riepilogo.to_excel(writer, sheet_name="Riepilogo", index=False)
            key_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="Key Items", index=False)
            mus_items[[codice_col, descr_col, valore_col]].to_excel(writer, sheet_name="MUS Items", index=False)
        output.seek(0)
        return output

    st.download_button(
        "Export Excel",
        data=export_excel(),
        file_name="riepilogo_key_mus.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

