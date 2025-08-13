
import streamlit as st
import pandas as pd
import io
from scripts.agent_workflow import detect_columns, run_full_workflow, AVAILABLE_FIELDS

st.set_page_config(page_title="Analisi Portafoglio – ONE-CLICK (Auto)", layout="wide")
st.title("🔎 Analisi Portafoglio – ONE-CLICK (Auto)")

uploaded = st.file_uploader("Carica il file portafoglio (Excel/CSV)", type=["xlsx","xls","csv"])
preset_name = ""
if uploaded is not None and uploaded.name:
    if "MERASSI" in uploaded.name.upper():
        preset_name = "MERASSI"

if uploaded is None:
    st.info("Carica un portafoglio: l'analisi parte automaticamente.")
else:
    raw = uploaded.read()
    buf = io.BytesIO(raw)

    # Analisi immediata
    with st.spinner("Analisi in corso..."):
        results = run_full_workflow(buf, auto_email="", user_map=None, preset_name=preset_name)

    # Output
    st.success("Analisi completata.")
    st.subheader("Riepilogo parsing")
    st.json(results["summary"])

    if results.get("input_preview") is not None:
        st.subheader("Anteprima input")
        st.dataframe(results["input_preview"].head(20))

    st.download_button("⬇️ Scarica Excel Master", data=results["excel_bytes"], file_name="Excel_Master_AAI.xlsx")
    st.download_button("⬇️ Scarica Report PDF (AAI)", data=results["pdf_bytes"], file_name="Report_AAI.pdf")

    if results["summary"].get("AUM_totale", 0) == 0:
        st.warning("AUM totale = 0. Se i nomi colonne sono molto particolari, dimmeli e li aggiungo al riconoscimento.")
