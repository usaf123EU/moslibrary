
import streamlit as st

st.set_page_config(page_title="EFESO Functional Cost Analysis – v10.2", layout="wide")

st.title("EFESO Functional Cost Analysis TOOLSET – v10.2")
st.caption("Funktionsmatrix • Kosten • Technik • Abweichungen")

uploaded_files = st.file_uploader(
    "Excel-Dateien (.xlsx/.xlsm) – je Produkt eine Datei",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

tabs = st.tabs(["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"])

with tabs[0]:
    st.subheader("Funktionsmatrix (H1 → zugehörige Nebenfunktionen)")
    st.markdown("Hier werden die Hauptfunktionen (H1) mit ihren Nebenfunktionen (H2) dargestellt.")
    st.text_area("Beschreibung / Kommentare – Funktionsmatrix", height=120)

with tabs[1]:
    st.subheader("Funktionenkosten")
    st.markdown("Zeigt Kosten nach Haupt- und Nebenfunktionen (Zeilen 7 und 8).")
    st.text_area("Beschreibung / Kommentare – Funktionenkosten", height=120)

with tabs[2]:
    st.subheader("Technische Bewertung")
    st.markdown("Bewertung von Nebenfunktionen (H2) nach Score & Gewichtung.")
    st.text_area("Beschreibung / Kommentare – Technik Bewertung", height=120)

with tabs[3]:
    st.subheader("Top Kostenabweichungen")
    st.markdown("Hier können zwei Produkte verglichen werden, um Abweichungen zu identifizieren.")
    st.text_area("Beschreibung / Kommentare – Abweichungen", height=120)

st.divider()
st.caption("EFESO • Version 10.2 • Vorlage für Funktions- & Kostenanalyse")
