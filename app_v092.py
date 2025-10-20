
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

st.set_page_config(
    page_title="EFESO Functional Cost Analysis TOOLSET",
    page_icon="🟧",
    layout="wide"
)

# ---------------------- Header ----------------------
st.markdown("# EFESO Functional Cost Analysis TOOLSET")
st.caption("Funktionsmatrix · Kosten · Technik · Abweichungen")

# ---------------------- Sidebar ----------------------
with st.sidebar:
    st.markdown("## Hinweise & Erläuterungen")

    st.markdown("### 1️⃣ Funktionsmatrix")
    st.write(
        "Die Funktionsmatrix zeigt die Zuordnung von Haupt- und Nebenfunktionen. "
        "Die Prozentwerte in den Kacheln geben die relative Gewichtung der Nebenfunktion "
        "innerhalb der jeweiligen Hauptfunktion an. "
        "Eine hohe Gewichtung bedeutet, dass diese Teilfunktion technisch oder wirtschaftlich besonders relevant ist."
    )

    st.markdown("### 2️⃣ Funktionenkosten")
    st.write(
        "Hauptfunktionskosten stammen aus **Zeile 7** des Blatts "
        "`SLAVE_Funktions-Kostenstruktur`, die Teilfunktionskosten aus **Zeile 8**. "
        "Die Diagramme zeigen zuerst die Kosten je Hauptfunktion; per Auswahl erfolgt der Drilldown "
        "auf die Kosten der zugeordneten Nebenfunktionen."
    )

    st.markdown("### 3️⃣ Technische Bewertung & Abweichungen")
    st.write(
        "Die technische Bewertung wird im Blatt `SLAVE_Techn.Bewertung` je Nebenfunktion in "
        "Spalte **R** geführt (B → R). Die App aggregiert zusätzlich nach Hauptfunktionen. "
        "Im Reiter **Top Kostenabweichung** siehst du die größten deltas zwischen zwei gewählten Produkten."
    )

    st.markdown("---")
    st.caption("Version: **v09.2**")


def col_index_from_label(lbl: str) -> int:
    \"\"\"Convert Excel column label to zero-based integer index.\"\"\"
    idx = 0
    for c in lbl:
        idx = idx*26 + (ord(c.upper()) - ord('A') + 1)
    return idx - 1

def parse_kostenstruktur(xls: bytes):
    # Read sheet without header
    df = pd.read_excel(xls, sheet_name='SLAVE_Funktions-Kostenstruktur', header=None, engine='openpyxl')
    # Start reading at column I (index 8)
    start_col = col_index_from_label('I')
    # Row indices (0-based)
    row_h1 = 0     # Zeile 1
    row_h2 = 1     # Zeile 2
    row_w1 = 3     # Zeile 4 (H1 Gewichte)
    row_w2 = 4     # Zeile 5 (H2 Gewichte)
    row_c1 = 6     # Zeile 7 (H1 Kosten)
    row_c2 = 7     # Zeile 8 (H2 Kosten)

    h1_list = []
    h2_rows = []  # rows for matrix
    h1_costs = []
    h2_costs = []

    current_h1 = None

    # iterate columns until all empty in row_h1 and row_h2 for a while
    # We'll go through all columns from start_col to end
    for c in range(start_col, df.shape[1]):
        h1 = df.iat[row_h1, c]
        h2 = df.iat[row_h2, c]

        # Define switching to new H1 when a new non-empty in row_h1 appears
        if pd.notna(h1) and str(h1).strip() != '':
            current_h1 = str(h1).strip()
            if current_h1 not in h1_list:
                h1_list.append(current_h1)
            # H1 cost may be in this same column (sum over group later). We will capture at first occurrence.
        # Capture H1 costs if present for the current column
        if current_h1 is not None and pd.notna(df.iat[row_c1, c]):
            val = df.iat[row_c1, c]
            if isinstance(val, str):
                try:
                    val = float(str(val).replace('€', '').replace(',', '.'))
                except Exception:
                    val = np.nan
            # Keep only first occurrence per H1 (prevents duplicates across group columns)
            if (len(h1_costs)==0) or (current_h1 not in [x[0] for x in h1_costs]):
                h1_costs.append((current_h1, float(val)))

        # H2 entries
        if current_h1 is not None and pd.notna(h2) and str(h2).strip() != '':
            h2name = str(h2).strip()
            # Gewicht aus row_w2
            w = df.iat[row_w2, c]
            try:
                w = float(str(w).replace('%', '').replace(',', '.'))
            except Exception:
                w = np.nan
            # Kosten aus row_c2
            kc = df.iat[row_c2, c]
            try:
                kc = float(str(kc).replace('€', '').replace(',', '.'))
            except Exception:
                kc = np.nan
            h2_rows.append({
                "Hauptfunktion": current_h1,
                "Nebenfunktion": h2name,
                "H2_Gewicht_%": w,
                "H2_Kosten": kc
            })
            h2_costs.append((current_h1, h2name, kc))

    h1_df = pd.DataFrame(h1_costs, columns=["Hauptfunktion", "H1_Kosten"])
    # Summaries might be duplicated depending on layout; group by and take first non-NaN
    h1_df = (h1_df.groupby("Hauptfunktion", as_index=False)
             .agg({"H1_Kosten":"first"}))

    h2_df = pd.DataFrame(h2_rows)
    return h1_df, h2_df

def parse_techsheet(xls: bytes):
    try:
        df = pd.read_excel(xls, sheet_name='SLAVE_Techn.Bewertung', header=None, engine='openpyxl')
    except Exception:
        return pd.DataFrame(columns=["Nebenfunktion","Score"])
    # Nebenfunktionsnamen in Spalte B (index 1), Score in Spalte R (index 17), Zeilen ab 5..
    names = df.iloc[4:, 1]  # B
    scores = df.iloc[4:, 17] if df.shape[1] > 17 else pd.Series([], dtype=float)  # R
    res = pd.DataFrame({"Nebenfunktion": names, "Score": scores})
    res = res.dropna(subset=["Nebenfunktion"]).copy()
    # Zahlformat
    def to_float(x):
        try:
            return float(str(x).replace(',', '.'))
        except Exception:
            return np.nan
    res["Score"] = res["Score"].apply(to_float)
    return res

@st.cache_data(show_spinner=False)
def parse_uploaded_file(uploaded_file):
    bytes_data = uploaded_file.read()
    name = uploaded_file.name.rsplit('.',1)[0]
    h1_df, h2_df = parse_kostenstruktur(pd.io.common.BytesIO(bytes_data))
    tech_df = parse_techsheet(pd.io.common.BytesIO(bytes_data))

    # Merge Technik Scores auf H2
    if not h2_df.empty and not tech_df.empty:
        h2_df = h2_df.merge(tech_df, on="Nebenfunktion", how="left")
    else:
        if "Score" not in h2_df.columns:
            h2_df["Score"] = np.nan

    # Aggregation: gewichteter Score je H1 (H2-Gewichte als Anteil verwenden)
    h2_temp = h2_df.dropna(subset=["H2_Gewicht_%"]).copy()
    if not h2_temp.empty:
        h2_temp["w"] = h2_temp["H2_Gewicht_%"] / 100.0
        h2_temp["wScore"] = h2_temp["w"] * h2_temp["Score"]
        tech_h1 = (h2_temp.groupby("Hauptfunktion", as_index=False)
                   .agg(TechScore_H1_weighted=("wScore","sum"),
                        TechScore_H1_mean=("Score","mean")))
    else:
        tech_h1 = pd.DataFrame(columns=["Hauptfunktion","TechScore_H1_weighted","TechScore_H1_mean"])

    product = {
        "name": name,
        "h1": h1_df,
        "h2": h2_df,
        "tech_h1": tech_h1
    }
    return product

# ---------------------- Upload ----------------------
st.markdown("##### Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei")
uploads = st.file_uploader("Drag and drop files here", type=["xlsx","xlsm"], accept_multiple_files=True, label_visibility="collapsed")

if not uploads:
    st.info("Bitte mindestens eine Excel-Datei hochladen.")
    st.stop()

products = [parse_uploaded_file(f) for f in uploads if f is not None and f.size > 0]
prod_names = [p["name"] for p in products]

# ---------------------- Tabs ----------------------
tabMatrix, tabKosten, tabTech, tabDelta = st.tabs(["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"])

LABEL_ANGLE = 0  # horizontale Achsenbeschriftung

# ---------------------- Funktionsmatrix ----------------------
with tabMatrix:
    st.subheader("Funktionsmatrix")
    pname = st.selectbox("Produkt wählen", prod_names, key="mx_prod")
    P = next(p for p in products if p["name"] == pname)

    mat = P["h2"][["Hauptfunktion", "Nebenfunktion", "H2_Gewicht_%"]]
    if mat.empty:
        st.warning("Keine H2-Daten gefunden.")
    else:
        h1_order = list(P["h1"]["Hauptfunktion"])
        h2_order = sorted(mat["Nebenfunktion"].dropna().unique().tolist())

        heat = (
            alt.Chart(mat)
            .mark_rect()
            .encode(
                x=alt.X("Nebenfunktion:N",
                        sort=h2_order,
                        title="Nebenfunktion",
                        axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                y=alt.Y("Hauptfunktion:N",
                        sort=h1_order,
                        title="Hauptfunktion"),
                color=alt.Color("H2_Gewicht_%:Q",
                                scale=alt.Scale(scheme="oranges"),
                                title="Gewicht (%)"),
                tooltip=[
                    "Hauptfunktion",
                    "Nebenfunktion",
                    alt.Tooltip("H2_Gewicht_%:Q", format=".0f")
                ],
            )
            .properties(height=540)
        )

        txt = (
            alt.Chart(mat)
            .mark_text(color="black", fontSize=11)
            .encode(
                x=alt.X("Nebenfunktion:N", sort=h2_order),
                y=alt.Y("Hauptfunktion:N", sort=h1_order),
                text=alt.Text("H2_Gewicht_%:Q", format=".0f"),
            )
        )

        st.altair_chart(heat + txt, use_container_width=True)

# ---------------------- Funktionenkosten ----------------------
with tabKosten:
    st.subheader("Kosten je Hauptfunktion (Zeile 7)")
    pname2 = st.selectbox("Produkt wählen", prod_names, key="cost_prod")
    P2 = next(p for p in products if p["name"] == pname2)

    h1c = P2["h1"].dropna()
    if h1c.empty:
        st.warning("Keine H1-Kosten gefunden.")
    else:
        chart = (
            alt.Chart(h1c)
            .mark_bar(size=24)
            .encode(
                x=alt.X("Hauptfunktion:N", sort=list(h1c["Hauptfunktion"]),
                        title="Hauptfunktion", axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                y=alt.Y("H1_Kosten:Q", title="Kosten [€]"),
                tooltip=["Hauptfunktion", alt.Tooltip("H1_Kosten:Q", format=".2f")]
            )
            .properties(height=320)
        )
        st.altair_chart(chart, use_container_width=True)

    st.markdown("**Drilldown: Kosten je Nebenfunktion (Zeile 8)**")
    # H2 drilldown
    if not P2["h2"].empty:
        h1_choices = list(P2["h1"]["Hauptfunktion"])
        sel_h1 = st.selectbox("Hauptfunktion wählen", h1_choices, key="dd_h1")
        dfx = P2["h2"][P2["h2"]["Hauptfunktion"] == sel_h1][["Nebenfunktion","H2_Kosten"]].dropna()
        if dfx.empty:
            st.info("Keine Nebenfunktionskosten gefunden.")
        else:
            dd = (
                alt.Chart(dfx)
                .mark_bar(size=18)
                .encode(
                    x=alt.X("Nebenfunktion:N", sort=list(dfx["Nebenfunktion"]), axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                    y=alt.Y("H2_Kosten:Q", title="Kosten [€]"),
                    tooltip=["Nebenfunktion", alt.Tooltip("H2_Kosten:Q", format=".2f")]
                )
                .properties(height=280)
            )
            st.altair_chart(dd, use_container_width=True)

# ---------------------- Technik Bewertung ----------------------
with tabTech:
    st.subheader("Technische Bewertung – Linienvergleich je Nebenfunktion (B → R)")
    # Linien: für eine Nebenfunktion Vergleich aller Produkte
    # Build Nebenfunktionsliste (Union)
    all_h2 = sorted(set(pd.concat([p["h2"]["Nebenfunktion"] for p in products if not p["h2"].empty]).dropna().tolist()))
    if not all_h2:
        st.warning("Keine Nebenfunktionen vorhanden.")
    else:
        sel_nf = st.selectbox("Nebenfunktion wählen", all_h2, key="tech_nf")
        rows = []
        for p in products:
            sub = p["h2"][p["h2"]["Nebenfunktion"] == sel_nf]
            if not sub.empty:
                sc = sub["Score"].iloc[0]
            else:
                sc = np.nan
            rows.append({"Produkt": p["name"], "Score": sc})
        dfl = pd.DataFrame(rows)
        line = (
            alt.Chart(dfl)
            .mark_line(point=True)
            .encode(
                x=alt.X("Produkt:N", sort=prod_names, axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                y=alt.Y("Score:Q", title="Technischer Score"),
                tooltip=["Produkt", alt.Tooltip("Score:Q", format=".2f")]
            )
            .properties(height=320)
        )
        st.altair_chart(line, use_container_width=True)

    st.markdown("**Aggregiert nach Hauptfunktion (Gruppierte Balken)**")
    # Gruppierte Balken: gewichteter Score je H1 pro Produkt
    bars = []
    for p in products:
        th1 = p["tech_h1"]
        if not th1.empty:
            tmp = th1.copy()
            tmp["Produkt"] = p["name"]
            bars.append(tmp[["Produkt","Hauptfunktion","TechScore_H1_weighted"]])
    if bars:
        dfb = pd.concat(bars, ignore_index=True)
        bar = (
            alt.Chart(dfb)
            .mark_bar(size=20)
            .encode(
                x=alt.X("Hauptfunktion:N", sort=list(dfb["Hauptfunktion"].unique()), axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                y=alt.Y("TechScore_H1_weighted:Q", title="Gewichteter Technischer Score"),
                color=alt.Color("Produkt:N"),
                tooltip=["Produkt","Hauptfunktion", alt.Tooltip("TechScore_H1_weighted:Q", format=".2f")]
            )
            .properties(height=320)
        )
        st.altair_chart(bar, use_container_width=True)
    else:
        st.info("Keine aggregierten technischen Scores verfügbar.")

# ---------------------- Top Kostenabweichung ----------------------
with tabDelta:
    st.subheader("Top Kostenabweichungen (Nebenfunktionen)")
    if len(products) < 2:
        st.info("Bitte mindestens zwei Produkte hochladen.")
    else:
        colA, colB = st.columns(2)
        with colA:
            pa = st.selectbox("Produkt A", prod_names, key="deltaA")
        with colB:
            pb = st.selectbox("Produkt B", prod_names, key="deltaB")
        A = next(p for p in products if p["name"] == pa)
        B = next(p for p in products if p["name"] == pb)

        h2A = A["h2"][["Hauptfunktion","Nebenfunktion","H2_Kosten"]].rename(columns={"H2_Kosten":"Cost_A"})
        h2B = B["h2"][["Hauptfunktion","Nebenfunktion","H2_Kosten"]].rename(columns={"H2_Kosten":"Cost_B"})
        merged = h2A.merge(h2B, on=["Hauptfunktion","Nebenfunktion"], how="outer")
        merged[["Cost_A","Cost_B"]] = merged[["Cost_A","Cost_B"]].fillna(0.0)
        merged["Delta"] = merged["Cost_B"] - merged["Cost_A"]

        top = merged.sort_values("Delta", ascending=False).head(10)
        bottom = merged.sort_values("Delta", ascending=True).head(10)
        show = pd.concat([top, bottom]).drop_duplicates().sort_values("Delta", ascending=False)

        st.dataframe(show, use_container_width=True)
        ch = (
            alt.Chart(show)
            .mark_bar(size=18)
            .encode(
                x=alt.X("Nebenfunktion:N", sort=list(show["Nebenfunktion"]), axis=alt.Axis(labelAngle=LABEL_ANGLE)),
                y=alt.Y("Delta:Q", title="Kosten-Differenz (B − A)"),
                color=alt.condition(alt.datum.Delta > 0, alt.value("#1f77b4"), alt.value("#d62728")),
                tooltip=["Hauptfunktion","Nebenfunktion", alt.Tooltip("Delta:Q", format=".2f")]
            )
            .properties(height=320)
        )
        st.altair_chart(ch, use_container_width=True)
