# app_v102.py
# EFESO Functional Cost Analysis TOOLSET — v10.2
# Includes: Swimlane Funktionsmatrix, Funktionen­kosten, Technik Bewertung (Nebenfunktionen, multi-product line),
# Top Kostenabweichung, plus Beschreibung/Kommentar-Felder je Tab + Export.

from typing import List, Tuple, Dict, Any
import math
import json

import pandas as pd
import numpy as np
import altair as alt
import streamlit as st

st.set_page_config(page_title="EFESO Functional Cost Analysis TOOLSET — v10.2", layout="wide")
st.markdown("## EFESO Functional Cost Analysis TOOLSET — v10.2")
st.caption("Funktionsmatrix · Funktionen­kosten · Technik (Nebenfunktionen) · Top Kostenabweichung — mit Kommentarfeldern")

H1_LABEL = "Hauptfunktion"
H2_LABEL = "Nebenfunktion"

# -------------------- Helpers --------------------
def excel_col_to_index(col: str) -> int:
    col = col.strip().upper()
    res = 0
    for ch in col:
        if not ('A' <= ch <= 'Z'):
            raise ValueError("Ungültiges Spaltenlabel")
        res = res * 26 + (ord(ch) - 64)
    return res - 1

START_COL = excel_col_to_index("I")
ROW_H1 = 0       # Zeile 1
ROW_H2 = 1       # Zeile 2
ROW_COST_H1 = 6  # Zeile 7
ROW_COST_H2 = 7  # Zeile 8

def parse_funktionskosten_sheet(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    # Returns (h1_costs, h2_costs)
    cols = df.columns.tolist()
    h1_map: Dict[int, str] = {}
    current_h1 = None

    # Map columns to current H1 based on Zeile 1
    for j in range(START_COL, len(cols)):
        v1 = df.iat[ROW_H1, j] if j < df.shape[1] else None
        v2 = df.iat[ROW_H2, j] if j < df.shape[1] else None
        v1s = str(v1).strip() if (v1 is not None and not pd.isna(v1)) else ""
        v2s = str(v2).strip() if (v2 is not None and not pd.isna(v2)) else ""
        if v1s and v1s.lower() != "nan":
            current_h1 = v1s
        if current_h1 and v2s and v2s.lower() != "nan":
            h1_map[j] = current_h1

    # H1-Kosten (erste Spalte pro H1 nehmen)
    h1_rows = []
    seen = set()
    for j in range(START_COL, len(cols)):
        v1 = df.iat[ROW_H1, j] if j < df.shape[1] else None
        v1s = str(v1).strip() if (v1 is not None and not pd.isna(v1)) else ""
        if v1s and v1s.lower() != "nan" and v1s not in seen:
            cost = df.iat[ROW_COST_H1, j] if j < df.shape[1] else None
            try:
                cost = float(str(cost).replace(",", "."))
            except Exception:
                cost = np.nan
            h1_rows.append((v1s, cost))
            seen.add(v1s)
    h1_costs = pd.DataFrame(h1_rows, columns=[H1_LABEL, "Kosten_H1"]).dropna(subset=["Kosten_H1"])

    # H2-Kosten je Spalte
    h2_rows = []
    for j in range(START_COL, len(cols)):
        v2 = df.iat[ROW_H2, j] if j < df.shape[1] else None
        v2s = str(v2).strip() if (v2 is not None and not pd.isna(v2)) else ""
        if not v2s or v2s.lower() == "nan":
            continue
        h1 = h1_map.get(j, None)
        cost = df.iat[ROW_COST_H2, j] if j < df.shape[1] else None
        try:
            cost = float(str(cost).replace(",", "."))
        except Exception:
            cost = np.nan
        if h1 and not pd.isna(cost):
            h2_rows.append((h1, v2s, cost))
    h2_costs = pd.DataFrame(h2_rows, columns=[H1_LABEL, H2_LABEL, "Kosten_H2"])

    return h1_costs, h2_costs

def find_sheet(xls: pd.ExcelFile, candidates: List[str]) -> str:
    for s in xls.sheet_names:
        ls = s.lower()
        if any(c in ls for c in candidates):
            return s
    return None

def load_product(file) -> Dict[str, Any]:
    # Read both sheets. Flexible naming for cost-sheet and tech-sheet.
    name = file.name.rsplit(".", 1)[0]
    xls = pd.ExcelFile(file, engine="openpyxl")

    # Kosten/Struktur Sheet: support both 'Slave_Funktions_kostenstruktur' and 'SLAVE_Funktions-Kostenstruktur'
    sheet_costs = find_sheet(xls, ["slave_funktions", "kostenstruktur", "funktions_kosten"])
    if sheet_costs is None:
        raise ValueError("Funktions-Kostenstruktur-Blatt nicht gefunden.")

    df_costs = xls.parse(sheet_costs, header=None)
    h1_costs, h2_costs = parse_funktionskosten_sheet(df_costs)

    # Technik Sheet optional
    sheet_tech = find_sheet(xls, ["techn", "bewert"])
    tech_df = pd.DataFrame(columns=[H2_LABEL, "Score"])
    if sheet_tech:
        df_tech = xls.parse(sheet_tech, header=None)
        # Spalte B (index 1) = Nebenfunktion, Spalte R (index 17) = Score
        names = df_tech.iloc[:, 1] if df_tech.shape[1] > 1 else pd.Series(dtype=str)
        scores = df_tech.iloc[:, 17] if df_tech.shape[1] > 17 else pd.Series(dtype=float)

        def to_num(x):
            try:
                return float(str(x).replace(",", "."))
            except Exception:
                return np.nan

        tech_df = pd.DataFrame({H2_LABEL: names.astype(str).str.strip(), "Score": scores.map(to_num)})
        tech_df = tech_df.dropna(subset=["Score"])
        tech_df = tech_df[tech_df[H2_LABEL].str.len() > 0]

    return {"Produkt": name, "h1_costs": h1_costs, "h2_costs": h2_costs, "tech": tech_df}

# -------------------- Upload --------------------
uploads = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei",
                           type=["xlsx", "xlsm"], accept_multiple_files=True)

if not uploads:
    st.info("Bitte Dateien hochladen.")
    st.stop()

products = []
errors = []
for f in uploads:
    try:
        products.append(load_product(f))
    except Exception as e:
        errors.append(f"**{f.name}**: {e}")

if errors:
    st.error("Einige Dateien konnten nicht verarbeitet werden:\n\n- " + "\n- ".join(errors))

if not products:
    st.stop()

# A helper to keep and export comments
def get_comments_state():
    if "efeso_comments" not in st.session_state:
        st.session_state["efeso_comments"] = {
            "Funktionsmatrix": "",
            "Funktionen­kosten": "",
            "Technik Bewertung (Nebenfunktionen)": "",
            "Top Kostenabweichung": "",
        }
    return st.session_state["efeso_comments"]

def export_comments_button():
    comments = get_comments_state()
    data = json.dumps(comments, ensure_ascii=False, indent=2).encode("utf-8")
    st.download_button("Kommentare exportieren (JSON)", data=data, file_name="EFESO_FKA_v102_kommentare.json", mime="application/json")

# -------------------- Tabs --------------------
tab_matrix, tab_costs, tab_tech, tab_top = st.tabs([
    "Funktionsmatrix",
    "Funktionen­kosten",
    "Technik Bewertung (Nebenfunktionen)",
    "Top Kostenabweichung",
])

# -------------------- Funktionsmatrix (Swimlane) --------------------
with tab_matrix:
    st.subheader("Funktionsmatrix (H1 → zugehörige Nebenfunktionen)")

    prod_names = [p["Produkt"] for p in products]
    psel = st.selectbox("Produkt wählen", prod_names, key="mx_prod")
    P = next(p for p in products if p["Produkt"] == psel)
    df = P["h2_costs"].copy()

    # Reihenfolge H1 wie in Datei (erste Vorkommen)
    h1_order = list(P["h1_costs"][H1_LABEL])

    # Reihenfolge H2 je H1: nach Kosten absteigend
    df["H2_order"] = df.groupby(H1_LABEL)["Kosten_H2"].rank(method="first", ascending=False) - 1
    df["H2_order"] = df["H2_order"].astype(int)

    base = alt.Chart(df)

    tiles = base.mark_bar(size=26, cornerRadius=3).encode(
        y=alt.Y(f"{H1_LABEL}:N", sort=h1_order, title=None),
        x=alt.X("H2_order:O", axis=alt.Axis(labels=False, ticks=False, title=None)),
        color=alt.Color("Kosten_H2:Q", scale=alt.Scale(scheme="orangered"), legend=alt.Legend(title="Kosten H2 [€]")),
        tooltip=[H1_LABEL, H2_LABEL, alt.Tooltip("Kosten_H2:Q", format=".2f")]
    ).properties(height=460)

    labels = base.mark_text(color="white", fontSize=11).encode(
        y=alt.Y(f"{H1_LABEL}:N", sort=h1_order, title=None),
        x="H2_order:O",
        text=alt.Text(f"{H2_LABEL}:N")
    )

    st.altair_chart(tiles + labels, use_container_width=True)

    # Beschreibung/Kommentare
    cm = get_comments_state()
    cm["Funktionsmatrix"] = st.text_area("Beschreibung / Kommentare – Funktionsmatrix", value=cm.get("Funktionsmatrix",""), height=140)
    export_comments_button()

# -------------------- Funktionen­kosten --------------------
with tab_costs:
    st.subheader("Funktionen­kosten")

    prod_names = [p["Produkt"] for p in products]
    psel = st.selectbox("Produkt wählen", prod_names, key="cost_prod")
    P = next(p for p in products if p["Produkt"] == psel)
    h1_costs = P["h1_costs"]
    h2_costs = P["h2_costs"]

    # H1-Kosten (Zeile 7)
    chart_h1 = (
        alt.Chart(h1_costs)
        .mark_bar(size=18)
        .encode(
            x=alt.X(f"{H1_LABEL}:N", sort=list(h1_costs[H1_LABEL]), axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("Kosten_H1:Q", axis=alt.Axis(title="Kosten [€]")),
            tooltip=[H1_LABEL, alt.Tooltip("Kosten_H1:Q", format=".2f")]
        )
        .properties(height=280, title="Kosten je Hauptfunktion (Zeile 7)")
    )
    st.altair_chart(chart_h1, use_container_width=True)

    # H2-Drilldown
    h1_sel = st.selectbox("Hauptfunktion wählen", list(h1_costs[H1_LABEL]), key="h1_dd")
    h2_df = h2_costs[h2_costs[H1_LABEL] == h1_sel][[H2_LABEL, "Kosten_H2"]]

    chart_h2 = (
        alt.Chart(h2_df)
        .mark_bar(size=16)
        .encode(
            x=alt.X(f"{H2_LABEL}:N", axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("Kosten_H2:Q", axis=alt.Axis(title="Kosten [€]")),
            tooltip=[H2_LABEL, alt.Tooltip("Kosten_H2:Q", format=".2f")]
        )
        .properties(height=280, title=f"{h1_sel} – Kosten je Nebenfunktion (Zeile 8)")
    )
    st.altair_chart(chart_h2, use_container_width=True)

    cm = get_comments_state()
    cm["Funktionen­kosten"] = st.text_area("Beschreibung / Kommentare – Funktionen­kosten", value=cm.get("Funktionen­kosten",""), height=140)
    export_comments_button()

# -------------------- Technik Bewertung (Nebenfunktionen) --------------------
with tab_tech:
    st.subheader("Technische Bewertung — Nebenfunktionen (B → R)")
    # Long DF über alle Produkte
    frames = []
    for p in products:
        t = p["tech"].copy()
        t["Produkt"] = p["Produkt"]
        frames.append(t)
    tech_long = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=[H2_LABEL, "Score", "Produkt"])

    # Reihenfolge Nebenfunktionen: nach der ersten Datei mit Tech-Tab
    h2_order = None
    for p in products:
        if not p["tech"].empty:
            h2_order = list(p["tech"][H2_LABEL])
            break
    if h2_order is None:
        h2_order = list(tech_long[H2_LABEL].unique())

    line = (
        alt.Chart(tech_long)
        .mark_line(point=True, strokeWidth=2)
        .encode(
            x=alt.X(f"{H2_LABEL}:N", sort=h2_order, axis=alt.Axis(labelAngle=0, title=None)),
            y=alt.Y("Score:Q", axis=alt.Axis(title="Technischer Score")),
            color=alt.Color("Produkt:N", legend=alt.Legend(title="Produkt")),
            tooltip=["Produkt", H2_LABEL, alt.Tooltip("Score:Q", format=".3f")]
        )
        .properties(height=340, title="Technischer Score je Nebenfunktion — alle Produkte")
        .interactive()
    )
    st.altair_chart(line, use_container_width=True)

    cm = get_comments_state()
    cm["Technik Bewertung (Nebenfunktionen)"] = st.text_area("Beschreibung / Kommentare – Technik Bewertung (Nebenfunktionen)", value=cm.get("Technik Bewertung (Nebenfunktionen)",""), height=140)
    export_comments_button()

# -------------------- Top Kostenabweichung --------------------
with tab_top:
    st.subheader("Top Kostenabweichung (Nebenfunktionen) — Paarweiser Vergleich")
    if len(products) < 2:
        st.info("Bitte mindestens zwei Produkte hochladen.")
    else:
        colA, colB = st.columns(2)
        with colA:
            a_name = st.selectbox("Produkt A", [p["Produkt"] for p in products], key="A_sel")
        with colB:
            b_name = st.selectbox("Produkt B", [p["Produkt"] for p in products if p["Produkt"] != a_name], key="B_sel")

        pa = next(p for p in products if p["Produkt"] == a_name)
        pb = next(p for p in products if p["Produkt"] == b_name)

        A = pa["h2_costs"][[H1_LABEL, H2_LABEL, "Kosten_H2"]].rename(columns={"Kosten_H2": "A"})
        B = pb["h2_costs"][[H1_LABEL, H2_LABEL, "Kosten_H2"]].rename(columns={"Kosten_H2": "B"})

        merged = pd.merge(A, B, on=[H1_LABEL, H2_LABEL], how="outer").fillna(0.0)
        merged["Delta (B-A)"] = merged["B"] - merged["A"]
        merged["absDelta"] = merged["Delta (B-A)"].abs()
        top = merged.sort_values("absDelta", ascending=False).head(10)

        st.dataframe(top[[H1_LABEL, H2_LABEL, "A", "B", "Delta (B-A)"]], use_container_width=True)

        chart = (
            alt.Chart(top.reset_index(drop=True))
            .mark_bar(size=18)
            .encode(
                x=alt.X(f"{H2_LABEL}:N", sort=list(top[H2_LABEL]), title="Nebenfunktion"),
                y=alt.Y("Delta (B-A):Q", title="Abweichung (B - A) [€]"),
                color=alt.condition(alt.datum["Delta (B-A)"] > 0, alt.value("#1f77b4"), alt.value("#d62728")),
                tooltip=[H1_LABEL, H2_LABEL, alt.Tooltip("A:Q", format=".2f"), alt.Tooltip("B:Q", format=".2f"), alt.Tooltip("Delta (B-A):Q", format=".2f")]
            )
            .properties(height=320, title=f"Top 10 Abweichungen: {b_name} vs {a_name}")
        )
        st.altair_chart(chart, use_container_width=True)

    cm = get_comments_state()
    cm["Top Kostenabweichung"] = st.text_area("Beschreibung / Kommentare – Top Kostenabweichung", value=cm.get("Top Kostenabweichung",""), height=140)
    export_comments_button()

st.markdown('<div style="opacity:.6;margin-top:.75rem;">© EFESO — v10.2</div>', unsafe_allow_html=True)
