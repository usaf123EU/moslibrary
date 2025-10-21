
import io
from typing import Dict, Tuple, List
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------
# Page config
# ----------------------------
st.set_page_config(
    page_title="EFESO Functional Cost Analysis TOOLSET — v10.3",
    layout="wide",
)

st.title("EFESO Functional Cost Analysis TOOLSET — v10.3")
st.caption("Funktionsmatrix · Kosten · Technik · Abweichungen (border-only swimlanes)")


# ----------------------------
# VISUAL CONSTANTS (thin bars)
# ----------------------------
BARGAP_H1 = 0.45   # thinner bars main chart
BARGAP_H2 = 0.55   # thinner bars drilldowns/tech
LINE_HEIGHT = 380


# ----------------------------
# Excel layout constants
# ----------------------------
ROW_H1_NAME   = 0   # Zeile 1
ROW_H2_NAME   = 1   # Zeile 2
ROW_H1_WEIGHT = 3   # Zeile 4
ROW_H2_WEIGHT = 4   # Zeile 5
ROW_H1_COST   = 6   # Zeile 7
ROW_H2_COST   = 7   # Zeile 8
COL_START     = 8   # Spalte I (0-indexed)

SHEET_COST = "SLAVE_Funktions-Kostenstruktur"
SHEET_TECH = "SLAVE_Techn.Bewertung"


# ----------------------------
# Utilities
# ----------------------------
def _to_float(x):
    try:
        s = str(x).replace("%", "").replace(",", ".").strip()
        return float(s) if s not in ("", "nan", "None") else np.nan
    except Exception:
        return np.nan

def read_excel_bytes(upload) -> Dict[str, pd.DataFrame]:
    x = pd.ExcelFile(upload, engine="openpyxl")
    out = {}
    for name in x.sheet_names:
        out[name] = x.parse(name, header=None)
    return out

def parse_cost_structure(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    df = df.fillna("")
    ncols = df.shape[1]
    col = COL_START
    h1_rows, h2_rows = [], []

    while col < ncols:
        h1 = str(df.iat[ROW_H1_NAME, col]).strip()
        if h1 == "" or h1.lower() == "nan":
            col += 1
            continue

        start_col = col
        col += 1
        while col < ncols and str(df.iat[ROW_H1_NAME, col]).strip() == "":
            col += 1
        end_col = col  # exclusive

        h1_weight = _to_float(df.iat[ROW_H1_WEIGHT, start_col])
        h1_cost   = _to_float(df.iat[ROW_H1_COST,   start_col])
        h1_rows.append([h1, h1_weight, h1_cost])

        for c in range(start_col, end_col):
            h2 = str(df.iat[ROW_H2_NAME, c]).strip()
            if not h2 or h2.lower() == "nan": 
                continue
            h2_weight = _to_float(df.iat[ROW_H2_WEIGHT, c])
            h2_cost   = _to_float(df.iat[ROW_H2_COST,   c])
            h2_rows.append([h1, h2, h2_weight, h2_cost])

    h1_df = pd.DataFrame(h1_rows, columns=["H1", "H1Weight_%", "H1Cost"])
    h2_df = pd.DataFrame(h2_rows, columns=["H1", "H2", "H2Weight_%", "H2Cost"])
    for c in ["H1Weight_%", "H1Cost"]:
        h1_df[c] = pd.to_numeric(h1_df[c], errors="coerce").fillna(0.0)
    for c in ["H2Weight_%", "H2Cost"]:
        h2_df[c] = pd.to_numeric(h2_df[c], errors="coerce").fillna(0.0)
    return h1_df, h2_df

def parse_tech_scores(df: pd.DataFrame) -> pd.DataFrame:
    # Spalte B (index=1) -> H2, Spalte R (index=17) -> Score
    df = df.fillna("")
    rows = []
    for r in range(df.shape[0]):
        h2 = str(df.iat[r, 1]).strip()
        if not h2 or h2.lower() == "nan": 
            continue
        score = _to_float(df.iat[r, 17])
        if pd.isna(score): 
            continue
        rows.append([h2, float(score)])
    return pd.DataFrame(rows, columns=["H2", "TechScore"])


# ----------------------------
# Upload
# ----------------------------
st.subheader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei")
uploads = st.file_uploader(
    "Drag and drop files here",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True,
    label_visibility="collapsed",
)

if not uploads:
    st.info("Bitte eine oder mehrere Produkt-Dateien hochladen.")
    st.stop()

products, errors = {}, []
for up in uploads:
    try:
        wb = read_excel_bytes(up)
        if SHEET_COST not in wb:
            errors.append(f"**{up.name}**: Blatt `{SHEET_COST}` fehlt.")
            continue
        h1_df, h2_df = parse_cost_structure(wb[SHEET_COST])
        tech_df = parse_tech_scores(wb[SHEET_TECH]) if SHEET_TECH in wb else pd.DataFrame(columns=["H2","TechScore"])
        products[up.name] = {"h1": h1_df, "h2": h2_df, "tech": tech_df}
    except Exception as e:
        errors.append(f"**{up.name}**: {e}")

if errors:
    st.warning("Einige Dateien hatten Probleme:\\n\\n- " + "\\n- ".join(errors))
if not products:
    st.error("Keine gültigen Daten.")
    st.stop()

product_names = list(products.keys())


# ----------------------------
# Tabs
# ----------------------------
tab_matrix, tab_cost, tab_tech, tab_top = st.tabs(
    ["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"]
)

# =============== Funktionsmatrix (border-only swimlanes) ===============
with tab_matrix:
    st.markdown("### Funktionsmatrix (H1 → zugehörige Nebenfunktionen)")
    pname = st.selectbox("Produkt wählen", product_names, key="mx_prod_v103")
    h1_df = products[pname]["h1"].copy()
    h2_df = products[pname]["h2"].copy()

    # Sort H1 by sum of H2 weights desc
    order = (
        h2_df.groupby("H1")["H2Weight_%"].sum()
        .sort_values(ascending=False).index.tolist()
    )
    h1_df["H1"] = pd.Categorical(h1_df["H1"], categories=order, ordered=True)
    h1_df = h1_df.sort_values("H1")

    TILE_HTML = \"\"\"
<div style="background: transparent; border: 1px solid #D0D0D0; border-radius: 10px; padding: 10px 12px; margin: 6px 0;">
  <div style="font-weight: 600; font-size: 14px;">{h2}</div>
  <div style="font-size: 12px; color: #555;">
     Anteil: <b>{w:.0f}%</b> &nbsp;|&nbsp; Kosten: <b>€ {c:.2f}</b>
  </div>
</div>
\"\"\"

    for _, h1row in h1_df.iterrows():
        h1 = h1row["H1"]
        sub = h2_df[h2_df["H1"] == h1].copy().sort_values("H2Weight_%", ascending=False)
        st.markdown(f"**{h1}**  —  Gewichtung: {h1row['H1Weight_%']:.0f}%  |  Kosten: € {h1row['H1Cost']:.2f}")

        cols = st.columns(max(1, min(6, len(sub))))
        i = 0
        for _, r2 in sub.iterrows():
            with cols[i % len(cols)]:
                st.markdown(TILE_HTML.format(h2=r2["H2"], w=r2["H2Weight_%"], c=r2["H2Cost"]), unsafe_allow_html=True)
            i += 1
    st.divider()

    st.markdown("**Beschreibung / Kommentare – Funktionsmatrix**")
    st.text_area("Kommentar", key="mx_comment_v103", height=110, label_visibility="collapsed")


# =============== Funktionenkosten ===============
with tab_cost:
    st.markdown("### Kosten je Hauptfunktion (Zeile 7)")
    pname = st.selectbox("Produkt wählen", product_names, key="cost_prod_v103")
    h1_df = products[pname]["h1"].copy()
    h2_df = products[pname]["h2"].copy()

    fig = px.bar(
        h1_df.sort_values("H1Cost", ascending=False),
        x="H1", y="H1Cost", text="H1Cost",
    )
    fig.update_traces(marker_color="#1261A0", texttemplate="€ %{y:.2f}", textposition="outside")
    fig.update_layout(
        bargap=BARGAP_H1,
        yaxis_title="Kosten (€)", xaxis_title="Hauptfunktion",
        height=420, margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Drilldown: Kosten je Nebenfunktion (Zeile 8)")
    h1_opt = st.selectbox("Hauptfunktion wählen", h1_df["H1"].tolist(), key="cost_h1_sel_v103")
    sub = h2_df[h2_df["H1"] == h1_opt].copy().sort_values("H2Cost", ascending=False)
    if sub.empty:
        st.info("Keine Nebenfunktionen gefunden.")
    else:
        fig2 = px.bar(sub, x="H2", y="H2Cost", text="H2Cost")
        fig2.update_traces(marker_color="#1f77b4", texttemplate="€ %{y:.2f}", textposition="outside")
        fig2.update_layout(
            bargap=BARGAP_H2,
            yaxis_title="Kosten (€)", xaxis_title=f"Nebenfunktionen – {h1_opt}",
            height=380, margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(fig2, use_container_width=True)

    st.markdown("**Beschreibung / Kommentare – Funktionenkosten**")
    st.text_area("Kommentar", key="cost_comment_v103", height=110, label_visibility="collapsed")


# =============== Technik Bewertung ===============
with tab_tech:
    st.markdown("### Technische Bewertung — Nebenfunktionen (B → R)")
    pname = st.selectbox("Produkt wählen", product_names, key="tech_prod_v103")
    h2_df = products[pname]["h2"]
    tech_df = products[pname]["tech"]

    merged = h2_df.merge(tech_df, on="H2", how="left")
    merged["TechScore"] = merged["TechScore"].fillna(0.0)

    fig = px.bar(
        merged.sort_values("TechScore", ascending=False),
        x="H2", y="TechScore", text="TechScore",
    )
    fig.update_traces(marker_color="#5b8e7d", texttemplate="%{y:.2f}", textposition="outside")
    fig.update_layout(
        bargap=BARGAP_H2,
        yaxis_title="Score", xaxis_title="Nebenfunktion",
        height=420, margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("#### Linienvergleich – TechScore je H2 (alle Produkte)")
    # Union of all H2
    all_h2 = sorted(set(pd.concat([products[p]["h2"]["H2"] for p in product_names]).unique()))
    line_df = []
    for p in product_names:
        tmp = products[p]["h2"][ ["H2"] ].merge(products[p]["tech"], on="H2", how="left")
        tmp["TechScore"] = tmp["TechScore"].fillna(0.0)
        tmp = tmp.set_index("H2").reindex(all_h2).reset_index()
        tmp["Produkt"] = p
        line_df.append(tmp)
    line_df = pd.concat(line_df, ignore_index=True)

    figl = px.line(line_df, x="H2", y="TechScore", color="Produkt", markers=True)
    figl.update_layout(
        yaxis_title="Score", xaxis_title="Nebenfunktion",
        height=LINE_HEIGHT, margin=dict(l=10, r=10, t=10, b=10),
    )
    st.plotly_chart(figl, use_container_width=True)

    st.markdown("**Beschreibung / Kommentare – Technik**")
    st.text_area("Kommentar", key="tech_comment_v103", height=110, label_visibility="collapsed")


# =============== Top Kostenabweichung ===============
with tab_top:
    st.markdown("### Top 10 Kostenabweichungen – Nebenfunktionen")
    if len(product_names) < 2:
        st.info("Bitte mindestens **zwei** Produkte hochladen.")
    else:
        colA, colB = st.columns(2)
        with colA:
            pA = st.selectbox("Produkt A", product_names, key="cmp_A_v103")
        with colB:
            pB = st.selectbox("Produkt B", product_names, key="cmp_B_v103")

        A = products[pA]["h2"][ ["H1","H2","H2Cost"] ].rename(columns={"H2Cost":"Cost_A"})
        B = products[pB]["h2"][ ["H1","H2","H2Cost"] ].rename(columns={"H2Cost":"Cost_B"})
        cmp = A.merge(B, on=["H1","H2"], how="outer").fillna(0.0)
        cmp["Delta"] = cmp["Cost_B"] - cmp["Cost_A"]

        top10 = cmp.reindex(cmp["Delta"].abs().sort_values(ascending=False).index)[:10]

        figd = go.Figure()
        figd.add_trace(go.Bar(name=pA, x=top10["H2"], y=top10["Cost_A"]))
        figd.add_trace(go.Bar(name=pB, x=top10["H2"], y=top10["Cost_B"]))
        figd.update_layout(
            barmode="group", bargap=0.35,
            yaxis_title="Kosten (€)", xaxis_title="Nebenfunktion",
            height=420, margin=dict(l=10, r=10, t=10, b=10),
        )
        st.plotly_chart(figd, use_container_width=True)

        st.dataframe(top10.reset_index(drop=True), use_container_width=True)

        st.markdown("**Beschreibung / Kommentare – Top Abweichungen**")
        st.text_area("Kommentar", key="diff_comment_v103", height=110, label_visibility="collapsed")
