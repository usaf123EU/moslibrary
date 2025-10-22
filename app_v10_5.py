
# app_v10_5.py
# EFESO – Functional Cost Analysis TOOLSET (v10.5 • full)
# - Robust uploader (always visible)
# - Parser for new Excel structure (rows 1/2/4/5/7/8 from column I)
# - Swimlanes (no fills, borders only) with yellow H1 badges (row 4)
# - Funktionenkosten (H1 bars; H2 drilldown colored by H1)
# - Technik Bewertung (Tech lines; Costs lines under it)
# - Top Kostenabweichung (bar + ranked table)
#
# Run: streamlit run app_v10_5.py

import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

st.set_page_config(page_title="EFESO – Functional Cost Analysis TOOLSET", layout="wide")
VERSION = "v10.5"

# ---------------- Helpers ----------------
def _to_num(x):
    if pd.isna(x): return np.nan
    if isinstance(x, (int, float, np.number)): return float(x)
    s = str(x).strip().replace(",", ".")
    s = re.sub(r"[^0-9.\-]", "", s)
    try:
        return float(s)
    except:
        return np.nan

def _to_pct(x):
    if pd.isna(x): return np.nan
    s = str(x).strip().replace(",", ".")
    if s.endswith("%"):
        s = s[:-1]
        return _to_num(s) / 100.0
    v = _to_num(s)
    if pd.isna(v): return np.nan
    return v/100.0 if v>1 else v

def find_sheet(xls, must_all):
    for name in xls.sheet_names:
        low = name.lower()
        if all(k in low for k in must_all):
            return name
    return None

def parse_cost_structure(xls):
    sheet = find_sheet(xls, ["funktions", "kosten"]) or xls.sheet_names[0]
    df = xls.parse(sheet, header=None, dtype=str).fillna("")
    # Excel rows -> 0-based indices:
    ROW_H1, ROW_H2 = 0, 1           # names
    ROW_W1, ROW_W2 = 3, 4           # weights
    ROW_C1, ROW_C2 = 6, 7           # costs
    START_COL = 8                   # column I (0-based)

    # find H1 block starts
    starts = []
    last = None
    for c in range(START_COL, min(500, df.shape[1])):
        lab = str(df.iat[ROW_H1, c]).strip()
        if lab:
            if lab != last:
                starts.append(c)
            last = lab
        else:
            last = None
    if not starts:
        return pd.DataFrame(columns=["H1","H1Weight","H1Cost"]), pd.DataFrame(columns=["H1","H2","H2Weight","H2Cost"])

    blocks = []
    for i,s in enumerate(starts):
        e = (starts[i+1]-1) if i+1 < len(starts) else df.shape[1]-1
        blocks.append((s,e))

    h1_rows, h2_rows = [], []
    for s,e in blocks:
        h1 = str(df.iat[ROW_H1, s]).strip()
        if not h1: continue
        h1_w = _to_pct(df.iat[ROW_W1, s])
        h1_c = _to_num(df.iat[ROW_C1, s])
        h1_rows.append([h1, h1_w, h1_c])
        for c in range(s, e+1):
            h2 = str(df.iat[ROW_H2, c]).strip()
            if not h2: continue
            h2_w = _to_pct(df.iat[ROW_W2, c])
            h2_c = _to_num(df.iat[ROW_C2, c])
            h2_rows.append([h1, h2, h2_w, h2_c])

    H1 = pd.DataFrame(h1_rows, columns=["H1","H1Weight","H1Cost"])
    H2 = pd.DataFrame(h2_rows, columns=["H1","H2","H2Weight","H2Cost"])
    if not H2.empty:
        H2 = H2.groupby(["H1","H2"], as_index=False).agg({"H2Weight":"max","H2Cost":"max"})
    return H1, H2

def parse_tech(xls):
    sheet = find_sheet(xls, ["techn", "bewert"])
    if sheet is None:
        return pd.DataFrame(columns=["H2","TechScore"])
    df = xls.parse(sheet, header=None, dtype=str).fillna("")
    B, R = (1 if df.shape[1]>1 else None), (17 if df.shape[1]>17 else None)
    rows = []
    for i in range(df.shape[0]):
        h2 = str(df.iat[i,B]).strip() if B is not None else ""
        sc = df.iat[i,R] if R is not None else ""
        if h2:
            v = _to_num(sc)
            if not pd.isna(v):
                rows.append([h2, v])
    return pd.DataFrame(rows, columns=["H2","TechScore"])

class Product:
    def __init__(self, name, H1, H2, TECH):
        self.name=name; self.H1=H1; self.H2=H2; self.TECH=TECH

# ---------------- UI ----------------
st.markdown("# EFESO – Functional Cost Analysis TOOLSET")
st.caption(f"Version {VERSION} • Vorlage für Funktions- & Kostenanalyse")

files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei",
                         type=["xlsx","xlsm"], accept_multiple_files=True)
if not files:
    st.info("Bitte laden Sie eine oder mehrere Excel-Dateien hoch.")
    st.stop()

products = {}
errors = []
for f in files:
    try:
        xls = pd.ExcelFile(f)
        H1,H2 = parse_cost_structure(xls)
        TECH = parse_tech(xls)
        # attach tech to H2
        if not TECH.empty and not H2.empty:
            H2 = H2.merge(TECH, on="H2", how="left")
        products[f.name] = Product(f.name, H1, H2, TECH)
    except Exception as e:
        errors.append(f"{f.name}: {e}")

if errors:
    with st.expander("Parsing-Hinweise", expanded=True):
        for m in errors: st.error(m)
if not products:
    st.stop()

names = list(products.keys())

tab1, tab2, tab3, tab4 = st.tabs(["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"])

# ---------------- Tab 1: Funktionsmatrix ----------------
with tab1:
    sel = st.selectbox("Produkt wählen", names, index=0)
    P = products[sel]
    H1, H2 = P.H1.copy(), P.H2.copy()
    st.caption("Kacheln mit Rahmen (ohne Füllfarbe). Gelbes Badge = H1-Gewichtung (Zeile 4). Rechts in jeder H2-Kachel: H2-Gewichtung (Zeile 5).")

    if H1.empty:
        st.info("Keine Hauptfunktionen gefunden.")
    else:
        cols = 3
        H1 = H1.reset_index(drop=True)
        for i in range(len(H1)):
            if i%cols==0: cols_list = st.columns(cols, gap="large")
            col = cols_list[i%cols]
            row = H1.iloc[i]
            h1, w1 = row["H1"], row["H1Weight"]
            w1s = "" if pd.isna(w1) else f"{int(round(w1*100))}%"
            with col:
                st.markdown(
                    f"<div style='border:1px solid #E0E0E0;border-radius:10px;padding:10px 12px;margin-bottom:10px;'>"
                    f"<div style='display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;'>"
                    f"<div style='font-weight:700'>{h1}</div>"
                    f"<div style='background:#FFD24D;color:#333;border:1px solid #CCAA00;border-radius:999px;padding:2px 8px;font-weight:700'>{w1s}</div>"
                    f"</div>", unsafe_allow_html=True
                )
                sub = H2[H2["H1"].eq(h1)].copy()
                if sub.empty:
                    st.markdown("<div style='color:#777'>–</div>", unsafe_allow_html=True)
                else:
                    for _,r in sub.iterrows():
                        h2 = r["H2"]
                        w2 = r["H2Weight"]; w2s = "–" if pd.isna(w2) else f"{int(round(w2*100))}%"
                        st.markdown(
                            f"<div style='border:1px solid #E6E6E6;border-radius:8px;padding:6px 8px;margin:6px 0;display:flex;justify-content:space-between;'>"
                            f"<span>{h2}</span><span style='color:#666'>{w2s}</span></div>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

# ---------------- Tab 2: Funktionenkosten ----------------
with tab2:
    sel2 = st.selectbox("Produkt wählen ", names, index=0, key="costprod")
    P = products[sel2]; H1,H2 = P.H1.copy(), P.H2.copy()

    st.subheader("Kosten je Hauptfunktion (Zeile 7)")
    if H1.empty:
        st.info("Keine H1-Kosten vorhanden.")
    else:
        fig = go.Figure(go.Bar(x=H1["H1"], y=H1["H1Cost"], marker_color="#1F5AA6", width=0.35))
        fig.update_layout(height=340, margin=dict(l=20,r=20,t=20,b=80), yaxis_title="Kosten Hauptfunktion")
        st.plotly_chart(fig, use_container_width=True)

    st.subheader("Drilldown: Kosten je Nebenfunktion (Zeile 8)")
    if H2.empty:
        st.info("Keine H2-Kosten vorhanden.")
    else:
        # color by H1
        h1_list = H1["H1"].tolist()
        cmap_colors = ["#1F5AA6","#F28C28","#0B3C7A","#FFB347","#8A8A8A","#D46A00","#B3B3B3"]
        cmap = {h: cmap_colors[i%len(cmap_colors)] for i,h in enumerate(h1_list)}
        H2c = H2.copy()
        H2c["Color"] = H2c["H1"].map(cmap)
        fig2 = go.Figure(go.Bar(x=H2c["H2"], y=H2c["H2Cost"], marker_color=H2c["Color"], width=0.35))
        fig2.update_layout(height=420, margin=dict(l=20,r=20,t=10,b=140), yaxis_title="Kosten Nebenfunktion")
        fig2.update_xaxes(tickangle=45)
        st.plotly_chart(fig2, use_container_width=True)

# ---------------- Tab 3: Technik Bewertung ----------------
with tab3:
    st.subheader("Technische Bewertung – Nebenfunktionen (H2)")
    # union axis
    all_h2 = sorted(set(h2 for P in products.values() for h2 in P.H2["H2"].astype(str).tolist()))
    if not all_h2:
        st.info("Keine H2 gefunden.")
    else:
        figt = go.Figure()
        palette = ["#1F5AA6","#F28C28","#0B3C7A","#FFB347","#8A8A8A","#D46A00","#B3B3B3","#6AA6FF","#FF8C66"]
        for i,(n,P) in enumerate(products.items()):
            ser = dict(P.H2[["H2","TechScore"]].dropna().values)
            y = [ser.get(h2, np.nan) for h2 in all_h2]
            figt.add_scatter(x=all_h2, y=y, mode="lines+markers", name=n, line=dict(color=palette[i%len(palette)], width=2))
        figt.update_layout(height=380, margin=dict(l=20,r=20,t=10,b=160), yaxis_title="TechScore")
        figt.update_xaxes(tickangle=45)
        st.plotly_chart(figt, use_container_width=True)

        st.subheader("Kosten (H2) – alle Produkte (Linien)")
        figk = go.Figure()
        for i,(n,P) in enumerate(products.items()):
            ser = dict(P.H2[["H2","H2Cost"]].dropna().values)
            y = [ser.get(h2, np.nan) for h2 in all_h2]
            figk.add_scatter(x=all_h2, y=y, mode="lines+markers", name=n, line=dict(color=palette[i%len(palette)], width=2, dash="dot"))
        figk.update_layout(height=360, margin=dict(l=20,r=20,t=10,b=160), yaxis_title="Kosten (H2)")
        figk.update_xaxes(tickangle=45)
        st.plotly_chart(figk, use_container_width=True)

# ---------------- Tab 4: Top Kostenabweichung ----------------
with tab4:
    st.subheader("Top Kostenabweichungen – Nebenfunktionen (H2)")
    a = st.selectbox("Produkt A", names, index=0)
    b = st.selectbox("Produkt B", names, index=min(1,len(names)-1))
    if a == b:
        st.info("Bitte zwei unterschiedliche Produkte wählen.")
    else:
        A, B = products[a], products[b]
        sA = A.H2.set_index("H2")["H2Cost"]
        sB = B.H2.set_index("H2")["H2Cost"]
        idx = sorted(set(sA.index)|set(sB.index))
        rows = []
        for key in idx:
            ca = float(sA.get(key, np.nan)) if key in sA else np.nan
            cb = float(sB.get(key, np.nan)) if key in sB else np.nan
            if np.isnan(ca) and np.isnan(cb): continue
            delta = (0 if np.isnan(ca) else ca) - (0 if np.isnan(cb) else cb)
            rows.append([key, ca, cb, delta, abs(delta)])
        dd = pd.DataFrame(rows, columns=["H2","Cost_A","Cost_B","Delta","AbsDelta"]).sort_values("AbsDelta", ascending=False)
        top10 = dd.head(10)
        figd = go.Figure(go.Bar(x=top10["H2"], y=top10["AbsDelta"], marker_color="#1F5AA6", width=0.35))
        figd.update_layout(height=380, margin=dict(l=20,r=20,t=10,b=160), yaxis_title="|Delta|")
        figd.update_xaxes(tickangle=45)
        st.plotly_chart(figd, use_container_width=True)
        st.markdown("**Ranking – größte Abweichungen (H2)**")
        st.dataframe(dd[["H2","Cost_A","Cost_B","Delta"]].reset_index(drop=True), use_container_width=True)

st.caption(f"© EFESO • Version {VERSION} • Vorlage für Funktions- & Kostenanalyse")
