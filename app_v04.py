import io, re, os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st

PRIMARY = "#f47c24"; DARK = "#2a2a2a"

st.set_page_config(page_title="FKA v04 — EFESO Style (robustes Logo + Kostenspalten)", layout="wide")

st.markdown(r'''
<style>
.efeso-header {background:#f47c24;padding:10px 16px;border-radius:10px;display:flex;align-items:center;gap:16px;margin-bottom:12px;}
.efeso-title {color:white;font-size:22px;font-weight:700;margin:0;}
.efeso-sub {color:#ffe9d6;font-size:13px;margin:0;}
.efeso-section-title {color:#2a2a2a;border-left:6px solid #f47c24;padding-left:10px;font-weight:700;margin:8px 0 6px 0;font-size:18px;}
.stTabs [data-baseweb="tab"] { font-size:15px;font-weight:600;color:#2a2a2a;}
.stTabs [aria-selected="true"] { color:#f47c24; }
</style>
''', unsafe_allow_html=True)

def get_placeholder_logo_path():
    try:
        base = Path(__file__).resolve().parent
    except Exception:
        base = Path(".").resolve()
    candidate = base / "assets" / "efeso_logo_placeholder.png"
    return str(candidate) if candidate.exists() else None

col_logo, col_head = st.columns([1,6])
with col_logo:
    st.caption("Logo")
    uploaded_logo = st.file_uploader("Logo (optional)", type=["png","jpg","jpeg"], label_visibility="collapsed", key="logo")
    if uploaded_logo is not None:
        st.image(uploaded_logo, use_column_width=True)
    else:
        ph = get_placeholder_logo_path()
        if ph:
            st.image(ph, use_column_width=True)
        else:
            st.markdown(f"<div style='color:{PRIMARY};font-weight:700;font-size:22px;'>EFESO</div>", unsafe_allow_html=True)
            st.caption("MANAGEMENT CONSULTANTS")

with col_head:
    st.markdown('<div class="efeso-header"><div><p class="efeso-title">FKA — Funktionskosten & Evaluierung</p><p class="efeso-sub">v04 · EFESO-Farben · robustes Logo & präzise Kostenspalten</p></div></div>', unsafe_allow_html=True)

files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)

FUNC_SHEET_HINTS = ["SLAVE_Funktions", "Funktion", "Function", "Kosten"]
TECH_SHEET_HINTS = ["Techn", "Bewertung", "Rating", "Score"]
START_SHEET_HINTS = ["SLAVE_START", "Start", "START"]

def _pick_sheet(xl, hints):
    names = xl.sheet_names
    for n in names:
        for h in hints:
            if n.strip().lower() == h.strip().lower():
                return n
    for n in names:
        for h in hints:
            if h.lower() in n.lower():
                return n
    return names[0]

def detect_cost_columns(df):
    header_zone = df.head(8).astype(str)
    euro_cols = [c for c in df.columns if header_zone[c].str.contains("€").any()]
    parsed_numeric_cols = []
    for c in df.columns:
        if c in euro_cols: 
            parsed_numeric_cols.append(c); 
            continue
        s = pd.to_numeric(df[c].astype(str).str.replace(".","", regex=False).str.replace(",",".", regex=False), errors="coerce")
        if s.notna().sum() >= max(3, int(0.5*len(s))):
            if not header_zone[c].str.contains(r"[A-Za-z]", regex=True).any() or header_zone[c].str.contains("€").any():
                parsed_numeric_cols.append(c)
    cost_cols = list(dict.fromkeys(euro_cols + parsed_numeric_cols))
    return cost_cols

def clean_numeric(series):
    if pd.api.types.is_numeric_dtype(series): 
        return series
    return pd.to_numeric(series.astype(str).str.replace(".","", regex=False).str.replace(",",".", regex=False), errors="coerce")

def rollup_costs(df, hier_cols, cost_cols):
    tmp = df.copy()
    for c in cost_cols:
        tmp[c] = clean_numeric(tmp[c])
    tmp["__row_cost__"] = tmp[cost_cols].sum(axis=1, skipna=True)
    res = {}
    if len(hier_cols)>=1:
        g1 = tmp.groupby([hier_cols[0]], dropna=False)["__row_cost__"].sum().reset_index()
        res["H1"] = g1.rename(columns={hier_cols[0]:"H1","__row_cost__":"TotalCost"})
    if len(hier_cols)>=2:
        g2 = tmp.groupby([hier_cols[0],hier_cols[1]], dropna=False)["__row_cost__"].sum().reset_index()
        res["H2"] = g2.rename(columns={hier_cols[0]:"H1",hier_cols[1]:"H2","__row_cost__":"TotalCost"})
    if len(hier_cols)>=3:
        g3 = tmp.groupby([hier_cols[0],hier_cols[1],hier_cols[2]], dropna=False)["__row_cost__"].sum().reset_index()
        res["H3"] = g3.rename(columns={hier_cols[0]:"H1",hier_cols[1]:"H2",hier_cols[2]:"H3","__row_cost__":"TotalCost"})
    return res

def parse_tech(df_raw):
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    crit_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ["kriter","criterion","metric"]):
            crit_col = c; break
    cat_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ["kategorie","category","group"]):
            cat_col = c; break
    weight_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ["gewicht","weight","wgt"]):
            weight_col = c; break
    score_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ["score","wert","points"]):
            if pd.api.types.is_numeric_dtype(df[c]):
                score_col = c; break
    if score_col is None:
        nums = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        score_col = nums[-1] if nums else None
    if crit_col is None:
        crit_col = "_Criterion"; df[crit_col] = np.arange(1, len(df)+1)
    if weight_col is None:
        weight_col = "_Weight"; df[weight_col] = 1.0
    else:
        df[weight_col] = pd.to_numeric(df[weight_col], errors="coerce").fillna(1.0)
    if score_col is None:
        df["_score"] = 0.0; score_col = "_score"
    else:
        df[score_col] = pd.to_numeric(df[score_col], errors="coerce").fillna(0.0)
    df["_Weighted"] = df[score_col]*df[weight_col]
    wsum = df[weight_col].sum() or 1.0
    overall = df["_Weighted"].sum()/wsum
    by_cat = None
    if cat_col:
        by_cat = df.groupby(cat_col, dropna=False).apply(
            lambda g: pd.Series({
                "Weight": g[weight_col].sum(),
                "WeightedScore": g["_Weighted"].sum(),
                "ScoreAvg": (g["_Weighted"].sum()/(g[weight_col].sum() or 1.0))
            })
        ).reset_index()
    return {"crit_col":crit_col,"weight_col":weight_col,"score_col":score_col,"overall":overall,"by_cat":by_cat,"table":df[[crit_col,weight_col,score_col,"_Weighted"]]}

def read_product(file):
    name = file.name.rsplit(".",1)[0]
    xl = pd.ExcelFile(file)
    func_sheet = _pick_sheet(xl, FUNC_SHEET_HINTS)
    tech_sheet = _pick_sheet(xl, TECH_SHEET_HINTS)
    start_sheet = _pick_sheet(xl, START_SHEET_HINTS)
    func_df = xl.parse(func_sheet)
    tech_df = xl.parse(tech_sheet)
    start_df = xl.parse(start_sheet)

    text_cols = [c for c in func_df.columns if not pd.api.types.is_numeric_dtype(func_df[c])]
    hier_cols = []
    for key in ["H1","H2","H3","Hauptfunktion","Teilfunktion","Unterfunktion"]:
        for c in func_df.columns:
            if key.lower() in str(c).lower() and c not in hier_cols: hier_cols.append(c)
    for c in func_df.columns:
        if c not in hier_cols and c in text_cols and len(hier_cols)<3: hier_cols.append(c)

    candidates = detect_cost_columns(func_df)
    st.session_state.setdefault("cost_map", {})
    with st.expander(f"⚙️ Kostenspalten für {name} (Auto: {len(candidates)} erkannt)"):
        st.caption("Regel: '€' im Header (erste Zeilen) + überwiegend numerisch. Hier kannst du exakt festlegen, welche Spalten als Kosten zählen.")
        manual = st.multiselect("Kostenspalten auswählen", func_df.columns.tolist(), default=candidates, key=f"costcols_{name}")
        st.session_state["cost_map"][name] = manual

    cost_cols = st.session_state["cost_map"][name]
    func_parsed = {"hier_cols": hier_cols, "cost_cols": cost_cols, "rollups": rollup_costs(func_df, hier_cols, cost_cols), "raw": func_df}

    tech_parsed = parse_tech(tech_df)
    meta = {}
    try:
        if start_df.shape[1] >= 2:
            for i in range(min(30, len(start_df))):
                k = str(start_df.iloc[i,0])
                v = start_df.iloc[i,1] if start_df.shape[1] > 1 else None
                if k and k.lower() not in ["nan","none",""]:
                    meta[k] = v
    except Exception:
        pass
    return {"name":name,"sheets":{"func":func_sheet,"tech":tech_sheet,"start":start_sheet},"func":func_parsed,"tech":tech_parsed,"meta":meta}

if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

products = [read_product(f) for f in files]
names = [p["name"] for p in products]

tab1, tab2, tab3, tab4 = st.tabs(["1) Funktionsstruktur", "2) Kosten & technische Evaluierung", "3) Vergleich (A vs B)", "4) Top 10 Abweichungen"])

with tab1:
    st.markdown('<div class="efeso-section-title">Funktionsstruktur (H1 > H2 > H3)</div>', unsafe_allow_html=True)
    for p in products:
        st.write(f"**{p['name']}** — Sheet: `{p['sheets']['func']}`")
        f = p["func"]["rollups"]
        if "H1" in f: st.caption("H1 (Hauptfunktionen)"); st.dataframe(f["H1"], use_container_width=True, height=220)
        if "H2" in f: st.caption("H2 (Teilfunktionen)"); st.dataframe(f["H2"], use_container_width=True, height=220)
        if "H3" in f: st.caption("H3 (Unterfunktionen)"); st.dataframe(f["H3"], use_container_width=True, height=220)

with tab2:
    st.markdown('<div class="efeso-section-title">Funktionskosten und technische Evaluierung</div>', unsafe_allow_html=True)
    sel = st.selectbox("Produkt wählen", names, index=0 if names else None, key="ktt")
    P = next(p for p in products if p["name"] == sel)
    if "H1" in P["func"]["rollups"]:
        h1 = P["func"]["rollups"]["H1"][["H1","TotalCost"]].copy().sort_values("H1")
        st.bar_chart(h1.set_index("H1"))
    st.write("**Technische Bewertung (gewichteter Score)**")
    st.write(pd.DataFrame([{"Produkt": P["name"], "Overall Tech Score": P["tech"]["overall"]}]))
    if P["tech"]["by_cat"] is not None:
        st.caption("Nach Kategorien")
        st.dataframe(P["tech"]["by_cat"], use_container_width=True, height=220)

with tab3:
    st.markdown('<div class="efeso-section-title">Auswahl & Vergleich</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen.")
    else:
        colA, colB = st.columns(2)
        with colA: product_a = st.selectbox("Produkt A", names, index=0, key="cmpA")
        with colB: product_b = st.selectbox("Produkt B", names, index=1, key="cmpB")
        if product_a == product_b:
            st.warning("Bitte zwei unterschiedliche Produkte wählen.")
        else:
            A = next(p for p in products if p["name"] == product_a)
            B = next(p for p in products if p["name"] == product_b)
            def sum_h1(p): return float(p["func"]["rollups"]["H1"]["TotalCost"].sum()) if "H1" in p["func"]["rollups"] else 0.0
            a_cost, b_cost = sum_h1(A), sum_h1(B)
            a_score, b_score = A["tech"]["overall"], B["tech"]["overall"]
            t1,t2,t3,t4 = st.columns(4)
            t1.metric(f"{A['name']} — Summe H1", f"{a_cost:,.0f} €")
            t2.metric(f"{B['name']} — Summe H1", f"{b_cost:,.0f} €", delta=f"{(b_cost-a_cost):,.0f} € vs A")
            t3.metric(f"{A['name']} — Tech", f"{a_score:.2f}")
            t4.metric(f"{B['name']} — Tech", f"{b_score:.2f}", delta=f"{(b_score-a_score):+.2f}")
            h1A = A["func"]["rollups"]["H1"].rename(columns={"TotalCost":"Cost_A"})
            h1B = B["func"]["rollups"]["H1"].rename(columns={"TotalCost":"Cost_B"})
            h1 = pd.merge(h1A, h1B, on="H1", how="outer").fillna(0.0)
            h1["Delta (B - A)"] = h1["Cost_B"] - h1["Cost_A"]
            st.dataframe(h1.sort_values("H1"), use_container_width=True, height=280)
            st.bar_chart(h1.set_index("H1")[["Cost_A","Cost_B"]])

with tab4:
    st.markdown('<div class="efeso-section-title">Top 10 Abweichungen</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen und im Tab 3 auswählen.")
    else:
        product_a = st.session_state.get("cmpA", names[0])
        product_b = st.session_state.get("cmpB", names[1] if len(names)>1 else names[0])
        if product_a == product_b:
            st.warning("Bitte zwei unterschiedliche Produkte im Tab 3 wählen.")
        else:
            A = next(p for p in products if p["name"] == product_a)
            B = next(p for p in products if p["name"] == product_b)
            def h23_table(p):
                f = p["func"]["rollups"]
                frames = []
                if "H2" in f: frames.append(f["H2"][["H1","H2","TotalCost"]].rename(columns={"TotalCost":"Cost"}))
                if "H3" in f: frames.append(f["H3"][["H1","H2","H3","TotalCost"]].rename(columns={"TotalCost":"Cost"}))
                if frames: out = pd.concat(frames, ignore_index=True)
                else: out = f["H1"][["H1","TotalCost"]].rename(columns={"TotalCost":"Cost"})
                return out
            tA = h23_table(A); tB = h23_table(B)
            tA["key"] = tA.apply(lambda r: " > ".join([str(r.get(c,"")) for c in ["H1","H2","H3"] if pd.notna(r.get(c,"")) and str(r.get(c,""))!=""]), axis=1)
            tB["key"] = tB.apply(lambda r: " > ".join([str(r.get(c,"")) for c in ["H1","H2","H3"] if pd.notna(r.get(c,"")) and str(r.get(c,""))!=""]), axis=1)
            tA = tA[["key","Cost"]].rename(columns={"Cost":"Cost_A"})
            tB = tB[["key","Cost"]].rename(columns={"Cost":"Cost_B"})
            diff = pd.merge(tA, tB, on="key", how="outer").fillna(0.0)
            diff["Delta (B - A)"] = diff["Cost_B"] - diff["Cost_A"]
            level = st.radio("Ebene", ["Gemischt (H2+H3)","Nur H2","Nur H3"], horizontal=True)
            if level == "Nur H2":
                diff = diff[diff["key"].str.count(">")==1]
            elif level == "Nur H3":
                diff = diff[diff["key"].str.count(">")>=2]
            topn = st.slider("Anzahl Top-Abweichungen", 5, 30, 10)
            diff_sorted = diff.sort_values("Delta (B - A)", key=lambda s: abs(s), ascending=False)
            topN = diff_sorted.head(topn)
            st.dataframe(topN, use_container_width=True, height=320)
            st.bar_chart(topN.set_index("key")[["Delta (B - A)"]])
            buff = io.BytesIO(); topN.to_csv(buff, index=False)
            st.download_button("Export Top-Abweichungen (CSV)", data=buff.getvalue(), file_name="top_differences.csv", mime="text/csv")
