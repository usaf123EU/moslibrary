import io
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st

PRIMARY = "#f47c24"; DARK="#2a2a2a"

st.set_page_config(page_title="FKA v06 — EFESO Style (H2-only Top10, H1=sum(H2))", layout="wide")

# ---------------------- CSS ----------------------
st.markdown(r'''
<style>
.top-logo { margin: 6px 0 4px 0; }
.sticky-header { position: sticky; top:0; z-index:999; background:#f47c24; border-radius:12px; padding:12px 16px; margin:4px 0 12px 0; box-shadow:0 2px 8px rgba(0,0,0,.1); }
.sticky-header h1{color:#fff;font-size:22px;margin:0;}
.sticky-header p{color:#ffe9d6;font-size:13px;margin:0;}
.section-title { color:#2a2a2a; border-left:6px solid #f47c24; padding-left:10px; font-weight:700; margin:14px 0 8px 0; font-size:18px;}
.stTabs [data-baseweb="tab"] { font-size:15px;font-weight:600;color:#2a2a2a;}
.stTabs [aria-selected="true"] { color:#f47c24; }
</style>
''', unsafe_allow_html=True)

# ---------- Top: Logo ----------
logo_col, _ = st.columns([1,6])
with logo_col:
    up_logo = st.file_uploader("Logo", type=["png","jpg","jpeg"], label_visibility="collapsed", key="logo")
    if up_logo is not None:
        st.image(up_logo, width=180)
    else:
        base = Path(__file__).resolve().parent
        ph = base / "assets" / "efeso_logo_placeholder.png"
        st.image(str(ph) if ph.exists() else None, width=180)

# ---------- Sticky header ----------
st.markdown('<div class="sticky-header"><h1>FKA — Funktionskosten & Evaluierung</h1><p>v06 · H2-only Top10 · H1=sum(H2)</p></div>', unsafe_allow_html=True)

# ---------- Upload ----------
files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)

# ------------------ Helpers ------------------
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
            parsed_numeric_cols.append(c); continue
        s = pd.to_numeric(df[c].astype(str).str.replace(".","", regex=False).str.replace(",",".", regex=False), errors="coerce")
        if s.notna().sum() >= max(3, int(0.5*len(s))):
            if not header_zone[c].str.contains(r"[A-Za-z]", regex=True).any() or header_zone[c].str.contains("€").any():
                parsed_numeric_cols.append(c)
    return list(dict.fromkeys(euro_cols + parsed_numeric_cols))

def clean_numeric(series):
    if pd.api.types.is_numeric_dtype(series): return series
    return pd.to_numeric(series.astype(str).str.replace(".","", regex=False).str.replace(",",".", regex=False), errors="coerce")

def parse_tech(df_raw):
    df = df_raw.copy(); df.columns=[str(c).strip() for c in df.columns]
    score_col=None; weight_col=None; crit_col=None; cat_col=None
    for c in df.columns:
        cl=str(c).lower()
        if crit_col is None and any(k in cl for k in ["kriter","criterion","metric"]): crit_col=c
        if cat_col is None and any(k in cl for k in ["kategorie","category","group"]): cat_col=c
        if weight_col is None and any(k in cl for k in ["gewicht","weight","wgt"]): weight_col=c
        if score_col is None and any(k in cl for k in ["score","wert","points"]) and pd.api.types.is_numeric_dtype(df[c]): score_col=c
    if score_col is None:
        nums=[c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        score_col = nums[-1] if nums else None
    if crit_col is None:
        crit_col="_Criterion"; df[crit_col]=np.arange(1,len(df)+1)
    if weight_col is None:
        weight_col="_Weight"; df[weight_col]=1.0
    else: df[weight_col]=pd.to_numeric(df[weight_col], errors="coerce").fillna(1.0)
    if score_col is None:
        df["_score"]=0.0; score_col="_score"
    else: df[score_col]=pd.to_numeric(df[score_col], errors="coerce").fillna(0.0)
    df["_Weighted"]=df[score_col]*df[weight_col]
    wsum=df[weight_col].sum() or 1.0
    overall=df["_Weighted"].sum()/wsum
    by_cat=None
    if cat_col:
        by_cat=df.groupby(cat_col, dropna=False).apply(
            lambda g: pd.Series({"Weight":g[weight_col].sum(),"WeightedScore":g["_Weighted"].sum(),"ScoreAvg":(g["_Weighted"].sum()/(g[weight_col].sum() or 1.0))})
        ).reset_index()
    return {"overall":overall,"by_cat":by_cat}

def read_product(file):
    name=file.name.rsplit(".",1)[0]
    xl=pd.ExcelFile(file)
    func_sheet=_pick_sheet(xl,FUNC_SHEET_HINTS)
    tech_sheet=_pick_sheet(xl,TECH_SHEET_HINTS)
    func_df=xl.parse(func_sheet)
    tech_df=xl.parse(tech_sheet)

    # hierarchy detection (prefer explicit names)
    h_cols=[]
    for key in ["H1","H2","Hauptfunktion","Teilfunktion","Nebenfunktion"]:
        for c in func_df.columns:
            if key.lower() in str(c).lower() and c not in h_cols: h_cols.append(c)
    text_cols=[c for c in func_df.columns if not pd.api.types.is_numeric_dtype(func_df[c])]
    for c in func_df.columns:
        if c not in h_cols and c in text_cols and len(h_cols)<2: h_cols.append(c)
    if len(h_cols)<2:
        # guarantee two cols
        h_cols = (h_cols + text_cols)[:2]

    # cost columns
    candidates = detect_cost_columns(func_df)
    st.session_state.setdefault("cost_map", {})
    with st.expander(f"⚙️ Kostenspalten für {name} (Auto: {len(candidates)} erkannt)"):
        manual = st.multiselect("Kostenspalten auswählen", func_df.columns.tolist(), default=candidates, key=f"costcols_{name}")
        st.session_state["cost_map"][name] = manual
    cost_cols = st.session_state["cost_map"][name]

    # compute H2 costs per (H1,H2)
    tmp=func_df.copy()
    for c in cost_cols: tmp[c]=clean_numeric(tmp[c])
    tmp["RowCost"]=tmp[cost_cols].sum(axis=1, skipna=True)
    H2 = tmp.groupby([h_cols[0],h_cols[1]], dropna=False)["RowCost"].sum().reset_index().rename(columns={h_cols[0]:"H1",h_cols[1]:"H2","RowCost":"H2Cost"})
    # H1 strictly as sum of H2
    H1 = H2.groupby("H1", dropna=False)["H2Cost"].sum().reset_index().rename(columns={"H2Cost":"H1Cost"})

    tech=parse_tech(tech_df)
    return {"name":name, "sheets":{"func":func_sheet,"tech":tech_sheet}, "H1":H1, "H2":H2, "tech":tech}

# ------------------ Ingest ------------------
if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

products=[read_product(f) for f in files]
names=[p["name"] for p in products]

tab1, tab2, tab3, tab4 = st.tabs(["1) Komposition (H1 aus H2)", "2) Kosten & Technische Evaluierung", "3) Vergleich (A vs B)", "4) Top 10 Abweichungen (H2)"])

with tab1:
    st.markdown('<div class="section-title">Komposition: H1 setzt sich aus H2 zusammen</div>', unsafe_allow_html=True)
    psel=st.selectbox("Produkt wählen", names, index=0 if names else None, key="comp")
    P=next(p for p in products if p["name"]==psel)
    st.caption("H1-Kosten = Summe der H2-Kosten (strikt).")
    st.dataframe(P["H1"], use_container_width=True, height=220)
    # composition view
    h1_list=sorted(P["H2"]["H1"].dropna().astype(str).unique().tolist())
    chosen_h1=st.selectbox("Hauptfunktion wählen", h1_list, key="h1x")
    h2=P["H2"][P["H2"]["H1"].astype(str)==chosen_h1][["H2","H2Cost"]].copy()
    total=h2["H2Cost"].sum() or 1.0
    h2["Anteil_%"]=(h2["H2Cost"]/total*100).round(1)
    h2=h2.sort_values("Anteil_%", ascending=False)
    st.dataframe(h2, use_container_width=True, height=320)
    st.bar_chart(h2.set_index("H2")[["Anteil_%"]])

with tab2:
    st.markdown('<div class="section-title">Funktionskosten (H1) & Technische Evaluierung</div>', unsafe_allow_html=True)
    sel=st.selectbox("Produkt wählen", names, index=0 if names else None, key="ktt")
    P=next(p for p in products if p["name"]==sel)
    st.bar_chart(P["H1"].set_index("H1")[["H1Cost"]])
    st.write("**Technische Bewertung (gewichteter Score)**")
    st.write(pd.DataFrame([{"Produkt": P["name"], "Overall Tech Score": P["tech"]["overall"]}]))

with tab3:
    st.markdown('<div class="section-title">Vergleich (A vs B) – H1 & H2</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen.")
    else:
        colA, colB = st.columns(2)
        with colA: a=st.selectbox("Produkt A", names, index=0, key="cmpA")
        with colB: b=st.selectbox("Produkt B", names, index=1, key="cmpB")
        if a==b: st.warning("Bitte zwei unterschiedliche Produkte wählen.")
        else:
            A=next(p for p in products if p["name"]==a)
            B=next(p for p in products if p["name"]==b)
            a_sum=float(A["H1"]["H1Cost"].sum()); b_sum=float(B["H1"]["H1Cost"].sum())
            c1,c2=st.columns(2)
            with c1: st.metric(f"{A['name']} — Summe H1", f"{a_sum:,.0f} €")
            with c2: st.metric(f"{B['name']} — Summe H1", f"{b_sum:,.0f} €", delta=f"{(b_sum-a_sum):,.0f} € vs A")
            # H1 compare
            h1=pd.merge(A["H1"].rename(columns={"H1Cost":"Cost_A"}),
                        B["H1"].rename(columns={"H1Cost":"Cost_B"}), on="H1", how="outer").fillna(0.0)
            h1["Delta (B - A)"]=h1["Cost_B"]-h1["Cost_A"]
            st.dataframe(h1.sort_values("H1"), use_container_width=True, height=300)
            st.bar_chart(h1.set_index("H1")[["Cost_A","Cost_B"]])
            # H2 compare (all H2 juxtaposed)
            st.subheader("Alle H2 gegenübergestellt")
            h2=pd.merge(A["H2"].rename(columns={"H2Cost":"Cost_A"}),
                        B["H2"].rename(columns={"H2Cost":"Cost_B"}), on=["H1","H2"], how="outer").fillna(0.0)
            h2["Delta (B - A)"]=h2["Cost_B"]-h2["Cost_A"]
            st.dataframe(h2.sort_values(["H1","H2"]), use_container_width=True, height=360)

with tab4:
    st.markdown('<div class="section-title">Top 10 Abweichungen (H2-basiert)</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen und im Tab 3 auswählen.")
    else:
        a=st.session_state.get("cmpA", names[0])
        b=st.session_state.get("cmpB", names[1] if len(names)>1 else names[0])
        if a==b: st.warning("Bitte zwei unterschiedliche Produkte im Tab 3 wählen.")
        else:
            A=next(p for p in products if p["name"]==a)
            B=next(p for p in products if p["name"]==b)
            h2=pd.merge(A["H2"].rename(columns={"H2Cost":"Cost_A"}),
                        B["H2"].rename(columns={"H2Cost":"Cost_B"}), on=["H1","H2"], how="outer").fillna(0.0)
            h2["Delta (B - A)"]=h2["Cost_B"]-h2["Cost_A"]
            h2["key"]=h2["H1"].astype(str)+" > "+h2["H2"].astype(str)
            topn=st.slider("Top-N", 5, 30, 10)
            top=h2.sort_values("Delta (B - A)", key=lambda s: abs(s), ascending=False).head(topn)
            st.dataframe(top[["key","Cost_A","Cost_B","Delta (B - A)"]], use_container_width=True, height=320)
            st.bar_chart(top.set_index("key")[["Delta (B - A)"]])
            import io
            buff=io.BytesIO(); top.to_csv(buff, index=False)
            st.download_button("Export Top-N (CSV)", data=buff.getvalue(), file_name="topN_H2_differences.csv", mime="text/csv")
