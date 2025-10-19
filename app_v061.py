
import io
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st

PRIMARY = "#f47c24"; DARK="#2a2a2a"

st.set_page_config(page_title="FKA v06.1 — EFESO Style (H2-only, H1=sum(H2), Header-Parser)", layout="wide")

st.markdown("""
<style>
.top-logo { margin: 6px 0 4px 0; }
.sticky-header { position: sticky; top:0; z-index:999; background:#f47c24; border-radius:12px; padding:12px 16px; margin:4px 0 12px 0; box-shadow:0 2px 8px rgba(0,0,0,.1); }
.sticky-header h1{color:#fff;font-size:22px;margin:0;}
.sticky-header p{color:#ffe9d6;font-size:13px;margin:0;}
.section-title { color:#2a2a2a; border-left:6px solid #f47c24; padding-left:10px; font-weight:700; margin:14px 0 8px 0; font-size:18px;}
.stTabs [data-baseweb="tab"] { font-size:15px;font-weight:600;color:#2a2a2a;}
.stTabs [aria-selected="true"] { color:#f47c24; }
</style>
""", unsafe_allow_html=True)

logo_col, _ = st.columns([1,6])
with logo_col:
    up_logo = st.file_uploader("Logo", type=["png","jpg","jpeg","gif"], label_visibility="collapsed", key="logo")
    if up_logo is not None:
        st.image(up_logo, width=180)
    else:
        base = Path(__file__).resolve().parent
        ph = base / "assets" / "efeso_logo_placeholder.png"
        if ph.exists():
            st.image(str(ph), width=180)
        else:
            st.markdown("<div style='color:#f47c24;font-weight:800;font-size:26px;'>EFESO</div>", unsafe_allow_html=True)
            st.caption("MANAGEMENT CONSULTANTS")

st.markdown('<div class="sticky-header"><h1>FKA — Funktionskosten & Evaluierung</h1><p>v06.1 · H2-only · H1 = Summe(H2) · Header-Parser</p></div>', unsafe_allow_html=True)

files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)

H1_ROW = 3
H2_ROW = 4

def _cell_str(df, r, c):
    try:
        v = df.iat[r, c]
        s = str(v).strip()
        return s if s.lower() not in ("nan", "none", "") else ""
    except Exception:
        return ""

def _cell_num(df, r, c):
    try:
        v = df.iat[r, c]
        if isinstance(v, str):
            v = v.replace(".", "").replace(",", ".")
        return pd.to_numeric(v, errors="coerce")
    except Exception:
        return np.nan

def parse_h1_h2_from_header(func_df):
    h1_blocks = []
    current_h1 = None
    ncols = func_df.shape[1]

    for c in range(ncols):
        name = _cell_str(func_df, H1_ROW, c)
        if name:
            if current_h1 is not None:
                h1_blocks.append((current_h1, block_start, c-1))
            current_h1 = name
            block_start = c
    if current_h1 is not None:
        h1_blocks.append((current_h1, block_start, ncols-1))

    rows_h2 = []
    rows_h1 = []
    for h1, c0, c1 in h1_blocks:
        h2_entries = []
        h2_count = 0
        for c in range(c0, c1+1):
            val = _cell_num(func_df, H2_ROW, c)
            if pd.notna(val) and val != 0:
                h2_count += 1
                h2_name = _cell_str(func_df, H2_ROW-1, c)
                if not h2_name:
                    h2_name = f"H2_{h2_count}"
                rows_h2.append({"H1": h1, "H2": h2_name, "H2Cost": float(val)})
                h2_entries.append(float(val))
        h1cost = sum(h2_entries) if h2_entries else 0.0
        rows_h1.append({"H1": h1, "H1Cost": h1cost})

    H2 = pd.DataFrame(rows_h2)
    H1 = pd.DataFrame(rows_h1).groupby("H1", as_index=False)["H1Cost"].sum()

    if not H2.empty:
        H2["H1"] = H2["H1"].astype(str)
        H2["H2"] = H2["H2"].astype(str)
    if not H1.empty:
        H1["H1"] = H1["H1"].astype(str)
    return H1, H2

def parse_tech(df_raw):
    df = df_raw.copy(); df.columns=[str(c).strip() for c in df.columns]
    score_col=None; weight_col=None
    for c in df.columns:
        cl=str(c).lower()
        if score_col is None and pd.api.types.is_numeric_dtype(df[c]) and any(k in cl for k in ["score","wert","points"]):
            score_col=c
        if weight_col is None and any(k in cl for k in ["gewicht","weight","wgt"]):
            weight_col=c
    if score_col is None:
        nums=[c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
        score_col = nums[-1] if nums else None
    if weight_col is None:
        weight_col="_Weight"; df[weight_col]=1.0
    else:
        df[weight_col]=pd.to_numeric(df[weight_col], errors="coerce").fillna(1.0)
    if score_col is None:
        df["_score"]=0.0; score_col="_score"
    else:
        df[score_col]=pd.to_numeric(df[score_col], errors="coerce").fillna(0.0)
    df["_Weighted"]=df[score_col]*df[weight_col]
    wsum=df[weight_col].sum() or 1.0
    overall=df["_Weighted"].sum()/wsum
    return {"overall": overall}

def read_product(file):
    name=file.name.rsplit(".",1)[0]
    xl=pd.ExcelFile(file)
    func_sheet = xl.sheet_names[0]
    for n in xl.sheet_names:
        if any(k in n.lower() for k in ["funktion","kosten","slave_funktions"]):
            func_sheet = n; break
    func_df = xl.parse(func_sheet, header=None)

    tech_sheet = None
    for n in xl.sheet_names:
        if any(k in n.lower() for k in ["techn", "bewert", "rating", "score"]):
            tech_sheet = n; break
    tech_df = xl.parse(tech_sheet) if tech_sheet else pd.DataFrame()

    H1, H2 = parse_h1_h2_from_header(func_df)
    tech = parse_tech(tech_df) if not tech_df.empty else {"overall": 0.0}
    return {"name": name, "sheets": {"func": func_sheet, "tech": tech_sheet or "-"}, "H1": H1, "H2": H2, "tech": tech}

if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

products=[read_product(f) for f in files]
names=[p["name"] for p in products]

tab1, tab2, tab3, tab4 = st.tabs(["1) Komposition (H1 aus H2)", "2) Kosten & Technische Evaluierung", "3) Vergleich (A vs B)", "4) Top 10 Abweichungen (H2)"])

with tab1:
    st.markdown('<div class="section-title">Komposition: H1 setzt sich aus H2 zusammen (dynamisch aus Zeile 5)</div>', unsafe_allow_html=True)
    psel=st.selectbox("Produkt wählen", names, index=0 if names else None, key="comp")
    P=next(p for p in products if p["name"]==psel)
    st.caption("H1-Kosten = Summe der H2-Kosten (strict, dynamisch je Block).")
    st.dataframe(P["H1"], use_container_width=True, height=220)
    if not P["H2"].empty:
        h1_list=sorted(P["H2"]["H1"].dropna().astype(str).unique().tolist())
        chosen_h1=st.selectbox("Hauptfunktion wählen", h1_list, key="h1x")
        h2=P["H2"][P["H2"]["H1"].astype(str)==chosen_h1][["H2","H2Cost"]].copy()
        total=h2["H2Cost"].sum() or 1.0
        h2["Anteil_%"]=(h2["H2Cost"]/total*100).round(1)
        h2=h2.sort_values("Anteil_%", ascending=False)
        st.dataframe(h2, use_container_width=True, height=320)
        st.bar_chart(h2.set_index("H2")[["Anteil_%"]])
    else:
        st.info("Keine H2-Zeile erkannt. Bitte prüfen, ob Row 5 Werte enthält.")

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
            h1=pd.merge(A["H1"].rename(columns={"H1Cost":"Cost_A"}),
                        B["H1"].rename(columns={"H1Cost":"Cost_B"}), on="H1", how="outer").fillna(0.0)
            h1["Delta (B - A)"]=h1["Cost_B"]-h1["Cost_A"]
            st.dataframe(h1.sort_values("H1"), use_container_width=True, height=280)
            st.bar_chart(h1.set_index("H1")[["Cost_A","Cost_B"]])
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
            buff=io.BytesIO(); top.to_csv(buff, index=False)
            st.download_button("Export Top-N (CSV)", data=buff.getvalue(), file_name="topN_H2_differences.csv", mime="text/csv")
