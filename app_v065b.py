import io
import pandas as pd
import numpy as np
import streamlit as st
import altair as alt
import re

PRIMARY = "#f47c24"; DARK="#2a2a2a"; BLUE="#0055A4"; LIGHTBLUE="#66B2FF"

st.set_page_config(page_title="FKA v06.5b — EFESO Style (no Plotly)", layout="wide")

st.markdown('''
<style>
.sticky-header { position: sticky; top:0; z-index:999; background:#f47c24; border-radius:12px; padding:12px 16px; margin:4px 0 12px 0; box-shadow:0 2px 8px rgba(0,0,0,.1); }
.sticky-header h1{color:#fff;font-size:22px;margin:0;}
.sticky-header p{color:#ffe9d6;font-size:13px;margin:0;}
.section-title { color:#2a2a2a; border-left:6px solid #f47c24; padding-left:10px; font-weight:700; margin:14px 0 8px 0; font-size:18px;}
.stTabs [data-baseweb="tab"] { font-size:15px;font-weight:600;color:#2a2a2a;}
.stTabs [aria-selected="true"] { color:#f47c24; border-bottom:3px solid #f47c24;}
</style>
''', unsafe_allow_html=True)

st.markdown('<div class="sticky-header"><h1>FKA — Funktionskosten & Evaluierung</h1><p>v06.5b · robuste H1/H2-Erkennung · ohne Plotly</p></div>', unsafe_allow_html=True)

files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)

HAUPTFUNKTION_NAME_ROW = 0    # Excel Zeile 1
NEBENFUNKTION_NAME_ROW = 1    # Excel Zeile 2
NEBENFUNKTION_WERT_ROW = 4    # Excel Zeile 5

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

def _has_letters(s: str) -> bool:
    return bool(re.search(r"[A-Za-zÄÖÜäöüß]", s or ""))

def _is_h1_label(s: str) -> bool:
    if not s:
        return False
    s = s.strip()
    # reine Zahl / Zahl mit Komma / Prozent ausschließen
    if re.fullmatch(r"\d+([.,]\d+)?%?", s):
        return False
    if len(s) < 3:
        return False
    return _has_letters(s)

def _is_h2_label(s: str) -> bool:
    return bool(s and len(s.strip()) >= 3 and _has_letters(s))

def parse_h1_h2_from_header(func_df):
    ncols = func_df.shape[1]

    # 1) Erste sinnvolle H1-Spalte finden
    start_col = None
    for c in range(ncols):
        if _is_h1_label(_cell_str(func_df, HAUPTFUNKTION_NAME_ROW, c)):
            start_col = c
            break
    if start_col is None:
        start_col = 0  # Fallback

    # 2) H1-Blöcke aufbauen
    h1_blocks = []
    current_h1 = None
    block_start = None
    for c in range(start_col, ncols):
        cand = _cell_str(func_df, HAUPTFUNKTION_NAME_ROW, c)
        if _is_h1_label(cand):
            if current_h1 is not None:
                h1_blocks.append((current_h1, block_start, c - 1))
            current_h1 = cand
            block_start = c
    if current_h1 is not None:
        h1_blocks.append((current_h1, block_start, ncols - 1))

    # 3) H2 sammeln; nur Blöcke mit mind. einer validen H2 behalten
    rows_h2, rows_h1 = [], []
    for h1, c0, c1 in h1_blocks:
        h2_costs = []
        for c in range(c0, c1 + 1):
            h2_name = _cell_str(func_df, NEBENFUNKTION_NAME_ROW, c)
            if not _is_h2_label(h2_name):
                continue
            val = _cell_num(func_df, NEBENFUNKTION_WERT_ROW, c)
            if pd.notna(val) and float(val) != 0.0:
                rows_h2.append({
                    "Hauptfunktion": h1,
                    "Nebenfunktion": h2_name,
                    "Kosten Nebenfunktion": float(val),
                })
                h2_costs.append(float(val))

        if h2_costs:
            rows_h1.append({
                "Hauptfunktion": h1,
                "Kosten Hauptfunktion": float(sum(h2_costs)),
            })

    H2 = pd.DataFrame(rows_h2)
    H1 = pd.DataFrame(rows_h1)
    if not H1.empty:
        H1 = H1.groupby("Hauptfunktion", as_index=False)["Kosten Hauptfunktion"].sum()
        H1["Hauptfunktion"] = H1["Hauptfunktion"].astype(str)
    if not H2.empty:
        H2["Hauptfunktion"] = H2["Hauptfunktion"].astype(str)
        H2["Nebenfunktion"] = H2["Nebenfunktion"].astype(str)
    return H1, H2

def parse_tech_sheet(xl):
    # try to find tech sheet
    tech_sheet = None
    for n in xl.sheet_names:
        if "techn" in n.lower() or "bewert" in n.lower():
            tech_sheet = n
            if "slave_techn" in n.lower():
                break
    if tech_sheet is None:
        return {"overall": 0.0, "table": pd.DataFrame()}

    df = xl.parse(tech_sheet, header=None)
    # Columns: B -> names (index 1), L -> weight% (index 11), R -> score (index 17)
    name_col = 1; weight_col = 11; score_col = 17
    names = df.iloc[:, name_col].astype(str)
    weights = pd.to_numeric(df.iloc[:, weight_col].replace(",", ".", regex=True), errors="coerce")
    scores = pd.to_numeric(df.iloc[:, score_col].replace(",", ".", regex=True), errors="coerce")

    t = pd.DataFrame({"Funktion": names, "Gewichtung_%": weights, "Score": scores})
    t = t[(t["Funktion"].str.strip() != "") & (t["Score"].notna()) & (t["Gewichtung_%"].notna())]
    if t.empty:
        return {"overall": 0.0, "table": pd.DataFrame()}

    wsum = t["Gewichtung_%"].sum()
    if not wsum or wsum == 0:
        t["Gewichtung_%"] = 1.0
        wsum = t["Gewichtung_%"].sum()
    t["Gewichtung_norm"] = t["Gewichtung_%"] / wsum
    t["Gewichteter Score"] = t["Score"] * t["Gewichtung_norm"]
    overall = t["Gewichteter Score"].sum()

    t_display = t[["Funktion", "Gewichtung_%", "Score", "Gewichteter Score"]].copy()
    t_display["Gewichtung_%"] = (t_display["Gewichtung_%"]).round(2)
    t_display["Gewichteter Score"] = t_display["Gewichteter Score"].round(3)

    return {"overall": float(overall), "table": t_display, "sheet": tech_sheet}

def read_product(file):
    name = file.name.rsplit(".",1)[0]
    xl = pd.ExcelFile(file)

    func_sheet = xl.sheet_names[0]
    for n in xl.sheet_names:
        if any(k in n.lower() for k in ["funktion","kosten","slave_funktions"]):
            func_sheet = n; break
    func_df = xl.parse(func_sheet, header=None)

    H1, H2 = parse_h1_h2_from_header(func_df)
    tech = parse_tech_sheet(xl)

    return {"name": name, "sheets":{"func":func_sheet,"tech": tech.get("sheet","-")}, "H1":H1, "H2":H2, "tech":tech}

if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

products=[read_product(f) for f in files]
names=[p["name"] for p in products]

tab1, tab2, tab3, tab4 = st.tabs([
    "1) Überblick Haupt- & Nebenfunktionen",
    "2) Kosten & Technische Evaluierung",
    "3) Vergleich (Produkte A vs B)",
    "4) Top 10 Abweichungen (Nebenfunktionen)"
])

def alt_bar(df, x, y, color=None, title=None, sort=None):
    enc = alt.Chart(df).mark_bar().encode(
        x=alt.X(x, sort=sort, axis=alt.Axis(labelAngle=0)),
        y=alt.Y(y, axis=alt.Axis(title=None)),
    )
    if color:
        enc = enc.encode(color=color)
    if title:
        enc = enc.properties(title=title)
    return enc.properties(width="container", height=340).configure_axis(
        grid=True, gridColor="#e6e6e6"
    )

with tab1:
    st.markdown('<div class="section-title">Überblick Hauptfunktionen & Deep-Dive Nebenfunktionen</div>', unsafe_allow_html=True)
    psel=st.selectbox("Produkt wählen", names, index=0 if names else None, key="comp")
    P=next(p for p in products if p["name"]==psel)

    if not P["H1"].empty:
        chart1 = alt_bar(P["H1"], x="Hauptfunktion:N", y="Kosten Hauptfunktion:Q", title="Kosten je Hauptfunktion")
        st.altair_chart(chart1, use_container_width=True)
    st.dataframe(P["H1"], use_container_width=True, height=240)

    if not P["H2"].empty:
        h1_list=sorted(P["H2"]["Hauptfunktion"].dropna().astype(str).unique().tolist())
        chosen_h1=st.selectbox("Hauptfunktion wählen", h1_list, key="h1x")
        h2=P["H2"][P["H2"]["Hauptfunktion"].astype(str)==chosen_h1][["Nebenfunktion","Kosten Nebenfunktion"]].copy()
        total=h2["Kosten Nebenfunktion"].sum() or 1.0
        h2["Anteil_%"]=(h2["Kosten Nebenfunktion"]/total*100).round(1)
        h2=h2.sort_values("Anteil_%", ascending=False)
        chart2 = alt_bar(h2, x="Nebenfunktion:N", y="Anteil_%:Q", title="Anteile Nebenfunktionen (%)", sort=None)
        st.altair_chart(chart2, use_container_width=True)
        st.dataframe(h2, use_container_width=True, height=320)
    else:
        st.info("Keine Nebenfunktions-Kosten in Zeile 5 erkannt.")

with tab2:
    st.markdown('<div class="section-title">Funktionskosten (H1) & Technische Evaluierung</div>', unsafe_allow_html=True)
    sel=st.selectbox("Produkt wählen", names, index=0 if names else None, key="ktt")
    P=next(p for p in products if p["name"]==sel)
    c1, c2 = st.columns([2,2])
    with c1:
        chart = alt_bar(P["H1"], x="Hauptfunktion:N", y="Kosten Hauptfunktion:Q", title="Kosten je Hauptfunktion")
        st.altair_chart(chart, use_container_width=True)
    with c2:
        tinfo = P["tech"]
        overall = tinfo.get("overall", 0.0)
        ttable = tinfo.get("table", pd.DataFrame())
        st.metric("Overall Tech Score (gewichtet)", f"{overall:.3f}")
        if not ttable.empty:
            chartt = alt_bar(ttable, x="Funktion:N", y="Gewichteter Score:Q", title="Technische Bewertung (gewichtet)")
            st.altair_chart(chartt, use_container_width=True)
            st.dataframe(ttable, use_container_width=True, height=320)
        else:
            st.info("Kein technisches Bewertungsblatt gefunden oder keine Daten (B=Name, L=Gewichtung, R=Score).")

with tab3:
    st.markdown('<div class="section-title">Vergleich (A vs B) – Haupt-/Nebenfunktionen</div>', unsafe_allow_html=True)
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
            h1 = pd.merge(A["H1"].rename(columns={"Kosten Hauptfunktion":"Cost_A"}),
                          B["H1"].rename(columns={"Kosten Hauptfunktion":"Cost_B"}), on="Hauptfunktion", how="outer").fillna(0.0)
            dfm = h1.melt(id_vars="Hauptfunktion", value_vars=["Cost_A","Cost_B"], var_name="Produkt", value_name="Kosten")
            dfm["Produkt"] = dfm["Produkt"].map({"Cost_A": a, "Cost_B": b})
            chart = alt.Chart(dfm).mark_bar().encode(
                x=alt.X("Hauptfunktion:N", axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Kosten:Q"),
                color=alt.Color("Produkt:N", scale=alt.Scale(range=[BLUE, LIGHTBLUE]))
            ).properties(title="Kostenvergleich A vs B", width="container", height=360).configure_axis(grid=True, gridColor="#e6e6e6")
            st.altair_chart(chart, use_container_width=True)

            st.subheader("Alle Nebenfunktionen gegenübergestellt")
            h2=pd.merge(A["H2"].rename(columns={"Kosten Nebenfunktion":"Cost_A"}),
                        B["H2"].rename(columns={"Kosten Nebenfunktion":"Cost_B"}), on=["Hauptfunktion","Nebenfunktion"], how="outer").fillna(0.0)
            h2["Delta (B - A)"]=h2["Cost_B"]-h2["Cost_A"]
            st.dataframe(h2.sort_values(["Hauptfunktion","Nebenfunktion"]), use_container_width=True, height=360)

with tab4:
    st.markdown('<div class="section-title">Top 10 Abweichungen (Nebenfunktionen)</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen und im Tab 3 auswählen.")
    else:
        a=st.session_state.get("cmpA", names[0])
        b=st.session_state.get("cmpB", names[1] if len(names)>1 else names[0])
        if a==b: st.warning("Bitte zwei unterschiedliche Produkte im Tab 3 wählen.")
        else:
            A=next(p for p in products if p["name"]==a)
            B=next(p for p in products if p["name"]==b)
            h2=pd.merge(A["H2"].rename(columns={"Kosten Nebenfunktion":"Cost_A"}),
                        B["H2"].rename(columns={"Kosten Nebenfunktion":"Cost_B"}), on=["Hauptfunktion","Nebenfunktion"], how="outer").fillna(0.0)
            h2["Delta (B - A)"]=h2["Cost_B"]-h2["Cost_A"]
            h2["Delta_abs"]=h2["Delta (B - A)"].abs()

            topn=st.slider("Top-N", 5, 30, 10)
            top=(h2.sort_values("Delta_abs", ascending=False).head(topn).copy())
            top.insert(0,"Rang", range(1, len(top)+1))
            top["key"]=top["Hauptfunktion"].astype(str)+" > "+top["Nebenfunktion"].astype(str)
            top["Label"]=top["Rang"].astype(str)+". "+top["key"]

            tbl = top[["Rang","key","Cost_A","Cost_B","Delta (B - A)"]].rename(columns={
                "key":"Hauptfunktion > Nebenfunktion",
                "Cost_A":f"Kosten {a}", "Cost_B":f"Kosten {b}", "Delta (B - A)":"Δ Kosten (B−A)"
            })
            st.dataframe(tbl, use_container_width=True, height=320)

            # Altair chart with sign-based colors
            top["_pos"] = (top["Delta (B - A)"] >= 0).astype(int)
            scale = alt.Scale(domain=[0,1], range=[PRIMARY, BLUE])  # orange negative, blue positive
            chart = alt.Chart(top).mark_bar().encode(
                x=alt.X("Label:N", sort=list(top["Label"]), axis=alt.Axis(labelAngle=0)),
                y=alt.Y("Delta (B - A):Q"),
                color=alt.Color("_pos:O", scale=scale, legend=None),
                tooltip=["key","Delta (B - A)"]
            ).properties(width="container", height=360).configure_axis(grid=True, gridColor="#e6e6e6")
            st.altair_chart(chart, use_container_width=True)
