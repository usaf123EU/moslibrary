
import io, os, re, base64
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

PRIMARY = "#f47c24"; DARK="#2a2a2a"; BLUE="#0055A4"; LIGHTBLUE="#66B2FF"
st.set_page_config(page_title="FKA v07.3 — EFESO", layout="wide", page_icon="📊")

def _efeso_logo_b64():
    svg_inline = \"\"\"<svg xmlns='http://www.w3.org/2000/svg' width='220' height='36' viewBox='0 0 220 36'>
      <rect width='220' height='36' fill='#ffffff'/>
      <circle cx='18' cy='18' r='10' fill='#f47c24'/>
      <text x='36' y='24' font-family='Arial, Helvetica, sans-serif' font-size='18' fill='#2a2a2a'>EFESO MANAGEMENT</text>
    </svg>\"\"\"
    path = os.path.join(os.path.dirname(__file__), "EFESO-Logo.svg")
    data = open(path, "rb").read() if os.path.exists(path) else svg_inline.encode("utf-8")
    return "data:image/svg+xml;base64," + base64.b64encode(data).decode()

logo_html = f"<img src='{_efeso_logo_b64()}' style='height:36px;vertical-align:middle;margin-right:12px;'/>"

st.markdown(\"\"\"
<style>
.headerbar { position: sticky; top:0; z-index:999; background:#fff; border-radius:12px;
             padding:10px 14px; margin:0 0 10px 0; box-shadow:0 2px 8px rgba(0,0,0,.06);
             display:flex; gap:12px; align-items:center; }
.headerbar h1{color:#2a2a2a;font-size:20px;margin:0;}
.headerbar small{color:#666;margin-left:6px;}
.section-title { color:#2a2a2a; border-left:6px solid #f47c24; padding-left:10px;
                 font-weight:700; margin:14px 0 8px 0; font-size:18px;}
.stTabs [data-baseweb="tab"]{font-size:15px;font-weight:600;color:#2a2a2a;}
.stTabs [aria-selected="true"]{color:#f47c24;border-bottom:3px solid #f47c24;}
</style>
\"\"\", unsafe_allow_html=True)

st.markdown(f\"\"\"
<div class="headerbar">
  {logo_html}
  <div>
    <h1>FKA — Funktionskosten & Evaluierung <small>v07.3</small></h1>
    <small>Fixer Funktionsbaum-Parser (Zeile 3 ab Spalte B) · H1-Fallback aus H2 · EFESO-Branding · Legende verbessert</small>
  </div>
</div>
\"\"\", unsafe_allow_html=True)

files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)
if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

HAUPTFUNKTION_NAME_ROW = 0
NEBENFUNKTION_NAME_ROW = 1
NEBENFUNKTION_WERT_ROW = 4

def _cell_str(df, r, c):
    try:
        v = df.iat[r, c]; s = str(v).strip()
        return s if s.lower() not in ("nan","none","") else ""
    except Exception: return ""

def _cell_num(df, r, c):
    try:
        v = df.iat[r, c]
        if isinstance(v,str): v = v.replace(".","").replace(",",".")
        return pd.to_numeric(v, errors="coerce")
    except Exception: return np.nan

def _has_letters(s): return bool(re.search(r"[A-Za-zÄÖÜäöüß]", s or ""))

def _is_h1_label(s):
    if not s: return False
    s=s.strip()
    if re.fullmatch(r"\d+([.,]\d+)?%?", s): return False
    return len(s)>=3 and _has_letters(s)

def _is_h2_label(s): return bool(s and len(s.strip())>=3 and _has_letters(s))

def parse_h1_h2_from_header(func_df):
    ncols = func_df.shape[1]
    start_col = 0
    for c in range(ncols):
        if _is_h1_label(_cell_str(func_df, HAUPTFUNKTION_NAME_ROW, c)):
            start_col = c; break

    h1_blocks=[]; current=None; cstart=None
    for c in range(start_col, ncols):
        cand=_cell_str(func_df, HAUPTFUNKTION_NAME_ROW, c)
        if _is_h1_label(cand):
            if current is not None: h1_blocks.append((current, cstart, c-1))
            current=cand; cstart=c
    if current is not None: h1_blocks.append((current, cstart, ncols-1))

    rows_h2=[]; rows_h1=[]
    for h1,c0,c1 in h1_blocks:
        h2_costs=[]
        for c in range(c0, c1+1):
            h2=_cell_str(func_df, NEBENFUNKTION_NAME_ROW, c)
            if not _is_h2_label(h2): continue
            val=_cell_num(func_df, NEBENFUNKTION_WERT_ROW, c)
            if pd.notna(val) and float(val)!=0.0:
                rows_h2.append({"Hauptfunktion":h1,"Nebenfunktion":h2,"Kosten Nebenfunktion":float(val)})
                h2_costs.append(float(val))
        if h2_costs:
            rows_h1.append({"Hauptfunktion":h1,"Kosten Hauptfunktion":float(sum(h2_costs))})

    H2=pd.DataFrame(rows_h2); H1=pd.DataFrame(rows_h1)
    if not H1.empty:
        H1=H1.groupby("Hauptfunktion",as_index=False)["Kosten Hauptfunktion"].sum()
        H1["Hauptfunktion"]=H1["Hauptfunktion"].astype(str)
    if not H2.empty:
        H2["Hauptfunktion"]=H2["Hauptfunktion"].astype(str)
        H2["Nebenfunktion"]=H2["Nebenfunktion"].astype(str)
    return H1,H2

def parse_tech_sheet(xl):
    tech_sheet=None
    for n in xl.sheet_names:
        if "techn" in n.lower() or "bewert" in n.lower():
            tech_sheet=n
            if "slave_techn" in n.lower(): break
    if tech_sheet is None: return {"overall":0.0,"table":pd.DataFrame()}
    df=xl.parse(tech_sheet, header=None)
    name_col=1; weight_col=11; score_col=17
    names=df.iloc[:,name_col].astype(str)
    weights=pd.to_numeric(df.iloc[:,weight_col].replace(",",".",regex=True), errors="coerce")
    scores=pd.to_numeric(df.iloc[:,score_col].replace(",",".",regex=True), errors="coerce")
    t=pd.DataFrame({"Funktion":names,"Gewichtung_%":weights,"Score":scores})
    t=t[(t["Funktion"].str.strip()!="") & t["Gewichtung_%"].notna() & t["Score"].notna()]
    return {"overall":0.0,"table":t,"sheet":tech_sheet}

def parse_funktionsbaum(xl):
    fb_sheet=None
    for n in xl.sheet_names:
        if "funktionsbaum" in n.lower():
            fb_sheet=n; break
    if fb_sheet is None:
        return pd.DataFrame(columns=["Hauptfunktion","Gewichtung_%"])

    df=xl.parse(fb_sheet, header=None)
    H1_ROW=2; START_COL=1; ncols=df.shape[1]
    h1_map=[]
    for c in range(START_COL, ncols):
        v=str(df.iat[H1_ROW, c]).strip()
        if v and v.lower() not in ("nan","none"):
            h1_map.append((c,v))
    if not h1_map: return pd.DataFrame(columns=["Hauptfunktion","Gewichtung_%"])

    best_row=None; best_cnt=-1
    for r in range(H1_ROW+2, min(H1_ROW+12, df.shape[0])):
        vals=pd.to_numeric(df.iloc[r].astype(str).str.replace("%","",regex=False).str.replace(",",".",regex=False), errors="coerce")
        cnt=int(vals.notna().sum())
        if cnt>best_cnt: best_cnt=cnt; best_row=r

    weights=[]
    if best_row is not None:
        rowvals=pd.to_numeric(df.iloc[best_row].astype(str).str.replace("%","",regex=False).str.replace(",",".",regex=False), errors="coerce")
        for c,label in h1_map:
            if c<len(rowvals):
                w=rowvals.iat[c]
                if pd.notna(w):
                    weights.append({"Hauptfunktion":label,"Gewichtung_%":float(w)})
    return pd.DataFrame(weights, columns=["Hauptfunktion","Gewichtung_%"])

def ensure_h1_df(d):
    if d is None or d.empty: return pd.DataFrame(columns=["Hauptfunktion","Kosten Hauptfunktion"])
    cols=set(d.columns)
    if "Hauptfunktion" not in cols or "Kosten Hauptfunktion" not in cols:
        return pd.DataFrame(columns=["Hauptfunktion","Kosten Hauptfunktion"])
    return d[["Hauptfunktion","Kosten Hauptfunktion"]].copy()

def ensure_h1_costs(product):
    H1=product.get("H1")
    if H1 is None or H1.empty:
        H2=product.get("H2", pd.DataFrame())
        if not H2.empty:
            H1=(H2.groupby("Hauptfunktion", as_index=False)["Kosten Nebenfunktion"]
                   .sum().rename(columns={"Kosten Nebenfunktion":"Kosten Hauptfunktion"}))
    return ensure_h1_df(H1)

def read_product(file):
    raw=file.name.rsplit(".",1)[0]
    safe=re.sub(r"[^A-Za-z0-9_\-]+","_",raw).strip("_")
    name=safe or raw
    xl=pd.ExcelFile(file)

    func_sheet=xl.sheet_names[0]
    for n in xl.sheet_names:
        if any(k in n.lower() for k in ["funktion","kosten","slave_funktions"]):
            func_sheet=n; break
    func_df=xl.parse(func_sheet, header=None)

    H1,H2=parse_h1_h2_from_header(func_df)
    tech=parse_tech_sheet(xl)
    weights=parse_funktionsbaum(xl)

    return {"name":name,"sheets":{"func":func_sheet,"tech":tech.get("sheet","-")},
            "H1":H1,"H2":H2,"tech":tech,"weights":weights}

products=[read_product(f) for f in files]
names=[p["name"] for p in products]

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "1) Überblick","2) Kosten & Tech (H1)","3) Vergleich A vs B","4) Top-Abweichungen (H2)","5) Kosten vs Gewichtung"
])

def alt_bar(df, x, y, title=None):
    ch=(alt.Chart(df).mark_bar().encode(
        x=alt.X(x, axis=alt.Axis(labelAngle=0)),
        y=alt.Y(y, axis=alt.Axis(title=None))
    ).properties(width="container", height=340))
    if title: ch=ch.properties(title=title)
    return ch.configure_axis(grid=True, gridColor="#e6e6e6")

with tab1:
    st.markdown('<div class="section-title">Überblick Haupt- & Nebenfunktionen</div>', unsafe_allow_html=True)
    psel=st.selectbox("Produkt wählen", names, key="t1sel")
    P=next(p for p in products if p["name"]==psel)
    P_h1=ensure_h1_costs(P)
    st.dataframe(P_h1, use_container_width=True, height=240)
    if not P["H2"].empty:
        h1_list=sorted(P["H2"]["Hauptfunktion"].unique().tolist())
        chosen=st.selectbox("Hauptfunktion", h1_list)
        h2=P["H2"][P["H2"]["Hauptfunktion"]==chosen][["Nebenfunktion","Kosten Nebenfunktion"]].copy()
        tot=h2["Kosten Nebenfunktion"].sum() or 1
        h2["Anteil_%"]=(h2["Kosten Nebenfunktion"]/tot*100).round(1)
        st.altair_chart(alt_bar(h2,"Nebenfunktion:N","Anteil_%:Q","Anteile Nebenfunktionen (%)"), use_container_width=True)
        st.dataframe(h2, use_container_width=True, height=300)

with tab2:
    st.markdown('<div class="section-title">Funktionskosten (H1) & Technische Evaluierung</div>', unsafe_allow_html=True)
    psel=st.selectbox("Produkt wählen", names, key="t2sel")
    P=next(p for p in products if p["name"]==psel)
    c1,c2=st.columns(2)
    with c1:
        P_h1=ensure_h1_costs(P)
        st.altair_chart(alt_bar(P_h1,"Hauptfunktion:N","Kosten Hauptfunktion:Q","Kosten je Hauptfunktion"), use_container_width=True)
    with c2:
        ttable=P["tech"]["table"]
        if not ttable.empty and not P_h1.empty:
            h1set=set(P_h1["Hauptfunktion"].astype(str))
            ttf=ttable[ttable["Funktion"].astype(str).isin(h1set)].copy()
            if not ttf.empty:
                wsum=ttf["Gewichtung_%"].sum()
                ttf["Gewichtung_norm"]=ttf["Gewichtung_%"]/wsum if wsum else 0.0
                ttf["Gewichteter Score"]=ttf["Score"]*ttf["Gewichtung_norm"]
                overall=float(ttf["Gewichteter Score"].sum())
                st.metric("Overall Tech Score (gewichtet, nur H1)", f"{overall:.3f}")
                chartt=(alt.Chart(ttf).mark_bar().encode(x=alt.X("Funktion:N", axis=alt.Axis(labelAngle=0)), y="Gewichteter Score:Q")
                        .properties(title="Technische Bewertung (H1, gewichtet)", width="container", height=340)
                        ).configure_axis(grid=True, gridColor="#e6e6e6")
                st.altair_chart(chartt, use_container_width=True)
                st.dataframe(ttf[["Funktion","Gewichtung_%","Score","Gewichteter Score"]], use_container_width=True, height=300)
            else:
                st.info("Im Bewertungsblatt wurden keine H1-Namen gefunden.")
        else:
            st.info("Kein technisches Bewertungsblatt gefunden. (B=Name, L=Gewichtung, R=Score)")

with tab3:
    st.markdown('<div class="section-title">Vergleich (A vs B) – nebeneinander</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen.")
    else:
        colA,colB=st.columns(2)
        with colA: a=st.selectbox("Produkt A", names, index=0, key="cmpA")
        with colB: b=st.selectbox("Produkt B", names, index=1, key="cmpB")
        if a==b:
            st.warning("Bitte zwei unterschiedliche Produkte wählen.")
        else:
            A=next(p for p in products if p["name"]==a); B=next(p for p in products if p["name"]==b)
            A_h1=ensure_h1_costs(A); B_h1=ensure_h1_costs(B)
            if A_h1.empty or B_h1.empty:
                st.error("In mindestens einer Datei konnten keine Hauptfunktionen erkannt werden.")
                st.stop()
            h1=pd.merge(A_h1.rename(columns={"Kosten Hauptfunktion":"Cost_A"}),
                        B_h1.rename(columns={"Kosten Hauptfunktion":"Cost_B"}),
                        on="Hauptfunktion", how="outer").fillna(0.0)
            dfm=h1.melt(id_vars="Hauptfunktion", value_vars=["Cost_A","Cost_B"], var_name="Produkt", value_name="Kosten")
            dfm["Produkt"]=dfm["Produkt"].map({"Cost_A":a,"Cost_B":b})
            chart=(alt.Chart(dfm).mark_bar().encode(
                x=alt.X("Hauptfunktion:N", axis=alt.Axis(labelAngle=0)),
                xOffset="Produkt:N",
                y="Kosten:Q",
                color=alt.Color("Produkt:N", scale=alt.Scale(range=[BLUE, LIGHTBLUE]))
            ).properties(title="Kostenvergleich A vs B (gruppiert)", width="container", height=360)
            ).configure_axis(grid=True, gridColor="#e6e6e6")
            st.altair_chart(chart, use_container_width=True)

with tab4:
    st.markdown('<div class="section-title">Top-Abweichungen (Nebenfunktionen)</div>', unsafe_allow_html=True)
    if len(products)<2:
        st.info("Bitte mind. zwei Produkte hochladen und im Tab 3 A/B wählen.")
    else:
        a=st.session_state.get("cmpA", names[0])
        b=st.session_state.get("cmpB", names[1] if len(names)>1 else names[0])
        if a==b:
            st.warning("Bitte zwei unterschiedliche Produkte im Tab 3 wählen.")
        else:
            A=next(p for p in products if p["name"]==a); B=next(p for p in products if p["name"]==b)
            A_h2=A["H2"] if not A["H2"].empty else pd.DataFrame(columns=["Hauptfunktion","Nebenfunktion","Kosten Nebenfunktion"])
            B_h2=B["H2"] if not B["H2"].empty else pd.DataFrame(columns=["Hauptfunktion","Nebenfunktion","Kosten Nebenfunktion"])
            h2=pd.merge(A_h2.rename(columns={"Kosten Nebenfunktion":"Cost_A"}),
                        B_h2.rename(columns={"Kosten Nebenfunktion":"Cost_B"}),
                        on=["Hauptfunktion","Nebenfunktion"], how="outer").fillna(0.0)
            if h2.empty:
                st.info("Keine Nebenfunktionen erkannt.")
            else:
                h2["Delta (B - A)"]=h2["Cost_B"]-h2["Cost_A"]
                h2["Delta_abs"]=h2["Delta (B - A)"].abs()
                topn=st.slider("Top-N",5,30,10)
                top=h2.sort_values("Delta_abs", ascending=False).head(topn).copy()
                top.insert(0,"Rang", range(1,len(top)+1))
                top["key"]=top["Hauptfunktion"].astype(str)+" > "+top["Nebenfunktion"].astype(str)
                tbl=top[["Rang","key","Cost_A","Cost_B","Delta (B - A)"]].rename(columns={"key":"Hauptfunktion > Nebenfunktion"})
                st.dataframe(tbl, use_container_width=True, height=300)
                top["Label"]=top["Rang"].astype(str)+". "+top["key"]
                top["_pos"]=(top["Delta (B - A)"]>=0).astype(int)
                chart=(alt.Chart(top).mark_bar().encode(
                    x=alt.X("Label:N", sort=list(top["Label"]), axis=alt.Axis(labelAngle=0)),
                    y="Delta (B - A):Q",
                    color=alt.Color("_pos:O", scale=alt.Scale(domain=[0,1], range=[PRIMARY, BLUE]), legend=None)
                ).properties(width="container", height=360)
                ).configure_axis(grid=True, gridColor="#e6e6e6")
                st.altair_chart(chart, use_container_width=True)

with tab5:
    st.markdown('<div class="section-title">Kosten vs Gewichtung (Funktionsbaum + H1-Kosten)</div>', unsafe_allow_html=True)
    all_h1=sorted(set().union(*[set(ensure_h1_costs(p)["Hauptfunktion"]) for p in products]))
    if not all_h1:
        st.info("Keine Hauptfunktionen erkannt.")
    else:
        wdf = products[0]["weights"]
        if wdf.empty:
            st.info("Kein Funktionsbaum mit Gewichten gefunden (Tab 'Funktionsbaum').")
        else:
            base=wdf.set_index("Hauptfunktion").reindex(all_h1).fillna(0).reset_index()
            base["Serie"]="Gewichtung [%]"
            lines=[]
            for p in products:
                df=ensure_h1_costs(p)
                if df.empty: continue
                df=df.set_index("Hauptfunktion").reindex(all_h1).fillna(0).reset_index()
                df["Produkt"]=p["name"]; lines.append(df)
            if not lines:
                st.info("Keine H1-Kosten vorhanden.")
            else:
                costs=pd.concat(lines, ignore_index=True)
                bars=(alt.Chart(base).mark_bar().encode(
                    x=alt.X("Hauptfunktion:N", axis=alt.Axis(labelAngle=45)),
                    y=alt.Y("Gewichtung_%:Q", axis=alt.Axis(title="Gewichtung [%]")),
                    color=alt.Color("Serie:N", scale=alt.Scale(range=["#cfcfcf"]), legend=alt.Legend(title=" "))
                ))
                lines=(alt.Chart(costs).mark_line(point=True, strokeWidth=2).encode(
                    x="Hauptfunktion:N",
                    y=alt.Y("Kosten Hauptfunktion:Q", axis=alt.Axis(title="Kosten [€]")),
                    color=alt.Color("Produkt:N", legend=alt.Legend(title="Produkt"))
                ))
                combo=(bars + lines).resolve_scale(y='independent', color='independent').properties(width="container", height=400)
                st.altair_chart(combo, use_container_width=True)
                st.caption("Graue Balken = Funktions-Gewichtung [%] (Funktionsbaum). Linien = Kosten je Produkt [€].")
