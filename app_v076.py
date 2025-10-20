
import os, re, base64
import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

PRIMARY = "#f47c24"; DARK="#2a2a2a"
st.set_page_config(page_title="FKA v07.6 - EFESO", layout="wide", page_icon="📊")

def _efeso_logo_b64():
    svg_inline = "<svg xmlns='http://www.w3.org/2000/svg' width='220' height='36' viewBox='0 0 220 36'><rect width='220' height='36' fill='#ffffff'/><circle cx='18' cy='18' r='10' fill='#f47c24'/><text x='36' y='24' font-family='Arial, Helvetica, sans-serif' font-size='18' fill='#2a2a2a'>EFESO MANAGEMENT</text></svg>"
    import base64
    return "data:image/svg+xml;base64," + base64.b64encode(svg_inline.encode("utf-8")).decode()

st.markdown("""
<div style="position:sticky;top:0;z-index:999;background:#fff;padding:10px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,.06);display:flex;gap:10px;align-items:center;margin-bottom:8px">
  <img src='{_efeso_logo_b64()}' style='height:36px'/>
  <div>
    <div style="font-weight:700;font-size:18px;color:#2a2a2a">FKA - Funktionskosten & Evaluierung <small>v07.6</small></div>
    <div style="color:#666">Fix: H1-Kosten=Zeile 7 · H2-Kosten=Zeile 8 · Technik: Spalte R zu Nebenfunktionen (B) · H1-Aggregat aus H2</div>
  </div>
</div>
""", unsafe_allow_html=True)

# --- Upload
files = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei",
                         type=["xlsx","xlsm"], accept_multiple_files=True)
if not files:
    st.info("Bitte mind. eine Excel-Datei hochladen.")
    st.stop()

# --- Fixed layout (0-based indices)
START_COL = 8           # I
H1_ROW = 0              # row1: H1
H2_ROW = 1              # row2: H2
H1_WEIGHT_ROW = 3       # row4: H1 weight
H2_WEIGHT_ROW = 4       # row5: H2 weight
ROW_H1_COST = 6         # row7: H1 cost
ROW_H2_COST = 7         # row8: H2 cost

import re
_nonnum = re.compile(r"[^0-9,\.-]+")
def to_float(x):
    if x is None: return np.nan
    s = str(x).strip()
    if not s or s.lower() in ("nan", "none"): return np.nan
    s = _nonnum.sub("", s)
    s = s.replace(".", "").replace(",", ".") if s.count(",")>0 and s.count(".")>0 else s.replace(",", ".")
    try: return float(s)
    except: return np.nan

def valid_label(s):
    if s is None: return False
    s = str(s).strip()
    if not s or s.lower() in ("nan","none"): return False
    return bool(re.search(r"[A-Za-zÄÖÜäöüß]", s))

def scan_h1_blocks(df, start_col=START_COL):
    ncols = df.shape[1]
    blocks, current, cstart = [], None, None
    for c in range(start_col, ncols):
        val = df.iat[H1_ROW, c]
        if valid_label(val):
            if current is not None:
                blocks.append((current, cstart, c-1))
            current, cstart = str(val).strip(), c
    if current is not None:
        blocks.append((current, cstart, ncols-1))
    return blocks

def parse_product(file):
    raw = file.name.rsplit('.',1)[0]
    safe = re.sub(r"[^A-Za-z0-9_\-]+","_", raw).strip("_") or raw
    name = safe

    xl = pd.ExcelFile(file)

    # --- function sheet
    func_sheet = None
    for n in xl.sheet_names:
        if "slave_funktions" in n.lower(): func_sheet = n; break
    if func_sheet is None:
        func_sheet = xl.sheet_names[0]
    df = xl.parse(func_sheet, header=None)

    blocks = scan_h1_blocks(df)

    # Weights & costs & mapping H2->H1
    H1w_rows, H2w_rows, H1c_rows, H2c_rows, map_rows = [], [], [], [], []
    for h1, c0, c1 in blocks:
        H1w_rows.append({"Hauptfunktion": h1, "Wichtigkeit H1_%": to_float(df.iat[H1_WEIGHT_ROW, c0])})
        H1c_rows.append({"Hauptfunktion": h1, "Kosten Hauptfunktion": to_float(df.iat[ROW_H1_COST, c0])})
        for c in range(c0, c1+1):
            h2 = df.iat[H2_ROW, c]
            if not valid_label(h2): 
                continue
            h2s = str(h2).strip()
            H2w_rows.append({"Hauptfunktion": h1, "Nebenfunktion": h2s, "Wichtigkeit H2_%": to_float(df.iat[H2_WEIGHT_ROW, c])})
            H2c_rows.append({"Hauptfunktion": h1, "Nebenfunktion": h2s, "Kosten Nebenfunktion": to_float(df.iat[ROW_H2_COST, c])})
            map_rows.append({"Nebenfunktion": h2s, "Hauptfunktion": h1})

    H1_weights = pd.DataFrame(H1w_rows)
    H1_costs   = pd.DataFrame(H1c_rows)
    H2_weights = pd.DataFrame(H2w_rows)
    H2_costs   = pd.DataFrame(H2c_rows)
    map_h2_to_h1 = pd.DataFrame(map_rows).drop_duplicates()

    # Recon
    H1_from_H2 = (H2_costs.groupby("Hauptfunktion", as_index=False)["Kosten Nebenfunktion"]
                  .sum().rename(columns={"Kosten Nebenfunktion": "Summe H2 (Zeile 8)"}))
    recon = pd.merge(H1_costs, H1_from_H2, on="Hauptfunktion", how="left")
    recon["Delta (H1-Z7 - Summe H2-Z8)"] = recon["Kosten Hauptfunktion"] - recon["Summe H2 (Zeile 8)"]

    # --- tech sheet
    tech_sheet=None
    for n in xl.sheet_names:
        ln=n.lower()
        if "techn" in ln or "bewert" in ln:
            tech_sheet=n
            if "slave_techn" in ln: break
    if tech_sheet is not None:
        tdf = xl.parse(tech_sheet, header=None)
        # B: col 1, R: col 17
        tech_h2 = tdf.iloc[:, [1,17]].copy()
        tech_h2.columns = ["Nebenfunktion", "TechScore"]
        tech_h2["Nebenfunktion"] = tech_h2["Nebenfunktion"].astype(str).str.strip()
        tech_h2 = tech_h2[tech_h2["Nebenfunktion"]!=""]
        tech_h2["TechScore"] = pd.to_numeric(tech_h2["TechScore"], errors="coerce")
        tech_h2 = tech_h2.dropna(subset=["TechScore"])
        # map to H1
        tech_h2 = pd.merge(tech_h2, map_h2_to_h1, on="Nebenfunktion", how="left")
        # agg per H1
        # ungewichtet
        tech_h1_mean = tech_h2.groupby("Hauptfunktion", as_index=False)["TechScore"].mean().rename(columns={"TechScore":"TechScore_H1_mean"})
        # H2-gewichtet (mit Wichtigkeit Zeile 5)
        tw = H2_weights.copy()
        tw["Wichtigkeit H2_%"] = pd.to_numeric(tw["Wichtigkeit H2_%"], errors="coerce").fillna(0.0)
        tech_w = pd.merge(tech_h2, tw, on=["Hauptfunktion","Nebenfunktion"], how="left")
        tech_w["w"] = tech_w["Wichtigkeit H2_%"] / 100.0
        tech_w["w"] = tech_w["w"].fillna(0.0)
        tech_w["wScore"] = tech_w["TechScore"] * tech_w["w"]
        tech_h1_w = (tech_w.groupby("Hauptfunktion", as_index=False)
                        .agg(wScore=("wScore","sum"), w=("w","sum")))
        tech_h1_w["TechScore_H1_weighted"] = tech_h1_w["wScore"] / tech_h1_w["w"].replace(0,np.nan)
        tech_h1 = pd.merge(tech_h1_mean, tech_h1_w[["Hauptfunktion","TechScore_H1_weighted"]], on="Hauptfunktion", how="left")
    else:
        tech_h2 = pd.DataFrame(columns=["Nebenfunktion","TechScore","Hauptfunktion"])
        tech_h1 = pd.DataFrame(columns=["Hauptfunktion","TechScore_H1_mean","TechScore_H1_weighted"])

    return {
        "name": name, "func_sheet": func_sheet, "tech_sheet": tech_sheet,
        "H1_weights": H1_weights, "H2_weights": H2_weights,
        "H1_costs": H1_costs, "H2_costs": H2_costs, "recon": recon,
        "tech_h2": tech_h2, "tech_h1": tech_h1
    }

products = [parse_product(f) for f in files]
names = [p["name"] for p in products]

tab0, tab1, tab2, tab3 = st.tabs(["0) Wichtigkeiten", "1) Überblick", "2) Technik (H2 & H1)", "3) Abgleich"])

with tab0:
    psel = st.selectbox("Produkt wählen", names, key="wsel")
    P = next(p for p in products if p["name"] == psel)
    st.subheader("Wichtigkeit Matrix (H2 je H1)")
    H2w = P["H2_weights"]
    if H2w.empty:
        st.info("Keine H2-Wichtigkeiten gefunden.")
    else:
        H2w["Wichtigkeit H2_%"] = pd.to_numeric(H2w["Wichtigkeit H2_%"], errors="coerce").fillna(0.0)
        h1_order = H2w["Hauptfunktion"].astype(str).drop_duplicates().tolist()
        h2_order = H2w["Nebenfunktion"].astype(str).value_counts().index.tolist()
        heat = alt.Chart(H2w).mark_rect().encode(
            x=alt.X("Hauptfunktion:N", sort=h1_order),
            y=alt.Y("Nebenfunktion:N", sort=h2_order),
            color=alt.Color("Wichtigkeit H2_%:Q", scale=alt.Scale(domain=[0,100], scheme="oranges"), title="Wichtigkeit [%]"),
            tooltip=["Hauptfunktion","Nebenfunktion", alt.Tooltip("Wichtigkeit H2_%:Q", format=".1f")]
        ).properties(height=430)
        lbl = alt.Chart(H2w).mark_text(baseline='middle').encode(
            x=alt.X("Hauptfunktion:N", sort=h1_order),
            y=alt.Y("Nebenfunktion:N", sort=h2_order),
            text=alt.Text("Wichtigkeit H2_%:Q", format=".0f"),
            color=alt.condition(alt.datum["Wichtigkeit H2_%"] >= 45, alt.value("white"), alt.value("#2a2a2a"))
        )
        st.altair_chart(heat + lbl, use_container_width=True)

with tab1:
    psel = st.selectbox("Produkt wählen", names, key="usel")
    P = next(p for p in products if p["name"] == psel)
    st.subheader("Kosten je Hauptfunktion (Zeile 7)")
    st.altair_chart(alt.Chart(P["H1_costs"]).mark_bar().encode(
        x=alt.X("Hauptfunktion:N", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("Kosten Hauptfunktion:Q", title="Kosten [€]")
    ).properties(height=320), use_container_width=True)

    st.subheader("Drilldown: Kosten je Nebenfunktion (Zeile 8)")
    H2c = P["H2_costs"]
    if not H2c.empty:
        sel = st.selectbox("Hauptfunktion", H2c["Hauptfunktion"].drop_duplicates().tolist())
        part = H2c[H2c["Hauptfunktion"]==sel]
        st.altair_chart(alt.Chart(part).mark_bar().encode(
            x=alt.X("Nebenfunktion:N", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Kosten Nebenfunktion:Q", title="Kosten [€]")
        ).properties(height=300), use_container_width=True)
        st.dataframe(part, use_container_width=True, height=260)

with tab2:
    psel = st.selectbox("Produkt wählen", names, key="tsel")
    P = next(p for p in products if p["name"] == psel)

    st.subheader("Technische Bewertung - Nebenfunktionen (B -> R)")
    t2 = P["tech_h2"]
    if t2.empty:
        st.info("Kein Technikblatt gefunden oder keine Werte in Spalte R.")
    else:
        st.dataframe(t2, use_container_width=True, height=340)
        st.altair_chart(alt.Chart(t2).mark_bar(color=PRIMARY).encode(
            x=alt.X("Nebenfunktion:N", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("TechScore:Q", title="Tech Score (-2..+2)"),
            color=alt.Color("Hauptfunktion:N", legend=None)
        ).properties(height=320), use_container_width=True)

    st.subheader("Technische Bewertung - Aggregiert nach Hauptfunktion")
    t1 = P["tech_h1"]
    if not t1.empty:
        t1melt = t1.melt(id_vars="Hauptfunktion", var_name="Metrik", value_name="Score")
        st.altair_chart(alt.Chart(t1melt).mark_bar().encode(
            x=alt.X("Hauptfunktion:N", axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Score:Q"),
            color=alt.Color("Metrik:N"),
            tooltip=["Hauptfunktion","Metrik", alt.Tooltip("Score:Q", format=".3f")]
        ).properties(height=320), use_container_width=True)
        st.dataframe(t1, use_container_width=True, height=240)

with tab3:
    psel = st.selectbox("Produkt wählen", names, key="rsel")
    P = next(p for p in products if p["name"] == psel)
    st.subheader("Abgleich H1 (Zeile 7) vs. Summe H2 (Zeile 8)")
    st.dataframe(P["recon"], use_container_width=True, height=360)
