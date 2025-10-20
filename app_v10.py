
# EFESO Functional Cost Analysis TOOLSET – v10
# Tabs: Funktionsmatrix (Swimlane), Funktionenkosten, Technik Bewertung, Top Kostenabweichung

import math
from typing import Any, Dict, List

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

st.set_page_config(page_title="EFESO Functional Cost Analysis TOOLSET v10", layout="wide")

st.markdown(
    '''
    <h2 style="margin-bottom:0.25rem;">EFESO Functional Cost Analysis TOOLSET</h2>
    <div style="opacity:.7;margin-bottom:.75rem;">v10 · Funktionsmatrix · Kosten · Technik · Abweichungen</div>
    ''',
    unsafe_allow_html=True,
)

# -------------------- Helpers --------------------

def _get_cell(df: pd.DataFrame, r: int, c: int):
    try:
        v = df.iat[r, c]
    except Exception:
        return None
    if pd.isna(v) or v == "":
        return None
    return v

def _to_float(x) -> float:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return float("nan")
    try:
        s = str(x).strip()
        if s.endswith("%"):
            s = s[:-1].replace(",", ".")
            return float(s)
        return float(s.replace(",", "."))
    except Exception:
        return float("nan")

def parse_structure_from_sheet(df: pd.DataFrame) -> pd.DataFrame:
    '''
    Erwartet Blatt 'SLAVE_Funktions-Kostenstruktur' ohne Header.
    Zeile 1  (r0): H1-Namen ab Spalte I (c8)
    Zeile 2  (r1): H2-Namen
    Zeile 4/5 (r3/r4): Gewichte der H2 in % (präferiert r4)
    Zeile 7  (r6): H1-Kosten
    Zeile 8  (r7): H2-Kosten
    '''
    start_col = 8
    R_H1, R_H2 = 0, 1
    R_W1, R_W2 = 3, 4
    R_C1, R_C2 = 6, 7

    rows = []
    empty_streak = 0
    cur_h1 = None
    cur_h1_cost = np.nan

    for c in range(start_col, start_col + 400):
        h1 = _get_cell(df, R_H1, c)
        h2 = _get_cell(df, R_H2, c)
        if not h1 and not h2:
            empty_streak += 1
            if empty_streak >= 12:
                break
            continue
        empty_streak = 0

        if h1:
            cur_h1 = str(h1).strip()
            cur_h1_cost = _to_float(_get_cell(df, R_C1, c))

        if h2 and cur_h1:
            w = _to_float(_get_cell(df, R_W2, c))
            if math.isnan(w):
                w = _to_float(_get_cell(df, R_W1, c))
            h2_cost = _to_float(_get_cell(df, R_C2, c))

            rows.append(
                {
                    "Hauptfunktion": cur_h1,
                    "Nebenfunktion": str(h2).strip(),
                    "H2_Gewicht_%": w,   # 0..100
                    "H2_Kosten": h2_cost,
                    "H1_Kosten": cur_h1_cost,
                }
            )
    out = pd.DataFrame(rows)
    if not out.empty:
        out["H2_Gewicht_%"] = out["H2_Gewicht_%"].fillna(0.0)
        if out["H2_Gewicht_%"].max() <= 1.0:
            out["H2_Gewicht_%"] = out["H2_Gewicht_%"] * 100.0
    return out

def parse_tech_from_sheet(df: pd.DataFrame, map_h2_to_h1: Dict[str, str]) -> pd.DataFrame:
    '''
    Versucht, Technik-Bewertung aus einem Blatt 'SLAVE_Techn.Bewertung' zu lesen.
    Heuristik (ohne Header):
      - Spalte B (c1): Nebenfunktionsnamen
      - Spalte R (c17): Score (z.B. -2..+2)
      - Spalte M (c12): Gewicht in % (optional)
    '''
    data = []
    for r in range(4, min(120, len(df))):
        name = _get_cell(df, r, 1)
        if name is None:
            continue
        name = str(name).strip()
        score = _to_float(_get_cell(df, r, 17))
        weight = _to_float(_get_cell(df, r, 12))
        if math.isnan(score):
            continue
        if math.isnan(weight):
            weight = float("nan")
        h1 = map_h2_to_h1.get(name, None)
        data.append({"Nebenfunktion": name, "TechScore": score, "Gew_%": weight, "Hauptfunktion": h1})
    return pd.DataFrame(data)

def load_products(files) -> List[Dict[str, Any]]:
    prods = []
    for f in files:
        name = f.name.rsplit('.', 1)[0]
        try:
            df_struct = pd.read_excel(f, sheet_name='SLAVE_Funktions-Kostenstruktur', header=None, engine='openpyxl')
        except Exception as e:
            st.warning(f"⚠️ {name}: Konnte Blatt 'SLAVE_Funktions-Kostenstruktur' nicht lesen ({e}).")
            continue
        df_long = parse_structure_from_sheet(df_struct)
        if df_long.empty:
            st.warning(f"⚠️ {name}: Keine Funktionsstruktur erkannt.")
            continue

        # Map H2->H1 für Technik
        h2_to_h1 = dict(zip(df_long['Nebenfunktion'], df_long['Hauptfunktion']))

        # Technik optional
        tech = pd.DataFrame()
        try:
            df_tech = pd.read_excel(f, sheet_name='SLAVE_Techn.Bewertung', header=None, engine='openpyxl')
            tech = parse_tech_from_sheet(df_tech, h2_to_h1)
        except Exception:
            pass

        prods.append({'name': name, 'structure': df_long, 'tech': tech})
    return prods

def bar_chart(df: pd.DataFrame, x: str, y: str, title: str = "", orient: str = "x", sort=None, color=None):
    if df.empty:
        return None
    if sort is None:
        sort = "-y" if orient == "x" else "-x"
    enc = {"x": alt.X(x, sort=sort), "y": alt.Y(y)} if orient == "x" else {"x": alt.X(x), "y": alt.Y(y, sort=sort)}
    if color:
        enc["color"] = alt.Color(color)
    chart = alt.Chart(df).mark_bar().encode(**enc).properties(height=320, title=title)
    return chart

# -------------------- Upload & Tabs --------------------

left, right = st.columns([1, 3])
with left:
    st.caption("Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei")
    files = st.file_uploader("Drag and drop files here", type=["xlsx", "xlsm"], accept_multiple_files=True)

tabs = st.tabs(["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"])
tab_matrix, tab_costs, tab_tech, tab_top = tabs

# -------------- Funktionsmatrix (Swimlane) --------------

with tab_matrix:
    st.subheader("Funktionsmatrix – Swimlane (H2 je H1, nach Gewicht sortiert)")
    if not files:
        st.info("Bitte Excel-Dateien hochladen.")
        st.stop()
    products = load_products(files)
    if not products:
        st.stop()
    names = [p["name"] for p in products]
    pname = st.selectbox("Produkt wählen", names, key="swim_prod")
    P = next(p for p in products if p["name"] == pname)
    df = P["structure"].copy()

    groups = df.groupby("Hauptfunktion", sort=False)
    h1s = list(groups.groups.keys())

    def chunk(lst, n):
        for i in range(0, len(lst), n):
            yield lst[i:i+n]

    for row in chunk(h1s, 3):
        cols = st.columns(len(row))
        for ci, h1 in enumerate(row):
            with cols[ci]:
                st.markdown(f"#### {h1}")
                g = groups.get_group(h1).sort_values("H2_Gewicht_%", ascending=False)
                for _, r in g.iterrows():
                    name = r["Nebenfunktion"]
                    pct = r["H2_Gewicht_%"]
                    cost = r.get("H2_Kosten", np.nan)
                    st.markdown(
                        f'''
                        <div style="border:1px solid rgba(0,0,0,.15);border-radius:6px;padding:.45rem .6rem;margin:.35rem 0;">
                          <div style="display:flex;justify-content:space-between;gap:.75rem;">
                            <div><b>{name}</b></div>
                            <div style="white-space:nowrap;">{pct:.0f}%{" · € {:,.0f}".format(cost) if not pd.isna(cost) else ""}</div>
                          </div>
                        </div>
                        ''',
                        unsafe_allow_html=True,
                    )
                st.divider()

# -------------- Funktionenkosten --------------

with tab_costs:
    st.subheader("Funktionenkosten")
    if not files:
        st.info("Bitte Excel-Dateien hochladen.")
        st.stop()
    if "products_cache" not in st.session_state:
        st.session_state["products_cache"] = load_products(files)
    products = st.session_state["products_cache"]
    if not products:
        st.stop()
    pname = st.selectbox("Produkt wählen", [p["name"] for p in products], key="cost_prod")
    P = next(p for p in products if p["name"] == pname)
    df = P["structure"].copy()

    # H1-Kosten (Zeile 7)
    df_h1 = df[["Hauptfunktion", "H1_Kosten"]].dropna().copy()
    df_h1 = df_h1.groupby("Hauptfunktion", as_index=False)["H1_Kosten"].max()
    c1 = bar_chart(df_h1, x="Hauptfunktion", y="H1_Kosten", title="Kosten je Hauptfunktion (Zeile 7)")
    if c1 is not None:
        st.altair_chart(c1, use_container_width=True)

    st.markdown("**Drilldown: Kosten je Nebenfunktion (Zeile 8)**")
    h1_sel = st.selectbox("Hauptfunktion wählen", df["Hauptfunktion"].unique(), key="h1_for_h2")
    df_h2 = df[df["Hauptfunktion"] == h1_sel][["Nebenfunktion", "H2_Kosten"]].dropna()
    c2 = bar_chart(df_h2, x="Nebenfunktion", y="H2_Kosten", title=f"{h1_sel} – H2-Kosten", sort="-y")
    if c2 is not None:
        st.altair_chart(c2, use_container_width=True)

# -------------- Technik Bewertung --------------

with tab_tech:
    st.subheader("Technische Bewertung (H2 → Score & gewichtet)")
    if not files:
        st.info("Bitte Excel-Dateien hochladen.")
        st.stop()
    if "products_cache" not in st.session_state:
        st.session_state["products_cache"] = load_products(files)
    products = st.session_state["products_cache"]
    if not products:
        st.stop()

    pname = st.selectbox("Produkt wählen", [p["name"] for p in products], key="tech_prod")
    P = next(p for p in products if p["name"] == pname)
    tech = P["tech"]
    if tech.empty:
        st.warning("Kein Technik-Blatt gefunden oder nicht lesbar (erwartet: 'SLAVE_Techn.Bewertung').")
        st.stop()

    # Weighted Score (wenn Gewicht vorhanden)
    tech["Gew_%_norm"] = tech["Gew_%"].apply(lambda v: v / 100.0 if (isinstance(v, (float, int)) and not math.isnan(v)) else np.nan)
    tech["TechScore_weighted"] = np.where(tech["Gew_%_norm"].notna(), tech["TechScore"] * tech["Gew_%_norm"], tech["TechScore"])

    c3 = bar_chart(tech[["Nebenfunktion", "TechScore"]], x="Nebenfunktion", y="TechScore", title="Technischer Score je Nebenfunktion", sort="-y")
    if c3 is not None:
        st.altair_chart(c3, use_container_width=True)

    # Aggregiert je H1
    if "Hauptfunktion" in tech.columns:
        agg = tech.groupby("Hauptfunktion", as_index=False)["TechScore_weighted"].sum()
        c4 = bar_chart(agg, x="Hauptfunktion", y="TechScore_weighted", title="Technische Bewertung – aggregiert nach Hauptfunktion", sort="-y")
        if c4 is not None:
            st.altair_chart(c4, use_container_width=True)

# -------------- Top Kostenabweichung --------------

with tab_top:
    st.subheader("Top Kostenabweichung (Nebenfunktionen)")
    if not files:
        st.info("Bitte Excel-Dateien hochladen.")
        st.stop()
    if "products_cache" not in st.session_state:
        st.session_state["products_cache"] = load_products(files)
    products = st.session_state["products_cache"]
    if len(products) < 2:
        st.info("Bitte mindestens zwei Dateien hochladen.")
        st.stop()

    colA, colB = st.columns(2)
    with colA:
        pA = st.selectbox("Produkt A", [p["name"] for p in products], key="top_A")
    with colB:
        pB = st.selectbox("Produkt B", [p["name"] for p in products if p["name"] != pA], key="top_B")

    A = next(p for p in products if p["name"] == pA)["structure"]
    B = next(p for p in products if p["name"] == pB)["structure"]

    a = A.groupby("Nebenfunktion", as_index=False)["H2_Kosten"].sum().rename(columns={"H2_Kosten": "Cost_A"})
    b = B.groupby("Nebenfunktion", as_index=False)["H2_Kosten"].sum().rename(columns={"H2_Kosten": "Cost_B"})
    m = pd.merge(a, b, on="Nebenfunktion", how="outer").fillna(0.0)
    m["Delta"] = m["Cost_B"] - m["Cost_A"]
    m["abs"] = m["Delta"].abs()
    top = m.sort_values("abs", ascending=False).head(10)

    chart = alt.Chart(top).mark_bar().encode(
        x=alt.X("Nebenfunktion", sort=top["Nebenfunktion"].tolist()),
        y=alt.Y("Delta", title="Kosten-Differenz (B - A)"),
        color=alt.condition(alt.datum.Delta > 0, alt.value("#0b84a5"), alt.value("#f25f5c"))
    ).properties(height=360, title=f"Top 10 Abweichungen: {pB} vs {pA}")
    st.altair_chart(chart, use_container_width=True)

st.markdown('<div style="opacity:.6;margin-top:.75rem;">© EFESO – v10</div>', unsafe_allow_html=True)
