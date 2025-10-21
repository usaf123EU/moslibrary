# app_v10_4.py
# EFESO – Functional Cost Analysis TOOLSET (v10.4)
# Upload je Produkt eine Excel (.xlsx/.xlsm)
# Struktur: Tab "SLAVE_Funktions-Kostenstruktur" ab Spalte I (Startcol=8):
# Zeile 1=H1-Namen, Zeile 2=H2-Namen, Zeile 4=H1-Gewicht %, Zeile 5=H2-Gewicht %,
# Zeile 7=H1-Kosten, Zeile 8=H2-Kosten
# Tech: Tab "SLAVE_Techn.Bewertung", Spalte B=H2-Name, Spalte R=Score

import io, re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(page_title="EFESO – Functional Cost Analysis TOOLSET", layout="wide")

# ---------- UI Header ----------
st.markdown("# EFESO – Functional Cost Analysis TOOLSET")
st.caption("Version v10.4 • Vorlage für Funktions- & Kostenanalyse")

# ---------- Helpers ----------
def _find_sheet(xls, must_include_all):
    names = [n for n in xls.sheet_names]
    for n in names:
        low = n.lower()
        if all(k in low for k in must_include_all):
            return n
    return None

START_COL = 8  # Spalte I (0-based)
ROW_H1, ROW_H2, ROW_W_H1, ROW_W_H2, ROW_C_H1, ROW_C_H2 = 0, 1, 3, 4, 6, 7

def _as_float(x):
    if pd.isna(x): return np.nan
    if isinstance(x, str):
        x = x.replace("%", "").replace(",", ".").strip()
    try:
        return float(x)
    except:
        return np.nan

def parse_cost_structure(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = _find_sheet(xls, ["funktions", "kosten"])
    if sheet is None:
        raise ValueError("Tab 'SLAVE_Funktions-Kostenstruktur' nicht gefunden.")
    df = xls.parse(sheet_name=sheet, header=None, dtype=object).fillna("")

    # nach rechts laufen ab START_COL und H1-Blöcke lesen
    col = START_COL
    blocks = []
    h1_costs = []
    h2_rows = []
    max_cols = df.shape[1]

    while col < max_cols:
        h1 = str(df.iat[ROW_H1, col]).strip()
        if h1 == "" or h1.lower() == "nan":
            col += 1
            continue

        w_h1 = _as_float(df.iat[ROW_W_H1, col])
        if w_h1 > 1.01: w_h1 = w_h1/100.0

        c_h1 = _as_float(df.iat[ROW_C_H1, col])
        h1_costs.append({"H1": h1, "Cost_H1": float(0 if np.isnan(c_h1) else c_h1)})

        # H2 bis zum nächsten H1 sammeln
        sub_h2 = []
        c2 = col
        while c2 < max_cols:
            next_h1 = str(df.iat[ROW_H1, c2]).strip()
            if c2 != col and next_h1 != "":
                break
            h2_name = str(df.iat[ROW_H2, c2]).strip()
            if h2_name != "":
                w_h2 = _as_float(df.iat[ROW_W_H2, c2])
                if w_h2 > 1.01: w_h2 = w_h2/100.0
                c_h2 = _as_float(df.iat[ROW_C_H2, c2])
                sub_h2.append({"name": h2_name, "w": float(0 if np.isnan(w_h2) else w_h2)})
                h2_rows.append({"H1": h1, "H2": h2_name,
                                "Weight_H2": float(0 if np.isnan(w_h2) else w_h2),
                                "Cost_H2": float(0 if np.isnan(c_h2) else c_h2)})
            c2 += 1

        blocks.append({"h1": h1, "h1_weight": float(0 if np.isnan(w_h1) else w_h1), "h2": sub_h2})
        col = c2

    df_h1_costs = pd.DataFrame(h1_costs)
    df_h2_costs = pd.DataFrame(h2_rows)
    return blocks, df_h1_costs, df_h2_costs

def parse_tech_scores(file_bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet = _find_sheet(xls, ["techn", "bew"])
    if sheet is None:
        return pd.Series(dtype=float)
    df = xls.parse(sheet_name=sheet, header=None, dtype=object).fillna("")
    names = df.iloc[:, 1]
    scores = df.iloc[:, 17] if df.shape[1] > 17 else pd.Series([])
    pairs = []
    for n, s in zip(names, scores):
        n = str(n).strip()
        if n and n.lower() not in ["", "nan", "nebenfunktion", "technical evaluation"]:
            try:
                val = float(str(s).replace(",", "."))
                pairs.append((n, val))
            except:
                continue
    if not pairs:
        return pd.Series(dtype=float)
    ser = pd.Series(dict(pairs))
    return ser

H1_COLORS = ["#1f77b4", "#2ca02c", "#ff7f0e", "#9467bd", "#8c564b", "#17becf", "#7f7f7f"]
def h1_color(h1, order):
    if h1 not in order: order.append(h1)
    return H1_COLORS[order.index(h1) % len(H1_COLORS)]

# (Rest identisch mit vorheriger Version 10.4)
