
# EFESO Functional Cost Analysis TOOLSET — v10.3-fixed
# (border-only swimlanes, thin bars, tech line chart, top-10 deltas)

import io, re
from typing import Dict, Any
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title='EFESO Functional Cost Analysis TOOLSET – v10.3', layout='wide')

def xl_col_to_idx(label: str) -> int:
    s = label.strip().upper()
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n - 1

START_COL = xl_col_to_idx('I')

def _is_empty(v):
    if v is None: return True
    if isinstance(v, float) and np.isnan(v): return True
    if isinstance(v, str) and v.strip() == "": return True
    return False

def _clean_num(x):
    if pd.isna(x): return np.nan
    if isinstance(x, (int, float, np.number)): return float(x)
    s = str(x).strip().replace('€','').replace('%','').replace(',','.')
    s = re.sub(r'[^0-9\.\-]', '', s)
    try: return float(s)
    except: return np.nan

def parse_product(file_name: str, content: bytes) -> Dict[str, Any]:
    xl = pd.ExcelFile(io.BytesIO(content), engine='openpyxl')
    sh = 'SLAVE_Funktions-Kostenstruktur'
    df = xl.parse(sh, header=None)
    h1_row, h2_row = 0, 1
    h1_w_row, h2_w_row = 3, 4
    h1_c_row, h2_c_row = 6, 7
    h1_blocks = []
    col = START_COL
    maxc = df.shape[1]
    while col < maxc:
        h1_name = df.iat[h1_row, col] if col < maxc else None
        if _is_empty(h1_name):
            if all(_is_empty(df.iat[h2_row, c]) for c in range(col, min(col+5, maxc))):
                break
            col += 1
            continue
        nxt = col + 1
        while nxt < maxc and _is_empty(df.iat[h1_row, nxt]): nxt += 1
        block_cols = list(range(col, nxt))
        h1_weight = _clean_num(df.iat[h1_w_row, col])
        h1_cost   = _clean_num(df.iat[h1_c_row, col])
        h2s = []
        for c in block_cols:
            h2_name = df.iat[h2_row, c]
            if _is_empty(h2_name): continue
            h2s.append({'name': str(h2_name).strip(),
                        'weight': _clean_num(df.iat[h2_w_row, c]),
                        'cost': _clean_num(df.iat[h2_c_row, c]),
                        'col': c})
        h1_blocks.append({'name': str(h1_name).strip(),
                          'weight': h1_weight,
                          'cost': h1_cost,
                          'h2': h2s,
                          'start_col': col,
                          'end_col': nxt-1})
        col = nxt

    tech = {}
    tname = 'SLAVE_Techn.Bewertung'
    if tname in xl.sheet_names:
        tdf = xl.parse(tname, header=None)
        for r in range(tdf.shape[0]):
            name = tdf.iat[r, 1] if 1 < tdf.shape[1] else None
            score = tdf.iat[r, 17] if 17 < tdf.shape[1] else None
            if not _is_empty(name):
                tech[str(name).strip()] = _clean_num(score)

    h1_costs = pd.DataFrame([{'H1': h['name'], 'Kosten_H1': h['cost'], 'Gewichtung_H1_%': h['weight']} for h in h1_blocks]).dropna(how='all')
    rows = []
    for h in h1_blocks:
        for h2 in h['h2']:
            rows.append({'H1': h['name'], 'H2': h2['name'], 'Kosten_H2': h2['cost'], 'Gewichtung_H2_%': h2['weight']})
    h2_costs = pd.DataFrame(rows).dropna(how='all')
    tech_rows = []
    for h in h1_blocks:
        for h2 in h['h2']:
            tech_rows.append({'H1': h['name'], 'H2': h2['name'], 'TechScore': tech.get(h2['name'], np.nan)})
    h2_tech = pd.DataFrame(tech_rows)

    def _w_h1(group: pd.DataFrame):
        vals = group['TechScore'].astype(float)
        w = group['Gewichtung_H2_%'].astype(float)
        if w.notna().sum() > 0 and w.sum() > 0:
            return (vals.fillna(0) * (w.fillna(0)/100.0)).sum()
        return vals.mean()

    if not h2_costs.empty:
        h1_tech_weighted = (
            h2_costs.merge(h2_tech[['H1','H2','TechScore']], on=['H1','H2'], how='left')
                    .groupby('H1').apply(_w_h1).reset_index(name='TechScore_H1_weighted')
        )
    else:
        h1_tech_weighted = pd.DataFrame(columns=['H1','TechScore_H1_weighted'])

    return {'file': file_name,
            'h1_blocks': h1_blocks,
            'h1_costs': h1_costs,
            'h2_costs': h2_costs,
            'h2_tech': h2_tech,
            'h1_tech_weighted': h1_tech_weighted}

st.markdown("# EFESO – Functional Cost Analysis TOOLSET  
**Version v10.3 (fixed)**", unsafe_allow_html=True)

uploaded = st.file_uploader("Excel-Dateien (.xlsx/.xlsm) – je Produkt eine Datei", type=["xlsx","xlsm"], accept_multiple_files=True)

if 'products' not in st.session_state:
    st.session_state.products = {}

if uploaded:
    for uf in uploaded:
        try:
            st.session_state.products[uf.name] = parse_product(uf.name, uf.getvalue())
        except Exception as e:
            st.warning(f"{uf.name}: {e}")

if not st.session_state.products:
    st.info("Bitte Dateien hochladen.")
    st.stop()

product_names = list(st.session_state.products.keys())
tabs = st.tabs(["Funktionsmatrix", "Funktionenkosten", "Technik Bewertung", "Top Kostenabweichung"])

with tabs[0]:
    st.subheader("Funktionsmatrix (H1 → zugehörige Nebenfunktionen)")
    prod = st.selectbox("Produkt wählen", product_names, key="sel_prod_matrix")
    pdata = st.session_state.products[prod]
    st.markdown("""
    <style>
    .tile-col { border: 1px solid #e2e2e2; padding: 8px 10px; border-radius: 10px; background: #fff; }
    .tile      { border: 1px solid #bcbcbc; padding: 6px 8px; border-radius: 6px; margin: 6px 0; }
    .tile .name { font-size: 0.9rem; line-height: 1.2; }
    .tile .pct  { font-size: 0.85rem; color: #444; float: right; }
    .h1-title  { font-weight: 600; margin-bottom: 6px; }
    </style>
    """, unsafe_allow_html=True)

    h1_blocks_sorted = sorted(pdata['h1_blocks'], key=lambda x: (x['cost'] if x['cost'] is not None else -1), reverse=True)
    ncols = min(4, max(1, len(h1_blocks_sorted)))
    rows = [h1_blocks_sorted[i:i+ncols] for i in range(0, len(h1_blocks_sorted), ncols)]
    for row in rows:
        cols = st.columns(len(row), gap="large")
        for c, h1 in zip(cols, row):
            with c:
                st.markdown(f"<div class='tile-col'><div class='h1-title'>{h1['name']}</div>", unsafe_allow_html=True)
                h2s_sorted = sorted(h1['h2'], key=lambda x: (x['weight'] if x['weight'] is not None else -1), reverse=True)
                for h2 in h2s_sorted:
                    pct = f"{h2['weight']:.0f}%" if pd.notna(h2['weight']) else ""
                    st.markdown(f"<div class='tile'><span class='name'>{h2['name']}</span><span class='pct'>{pct}</span></div>", unsafe_allow_html=True)
                st.markdown("</div>", unsafe_allow_html=True)

with tabs[1]:
    st.subheader("Funktionenkosten")
    prod = st.selectbox("Produkt wählen", product_names, key="sel_prod_costs")
    pdata = st.session_state.products[prod]
    if not pdata['h1_costs'].empty:
        fig1 = px.bar(pdata['h1_costs'].sort_values('Kosten_H1', ascending=False), x='H1', y='Kosten_H1', title='Kosten je Hauptfunktion (Zeile 7)')
        fig1.update_layout(bargap=0.6, bargroupgap=0.7, showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)
    st.caption("Drilldown: Kosten je Nebenfunktion (Zeile 8)")
    h1s = [h['name'] for h in pdata['h1_blocks']]
    sel_h1 = st.selectbox("Hauptfunktion wählen", ["alle"] + h1s, key="h1_dd")
    df_h2 = pdata['h2_costs'].copy()
    if sel_h1 != "alle":
        df_h2 = df_h2[df_h2['H1'] == sel_h1]
    if not df_h2.empty:
        fig2 = px.bar(df_h2.sort_values('Kosten_H2', ascending=False), x='H2', y='Kosten_H2', color='H1')
        fig2.update_layout(bargap=0.6, bargroupgap=0.7, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

with tabs[2]:
    st.subheader("Technische Bewertung – Nebenfunktionen (H2)")
    blocks = []
    for pname, p in st.session_state.products.items():
        blocks.append(p['h2_tech'][['H2','TechScore']].assign(Produkt=pname))
    if blocks:
        tech_all = pd.concat(blocks, ignore_index=True)
        order = tech_all.groupby('H2')['TechScore'].mean().sort_values(ascending=False).index.tolist()
        tech_all['H2'] = pd.Categorical(tech_all['H2'], categories=order, ordered=True)
        fig = px.line(tech_all.sort_values('H2'), x='H2', y='TechScore', color='Produkt', markers=True)
        fig.update_layout(xaxis_tickangle=-45, legend_title_text='Produkt')
        st.plotly_chart(fig, use_container_width=True)

with tabs[3]:
    st.subheader("Top Kostenabweichungen – Nebenfunktionen (H2)")
    if len(product_names) < 2:
        st.info("Bitte mindestens zwei Produkte hochladen.")
    else:
        a = st.selectbox("Produkt A", product_names, index=0, key="prodA")
        b = st.selectbox("Produkt B", product_names, index=1, key="prodB")
        A = st.session_state.products[a]['h2_costs'][['H2','Kosten_H2']].rename(columns={'Kosten_H2':'A'})
        B = st.session_state.products[b]['h2_costs'][['H2','Kosten_H2']].rename(columns={'Kosten_H2':'B'})
        comp = A.merge(B, on='H2', how='outer')
        comp['A'] = comp['A'].astype(float)
        comp['B'] = comp['B'].astype(float)
        comp['Delta'] = (comp['A'] - comp['B']).abs()
        top10 = comp.sort_values('Delta', ascending=False).head(10)
        if not top10.empty:
            fig = px.bar(top10, x='H2', y='Delta', title=f'Top 10 Abweichungen: {a} vs {b}')
            fig.update_layout(bargap=0.6, bargroupgap=0.7, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
