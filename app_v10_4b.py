
# -*- coding: utf-8 -*-
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title='EFESO – Functional Cost Analysis TOOLSET', layout='wide')
st.markdown('# EFESO – Functional Cost Analysis TOOLSET')
st.markdown('Version v10.4b • Vorlage für Funktions- & Kostenanalyse')

files = st.file_uploader('Excel-Dateien (.xlsx/.xlsm) — je Produkt eine Datei', type=['xlsx','xlsm'], accept_multiple_files=True)
if not files:
    st.info('Bitte laden Sie eine oder mehrere Excel-Dateien hoch.')
    st.stop()

def col_label_to_idx(label: str) -> int:
    label = label.strip().upper()
    idx = 0
    for ch in label:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1

COL_START = col_label_to_idx('I')

def find_sheet(xls: pd.ExcelFile, candidates):
    for name in xls.sheet_names:
        low = name.lower().strip()
        for c in candidates:
            if c in low:
                return name
    return None

def parse_product(file) -> dict:
    xls = pd.ExcelFile(file)
    cost_sheet = find_sheet(xls, ['slave_funktions-kostenstruktur', 'slave_funktions_kostenstruktur', 'funk'])
    tech_sheet = find_sheet(xls, ['slave_techn.bewertung', 'slave_techn_bewertung', 'techn'])
    if not cost_sheet:
        raise ValueError('Kostenstruktur-Blatt nicht gefunden.')
    df = pd.read_excel(xls, sheet_name=cost_sheet, header=None)
    row_h1 = df.iloc[1, COL_START:]
    row_h2 = df.iloc[2, COL_START:]
    row_w1 = df.iloc[4, COL_START:]
    row_w2 = df.iloc[5, COL_START:]
    row_c1 = df.iloc[7, COL_START:]
    row_c2 = df.iloc[8, COL_START:]
    frame = pd.DataFrame({
        'col': np.arange(COL_START, COL_START + len(row_h1)),
        'H1': row_h1.values, 'H2': row_h2.values,
        'W1': row_w1.values, 'W2': row_w2.values,
        'C1': row_c1.values, 'C2': row_c2.values,
    })
    frame['H1'] = frame['H1'].replace('', np.nan).fillna(method='ffill')
    frame = frame[frame['H2'].astype(str).str.strip().ne('')]
    for col in ['W1','W2','C1','C2']:
        frame[col] = pd.to_numeric(frame[col], errors='coerce').fillna(0.0)

    h1_first = (frame.groupby('H1')[['C1','W1']].first()
                        .rename(columns={'C1':'H1Cost','W1':'H1Weight'}).reset_index())
    h2_table = frame[['H1','H2','W2','C2']].rename(columns={'W2':'H2Weight','C2':'H2Cost'})

    tech = None
    if tech_sheet:
        tdf = pd.read_excel(xls, sheet_name=tech_sheet, header=None)
        col_H2 = tdf.iloc[:, 1].astype(str).str.strip()
        col_score = pd.to_numeric(tdf.iloc[:, 17], errors='coerce')
        tech = pd.DataFrame({'H2': col_H2, 'TechScore': col_score})
        tech = tech[tech['H2'].isin(h2_table['H2'].astype(str))]
    else:
        tech = pd.DataFrame(columns=['H2','TechScore'])
    h2_table = h2_table.merge(tech, on='H2', how='left')
    return {'name': getattr(file, 'name', 'Produkt'), 'h1': h1_first, 'h2': h2_table}

products, errors = [], []
for f in files:
    try:
        products.append(parse_product(f))
    except Exception as e:
        errors.append(f"{getattr(f,'name','Datei')}: {e}")
if errors:
    st.warning('Einige Dateien konnten nicht gelesen werden:\n- ' + '\n- '.join(errors))
if not products:
    st.error('Keine gültigen Dateien verarbeitet.')
    st.stop()

product_names = [p['name'] for p in products]
name_to_idx = {p['name']: i for i,p in enumerate(products)}
palette = px.colors.qualitative.Set1 + px.colors.qualitative.Safe + px.colors.qualitative.Plotly

tab1, tab2, tab3, tab4 = st.tabs(['Funktionsmatrix', 'Funktionenkosten', 'Technik Bewertung', 'Top Kostenabweichung'])

with tab1:
    st.subheader('Funktionsmatrix')
    sel = st.selectbox('Produkt wählen', product_names, index=0)
    data = products[name_to_idx[sel]]
    h1_df, h2_df = data['h1'], data['h2']
    st.caption('Kacheln haben nur Rahmen (keine Farben). Prozentangabe = H2-Anteil innerhalb der H1.')
    H1s = h1_df['H1'].tolist()
    cols_per_row = 3
    for i in range(0, len(H1s), cols_per_row):
        cols = st.columns(cols_per_row, gap='large')
        for c, h1 in zip(cols, H1s[i:i+cols_per_row]):
            with c:
                w = float(h1_df.loc[h1_df['H1']==h1, 'H1Weight'].fillna(0).values[0])
                cst = float(h1_df.loc[h1_df['H1']==h1, 'H1Cost'].fillna(0).values[0])
                st.markdown(f"#### {h1}  <span style='font-size:0.9rem;color:#999'>(Gewicht: {w:.0f}%, Kosten: {cst:.2f})</span>", unsafe_allow_html=True)
                sub = h2_df[h2_df['H1']==h1][['H2','H2Weight']]
                if sub.empty:
                    st.write('–')
                else:
                    for _, r in sub.iterrows():
                        st.markdown(
                            f"""
                            <div style='border:1px solid #c9c9c9;border-radius:6px;padding:8px 10px;margin:6px 0;'>
                                <div style='display:flex;justify-content:space-between;align-items:center;'>
                                    <div style='font-size:0.95rem;'>{r['H2']}</div>
                                    <div style='font-size:0.9rem;color:#666'>{r['H2Weight']:.0f}%</div>
                                </div>
                            </div>
                            """, unsafe_allow_html=True)

    st.markdown('**Beschreibung / Kommentare – Funktionsmatrix**')
    st.text_area('', '', height=120)

with tab2:
    st.subheader('Funktionenkosten')
    sel2 = st.selectbox('Produkt wählen', product_names, index=0, key='sel_cost')
    data2 = products[name_to_idx[sel2]]
    h1, h2 = data2['h1'], data2['h2']
    fig1 = go.Figure()
    fig1.add_bar(x=h1['H1'], y=h1['H1Cost'], marker_color='#0B66FF')
    fig1.update_layout(height=320, margin=dict(l=20,r=20,t=30,b=60), xaxis_title='H1', yaxis_title='Kosten_H1')
    st.plotly_chart(fig1, use_container_width=True)

    st.caption('Drilldown: Kosten je Nebenfunktion (Zeile 8)')
    options = ['alle'] + h1['H1'].tolist()
    sel_h1 = st.selectbox('Hauptfunktion wählen', options, index=0, key='sel_h1')
    h2_view = h2 if sel_h1 == 'alle' else h2[h2['H1']==sel_h1]
    fig2 = go.Figure()
    h1_list = h1['H1'].tolist()
    h1_colors = {name: px.colors.qualitative.Set3[i % len(px.colors.qualitative.Set3)] for i, name in enumerate(h1_list)}
    for group, sub in h2_view.groupby('H1'):
        fig2.add_bar(x=sub['H2'], y=sub['H2Cost'], name=group, marker_color=h1_colors.get(group, '#999'))
    fig2.update_layout(barmode='group', height=400, margin=dict(l=20,r=20,t=10,b=150))
    fig2.update_xaxes(tickangle=45)
    fig2.update_yaxes(title='Kosten_H2')
    st.plotly_chart(fig2, use_container_width=True)

with tab3:
    st.subheader('Technische Bewertung – Nebenfunktionen (H2)')
    all_h2 = pd.Index([])
    for p in products:
        all_h2 = all_h2.union(p['h2']['H2'].astype(str).dropna().unique())
    all_h2 = pd.Index(sorted(all_h2))

    fig3 = go.Figure()
    for i, p in enumerate(products):
        df = p['h2'][['H2','TechScore']].copy()
        df = df.groupby('H2', as_index=False).agg({'TechScore':'mean'})
        df = df.set_index('H2').reindex(all_h2).reset_index()
        fig3.add_trace(go.Scatter(x=df['H2'], y=df['TechScore'], mode='lines+markers', name=p['name'], line=dict(color=(px.colors.qualitative.Set1 + px.colors.qualitative.Safe + px.colors.qualitative.Plotly)[i % 28], width=2)))
    fig3.update_layout(height=340, margin=dict(l=20,r=20,t=20,b=160))
    fig3.update_xaxes(tickangle=45)
    fig3.update_yaxes(title='TechScore')
    st.plotly_chart(fig3, use_container_width=True)

    st.caption('Kosten je Nebenfunktion – Linien je Produkt')
    fig4 = go.Figure()
    for i, p in enumerate(products):
        dfc = p['h2'][['H2','H2Cost']].copy()
        dfc = dfc.groupby('H2', as_index=False).agg({'H2Cost':'sum'}).set_index('H2').reindex(all_h2).reset_index()
        fig4.add_trace(go.Scatter(x=dfc['H2'], y=dfc['H2Cost'], mode='lines+markers', name=p['name'], line=dict(color=(px.colors.qualitative.Set1 + px.colors.qualitative.Safe + px.colors.qualitative.Plotly)[i % 28], width=2, dash='dot')))
    fig4.update_layout(height=340, margin=dict(l=20,r=20,t=20,b=160))
    fig4.update_xaxes(tickangle=45)
    fig4.update_yaxes(title='Kosten_H2')
    st.plotly_chart(fig4, use_container_width=True)

with tab4:
    st.subheader('Top Kostenabweichungen – Nebenfunktionen (H2)')
    colA, colB = st.columns(2)
    with colA:
        A = st.selectbox('Produkt A', product_names, index=0, key='A_sel')
    with colB:
        B = st.selectbox('Produkt B', product_names, index=min(1, len(product_names)-1), key='B_sel')
    pa = products[name_to_idx[A]]['h2'][['H2','H2Cost']].groupby('H2', as_index=False).sum().rename(columns={'H2Cost':'CostA'})
    pb = products[name_to_idx[B]]['h2'][['H2','H2Cost']].groupby('H2', as_index=False).sum().rename(columns={'H2Cost':'CostB'})
    comp = pa.merge(pb, on='H2', how='outer').fillna(0.0)
    comp['Delta'] = (comp['CostA'] - comp['CostB']).abs()
    comp = comp.sort_values('Delta', ascending=False)
    top10 = comp.head(10)
    fig5 = go.Figure()
    fig5.add_bar(x=top10['H2'], y=top10['Delta'], marker_color='#0B66FF')
    fig5.update_layout(height=360, margin=dict(l=20,r=20,t=20,b=150))
    fig5.update_xaxes(tickangle=35); fig5.update_yaxes(title='Delta (|CostA - CostB|)')
    st.plotly_chart(fig5, use_container_width=True)
    st.caption('Ranking – größte zu kleinste Abweichung')
    st.dataframe(top10.reset_index(drop=True), use_container_width=True)

st.markdown('© EFESO • Version v10.4b – Vorlage für Funktions- & Kostenanalyse')
