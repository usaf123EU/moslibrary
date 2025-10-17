import io
import re
import pandas as pd
import numpy as np
import streamlit as st

st.set_page_config(page_title="FKA Prototype — Function Cost & Tech Scoring", layout="wide")

st.title("FKA Prototype — Function Cost & Tech Scoring")
st.caption("Upload one or more Excel files (.xlsx/.xlsm) with sheets similar to SLAVE_START, SLAVE_Funktions-Kostenstruktur, SLAVE_Techn.Bewertung. The app consolidates function costs (H1/H2/H3) and technical scores across products.")

files = st.file_uploader("Upload Excel files (multiple allowed)", type=["xlsx","xlsm"], accept_multiple_files=True)

#############################
# Helpers
#############################
HIER_COL_CANDIDATES = ["H1","H2","H3","Hauptfunktion","Teilfunktion","Unterfunktion","F1","F2","F3"]
FUNC_SHEET_HINTS = ["SLAVE_Funktions", "Funktionen", "Funktion", "Function", "Kosten"]
TECH_SHEET_HINTS = ["Techn", "Bewertung", "Rating", "Score"]
START_SHEET_HINTS = ["SLAVE_START", "Start", "START"]

def _pick_sheet(xl, hints):
    names = xl.sheet_names
    # prioritize exact-like matches
    for n in names:
        for h in hints:
            if n.lower() == h.lower():
                return n
    # then fuzzy contains
    for n in names:
        for h in hints:
            if h.lower() in n.lower():
                return n
    # fallback: first sheet
    return names[0]

def _coerce_numeric(df):
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = pd.to_numeric(df[c], errors="ignore")
    return df

def _possible_hierarchy_cols(df):
    cols = []
    for c in df.columns:
        if any(k.lower() in c.lower() for k in ["h1","h2","h3","haupt","teil","unter","funktion","func"]):
            cols.append(c)
    # keep order by appearance
    uniq = []
    for c in cols:
        if c not in uniq:
            uniq.append(c)
    return uniq[:3]

def parse_function_costs(df_raw):
    # try to detect hierarchy columns and cost columns
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    hier_cols = _possible_hierarchy_cols(df)
    if not hier_cols:
        return None, "No hierarchy columns (H1/H2/H3) detected"
    # limit to 3 levels where possible
    if len(hier_cols) == 1:
        hier_cols = [hier_cols[0]]
    elif len(hier_cols) == 2:
        hier_cols = hier_cols[:2]
    else:
        hier_cols = hier_cols[:3]

    # cost columns = numeric columns that are not hierarchy
    non_hier = [c for c in df.columns if c not in hier_cols]
    # Try to keep only numeric-esque columns for cost
    cost_candidates = []
    for c in non_hier:
        if pd.api.types.is_numeric_dtype(df[c]):
            cost_candidates.append(c)
        else:
            # try parse numbers from strings (e.g., "1.234,56" or "1234.56")
            as_num = pd.to_numeric(df[c].astype(str).str.replace(".","", regex=False).str.replace(",",".", regex=False), errors="coerce")
            if as_num.notna().sum() > 0:
                df[c] = as_num
                cost_candidates.append(c)
    if not cost_candidates:
        # If still nothing, just return as-is
        return None, "No numeric cost columns detected"
    df["__row_cost__"] = df[cost_candidates].sum(axis=1, skipna=True)

    # clean hierarchy text
    for h in hier_cols:
        df[h] = df[h].astype(str).str.strip()
        df.loc[df[h].isin(["","nan","None"]),"__isblank__"] = True

    # Roll-ups
    roll = {}
    def _roll(group_cols):
        g = df.groupby(group_cols, dropna=False)["__row_cost__"].sum().reset_index()
        g = g.rename(columns={"__row_cost__":"TotalCost"})
        return g

    if len(hier_cols) == 3:
        roll["H3"] = _roll(hier_cols)
        roll["H2"] = _roll(hier_cols[:2])
        roll["H1"] = _roll(hier_cols[:1])
    elif len(hier_cols) == 2:
        roll["H2"] = _roll(hier_cols)
        roll["H1"] = _roll(hier_cols[:1])
    else:
        roll["H1"] = _roll(hier_cols[:1])

    return {"hier_cols": hier_cols, "cost_cols": cost_candidates, "rollups": roll, "raw": df}, None

def parse_tech(df_raw):
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # Try to detect columns
    # Typical patterns: Criterion, Category, Weight, Score (or raw sub-scores to aggregate)
    # We'll be permissive and look for likely fields:
    crit_col = None
    for c in df.columns:
        if any(k in c.lower() for k in ["kriter", "criterion", "metric"]):
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
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    # prefer a column literally named score-like
    for c in df.columns:
        if any(k in c.lower() for k in ["score","wert","points"]):
            if pd.api.types.is_numeric_dtype(df[c]):
                score_col = c; break
    if score_col is None and numeric_cols:
        # fallback: last numeric col
        score_col = numeric_cols[-1]

    if crit_col is None:
        # If no explicit criteria column, create a row index
        crit_col = "_Criterion"
        df[crit_col] = np.arange(1, len(df)+1)

    # Normalize missing weight -> 1
    if weight_col is None:
        weight_col = "_Weight"
        df[weight_col] = 1.0
    else:
        df[weight_col] = pd.to_numeric(df[weight_col], errors="coerce").fillna(1.0)

    # Ensure score numeric
    df[score_col] = pd.to_numeric(df[score_col], errors="coerce")

    # Weighted scores
    df["_Weighted"] = df[score_col] * df[weight_col]
    total_weight = df[weight_col].sum() if df[weight_col].sum() != 0 else 1.0
    overall = df["_Weighted"].sum() / total_weight

    # Aggregation by category (optional)
    by_cat = None
    if cat_col:
        by_cat = df.groupby(cat_col, dropna=False).apply(
            lambda g: pd.Series({
                "Weight": g[weight_col].sum(),
                "WeightedScore": g["_Weighted"].sum(),
                "ScoreAvg": (g["_Weighted"].sum() / (g[weight_col].sum() if g[weight_col].sum()!=0 else 1.0))
            })
        ).reset_index()

    return {
        "crit_col": crit_col,
        "weight_col": weight_col,
        "score_col": score_col,
        "overall": overall,
        "by_cat": by_cat,
        "table": df[[crit_col, weight_col, score_col, "_Weighted"]]
    }

def read_product(file):
    name = file.name.rsplit(".",1)[0]
    xl = pd.ExcelFile(file)
    # Identify sheets
    func_sheet = _pick_sheet(xl, FUNC_SHEET_HINTS)
    tech_sheet = _pick_sheet(xl, TECH_SHEET_HINTS)
    start_sheet = _pick_sheet(xl, START_SHEET_HINTS)
    func_df = xl.parse(func_sheet)
    tech_df = xl.parse(tech_sheet)
    start_df = xl.parse(start_sheet)

    func_parsed, err = parse_function_costs(func_df)
    tech_parsed = parse_tech(tech_df)

    meta = {}
    # try to pull some metadata from start sheet (first row or key-value)
    try:
        if start_df.shape[1] >= 2:
            # Attempt key-value pairs
            for i in range(min(30, len(start_df))):
                k = str(start_df.iloc[i,0])
                v = start_df.iloc[i,1] if start_df.shape[1] > 1 else None
                if k and k.lower() not in ["nan","none",""]:
                    meta[k] = v
    except Exception:
        pass

    return {
        "name": name,
        "sheets": {"func": func_sheet, "tech": tech_sheet, "start": start_sheet},
        "func": func_parsed,
        "tech": tech_parsed,
        "meta": meta
    }, err

#############################
# UI
#############################
if not files:
    st.info("Upload at least one Excel file to begin.")
    st.stop()

products = []
errors = []
for f in files:
    try:
        p, err = read_product(f)
        products.append(p)
        if err:
            errors.append(f"{f.name}: {err}")
    except Exception as e:
        errors.append(f"{f.name}: {e}")

if errors:
    with st.expander("Parsing notes / errors"):
        for e in errors:
            st.warning(e)

# Summary
st.subheader("Products loaded")
st.write(pd.DataFrame([{"Product": p["name"], "FunctionSheet": p["sheets"]["func"], "TechSheet": p["sheets"]["tech"], "StartSheet": p["sheets"]["start"]} for p in products]))

# Consolidated function roll-ups (H1-level)
h1_frames = []
for p in products:
    func = p["func"]
    if func and "rollups" in func and "H1" in func["rollups"]:
        h1 = func["rollups"]["H1"].copy()
        h1.columns = [*h1.columns[:-1], "TotalCost"]
        # rename hierarchy col to generic "H1"
        h1_col = func["hier_cols"][0]
        h1 = h1.rename(columns={h1_col:"H1"})
        h1["Product"] = p["name"]
        h1_frames.append(h1)
if h1_frames:
    h1_all = pd.concat(h1_frames, ignore_index=True)
    st.subheader("Function Costs — H1 Roll-up (by Product)")
    st.dataframe(h1_all, use_container_width=True, height=360)
    # Pivot for quick comparison
    pivot = h1_all.pivot_table(index="H1", columns="Product", values="TotalCost", aggfunc="sum", fill_value=0).reset_index()
    st.caption("H1 cost comparison")
    st.dataframe(pivot, use_container_width=True, height=280)
    st.bar_chart(pivot.set_index("H1"))
else:
    st.info("No H1 roll-ups detected yet. Check your sheets/columns.")

# Technical scores
tech_rows = []
bycat_frames = []
for p in products:
    t = p["tech"]
    tech_rows.append({"Product": p["name"], "OverallTechScore": t["overall"]})
    if t["by_cat"] is not None:
        dfc = t["by_cat"].copy()
        dfc["Product"] = p["name"]
        bycat_frames.append(dfc)

st.subheader("Technical Scores — Overall")
st.write(pd.DataFrame(tech_rows))

if bycat_frames:
    st.subheader("Technical Scores — by Category (weighted average)")
    bycat = pd.concat(bycat_frames, ignore_index=True)
    st.dataframe(bycat, use_container_width=True, height=360)

# Exports for Think-Cell (CSV)
st.subheader("Exports")
exp = {}
if h1_frames:
    exp["thinkcell_h1_costs.csv"] = h1_all
if bycat_frames:
    exp["thinkcell_tech_by_category.csv"] = bycat
exp["tech_overall.csv"] = pd.DataFrame(tech_rows)

for fname, df in exp.items():
    buff = io.BytesIO()
    df.to_csv(buff, index=False)
    st.download_button(f"Download {fname}", data=buff.getvalue(), file_name=fname, mime="text/csv")

st.caption("Tip: Link these CSVs in think-cell as data sources to keep PowerPoint charts up to date.")
