# FKA Prototype — Function Cost & Tech Scoring

This Streamlit app ingests one or more Excel files that follow the idea of:
- `SLAVE_Funktions-Kostenstruktur` (function hierarchy + cost allocation)
- `SLAVE_Techn.Bewertung` (technical scoring)
- `SLAVE_START` (metadata)

It auto-detects sheet names (fuzzy), rolls up costs at H1/H2/H3, computes weighted technical scores, consolidates across products, and provides CSV exports suitable for think-cell linking.

## Run locally
```
pip install -r requirements.txt
streamlit run app.py
```
Upload your `.xlsm`/`.xlsx` files and explore the consolidated outputs.
