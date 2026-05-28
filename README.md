# Streamlit project

Minimal Python `streamlit` starter for 1080 reporting.

## Run

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

The app signs users in with their own `1080 API key`, validates it with a lightweight 1080 API request, and caches loaded clients in browser local storage.

The current MVP also supports:
- session lookup for a selected client
- session detail loading
- automatic `Running (LR)` FV profile fetch
- automatic `15-0-5` split profile fetch
- non-normative FV PDF export for a selected run with optional uploaded logo
- quadrant-based FV PDF export based on cohort medians in [`data/fv_norms.xlsx`](/Users/smola/Projects/streamlit-romansvantner/app/data/fv_norms.xlsx)

The quadrant reference file is expected to contain at least `category`, `f0_median`, and `v0_median`. The bundled file can be regenerated from source data with `../norm.py`.

## Optional environment variables

```bash
export API1080_BASE_URL=https://publicapi.1080motion.com
```

## Files

- `app.py` - main Streamlit app
- `requirements.txt` - Python dependencies
- `.gitignore` - common ignored files
- `v1.json` - local OpenAPI documentation for the 1080 API
- `data/fv_norms.xlsx` - editable FV quadrant reference medians for categories such as `U14`, `U15`, `U16`, `U17`, `U19`
