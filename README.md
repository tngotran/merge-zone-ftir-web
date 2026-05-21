# FTIR Zone Merger

Web app for merging FTIR `.dpt` spectra into a per-zone summary workbook.
Replaces the original local Python + Excel + VBA pipeline.

## What it does

1. Accepts multiple `.dpt` files via browser upload.
2. Parses each spectrum (UTF-8 / UTF-16, comma / tab / whitespace separators).
3. For each file, computes a peak-ratio metric at two wavenumber pairs (1590/2242 and 1595/2243).
4. Groups files by zone (detected from filename: `zone 1`, `zone2`, `z3`, etc.; or treats exactly 4 unzoned files as Zone 1).
5. Builds a multi-sheet `.xlsx` (one sheet per zone) and offers it for download.

## Using the deployed app

Visit the deployment URL. Upload `.dpt` files, click **Process files**, download the result.

## Running locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Opens at http://localhost:8501.

## Running tests

```bash
source .venv/bin/activate
pip install pytest
pytest tests/ -v
```

## Deployment to Streamlit Community Cloud

1. Push this repository to GitHub.
2. Sign in at <https://share.streamlit.io>.
3. **New app** → select the repo → branch `main` → main file `app.py` → **Deploy**.
4. Streamlit installs `requirements.txt` and gives you a public URL.

## Relationship to the original local pipeline

The original `main_convert_dpt_2_excel.py` and `LumosTemplateProtected.xlsm` are kept intact in this repository. They continue to work for anyone wanting to run the pipeline locally with desktop Excel. The web app (`app.py` + `pipeline.py`) is a pure-Python re-implementation that does not require Excel.
