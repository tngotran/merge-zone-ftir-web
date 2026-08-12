# FTIR Tools

Web app with two tools, selectable from the sidebar:

1. **FTIR Zone Merger** — merges FTIR `.dpt` spectra into a per-zone summary workbook.
2. **DAT to Excel** — converts SAXS/EDF-style `.dat` files into Excel workbooks.

## FTIR Zone Merger

Replaces the original local Python + Excel + VBA pipeline.

### What it does

1. Accepts multiple `.dpt` files via browser upload.
2. Parses each spectrum (UTF-8 / UTF-16, comma / tab / whitespace separators).
3. For each file, computes a peak-ratio metric at two wavenumber pairs (1590/2242 and 1595/2243).
4. Groups files by zone (detected from filename: `zone 1`, `zone2`, `z3`, etc.; or treats exactly 4 unzoned files as Zone 1).
5. Builds a multi-sheet `.xlsx` (one sheet per zone) and offers it for download.

Upload `.dpt` files, click **Process files**, download the result.

## DAT to Excel

Ports the local `dat_to_excel_converter.py` folder-walking script to browser upload.

### What it does

1. Accepts multiple `.dat` files via browser upload.
2. Reads each file's `#` comment header into a **Metadata** sheet (Parameter / Value pairs).
3. Detects the column layout — `psi(°)` / `Intensity(a.u.)` / `Sigma_I(a.u.)`, or
   `q(A-1)` / `I(q)` / `Sig(q)` — falling back to column count when the file has no
   named header line.
4. Writes the numeric block to a **Data** sheet with those column names.
5. Bundles one `.xlsx` per input file into a single `.zip` for download.

Files that yield no readable data are skipped with a logged reason rather than failing
the batch, and same-named files from different folders get a `_2`, `_3` suffix so they
don't collide inside the zip.

Upload `.dat` files, click **Convert files**, download the `.zip`.

## Using the deployed app

Visit the deployment URL and pick a tool from the sidebar.

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

## Layout

| File | Role |
| --- | --- |
| `app.py` | Entry point / router (`st.navigation`). Deployment main file. |
| `zone_merger_page.py` | Zone Merger UI |
| `dat_converter_page.py` | DAT to Excel UI |
| `pipeline.py` | Zone-merge logic (pure functions) |
| `dat_pipeline.py` | DAT conversion logic (pure functions) |

## Relationship to the original local pipelines

The original `main_convert_dpt_2_excel.py`, `LumosTemplateProtected.xlsm`, and
`dat_to_excel_converter.py` are kept intact in this repository. They continue to work
for anyone wanting to run the pipelines locally. The web app is a pure-Python
re-implementation: the zone merger does not require Excel, and the DAT converter works
on uploaded bytes instead of walking folders on disk.

`dat_pipeline.py` is verified byte-for-byte against `dat_to_excel_converter.py` across
all 744 `.dat` files in the source dataset — identical Metadata and Data sheet contents.
