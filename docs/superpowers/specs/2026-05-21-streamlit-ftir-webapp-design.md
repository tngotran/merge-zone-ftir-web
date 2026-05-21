# Streamlit FTIR Zone Merger — Web App Design

**Date:** 2026-05-21
**Status:** Approved (pending user review of this document)

## Goal

Replace the local Python + Excel/xlwings + VBA macro pipeline with a deployable
web application that lets one user upload FTIR `.dpt` files in a browser and
download a final merged `.xlsx`. Hosted on Streamlit Community Cloud (free).

## Why this is feasible

The original pipeline opens `LumosTemplateProtected.xlsm` via xlwings and runs
a VBA macro that computes a single number from the spectrum. xlwings requires
a desktop Excel installation, which is the blocker for web deployment.

VBA was extracted with `olevba`. The macro is trivial:

```vba
Z1 = INDEX(B2:B2000, MATCH(1590, A2:A2000, -1))   ' Sheet2 variant; Sheet1 uses 1580
Z2 = INDEX(B2:B2000, MATCH(2242, A2:A2000, -1))
Z3 = ((0.29 * Z1) / ((0.29 * Z1) + Z2)) * 100
I5 = Z3
```

This duplicates a calculation already in `main_convert_dpt_2_excel.py:170`
(which uses 1595 / 2243 instead of 1590 / 2242). Porting the macro to Python
removes the entire desktop-Excel dependency.

## Numerical fidelity

The current pipeline appends BOTH the macro's result (1590 / 2242) AND
Python's own result (1595 / 2243) into `result_l`, then averages all entries
for the trailing column. The web port preserves this dual-calculation
behavior exactly. Output `.xlsx` files from the web app should be numerically
identical to the current pipeline's output.

## Architecture

Single-file Streamlit app, all processing in-process. No database, no
external storage, no background workers. Files held in memory only; never
written to the server's filesystem.

```
Browser (single user)
   |
   v   HTTPS
Streamlit Community Cloud (free tier)
   |
   v
app.py
 - file uploader (multi .dpt)
 - calls pipeline.process_dpt_files()
 - shows progress log
 - serves final .xlsx via st.download_button
```

### Repo layout

```
merge_zone_FTIR/
├── app.py                    # Streamlit UI + entry point
├── pipeline.py               # Pure-Python processing (no Streamlit imports)
├── requirements.txt
├── tests/
│   ├── test_pipeline.py
│   └── fixtures/             # sample .dpt files + golden .xlsx
├── README.md
└── .streamlit/config.toml    # optional UI tweaks
```

`pipeline.py` is deliberately separated from `app.py` so it's testable
without Streamlit and reusable as a CLI later if needed.

## Pipeline (`pipeline.py`)

Public entry point:

```python
def process_dpt_files(
    dpt_files: list[tuple[str, bytes]],   # (filename, content)
    progress_callback=None,                # optional, called with str messages
) -> tuple[str, bytes]:                    # (output_filename, xlsx_bytes)
```

### Step 1 — Parse `.dpt` into DataFrames

Port of `main_convert_dpt_2_excel.py:23-51`. Preserve:

- Encoding detection (UTF-16 BOM → `utf-16`, else `utf-8`)
- Separator detection (`,` / `\t` / whitespace)
- Numeric-vs-text header heuristic
- Skip files under 10 bytes

Output: `{filename → DataFrame}`. No intermediate `.xlsx` files are written
(they only existed in the original because xlwings needed Excel files).

### Step 2 — Macro replacement (pure Python)

```python
def compute_macro_result(df: pd.DataFrame,
                        wn_a: float = 1590,
                        wn_b: float = 2242) -> float:
    col0 = pd.to_numeric(df.iloc[:, 0], errors='coerce')
    z1 = float(df.iloc[(col0 - wn_a).abs().idxmin(), 1])
    z2 = float(df.iloc[(col0 - wn_b).abs().idxmin(), 1])
    return ((0.29 * z1) / ((0.29 * z1) + z2)) * 100
```

Mirrors the extracted VBA exactly.

### Step 3 — Per-zone merge

Port of `main_convert_dpt_2_excel.py:78-226`. Preserve:

- Case-insensitive zone-name matching: `zone N`, `zoneN`, `zN`, `z N`,
  `zone N.`, etc.
- "Exactly 4 non-merged files → treat all as Zone 1" fallback
- Per-file column layout in the merged output:
  `[col_A, col_B, macro_result, blank, value_1595, max_2243, python_result, blank]`
- Append BOTH macro result (1590/2242) and Python result (1595/2243) into
  `result_l`
- Trailing column = `mean(result_l)`

Output: `{zone_num → merged DataFrame}`.

### Step 4 — Combine into final multi-sheet workbook

Port of `main_convert_dpt_2_excel.py:233-250`. Write each zone's merged
DataFrame to its own sheet (`ZONE 1`, `ZONE 2`, ...) in a single `.xlsx`,
named `<first_dpt_stem>_FINAL_OUTPUT.xlsx`. Return as bytes.

### Coexistence with original code

The original files MUST remain untouched:

- `main_convert_dpt_2_excel.py` — keep as-is
- `LumosTemplateProtected.xlsm` — keep as-is
- `readme.txt` — keep as-is
- `old/`, `FTIR_19May2026.zip` — keep as-is

The new web app is **additive**. `app.py` and `pipeline.py` are new files
that live alongside the existing pipeline. The existing local pipeline
continues to work for anyone who wants to run it that way. Nothing in the
original pipeline is modified or deleted.

### What the new pipeline does NOT use (vs. the original)

- No `xlwings`, no `wb.macro(...)`, no `wb.sheets[1].range(...)`,
  no `wb.save()`, no `wb.close()`
- No reads/writes of `LumosTemplateProtected.xlsm`
- No intermediate per-file `.xlsx` writes to disk

## UI (`app.py`)

Single page, top-to-bottom:

1. Title: "FTIR Zone Merger"
2. `st.file_uploader(accept_multiple_files=True, type=["dpt"])`
3. List of queued filenames
4. `[ Process files ]` button — disabled until at least one file uploaded
5. Live log area populated by `progress_callback` passed into pipeline
6. `[ Download FINAL_OUTPUT.xlsx ]` (`st.download_button` with returned bytes)
7. Small `[ Clear and start over ]` button to reset session state

### Behavior

- Uploaded files stay in memory as Streamlit `UploadedFile` objects;
  pipeline receives `(filename, bytes)` tuples. Nothing touches disk.
- Pipeline is pure — does not import Streamlit. The `progress_callback`
  decouples them.
- Malformed/empty `.dpt` files: skip and log (same as today).
- Zero valid files: show error banner, no download button.

### Deliberately omitted (YAGNI)

- No login, no history, no settings panel, no zone-rule configuration.
- No password gate (can be added in ~2 lines if requested later).

## Testing

Three tests, written first per TDD:

1. **`test_compute_macro_result`** — hand-built 2-column DataFrame with known
   values around 1590 and 2242; assert output matches formula by hand.

2. **`test_pipeline_snapshot`** — feed the new pipeline a fixture set of
   `.dpt` files; compare resulting `.xlsx` numerically against a "golden"
   output generated **once** from the current xlwings pipeline before its
   removal. This is the safety net that proves the port preserves behavior.

3. **`test_zone_detection`** — list of fake filenames (`Zone 1.dpt`,
   `zone2.dpt`, `Z 3 sample.dpt`, `merged_zone1.dpt`, etc.); assert each
   lands in the right zone bucket or is correctly excluded.

## Deployment

- New public GitHub repo (e.g., `merge-zone-ftir-web`)
- Push `app.py`, `pipeline.py`, `requirements.txt`, `tests/`, `README.md`
- Connect repo to https://share.streamlit.io, select `app.py` as entry point
- Result: public URL `https://<app-name>.streamlit.app`

### `requirements.txt`

```
streamlit==1.39.0
pandas==2.2.3
openpyxl==3.1.5
```

### `README.md` sections

1. What it does
2. How to use the web app
3. How to run locally (`pip install -r requirements.txt && streamlit run app.py`)

## Risks and limits

- **Idle sleep**: Streamlit Community Cloud apps sleep after ~7 days idle.
  First visit after sleep takes ~30s to wake. Acceptable for one user.
- **Free tier RAM**: 1 GB. FTIR `.dpt` files are KB-sized; not a concern.
- **Public URL**: anyone with the link can use the app. The user confirmed
  this is OK. If the URL becomes a concern later, Streamlit supports
  password-gating in a few lines.

## Out of scope

- Authentication / multi-user support
- Persistent history of runs
- Configurable wavenumber inputs from the UI
- Any change to the numerical algorithm (we are preserving current behavior
  exactly, not improving it)
