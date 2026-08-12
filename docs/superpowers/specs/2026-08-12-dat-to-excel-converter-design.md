# DAT to Excel Converter — Design

Date: 2026-08-12
Branch: `feat/dat-to-excel-converter`

## Goal

Add a second tool to the Streamlit app: convert SAXS/EDF-style `.dat` files to
Excel workbooks. Ports `/Users/t.ngo/Desktop/dat2ex/dat_to_excel_converter.py`
(a local folder-walking CLI script) to a browser upload workflow.

## Source material

The CLI script reads `.dat` files with an EDF-style header — a block of
`# key value` comment lines — followed by whitespace-separated numeric columns.
It auto-detects two known column layouts:

- `psi(°)` / `Intensity(a.u.)` / `Sigma_I(a.u.)`
- `q(A-1)` / `I(q)` / `Sig(q)`

Falling back, when no named header line is present, to column count: 3 columns →
the `q` layout, 6 columns → `psi` + `q` concatenated, anything else → generic
`Column_N` names.

Each `.dat` becomes one `.xlsx` with a **Metadata** sheet (Parameter/Value pairs
from the header) and a **Data** sheet (the numeric block with proper column
names).

A survey of the source data found 294 files in the `psi` format and 135 in the
`q` format across 744 total `.dat` files (~36 MB, largest single file 74 KB).

## Architecture

```
app.py                  router (st.navigation), Streamlit Cloud entry point
zone_merger_page.py     existing app.py UI, relocated unmodified
dat_converter_page.py   new UI
pipeline.py             untouched
dat_pipeline.py         new: pure conversion functions, no Streamlit, no disk I/O
dat_to_excel_converter.py  original CLI script, vendored verbatim for reference
tests/test_dat_pipeline.py
tests/fixtures/sample_dat/
```

`app.py` remains the deploy entry point so no Streamlit Cloud config changes.
It becomes a thin `st.navigation` router over two `st.Page`s. This requires
streamlit >= 1.36; the project pins >= 1.39.

`dat_pipeline.py` mirrors the shape of the existing `pipeline.py`: pure
functions, bytes in and bytes out, an optional `progress_callback` for status
messages.

```python
detect_column_headers(lines) -> (headers, data_start_line)
parse_dat(content: bytes) -> dict | None
dat_to_xlsx_bytes(parsed) -> bytes
convert_dat_files(files, progress_callback) -> (zip_name, zip_bytes)
```

## What carries over unchanged

The conversion logic is preserved exactly: header detection (both named formats
plus the 3-column and 6-column numeric fallbacks), metadata extraction
(`# key value`, skipping lines that are really column headers), padding ragged
rows out to `max_cols`, and the two-sheet Metadata + Data output layout.

## What is dropped, and why

- `install_openpyxl()` — runs `pip install` in a subprocess. openpyxl is already
  declared in `requirements.txt`; a subprocess installer must never ship to
  Streamlit Cloud.
- `process_folder` / `process_all_folders` / `main` — filesystem globbing,
  replaced by the list of uploaded files.
- `print()` — replaced by the `progress_callback` pattern `pipeline.py` already
  uses.

## Output packaging

A browser cannot receive a directory, so each `.dat` becomes its own `.xlsx` and
all of them are bundled into a single `.zip` download named
`{first_file_stem}_CONVERTED.zip`. This preserves the CLI's one-workbook-per-file
output shape and its filenames.

## Edge cases the CLI never hits

1. **Empty workbook.** If a `.dat` yields neither metadata nor data rows,
   `pd.ExcelWriter` closes with zero sheets and openpyxl raises
   `At least one sheet must be visible`. Such files are skipped with a logged
   reason instead of failing the batch.
2. **Duplicate filenames.** The `deliver/`, `archived/`, and `27072026/` folders
   can hold identically-named `.dat` files. On disk they land in separate output
   directories; inside one zip they would collide. Duplicates get a `_2`, `_3`
   suffix.
3. **Per-file error isolation.** A malformed file logs and skips; the rest of the
   batch still converts. The run ends with an `N converted, M skipped` summary.

## UI

`dat_converter_page.py` follows the Zone Merger's interaction model exactly: a
keyed multi-file uploader (the key increments to reset), a list of queued files,
`Convert` and `Reset` buttons in columns, a live progress log, a download button,
and a "Convert other files" reset. Target batch size is a few dozen files, so
`maxUploadSize` stays at 50 MB.

## Testing

`tests/test_dat_pipeline.py` covers both named header formats, the 3-column and
6-column numeric fallbacks, metadata parsing, ragged-row padding, the empty-file
skip, duplicate-name suffixing, zip structure, and the progress callback. Real
`.dat` files from `dat2ex`, trimmed to a header plus ~20 data rows, serve as
fixtures in `tests/fixtures/sample_dat/`.
