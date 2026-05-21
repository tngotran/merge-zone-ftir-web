# Streamlit FTIR Zone Merger — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Streamlit web app that accepts `.dpt` file uploads, runs the FTIR zone-merge pipeline in pure Python (no Excel/xlwings), and serves a downloadable `_FINAL_OUTPUT.xlsx`. The original local pipeline files (`main_convert_dpt_2_excel.py`, `LumosTemplateProtected.xlsm`, etc.) remain untouched — the new code is purely additive.

**Architecture:** Single Streamlit app (`app.py`) calling a pure-Python pipeline module (`pipeline.py`) that replaces the xlwings/VBA-macro step with an equivalent Python function. Everything runs in-process and in-memory. Deployable free to Streamlit Community Cloud.

**Tech Stack:** Python 3.11+, Streamlit, pandas, openpyxl, pytest.

---

## Files this plan creates

All NEW files. None of the existing files are modified.

| Path | Purpose |
|------|---------|
| `pipeline.py` | Pure-Python pipeline: parse `.dpt`, compute macro result, zone-merge, build workbook |
| `app.py` | Streamlit UI: file uploader, progress log, download button |
| `requirements.txt` | Pinned deps for Streamlit Community Cloud |
| `.gitignore` | Standard Python ignores |
| `.streamlit/config.toml` | Minimal UI theme config |
| `README.md` | Usage + deployment instructions |
| `tests/__init__.py` | Empty, marks tests as a package |
| `tests/test_pipeline.py` | Unit + snapshot tests |
| `tests/fixtures/sample_dpt/*.dpt` | Sample `.dpt` files extracted from `old/1.zip` |
| `tests/fixtures/golden_output.xlsx` | Reference output from original pipeline (user generates once) |

## Key design constraint: numerical fidelity vs. Excel MATCH semantics

The original VBA macro used `MATCH(value, range, -1)` which has descending-sort semantics. Our pure-Python replacement uses `(col - target).abs().idxmin()` — the "closest by absolute difference" approach already used by the original Python code at `main_convert_dpt_2_excel.py:145`. On well-formed ascending-sorted spectra with dense sample points, both approaches return the same row. The snapshot test (Task 12) verifies fidelity against the user's actual data.

---

## Task 1: Initialize git repo and scaffolding

**Files:**
- Create: `.gitignore`
- Create: `requirements.txt`
- Create: `tests/__init__.py`
- Create: `tests/fixtures/.gitkeep`
- Create: `.streamlit/config.toml`

- [ ] **Step 1: Initialize git repo**

Run from project root:
```bash
git init
git add main_convert_dpt_2_excel.py LumosTemplateProtected.xlsm readme.txt FTIR_19May2026.zip old docs
git commit -m "chore: snapshot existing project before web port"
```

This captures all existing files so the original local pipeline is preserved in history before any additive work.

- [ ] **Step 2: Create `.gitignore`**

Write to `.gitignore`:
```
# Python
__pycache__/
*.py[cod]
*$py.class
.pytest_cache/
.venv/
venv/
*.egg-info/

# OS
.DS_Store
Thumbs.db

# Editor
.vscode/
.idea/

# Streamlit
.streamlit/secrets.toml

# Temp test outputs
tests/output/
```

- [ ] **Step 3: Create `requirements.txt`**

Write to `requirements.txt`:
```
streamlit==1.39.0
pandas==2.2.3
openpyxl==3.1.5
```

(pytest is a dev-only dep, installed separately during development.)

- [ ] **Step 4: Create `tests/__init__.py`**

Write empty file (zero bytes) to `tests/__init__.py`.

- [ ] **Step 5: Create `tests/fixtures/.gitkeep`**

Write empty file (zero bytes) to `tests/fixtures/.gitkeep`. This keeps the fixtures directory in git before fixtures are added.

- [ ] **Step 6: Create `.streamlit/config.toml`**

Write to `.streamlit/config.toml`:
```toml
[theme]
base = "light"

[server]
maxUploadSize = 50

[browser]
gatherUsageStats = false
```

- [ ] **Step 7: Install dev dependencies locally**

Run:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt pytest
```
Expected: all packages install without errors.

- [ ] **Step 8: Commit**

```bash
git add .gitignore requirements.txt tests/__init__.py tests/fixtures/.gitkeep .streamlit/config.toml
git commit -m "feat: scaffold streamlit project structure"
```

---

## Task 2: Extract sample `.dpt` fixtures

**Files:**
- Create: `tests/fixtures/sample_dpt/<two .dpt files>`

- [ ] **Step 1: Extract fixture files from `old/1.zip`**

Run from project root:
```bash
mkdir -p tests/fixtures/sample_dpt
unzip -j old/1.zip "1/*.dpt" -d tests/fixtures/sample_dpt -x "__MACOSX/*"
ls tests/fixtures/sample_dpt/
```
Expected: 2 `.dpt` files listed.

- [ ] **Step 2: Rename fixtures for zone-detection testing**

The fixtures need to exercise the zone-naming logic. Rename whatever was extracted to Zone 1 form:
```bash
cd tests/fixtures/sample_dpt
files=( *.dpt )
echo "Found ${#files[@]} .dpt files"
mv -- "${files[0]}" "sample_zone1_a.dpt"
mv -- "${files[1]}" "sample_zone1_b.dpt"
ls
cd -
```
Expected: two files named `sample_zone1_a.dpt` and `sample_zone1_b.dpt`. If fewer than 2 files were extracted, stop and investigate `old/1.zip` contents.

- [ ] **Step 3: Commit fixtures**

```bash
git add tests/fixtures/sample_dpt/
git commit -m "test: add sample .dpt fixtures from old/1.zip"
```

---

## Task 3: TDD — `compute_macro_result()` (macro replacement)

**Files:**
- Create: `pipeline.py`
- Create: `tests/test_pipeline.py`

- [ ] **Step 1: Write failing test**

Write to `tests/test_pipeline.py`:
```python
import pandas as pd
import pytest
from pipeline import compute_macro_result


def test_compute_macro_result_exact_match_points():
    """When the spectrum contains exact wavenumber matches, returns the formula's value."""
    df = pd.DataFrame({
        0: [1000.0, 1500.0, 1590.0, 1800.0, 2000.0, 2242.0, 2500.0],
        1: [0.10,   0.20,   0.50,   0.40,   0.30,   0.80,   0.05],
    })
    result = compute_macro_result(df, wn_a=1590, wn_b=2242)
    # z1 = 0.5 (row at 1590), z2 = 0.8 (row at 2242)
    expected = (0.29 * 0.5) / ((0.29 * 0.5) + 0.8) * 100
    assert result == pytest.approx(expected, abs=1e-9)


def test_compute_macro_result_picks_closest_when_no_exact_match():
    """When wavenumbers don't appear exactly, picks the row with smallest abs difference."""
    df = pd.DataFrame({
        0: [1585.0, 1592.0, 2240.0, 2245.0],  # no exact 1590 or 2242
        1: [0.10,   0.50,   0.80,   0.20],
    })
    # Closest to 1590 is 1592 (|diff|=2), closest to 2242 is 2240 (|diff|=2)
    result = compute_macro_result(df, wn_a=1590, wn_b=2242)
    expected = (0.29 * 0.50) / ((0.29 * 0.50) + 0.80) * 100
    assert result == pytest.approx(expected, abs=1e-9)


def test_compute_macro_result_default_wavenumbers():
    """Defaults are wn_a=1590, wn_b=2242."""
    df = pd.DataFrame({
        0: [1590.0, 2242.0],
        1: [1.0,    1.0],
    })
    result = compute_macro_result(df)
    expected = (0.29 * 1.0) / ((0.29 * 1.0) + 1.0) * 100
    assert result == pytest.approx(expected, abs=1e-9)
```

- [ ] **Step 2: Run test, verify it fails**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: ImportError / ModuleNotFoundError: No module named 'pipeline'.

- [ ] **Step 3: Implement `compute_macro_result()`**

Write to `pipeline.py`:
```python
"""Pure-Python FTIR zone-merge pipeline. No Excel, no xlwings."""

from __future__ import annotations

import pandas as pd


def compute_macro_result(
    df: pd.DataFrame,
    wn_a: float = 1590,
    wn_b: float = 2242,
) -> float:
    """Replicates the VBA macro: ((0.29*z1)/((0.29*z1)+z2))*100

    z1 = value in column 1 of the row whose column 0 is closest to wn_a.
    z2 = value in column 1 of the row whose column 0 is closest to wn_b.
    Defaults (1590, 2242) match the macro's Sheet2 variant, which is the
    sheet the original Python pipeline reads from.
    """
    col0 = pd.to_numeric(df.iloc[:, 0], errors='coerce')
    z1 = float(df.iloc[(col0 - wn_a).abs().idxmin(), 1])
    z2 = float(df.iloc[(col0 - wn_b).abs().idxmin(), 1])
    return ((0.29 * z1) / ((0.29 * z1) + z2)) * 100
```

- [ ] **Step 4: Run test, verify it passes**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add pipeline.py tests/test_pipeline.py
git commit -m "feat: add compute_macro_result (pure-python replacement for VBA macro)"
```

---

## Task 4: TDD — `parse_dpt()` (.dpt parsing)

**Files:**
- Modify: `pipeline.py`
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: Add failing test**

Append to `tests/test_pipeline.py`:
```python
import io
from pipeline import parse_dpt


def test_parse_dpt_comma_separated_no_header():
    """Comma-separated numeric data with no header row."""
    content = b"1000.5,0.123\n1001.0,0.124\n1002.0,0.125\n"
    df = parse_dpt(content)
    assert df.shape == (3, 2)
    assert df.iloc[0, 0] == pytest.approx(1000.5)
    assert df.iloc[0, 1] == pytest.approx(0.123)


def test_parse_dpt_tab_separated_no_header():
    """Tab-separated numeric data with no header row."""
    content = b"1000.5\t0.123\n1001.0\t0.124\n"
    df = parse_dpt(content)
    assert df.shape == (2, 2)
    assert df.iloc[1, 0] == pytest.approx(1001.0)


def test_parse_dpt_whitespace_separated():
    """Whitespace-separated numeric data."""
    content = b"1000.5 0.123\n1001.0 0.124\n"
    df = parse_dpt(content)
    assert df.shape == (2, 2)


def test_parse_dpt_with_text_header():
    """First row is text → treat as header and skip."""
    content = b"wavenumber,intensity\n1000.5,0.123\n1001.0,0.124\n"
    df = parse_dpt(content)
    assert df.shape == (2, 2)
    assert df.iloc[0, 0] == pytest.approx(1000.5)


def test_parse_dpt_utf16_bom():
    """UTF-16 with BOM is decoded correctly."""
    text = "1000.5,0.123\n1001.0,0.124\n"
    content = text.encode("utf-16")  # adds BOM
    df = parse_dpt(content)
    assert df.shape == (2, 2)
    assert df.iloc[0, 0] == pytest.approx(1000.5)


def test_parse_dpt_too_small_returns_none():
    """Files under 10 bytes return None (skip signal)."""
    assert parse_dpt(b"1,2") is None


def test_parse_dpt_real_fixture():
    """Parses an actual .dpt fixture file end-to-end."""
    from pathlib import Path
    fixture_dir = Path(__file__).parent / "fixtures" / "sample_dpt"
    dpt_files = sorted(fixture_dir.glob("*.dpt"))
    assert len(dpt_files) >= 1, "Fixture .dpt files missing — re-run Task 2"
    content = dpt_files[0].read_bytes()
    df = parse_dpt(content)
    assert df is not None
    assert df.shape[1] == 2
    assert df.shape[0] > 100  # FTIR spectra are typically thousands of points
```

- [ ] **Step 2: Run tests, verify they fail**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: `compute_macro_result` tests pass (3); `parse_dpt` tests fail with ImportError.

- [ ] **Step 3: Implement `parse_dpt()`**

Append to `pipeline.py`:
```python
import io
from typing import Optional


def parse_dpt(content: bytes) -> Optional[pd.DataFrame]:
    """Parse a .dpt file's raw bytes into a 2-column DataFrame.

    Returns None for files under 10 bytes (treated as empty/corrupt — same
    behavior as the original pipeline).

    Handles:
    - UTF-16 (BOM-detected) and UTF-8 encodings
    - Comma, tab, or whitespace separators
    - Optional first-row text header
    """
    if len(content) < 10:
        return None

    # Encoding detection
    if content[:2] in (b'\xff\xfe', b'\xfe\xff'):
        encoding = 'utf-16'
    else:
        encoding = 'utf-8'

    text = content.decode(encoding)
    buf = io.StringIO(text)
    first_line = buf.readline()
    buf.seek(0)

    # Separator detection
    if ',' in first_line:
        sep = ','
    elif '\t' in first_line:
        sep = '\t'
    else:
        sep = r'\s+'

    # Numeric-vs-text header heuristic: try to parse the first column as float.
    # If it works, the file has no header. If not, treat row 0 as a header.
    try:
        first_field = first_line.strip().split(',')[0] if ',' in first_line else first_line.strip().split()[0]
        float(first_field)
        header = None
    except ValueError:
        header = 0

    df = pd.read_csv(
        buf,
        sep=sep,
        engine='python',
        header=header,
        names=['Column1', 'Column2'],
        encoding=None,  # already decoded
    )
    return df
```

- [ ] **Step 4: Run tests, verify they pass**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: all `parse_dpt` tests pass (7 new ones).

- [ ] **Step 5: Commit**

```bash
git add pipeline.py tests/test_pipeline.py
git commit -m "feat: add parse_dpt for parsing .dpt files from bytes"
```

---

## Task 5: TDD — `zone_for_filename()` (zone detection)

**Files:**
- Modify: `pipeline.py`
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: Add failing tests**

Append to `tests/test_pipeline.py`:
```python
from pipeline import zone_for_filename


def test_zone_for_filename_standard_forms():
    """Recognizes 'zone N', 'zoneN', 'zN', etc. case-insensitively."""
    assert zone_for_filename("Sample Zone 1.dpt") == 1
    assert zone_for_filename("sample_zone2_a.dpt") == 2
    assert zone_for_filename("data Zone3.dpt") == 3
    assert zone_for_filename("DATA ZONE 4.dpt") == 4
    assert zone_for_filename("file z5 thing.dpt") == 5
    assert zone_for_filename("Z6_sample.dpt") == 6


def test_zone_for_filename_decimal_variants():
    """Recognizes 'zone N.' forms (e.g., 'Zone 4.0')."""
    assert zone_for_filename("Sample Zone 4.0.dpt") == 4


def test_zone_for_filename_no_zone_returns_none():
    """Returns None when no zone pattern matches."""
    assert zone_for_filename("random_file.dpt") is None
    assert zone_for_filename("EXTRACT_sample.dpt") is None


def test_zone_for_filename_skips_merged_files():
    """Files with 'merged' in the name are skipped (return None)."""
    assert zone_for_filename("ZONE_1_merged.xlsx") is None
    assert zone_for_filename("zone_2_MERGED.dpt") is None
```

- [ ] **Step 2: Run tests, verify they fail**

Run:
```bash
pytest tests/test_pipeline.py::test_zone_for_filename_standard_forms -v
```
Expected: ImportError on `zone_for_filename`.

- [ ] **Step 3: Implement `zone_for_filename()`**

Append to `pipeline.py`:
```python
def zone_for_filename(filename: str) -> Optional[int]:
    """Return the zone number (1-6) for a filename, or None if no zone matches.

    Recognizes (case-insensitive): 'zone N', 'zoneN', 'zN', 'z N',
    'zone N.', 'zoneN.', 'zN.', 'z N.'. Files containing 'merged' are
    skipped (return None) so already-merged outputs aren't reprocessed.
    """
    fname_lower = filename.lower()
    if 'merged' in fname_lower:
        return None
    for zone_num in range(1, 7):
        patterns = [
            f'zone {zone_num}',
            f'zone{zone_num}',
            f'z{zone_num}',
            f'z {zone_num}',
            f'z{zone_num}.',
            f'z {zone_num}.',
            f'zone {zone_num}.',
            f'zone{zone_num}.',
        ]
        if any(p in fname_lower for p in patterns):
            return zone_num
    return None
```

- [ ] **Step 4: Run tests, verify they pass**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add pipeline.py tests/test_pipeline.py
git commit -m "feat: add zone_for_filename for zone-name detection"
```

---

## Task 6: TDD — `merge_zone()` (per-zone column assembly)

**Files:**
- Modify: `pipeline.py`
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: Add failing test**

Append to `tests/test_pipeline.py`:
```python
from pipeline import merge_zone


def test_merge_zone_two_files_column_layout():
    """Per-file: 8 columns appended. Plus trailing mean column at the end."""
    # Each "file" is (name, df). Build two simple spectra.
    df_a = pd.DataFrame({
        0: [1590.0, 2242.0, 1595.0, 2243.0],
        1: [0.5,    0.8,    0.5,    0.8],
    })
    df_b = pd.DataFrame({
        0: [1590.0, 2242.0, 1595.0, 2243.0],
        1: [1.0,    2.0,    1.0,    2.0],
    })
    files = [("file_a.dpt", df_a), ("file_b.dpt", df_b)]
    merged = merge_zone(files)
    # 2 files * 8 columns each + 1 trailing mean = 17 columns
    assert merged.shape[1] == 17

    # Trailing column is the mean of result_l.
    # For each file, result_l gets 2 values appended (macro + python).
    # Both calculations use the same formula here because all wavenumbers
    # match exactly to the same row, so macro_result == python_result.
    # For df_a: (0.29*0.5)/((0.29*0.5)+0.8)*100
    # For df_b: (0.29*1.0)/((0.29*1.0)+2.0)*100
    r_a = (0.29 * 0.5) / ((0.29 * 0.5) + 0.8) * 100
    r_b = (0.29 * 1.0) / ((0.29 * 1.0) + 2.0) * 100
    expected_mean = (r_a + r_a + r_b + r_b) / 4
    assert float(merged.iloc[0, -1]) == pytest.approx(expected_mean, abs=1e-9)


def test_merge_zone_empty_returns_empty():
    """Zero files → empty DataFrame."""
    merged = merge_zone([])
    assert merged.empty
```

- [ ] **Step 2: Run tests, verify they fail**

Run:
```bash
pytest tests/test_pipeline.py::test_merge_zone_two_files_column_layout -v
```
Expected: ImportError on `merge_zone`.

- [ ] **Step 3: Implement `merge_zone()`**

Append to `pipeline.py`:
```python
def merge_zone(files: list[tuple[str, pd.DataFrame]]) -> pd.DataFrame:
    """Build the per-zone merged DataFrame.

    For each file, appends columns: [col_A, col_B, macro_result, blank,
    value@1595, max@2243, python_result, blank]. After all files, appends a
    final column whose first cell is the mean of all per-file results
    (both macro and python calculations contribute to the mean).
    """
    if not files:
        return pd.DataFrame()

    dfs = [df for _, df in files]
    max_rows = max(df.shape[0] for df in dfs)

    merged_cols: list = []
    result_l: list[float] = []

    for df in dfs:
        if df.shape[0] < max_rows:
            df = df.reindex(range(max_rows), fill_value='')

        merged_cols.append(df.iloc[:, 0])
        merged_cols.append(df.iloc[:, 1])

        # Macro replacement: wavenumbers 1590 / 2242
        macro_result = compute_macro_result(df, wn_a=1590, wn_b=2242)
        result_l.append(macro_result)
        merged_cols.append([macro_result] + [''] * (max_rows - 1))
        merged_cols.append([''] * max_rows)  # blank column

        # Python's own calculation: wavenumbers 1595 / 2243
        col0_values = pd.to_numeric(df.iloc[:, 0], errors='coerce')
        idx_1595 = (col0_values - 1595).abs().idxmin()
        value_1595 = float(df.iloc[idx_1595, 1])
        idx_2243 = (col0_values - 2243).abs().idxmin()
        max_local = float(df.iloc[idx_2243, 1])
        python_result = ((0.29 * value_1595) / ((0.29 * value_1595) + max_local)) * 100
        result_l.append(python_result)

        for val in [value_1595, max_local, python_result]:
            merged_cols.append([val] + [''] * (max_rows - 1))
        merged_cols.append([''] * max_rows)

    # Trailing column: mean of all result_l values
    mean_val = sum(result_l) / len(result_l) if result_l else ''
    merged_cols.append([mean_val] + [''] * (max_rows - 1))

    return pd.DataFrame({i: col for i, col in enumerate(merged_cols)})
```

- [ ] **Step 4: Run tests, verify they pass**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add pipeline.py tests/test_pipeline.py
git commit -m "feat: add merge_zone for per-zone column assembly"
```

---

## Task 7: TDD — `process_dpt_files()` (full pipeline)

**Files:**
- Modify: `pipeline.py`
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: Add failing test**

Append to `tests/test_pipeline.py`:
```python
from pipeline import process_dpt_files
from openpyxl import load_workbook


def test_process_dpt_files_with_4_unzoned_files_treats_as_zone_1(tmp_path):
    """If exactly 4 non-merged files have no zone name, all become Zone 1."""
    # Build 4 minimal .dpt blobs with no 'zone' in filename
    content = b"1590,0.5\n2242,0.8\n1595,0.5\n2243,0.8\n"
    files = [(f"sample_{i}.dpt", content) for i in range(4)]
    filename, xlsx_bytes = process_dpt_files(files)
    assert filename.endswith("_FINAL_OUTPUT.xlsx")
    assert xlsx_bytes is not None and len(xlsx_bytes) > 0

    # Verify the output is a real xlsx with a ZONE 1 sheet
    out_path = tmp_path / "out.xlsx"
    out_path.write_bytes(xlsx_bytes)
    wb = load_workbook(out_path)
    assert "ZONE 1" in wb.sheetnames


def test_process_dpt_files_groups_by_zone_name():
    """Files named with zone numbers are grouped accordingly."""
    content = b"1590,0.5\n2242,0.8\n1595,0.5\n2243,0.8\n"
    files = [
        ("sample_zone1_a.dpt", content),
        ("sample_zone1_b.dpt", content),
        ("sample_zone2_a.dpt", content),
        ("sample_zone2_b.dpt", content),
    ]
    filename, xlsx_bytes = process_dpt_files(files)

    import io
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    assert "ZONE 1" in wb.sheetnames
    assert "ZONE 2" in wb.sheetnames


def test_process_dpt_files_progress_callback_called():
    """Progress callback receives status messages during processing."""
    content = b"1590,0.5\n2242,0.8\n"
    files = [(f"sample_{i}.dpt", content) for i in range(4)]
    messages: list[str] = []
    process_dpt_files(files, progress_callback=messages.append)
    assert len(messages) > 0
    # At least one message should mention zone or parse/merge work
    assert any("zone" in m.lower() or "parse" in m.lower() or "merge" in m.lower() for m in messages)


def test_process_dpt_files_empty_input_raises():
    """No valid files → raises ValueError."""
    with pytest.raises(ValueError):
        process_dpt_files([])


def test_process_dpt_files_all_files_corrupt_raises():
    """All files under 10 bytes → raises ValueError."""
    files = [(f"sample_{i}.dpt", b"x") for i in range(3)]
    with pytest.raises(ValueError):
        process_dpt_files(files)
```

- [ ] **Step 2: Run tests, verify they fail**

Run:
```bash
pytest tests/test_pipeline.py::test_process_dpt_files_with_4_unzoned_files_treats_as_zone_1 -v
```
Expected: ImportError on `process_dpt_files`.

- [ ] **Step 3: Implement `process_dpt_files()`**

Append to `pipeline.py`:
```python
import io as _io
from typing import Callable


def process_dpt_files(
    dpt_files: list[tuple[str, bytes]],
    progress_callback: Optional[Callable[[str], None]] = None,
) -> tuple[str, bytes]:
    """Run the full pipeline end-to-end.

    Returns: (output_filename, xlsx_bytes) suitable for st.download_button.
    Raises ValueError if no valid spectra could be parsed.
    """
    def _log(msg: str) -> None:
        if progress_callback is not None:
            progress_callback(msg)

    if not dpt_files:
        raise ValueError("No .dpt files provided")

    # Step 1: parse all .dpt files
    parsed: list[tuple[str, pd.DataFrame]] = []
    for name, content in dpt_files:
        df = parse_dpt(content)
        if df is None:
            _log(f"Skipped {name} (too small or empty)")
            continue
        parsed.append((name, df))
        _log(f"Parsed {name} ({df.shape[0]} rows)")

    if not parsed:
        raise ValueError("No valid .dpt files (all empty or corrupt)")

    # Determine output filename from the first parsed file
    first_stem = parsed[0][0].rsplit('.', 1)[0]
    output_filename = f"{first_stem}_FINAL_OUTPUT.xlsx"

    # Step 2: zone assignment
    # Special case: exactly 4 non-merged files with no zone match → all Zone 1
    zone_assignments: dict[int, list[tuple[str, pd.DataFrame]]] = {}
    detected_zones = [zone_for_filename(name) for name, _ in parsed]
    if len(parsed) == 4 and all(z is None for z in detected_zones):
        _log("Exactly 4 unzoned files — treating all as Zone 1")
        zone_assignments[1] = parsed
    else:
        for (name, df), zone in zip(parsed, detected_zones):
            if zone is None:
                _log(f"No zone matched for {name} — skipping")
                continue
            zone_assignments.setdefault(zone, []).append((name, df))

    if not zone_assignments:
        raise ValueError("No files matched any zone")

    # Step 3: merge each zone
    merged_by_zone: dict[int, pd.DataFrame] = {}
    for zone_num in sorted(zone_assignments.keys()):
        files_in_zone = zone_assignments[zone_num]
        _log(f"Merging ZONE {zone_num} ({len(files_in_zone)} files)")
        merged_by_zone[zone_num] = merge_zone(files_in_zone)

    # Step 4: combine into multi-sheet workbook
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        for zone_num in sorted(merged_by_zone.keys()):
            merged_by_zone[zone_num].to_excel(
                writer,
                sheet_name=f"ZONE {zone_num}",
                index=False,
                header=False,
            )
    _log(f"Final workbook ready: {output_filename}")
    return output_filename, buf.getvalue()
```

- [ ] **Step 4: Run tests, verify they pass**

Run:
```bash
pytest tests/test_pipeline.py -v
```
Expected: all tests pass (~20 total).

- [ ] **Step 5: Commit**

```bash
git add pipeline.py tests/test_pipeline.py
git commit -m "feat: add process_dpt_files end-to-end pipeline orchestrator"
```

---

## Task 8: Smoke test against real fixture data

**Files:**
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: Add fixture-based smoke test**

Append to `tests/test_pipeline.py`:
```python
def test_process_dpt_files_smoke_on_real_fixtures():
    """Run the pipeline on the real .dpt fixtures and verify it produces a valid xlsx."""
    from pathlib import Path
    import io as _io2

    fixture_dir = Path(__file__).parent / "fixtures" / "sample_dpt"
    dpt_paths = sorted(fixture_dir.glob("*.dpt"))
    assert len(dpt_paths) >= 1, "Fixture .dpt files missing"

    files = [(p.name, p.read_bytes()) for p in dpt_paths]
    filename, xlsx_bytes = process_dpt_files(files)

    # Output is a real xlsx
    wb = load_workbook(_io2.BytesIO(xlsx_bytes))
    assert len(wb.sheetnames) >= 1
    # Each sheet has some data
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        assert ws.max_row > 0
        assert ws.max_column > 0
```

- [ ] **Step 2: Run test, verify it passes**

Run:
```bash
pytest tests/test_pipeline.py::test_process_dpt_files_smoke_on_real_fixtures -v
```
Expected: PASS.

- [ ] **Step 3: Commit**

```bash
git add tests/test_pipeline.py
git commit -m "test: smoke test pipeline against real .dpt fixtures"
```

---

## Task 9: Optional — Snapshot test against original pipeline output

This task is **optional** — it requires running the original xlwings pipeline once to capture a golden output. If the user can't or won't run Excel, skip to Task 10. The unit + smoke tests above are sufficient for confidence in the port; this snapshot test is the gold-standard fidelity check.

**Files:**
- Create: `tests/fixtures/golden_output.xlsx`
- Modify: `tests/test_pipeline.py`

- [ ] **Step 1: User generates golden output from original pipeline**

User runs (Mac, with Excel installed):
```bash
# In a clean directory:
mkdir -p /tmp/ftir_golden/5
cp tests/fixtures/sample_dpt/*.dpt /tmp/ftir_golden/5/
# Adjust folder_path in main_convert_dpt_2_excel.py to /tmp/ftir_golden/5
# Then:
python3 main_convert_dpt_2_excel.py
# Excel will prompt for macro permissions — click Allow.
# Copy the resulting *_FINAL_OUTPUT.xlsx into tests/fixtures/:
cp /tmp/ftir_golden/5/*_FINAL_OUTPUT.xlsx tests/fixtures/golden_output.xlsx
```

If this step is impractical, document why in a comment in `tests/test_pipeline.py` and proceed to Task 10. The snapshot test below is skipped if the golden file is absent.

- [ ] **Step 2: Add snapshot test**

Append to `tests/test_pipeline.py`:
```python
import io as _io3
from pathlib import Path
import math


def _numeric_values(ws):
    """Yield all numeric cell values from a worksheet."""
    for row in ws.iter_rows(values_only=True):
        for v in row:
            if isinstance(v, (int, float)):
                yield float(v)


def test_pipeline_matches_golden_output():
    """Numerically compare new-pipeline output to a saved golden xlsx from the original."""
    fixture_dir = Path(__file__).parent / "fixtures"
    golden_path = fixture_dir / "golden_output.xlsx"
    if not golden_path.exists():
        pytest.skip("Golden output not present — see Task 9 step 1")

    dpt_paths = sorted((fixture_dir / "sample_dpt").glob("*.dpt"))
    files = [(p.name, p.read_bytes()) for p in dpt_paths]
    _, xlsx_bytes = process_dpt_files(files)

    new_wb = load_workbook(_io3.BytesIO(xlsx_bytes))
    golden_wb = load_workbook(golden_path)

    assert set(new_wb.sheetnames) == set(golden_wb.sheetnames), \
        f"Sheet names differ: new={new_wb.sheetnames} golden={golden_wb.sheetnames}"

    for sheet_name in new_wb.sheetnames:
        new_vals = list(_numeric_values(new_wb[sheet_name]))
        golden_vals = list(_numeric_values(golden_wb[sheet_name]))
        assert len(new_vals) == len(golden_vals), \
            f"Sheet {sheet_name}: cell-count mismatch ({len(new_vals)} vs {len(golden_vals)})"
        for i, (n, g) in enumerate(zip(new_vals, golden_vals)):
            # Tolerate tiny float diff. Mean column at the end may differ
            # slightly if Excel MATCH(-1) picked a different row than abs-idxmin.
            assert math.isclose(n, g, rel_tol=1e-6, abs_tol=1e-6), \
                f"Sheet {sheet_name}, value index {i}: new={n} golden={g}"
```

- [ ] **Step 3: Run test**

Run:
```bash
pytest tests/test_pipeline.py::test_pipeline_matches_golden_output -v
```
Expected (if golden present): PASS, or a specific cell mismatch that reveals a real discrepancy.
Expected (if golden absent): SKIPPED.

If the test fails with mismatches, investigate before proceeding — the discrepancy may indicate the Excel `MATCH(-1)` semantics matter for this dataset. Common fix: adjust `compute_macro_result` to use a "next value at or below target with descending-sort assumption" lookup instead of `abs().idxmin()`. Discuss with user before changing.

- [ ] **Step 4: Commit**

```bash
git add tests/fixtures/golden_output.xlsx tests/test_pipeline.py 2>/dev/null || git add tests/test_pipeline.py
git commit -m "test: add snapshot test against original pipeline golden output"
```

---

## Task 10: Build the Streamlit UI (`app.py`)

**Files:**
- Create: `app.py`

- [ ] **Step 1: Write `app.py`**

Write to `app.py`:
```python
"""Streamlit UI for the FTIR zone-merge pipeline."""

import streamlit as st

from pipeline import process_dpt_files


st.set_page_config(page_title="FTIR Zone Merger", page_icon="🧪", layout="centered")
st.title("FTIR Zone Merger")
st.write(
    "Upload one or more `.dpt` files. The app parses each spectrum, "
    "computes the standard peak-ratio metric, merges per-zone, and gives "
    "you a single multi-sheet Excel file to download."
)

# Session state init
if "result" not in st.session_state:
    st.session_state.result = None  # tuple (filename, bytes) or None
if "log" not in st.session_state:
    st.session_state.log = []  # list of str

uploaded = st.file_uploader(
    "Drop .dpt files here (or click to browse)",
    accept_multiple_files=True,
    type=["dpt"],
)

if uploaded:
    st.write(f"**{len(uploaded)} file(s) queued:**")
    for f in uploaded:
        st.write(f"- {f.name}")

col_run, col_reset = st.columns([3, 1])
with col_run:
    run_clicked = st.button("Process files", type="primary", disabled=not uploaded)
with col_reset:
    if st.button("Reset"):
        st.session_state.result = None
        st.session_state.log = []
        st.rerun()

if run_clicked:
    st.session_state.log = []
    st.session_state.result = None

    log_placeholder = st.empty()

    def on_progress(msg: str) -> None:
        st.session_state.log.append(msg)
        log_placeholder.code("\n".join(st.session_state.log))

    files_payload = [(f.name, f.getvalue()) for f in uploaded]
    try:
        with st.spinner("Processing..."):
            filename, xlsx_bytes = process_dpt_files(
                files_payload,
                progress_callback=on_progress,
            )
        st.session_state.result = (filename, xlsx_bytes)
        st.success("Done.")
    except ValueError as e:
        st.error(f"Pipeline error: {e}")
    except Exception as e:
        st.error(f"Unexpected error: {e}")
        st.exception(e)

if st.session_state.log and not run_clicked:
    st.code("\n".join(st.session_state.log))

if st.session_state.result is not None:
    fname, data = st.session_state.result
    st.download_button(
        label=f"Download {fname}",
        data=data,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
```

- [ ] **Step 2: Run Streamlit locally**

Run:
```bash
source .venv/bin/activate
streamlit run app.py
```
Expected: opens browser at http://localhost:8501 showing the FTIR Zone Merger page.

- [ ] **Step 3: Manual smoke test in browser**

In the browser:
1. Upload the two `.dpt` files from `tests/fixtures/sample_dpt/`.
2. Click **Process files**.
3. Verify the log shows "Parsed", "Merging ZONE 1", and "Final workbook ready".
4. Click the **Download** button — confirm an `.xlsx` file downloads.
5. Open the downloaded `.xlsx` in Excel/Numbers/LibreOffice. Verify it has at least one `ZONE` sheet with data.

If any step fails, the issue is in `app.py` or `pipeline.py` — debug before continuing.

- [ ] **Step 4: Stop Streamlit and commit**

Stop with Ctrl+C, then:
```bash
git add app.py
git commit -m "feat: add streamlit ui for upload/process/download"
```

---

## Task 11: README

**Files:**
- Create: `README.md`

- [ ] **Step 1: Write `README.md`**

Write to `README.md`:
```markdown
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
```

- [ ] **Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add README with usage and deployment instructions"
```

---

## Task 12: Final verification + push to GitHub

- [ ] **Step 1: Run full test suite**

```bash
source .venv/bin/activate
pytest tests/ -v
```
Expected: all tests pass (the snapshot test may be SKIPPED if Task 9 was skipped).

- [ ] **Step 2: Verify Streamlit app still launches cleanly**

```bash
streamlit run app.py
```
Open browser, do a quick upload/process/download cycle with the fixture files. Stop with Ctrl+C.

- [ ] **Step 3: Push to GitHub**

Create a new public GitHub repo (via web UI or `gh repo create`). Then:
```bash
git remote add origin <repo-url>
git branch -M main
git push -u origin main
```

- [ ] **Step 4: Deploy on Streamlit Community Cloud**

1. Go to <https://share.streamlit.io> and sign in with GitHub.
2. Click **New app**, pick the repo, branch `main`, main file path `app.py`.
3. Click **Deploy**. First build takes ~2 minutes.
4. Copy the resulting URL.

- [ ] **Step 5: Smoke test the deployed app**

Open the deployed URL. Upload the same fixture `.dpt` files. Verify the same output as the local test. If working, the project is done.
