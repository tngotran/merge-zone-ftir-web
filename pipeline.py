"""Pure-Python FTIR zone-merge pipeline. No Excel, no xlwings."""

from __future__ import annotations

import io
from typing import Optional

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
