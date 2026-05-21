"""Pure-Python FTIR zone-merge pipeline. No Excel, no xlwings."""

from __future__ import annotations

import io
from typing import Optional, Callable

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
    # IndexError means the first line is blank/whitespace; treat it as no header
    # so skip_blank_lines=True below will discard it and keep all data rows.
    try:
        first_field = first_line.strip().split(',')[0] if ',' in first_line else first_line.strip().split()[0]
        float(first_field)
        header = None
    except IndexError:
        header = None  # blank first line — let skip_blank_lines handle it
    except ValueError:
        header = 0  # text header row — skip it

    df = pd.read_csv(
        buf,
        sep=sep,
        engine='python',
        header=header,
        names=['Column1', 'Column2'],
        skip_blank_lines=True,
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
    buf = io.BytesIO()
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
