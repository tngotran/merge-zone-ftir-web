"""Pure-Python DAT-to-Excel conversion pipeline. No disk I/O, no Streamlit.

Ported from the local CLI script `dat_to_excel_converter.py`, which walks
folders on disk. The conversion logic (header detection, metadata extraction,
row padding, the Metadata + Data sheet layout) is preserved; the folder
globbing, the `print()` reporting, and the `pip install openpyxl` bootstrap are
replaced by uploaded bytes, a progress callback, and requirements.txt.
"""

from __future__ import annotations

import io
import zipfile
from typing import Callable, Optional

import pandas as pd

PSI_HEADERS = ['psi(°)', 'Intensity(a.u.)', 'Sigma_I(a.u.)']
Q_HEADERS = ['q(A-1)', 'I(q)', 'Sig(q)']


def detect_column_headers(lines: list[str]) -> tuple[list[str], int]:
    """Find the column names and the index of the first data line.

    Recognizes the two named layouts written by the beamline software. Failing
    that, scans for the first numeric row and infers names from its column
    count: 3 columns is the q layout, 6 is psi + q concatenated, anything else
    gets generic names.
    """
    for i, raw in enumerate(lines):
        line = raw.strip()
        if 'psi(°)' in line and 'Intensity(a.u.)' in line:
            return list(PSI_HEADERS), i + 1
        if 'q(A-1)' in line and 'I(q)' in line:
            return list(Q_HEADERS), i + 1

    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line or line.startswith('#'):
            continue
        try:
            numbers = [float(x) for x in line.split()]
        except ValueError:
            continue
        if len(numbers) >= 2:
            if len(numbers) == 3:
                return list(Q_HEADERS), i
            if len(numbers) == 6:
                return PSI_HEADERS + Q_HEADERS, i
            return [f'Column_{j + 1}' for j in range(len(numbers))], i

    return ['Column_1', 'Column_2', 'Column_3'], 0


def parse_dat(content: bytes) -> Optional[dict]:
    """Parse a .dat file's raw bytes into metadata, data rows, and headers.

    Returns None when there is nothing to parse, which callers treat as a skip
    signal.
    """
    if not content:
        return None

    if content[:2] in (b'\xff\xfe', b'\xfe\xff'):
        text = content.decode('utf-16', errors='ignore')
    else:
        text = content.decode('utf-8', errors='ignore')

    lines = text.splitlines()
    if not any(line.strip() for line in lines):
        return None

    headers, data_start_line = detect_column_headers(lines)

    # Metadata lives in the '# key value' lines above the data block. The
    # '####...' banner has no space after the '#', so it never matches.
    metadata: dict[str, str] = {}
    for line in lines[:data_start_line]:
        line = line.strip()
        if not line.startswith('# ') or len(line) <= 2:
            continue
        if any(h in line for h in ('psi(°)', 'q(A-1)', 'Intensity', 'I(q)')):
            continue
        parts = line[2:].split(None, 1)
        if len(parts) == 2:
            metadata[parts[0]] = parts[1]

    data: list[list[float]] = []
    for line in lines[data_start_line:]:
        line = line.strip()
        if not line or line.startswith('#'):
            continue
        try:
            numbers = [float(x) for x in line.split()]
        except ValueError:
            continue
        if len(numbers) >= 2:
            data.append(numbers)

    return {'metadata': metadata, 'data': data, 'headers': headers}


def dat_to_xlsx_bytes(parsed: dict) -> Optional[bytes]:
    """Render parsed .dat content as a two-sheet workbook.

    Returns None when there is neither metadata nor data. Writing that case
    would close a workbook with zero sheets, which openpyxl rejects with
    "At least one sheet must be visible".
    """
    metadata = parsed.get('metadata') or {}
    data = parsed.get('data') or []
    max_cols = max((len(row) for row in data), default=0)

    if not metadata and max_cols == 0:
        return None

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        if metadata:
            pd.DataFrame(
                list(metadata.items()), columns=['Parameter', 'Value']
            ).to_excel(writer, sheet_name='Metadata', index=False)

        if max_cols:
            headers = list(parsed.get('headers') or [])
            if len(headers) < max_cols:
                headers.extend(
                    f'Column_{i + 1}' for i in range(len(headers), max_cols)
                )
            else:
                headers = headers[:max_cols]

            padded = [
                (row + [None] * (max_cols - len(row)))[:max_cols] for row in data
            ]
            pd.DataFrame(padded, columns=headers).to_excel(
                writer, sheet_name='Data', index=False
            )

    return buf.getvalue()


def convert_dat_files(
    dat_files: list[tuple[str, bytes]],
    progress_callback: Optional[Callable[[str], None]] = None,
) -> tuple[str, bytes]:
    """Convert uploaded .dat files into a zip of one .xlsx per input.

    Returns: (zip_filename, zip_bytes) suitable for st.download_button.
    Raises ValueError if no file could be converted.
    """
    def _log(msg: str) -> None:
        if progress_callback is not None:
            progress_callback(msg)

    if not dat_files:
        raise ValueError("No .dat files provided")

    seen_stems: dict[str, int] = {}
    converted = 0
    skipped = 0

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for name, content in dat_files:
            try:
                parsed = parse_dat(content)
                xlsx_bytes = dat_to_xlsx_bytes(parsed) if parsed else None
                reason = None if xlsx_bytes else "no readable data"
            except Exception as e:  # one bad file must not kill the batch
                xlsx_bytes = None
                reason = str(e)

            if xlsx_bytes is None:
                _log(f"Skipped {name} ({reason})")
                skipped += 1
                continue

            stem = name.rsplit('.', 1)[0]
            count = seen_stems.get(stem, 0) + 1
            seen_stems[stem] = count
            out_name = f"{stem}.xlsx" if count == 1 else f"{stem}_{count}.xlsx"

            zf.writestr(out_name, xlsx_bytes)
            converted += 1
            _log(
                f"Converted {name} → {out_name} "
                f"({len(parsed['data'])} rows, {' | '.join(parsed['headers'][:3])})"
            )

    if converted == 0:
        raise ValueError(
            "No .dat files could be converted (all empty or unreadable)"
        )

    _log(f"Done: {converted} converted, {skipped} skipped")

    zip_name = f"{dat_files[0][0].rsplit('.', 1)[0]}_CONVERTED.zip"
    return zip_name, buf.getvalue()
