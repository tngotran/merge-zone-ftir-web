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
