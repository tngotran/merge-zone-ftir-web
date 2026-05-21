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
