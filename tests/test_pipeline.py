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
