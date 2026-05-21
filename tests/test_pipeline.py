import pandas as pd
import pytest
from pipeline import compute_macro_result
import io
from pipeline import parse_dpt


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
