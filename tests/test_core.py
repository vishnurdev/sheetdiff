from pathlib import Path

import pandas as pd
import pytest

from sheetdiff import SheetDiffError, compare, compare_files, write_excel, write_html
from sheetdiff.cli import main

SAMPLES = Path(__file__).resolve().parent.parent / "samples"


def df(rows, cols=("id", "name", "price")):
    return pd.DataFrame(rows, columns=list(cols))


def test_identical_tables():
    a = df([[1, "Pen", 1.5], [2, "Book", 9.0]])
    r = compare(a, a.copy(), "id")
    assert r.is_identical
    assert r.changed.empty and r.added.empty and r.removed.empty


def test_added_removed_and_changed():
    a = df([[1, "Pen", 1.5], [2, "Book", 9.0], [3, "Cup", 4.0]])
    b = df([[1, "Pen", 2.0], [2, "Book", 9.0], [4, "Lamp", 20.0]])
    r = compare(a, b, "id")
    assert r.added["id"].tolist() == ["4"]
    assert r.removed["id"].tolist() == ["3"]
    assert len(r.changed) == 1
    row = r.changed.iloc[0]
    assert row["Column"] == "price" and row["Difference"] == pytest.approx(0.5)


def test_tolerance_ignores_small_numeric_changes():
    a = df([[1, "Pen", 1.50]])
    b = df([[1, "Pen", 1.504]])
    assert not compare(a, b, "id").is_identical
    assert compare(a, b, "id", tolerance=0.01).is_identical


def test_ignore_case():
    a = df([[1, "Pen", 1.5]])
    b = df([[1, "PEN", 1.5]])
    assert not compare(a, b, "id").is_identical
    assert compare(a, b, "id", ignore_case=True).is_identical


def test_numeric_keys_match_across_formats():
    a = df([[1001.0, "Pen", 1.5]])
    b = df([["1001", "Pen", 1.5]])
    assert compare(a, b, "id").is_identical


def test_multi_column_key():
    cols = ("state", "coverage", "rate")
    a = df([["TX", "Dwelling", 100], ["CA", "Dwelling", 120]], cols)
    b = df([["TX", "Dwelling", 100], ["CA", "Dwelling", 125]], cols)
    r = compare(a, b, ["state", "coverage"])
    assert r.changed[["state", "coverage"]].values.tolist() == [["CA", "Dwelling"]]


def test_missing_key_raises():
    with pytest.raises(SheetDiffError):
        compare(df([[1, "a", 1]]), df([[1, "a", 1]]), "nope")


def test_column_differences_and_duplicates():
    a = pd.DataFrame({"id": [1, 1, 2], "x": [1, 1, 2]})
    b = pd.DataFrame({"id": [1, 2], "y": [1, 2]})
    r = compare(a, b, "id")
    assert r.columns_only_left == ["x"] and r.columns_only_right == ["y"]
    assert r.duplicate_keys_left == 1
    assert not r.is_identical


def test_totals():
    a = df([[1, "Pen", 1.5], [2, "Book", 9.0]])
    b = df([[1, "Pen", 2.5], [2, "Book", 9.0]])
    r = compare(a, b, "id", left_name="old", right_name="new")
    row = r.totals.set_index("Column").loc["price"]
    assert row["old"] == 10.5 and row["new"] == 11.5 and row["new − old"] == 1.0


def test_sample_files_and_reports(tmp_path):
    r = compare_files(SAMPLES / "rates_2025.csv", SAMPLES / "rates_2026.csv", "rate_id")
    assert r.added["rate_id"].tolist() == ["R1009"]
    assert r.removed["rate_id"].tolist() == ["R1008"]
    assert set(r.changed["rate_id"]) == {"R1001", "R1003", "R1004", "R1006"}

    xlsx = write_excel(r, tmp_path / "report.xlsx")
    sheets = pd.read_excel(xlsx, sheet_name=None)
    assert set(sheets) == {"Summary", "Totals", "Changed", "Added", "Removed"}

    page = write_html(r, tmp_path / "report.html").read_text(encoding="utf-8")
    assert "R1009" in page and "Differences found" in page


def test_excel_input_roundtrip(tmp_path):
    a = df([[1, "Pen", 1.5]])
    a.to_excel(tmp_path / "a.xlsx", index=False)
    a.to_csv(tmp_path / "b.csv", index=False)
    assert compare_files(tmp_path / "a.xlsx", tmp_path / "b.csv", "id").is_identical


def test_cli_exit_codes(tmp_path, capsys):
    left, right = str(SAMPLES / "rates_2025.csv"), str(SAMPLES / "rates_2026.csv")
    assert main([left, left, "-k", "rate_id", "--fail-on-diff", "-q"]) == 0
    assert main([left, right, "-k", "rate_id", "--fail-on-diff", "-q"]) == 1
    assert main([left, right, "-k", "missing", "-q"]) == 2
    assert main([left, right, "-k", "rate_id", "-q", "--html", str(tmp_path / "r.html")]) == 0
    assert (tmp_path / "r.html").exists()
