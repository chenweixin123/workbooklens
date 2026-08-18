from __future__ import annotations

import pytest

from workbooklens.formulas import (
    UnsupportedFormulaError,
    analyze_formula,
    normalize_formula,
    translate_formula,
)


def test_relative_copy_has_same_structural_signature() -> None:
    assert normalize_formula("=A1+$B$2+C$3+$D4", "E5") == normalize_formula(
        "=B2+$B$2+D$3+$D5", "F6"
    )


def test_signature_understands_quoted_sheet_and_ranges() -> None:
    signature = normalize_formula("=SUM('Detail Data'!A2:B4)", "C5")
    assert "'DETAIL DATA'!" in signature
    assert "R[-3]C[-2]:R[-1]C[-1]" in signature


def test_translate_mixed_references() -> None:
    translated = translate_formula("=A1+$B1+C$3+$D$4", "E5", "F7")
    assert translated == "=B3+$B3+D$3+$D$4"


@pytest.mark.parametrize(
    "formula,reason",
    [
        ("='[Budget.xlsx]Plan'!A1", "external"),
        ("=SUM(Table1[Amount])", "structured"),
        ("=FILTER(A1:A3,B1:B3=1)", "dynamic or advanced"),
        ("=_xlfn._xlws.FILTER(A1:A3,B1:B3=1)", "dynamic or advanced"),
    ],
)
def test_translate_rejects_unsupported_references(formula: str, reason: str) -> None:
    with pytest.raises(UnsupportedFormulaError, match=reason):
        translate_formula(formula, "A1", "A2")


def test_formula_features_are_nonexecuting_and_explicit() -> None:
    features = analyze_formula("=OFFSET(A1,1,0)+NOW()+SUM(B:B)+'[Book.xlsx]S'!C1")
    assert features.volatile_functions == ("OFFSET", "NOW")
    assert features.has_whole_column_reference
    assert features.external_references == ("'[Book.xlsx]S'!C1",)


def test_normalize_rejects_non_formula() -> None:
    with pytest.raises(ValueError, match="beginning"):
        normalize_formula("A1", "B2")
