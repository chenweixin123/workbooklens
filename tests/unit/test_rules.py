from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.models import PatchKind
from workbooklens.scanner import ScanResult, scan_workbook


@pytest.fixture(scope="module")
def demo_scan(tmp_path_factory: pytest.TempPathFactory) -> ScanResult:
    directory = tmp_path_factory.mktemp("rules-demo")
    path = directory / "demo.xlsx"
    generate_demo_workbook(path)
    config = {"keys": [{"sheet": "Sales", "range": "A2:A22", "ignore_blank": True}]}
    return scan_workbook(path, config=config)


def test_demo_exercises_fourteen_rules_and_all_patch_kinds(demo_scan: ScanResult) -> None:
    rule_ids = {finding.rule_id for finding in demo_scan.findings}
    assert rule_ids == {
        "WL001_BROKEN_REFERENCE",
        "WL003_BLANK_IN_FORMULA_BAND",
        "WL004_HARDCODED_VALUE_IN_FORMULA_BAND",
        "WL005_SUSPICIOUS_SUM_BOUNDARY",
        "WL006_NUMERIC_TEXT",
        "WL007_STYLE_OUTLIER",
        "WL008_HIDDEN_NONEMPTY_DATA",
        "WL009_EXTERNAL_LINK",
        "WL010_VOLATILE_OR_FRAGILE_FUNCTION",
        "WL011_ERROR_CELL",
        "WL012_DUPLICATE_CONFIGURED_KEY",
        "WL013_BROKEN_DEFINED_NAME",
        "WL014_MERGED_CELL_IN_DATA_REGION",
        "WL015_INCONSISTENT_DATA_VALIDATION",
    }
    assert {patch.kind for patch in demo_scan.patches} == set(PatchKind)
    assert all(patch.safe and float(patch.confidence) >= 0.95 for patch in demo_scan.patches)


def test_formula_pattern_outlier_rule_and_safe_proposal(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Band"
    for column in range(2, 22):
        worksheet.cell(1, column, column)
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["K2"] = "=K1*3"
    path = tmp_path / "outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = [
        finding for finding in scan.findings if finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER"
    ]
    assert len(findings) == 1
    assert findings[0].location == "K2"
    patches = [patch for patch in scan.patches if patch.cell == "K2"]
    assert len(patches) == 1
    assert patches[0].kind == PatchKind.SET_FORMULA
    assert patches[0].after == "=K1*2"


def test_unsupported_formula_suppresses_automatic_band_repair(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Band"
    for column in range(2, 22):
        worksheet.cell(1, column, column)
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["K2"] = "=Table1[Amount]"
    path = tmp_path / "unsupported-band.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" for finding in scan.findings)
    assert not any(patch.cell == "K2" for patch in scan.patches)


def test_numeric_text_excludes_leading_zero_identifiers(demo_scan: ScanResult) -> None:
    locations = {
        finding.location
        for finding in demo_scan.findings
        if finding.rule_id == "WL006_NUMERIC_TEXT"
    }
    assert locations == {"B10"}


def test_hidden_adjacent_row_suppresses_sum_boundary(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 10):
        worksheet.cell(row, 2, row)
    worksheet["B10"] = "=SUM(B2:B8)"
    worksheet.row_dimensions[9].hidden = True
    path = tmp_path / "hidden-boundary.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    assert not any(finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY" for finding in scan.findings)


def test_sum_boundary_does_not_cross_explicit_sheet_reference(tmp_path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    assert summary is not None
    summary.title = "Summary"
    source = workbook.create_sheet("Source")
    for row in range(2, 10):
        source.cell(row, 2, row)
    summary["B9"] = 999
    summary["B10"] = "=SUM(Source!B2:B8)"
    path = tmp_path / "cross-sheet-sum.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    assert not any(finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY" for finding in scan.findings)


def test_sum_boundary_requires_literal_numeric_adjacent_cell(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 9):
        worksheet.cell(row, 2, row)
    worksheet["B9"] = "=1+1"
    worksheet["B10"] = "=SUM(B2:B8)"
    path = tmp_path / "formula-adjacent-sum.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    assert not any(finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY" for finding in scan.findings)


def test_scientific_notation_text_is_not_auto_converted(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Measure", "Context"])
    for row in range(2, 10):
        worksheet.append([row, f"R{row}"])
    worksheet["A5"] = "1e3"
    path = tmp_path / "scientific-text.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A5"
        for finding in scan.findings
    )


def test_hidden_grouped_columns_report_full_range(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["B1"] = "hidden B"
    worksheet["C1"] = "hidden C"
    worksheet.column_dimensions.group("B", "C", hidden=True)
    path = tmp_path / "hidden-columns.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    locations = {
        finding.location
        for finding in scan.findings
        if finding.rule_id == "WL008_HIDDEN_NONEMPTY_DATA"
    }
    assert "B:C" in locations


def test_merged_header_is_not_reported_as_data_body_merge(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.merge_cells("A1:B1")
    worksheet["A1"] = "Header"
    for row in range(2, 8):
        worksheet.cell(row, 1, row)
        worksheet.cell(row, 2, row * 2)
    path = tmp_path / "merged-header.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL014_MERGED_CELL_IN_DATA_REGION" for finding in scan.findings
    )


def test_all_fifteen_builtin_rule_ids_are_stable(demo_scan: ScanResult, tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for column in range(1, 21):
        worksheet.cell(1, column, column)
        worksheet.cell(2, column, f"={worksheet.cell(1, column).coordinate}*2")
    worksheet["J2"] = "=J1*3"
    path = tmp_path / "rule-two.xlsx"
    workbook.save(path)
    workbook.close()
    second_scan = scan_workbook(path)
    all_ids = {finding.rule_id for finding in demo_scan.findings} | {
        finding.rule_id for finding in second_scan.findings
    }
    assert all_ids == {
        f"WL{number:03d}_{suffix}"
        for number, suffix in enumerate(
            [
                "BROKEN_REFERENCE",
                "FORMULA_PATTERN_OUTLIER",
                "BLANK_IN_FORMULA_BAND",
                "HARDCODED_VALUE_IN_FORMULA_BAND",
                "SUSPICIOUS_SUM_BOUNDARY",
                "NUMERIC_TEXT",
                "STYLE_OUTLIER",
                "HIDDEN_NONEMPTY_DATA",
                "EXTERNAL_LINK",
                "VOLATILE_OR_FRAGILE_FUNCTION",
                "ERROR_CELL",
                "DUPLICATE_CONFIGURED_KEY",
                "BROKEN_DEFINED_NAME",
                "MERGED_CELL_IN_DATA_REGION",
                "INCONSISTENT_DATA_VALIDATION",
            ],
            start=1,
        )
    }
