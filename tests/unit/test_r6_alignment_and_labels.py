from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment

from workbooklens.rules import RuleRegistry
from workbooklens.rules.builtin import TextDisplayRiskRule
from workbooklens.rules.formula_semantics import UnlabeledAggregateRoleRule
from workbooklens.scanner import ScanResult, scan_workbook


def _scan(tmp_path: Path, workbook: Workbook, rule: object, name: str) -> ScanResult:
    path = tmp_path / name
    workbook.save(path)
    workbook.close()
    return scan_workbook(path, registry=RuleRegistry([rule]))  # type: ignore[list-item]


def _three_column_table() -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["ID", "Name", "Notes"])
    for row in range(2, 9):
        worksheet.append([f"R{row:03d}", f"Name {row}", "ok"])
    worksheet.column_dimensions["B"].width = 200
    worksheet.column_dimensions["C"].width = 4
    return workbook


def test_right_aligned_text_at_right_table_edge_overflows_left_inside_table(
    tmp_path: Path,
) -> None:
    workbook = _three_column_table()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["B5"] = None
    worksheet["C5"] = "Right-aligned note overflows into a wide blank cell on its left."
    worksheet["C5"].alignment = Alignment(horizontal="right")

    scan = _scan(tmp_path, workbook, TextDisplayRiskRule(), "right-aligned-overflow.xlsx")

    assert not [
        finding
        for finding in scan.findings
        if finding.rule_id == "WL016_TEXT_DISPLAY_RISK" and finding.location == "C5"
    ]


def _summary_book() -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Quantity", "Category", "Sales", "Gross"])
    for row in range(2, 9):
        worksheet.append([f"Item {row}", row, "A", row * 10, row * 12])
    worksheet["B10"] = "=AVERAGE(B2:B8)"
    worksheet["C10"] = "TOTAL"
    worksheet["D10"] = "=SUM(D2:D8)"
    worksheet["E10"] = "=SUM(E2:E8)"
    return workbook


def test_nearby_upper_left_metric_label_suppresses_unlabeled_aggregate_warning(
    tmp_path: Path,
) -> None:
    workbook = _summary_book()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Average quantity"

    scan = _scan(tmp_path, workbook, UnlabeledAggregateRoleRule(), "labelled-average.xlsx")

    assert not [
        finding for finding in scan.findings if finding.rule_id == "WL052_UNLABELED_AGGREGATE_ROLE"
    ]


def test_unrelated_nearby_summary_label_does_not_hide_unlabeled_average(
    tmp_path: Path,
) -> None:
    workbook = _summary_book()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Total sales"

    scan = _scan(tmp_path, workbook, UnlabeledAggregateRoleRule(), "unrelated-total-label.xlsx")

    assert [(finding.rule_id, finding.location) for finding in scan.findings] == [
        ("WL052_UNLABELED_AGGREGATE_ROLE", "B10")
    ]
