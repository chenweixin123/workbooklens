from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font

from workbooklens.models import Severity
from workbooklens.rules import RuleRegistry
from workbooklens.rules.layout_geometry import OrphanMicroLabelRule
from workbooklens.scanner import ScanResult, scan_workbook


def _workbook() -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.merge_cells("B1:F2")
    worksheet["B1"] = "Quarterly Sales"
    worksheet["B1"].font = Font(size=18, bold=True)
    worksheet.append([])
    worksheet.append(["ID", "Date", "Customer", "Region", "Sales", "Status"])
    for row in range(5, 14):
        worksheet.append(
            [
                f"R{row:03d}",
                f"2026-08-{row:02d}",
                f"Customer {row}",
                "West",
                row * 100,
                "Open",
            ]
        )
    return workbook


def _scan(tmp_path: Path, workbook: Workbook, name: str) -> ScanResult:
    path = tmp_path / name
    workbook.save(path)
    workbook.close()
    return scan_workbook(path, registry=RuleRegistry([OrphanMicroLabelRule()]))


def test_unique_tiny_vivid_label_beside_title_is_review_only(tmp_path: Path) -> None:
    workbook = _workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "REVIEW?"
    worksheet["A1"].font = Font(size=7, color="FFFF0000")

    scan = _scan(tmp_path, workbook, "orphan-micro-label.xlsx")

    assert [(finding.rule_id, finding.location) for finding in scan.findings] == [
        ("WL056_ORPHAN_MICRO_LABEL", "A1")
    ]
    finding = scan.findings[0]
    assert finding.severity == Severity.INFO
    assert finding.confidence.root == 0.86
    assert finding.evidence.observed == {
        "font_size": 7.0,
        "font_rgb": "FF0000",
        "adjacent_title_range": "B1:F2",
    }
    assert finding.evidence.details["proof_level"] == "advisory"
    assert finding.patch_ids == []
    assert scan.patches == []


def test_small_footnote_below_table_is_not_an_orphan_title_label(tmp_path: Path) -> None:
    workbook = _workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A15"] = "Source: reviewed ledger"
    worksheet["A15"].font = Font(size=7, color="FF666666")

    scan = _scan(tmp_path, workbook, "small-footnote.xlsx")

    assert not scan.findings


def test_normal_size_red_notice_beside_title_is_not_a_micro_label(tmp_path: Path) -> None:
    workbook = _workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "Important"
    worksheet["A1"].font = Font(size=11, color="FFFF0000")

    scan = _scan(tmp_path, workbook, "normal-red-notice.xlsx")

    assert not scan.findings


def test_multiple_matching_title_labels_are_not_treated_as_one_orphan(tmp_path: Path) -> None:
    workbook = _workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for coordinate in ("A1", "G1"):
        worksheet[coordinate] = "Tag"
        worksheet[coordinate].font = Font(size=7, color="FFFF0000")

    scan = _scan(tmp_path, workbook, "paired-micro-labels.xlsx")

    assert not scan.findings
