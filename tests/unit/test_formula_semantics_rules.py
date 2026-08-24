from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace

import pytest
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.formulas.ir import (
    WorksheetContentIndex,
    parse_formula_ir,
    worksheet_content_index,
)
from workbooklens.models import Severity
from workbooklens.rules import RuleRegistry
from workbooklens.rules.formula_semantics import (
    AggregateRangeCoverageRule,
    CrossSheetReferenceValidityRule,
    DegenerateFormulaRule,
    FormulaFormatRoleMismatchRule,
    FormulaInNotesColumnRule,
    _formula_cells,
)
from workbooklens.scanner import ScanResult, scan_workbook


def _scan(
    tmp_path: Path,
    workbook: Workbook,
    rule: object,
    *,
    name: str = "formula-semantics.xlsx",
) -> ScanResult:
    path = tmp_path / name
    workbook.save(path)
    workbook.close()
    return scan_workbook(path, registry=RuleRegistry([rule]))  # type: ignore[list-item]


def _amount_table(workbook: Workbook) -> None:
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Sales"
    worksheet.append(["Item", "Amount"])
    for index in range(1, 8):
        worksheet.append([f"Item {index}", index * 10])


def _cross_sheet_book(*, blank_row: int | None = None) -> Workbook:
    workbook = Workbook()
    source = workbook.active
    assert source is not None
    source.title = "Source Data"
    source.append(["ID", "Amount", "Category"])
    for row in range(2, 9):
        source.append([row - 1, None if row == blank_row else row * 10, "A"])
    dashboard = workbook.create_sheet("Dashboard")
    dashboard["A1"] = "Result"
    return workbook


def test_formula_ir_resolves_quoted_cross_sheet_aggregate() -> None:
    ir = parse_formula_ir(
        "=SUM('Sales Data'!$B$2:$B$8)",
        origin_sheet="Dashboard",
        origin_coordinate="B5",
    )

    assert len(ir.references) == 1
    reference = ir.references[0]
    assert reference.sheet == "Sales Data"
    assert reference.is_cross_sheet
    assert reference.local_range == "B2:B8"
    assert [(item.function, item.reference.raw) for item in ir.aggregate_references] == [
        ("SUM", "'Sales Data'!$B$2:$B$8")
    ]


@pytest.mark.parametrize(
    ("formula", "kind"),
    [
        ("=A1/A1", "self_division"),
        ("=SUM(A1:A3)-SUM(A1:A3)", "self_subtraction"),
        ("=A1*0", "multiply_by_zero"),
        ("=0*A1", "multiply_by_zero"),
    ],
)
def test_formula_ir_identifies_only_top_level_degeneracy(formula: str, kind: str) -> None:
    ir = parse_formula_ir(formula, origin_sheet="Sheet", origin_coordinate="B1")

    assert ir.degeneracy is not None
    assert ir.degeneracy.kind == kind


@pytest.mark.parametrize("formula", ["=IF(A1=0,0,A1/A1)", "=A1/A1+1", '="A1/A1"'])
def test_formula_ir_does_not_flag_nested_or_display_text(formula: str) -> None:
    assert parse_formula_ir(formula, "Sheet", "B1").degeneracy is None


def test_aggregate_range_reports_omitted_tail_rows(tmp_path: Path) -> None:
    workbook = Workbook()
    _amount_table(workbook)
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Total"
    worksheet["B9"] = "=SUM(B2:B6)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert [(finding.location, finding.rule_id) for finding in scan.findings] == [
        ("B9", "WL042_AGGREGATE_RANGE_COVERAGE")
    ]
    assert scan.findings[0].evidence.details["excluded_cells"] == ["B7", "B8"]
    assert scan.findings[0].evidence.expected == {"reviewed_reference": "B2:B8"}
    assert scan.patches == []


def test_aggregate_range_accepts_complete_table_body(tmp_path: Path) -> None:
    workbook = Workbook()
    _amount_table(workbook)
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Total"
    worksheet["B9"] = "=SUM(B2:B8)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert scan.findings == []


def test_aggregate_range_does_not_guess_an_intentional_window(tmp_path: Path) -> None:
    workbook = Workbook()
    _amount_table(workbook)
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Last five periods"
    worksheet["B9"] = "=SUM(B4:B8)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert scan.findings == []


def test_aggregate_range_catches_one_omitted_leading_row(tmp_path: Path) -> None:
    workbook = Workbook()
    _amount_table(workbook)
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A9"] = "Grand total"
    worksheet["B9"] = "=SUM(B3:B8)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert len(scan.findings) == 1
    assert scan.findings[0].evidence.details["excluded_cells"] == ["B2"]


def test_cross_sheet_metric_label_reports_truncated_aggregate(tmp_path: Path) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["A2"] = "Gross Sales"
    dashboard["B2"] = "=SUM('Source Data'!B2:B5)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert [(finding.location, finding.rule_id) for finding in scan.findings] == [
        ("B2", "WL042_AGGREGATE_RANGE_COVERAGE")
    ]
    assert scan.findings[0].evidence.details["excluded_cells"] == ["B6", "B7", "B8"]


@pytest.mark.parametrize("label", ["Q1 Sales", "Rolling Sales", "Last period revenue"])
def test_cross_sheet_window_label_does_not_imply_full_range(tmp_path: Path, label: str) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["A2"] = label
    dashboard["B2"] = "=SUM('Source Data'!B2:B5)"

    scan = _scan(tmp_path, workbook, AggregateRangeCoverageRule())

    assert scan.findings == []


def test_cross_sheet_reference_treats_blank_future_input_as_advisory(tmp_path: Path) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Source Data'!F999"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert len(scan.findings) == 1
    assert scan.findings[0].location == "B1"
    assert scan.findings[0].evidence.details["reason"] == "blank_outside_content"
    assert scan.findings[0].severity == Severity.INFO
    assert float(scan.findings[0].confidence) == 0.7
    assert scan.findings[0].evidence.details["proof_level"] == "advisory"
    assert "planned future input" in scan.findings[0].explanation
    assert scan.patches == []


def test_cross_sheet_missing_worksheet_remains_proven_error(tmp_path: Path) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Missing Input'!A1"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert len(scan.findings) == 1
    finding = scan.findings[0]
    assert finding.evidence.details["reason"] == "missing_worksheet"
    assert finding.severity == Severity.ERROR
    assert float(finding.confidence) == 1.0
    assert finding.evidence.details["proof_level"] == "strong_structural"


def test_sparse_content_index_queries_are_sorted_bounded_and_cached() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "Header"
    worksheet["C1"] = "Other"
    for row in range(2, 42):
        worksheet.cell(row, 2, row)
    worksheet["D5"] = "Flag"

    index = WorksheetContentIndex.from_worksheet(worksheet)

    assert index.content_bounds == (1, 41, 1, 4)
    assert [cell.coordinate for cell in index.iter_nonblank_row(1)] == ["A1", "C1"]
    assert [cell.coordinate for cell in index.iter_nonblank_column(2, min_row=3, max_row=5)] == [
        "B3",
        "B4",
        "B5",
    ]
    assert [
        cell.coordinate
        for cell in index.iter_nonblank_rectangle(
            min_row=1,
            max_row=41,
            min_column=2,
            max_column=2,
            limit=32,
        )
    ] == [f"B{row}" for row in range(2, 34)]

    context = SimpleNamespace(analysis_cache={})
    assert worksheet_content_index(context, worksheet) is worksheet_content_index(
        context, worksheet
    )


def test_formula_cells_are_sorted_cached_and_do_not_build_content_index(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["C3"] = "=A1"
    worksheet["A1"] = "=1"
    worksheet["B2"] = "=2"
    calls: list[str] = []

    def unexpected_index(
        cls: type[WorksheetContentIndex], sheet: Worksheet
    ) -> WorksheetContentIndex:
        calls.append(sheet.title)
        raise AssertionError("formula collection must not build a content index")

    monkeypatch.setattr(WorksheetContentIndex, "from_worksheet", classmethod(unexpected_index))
    context = SimpleNamespace(analysis_cache={})

    first = _formula_cells(context, worksheet)
    second = _formula_cells(context, worksheet)

    assert first is second
    assert [cell.coordinate for cell in first] == ["A1", "B2", "C3"]
    assert calls == []
    assert context.analysis_cache["formula-semantics-formula-cells-v1"][worksheet.title] is first


@pytest.mark.parametrize("limit", [0, -1])
def test_sparse_iterators_return_no_cells_for_nonpositive_limits(limit: int) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "first"
    worksheet["B1"] = "row peer"
    worksheet["A2"] = "column peer"
    index = WorksheetContentIndex.from_worksheet(worksheet)

    assert list(index.iter_nonblank_row(1, limit=limit)) == []
    assert list(index.iter_nonblank_column(1, limit=limit)) == []
    assert (
        list(
            index.iter_nonblank_rectangle(
                min_row=1,
                max_row=2,
                min_column=1,
                max_column=2,
                limit=limit,
            )
        )
        == []
    )


def test_sparse_axis_indexes_use_flat_tuples_and_support_tail_queries() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    cells = [worksheet.cell(row, 1, row) for row in range(1, 5001)]

    index = WorksheetContentIndex.from_cells(worksheet, cells)

    assert isinstance(index._rows._cells, tuple)
    assert isinstance(index._columns._cells, tuple)
    assert len(index._rows._cells) == 5000
    assert len(index._columns._cells) == 5000
    assert not hasattr(index, "_row_columns")
    assert not hasattr(index, "_column_rows")
    assert [cell.coordinate for cell in index.iter_nonblank_row(5000)] == ["A5000"]
    assert [
        cell.coordinate for cell in index.iter_nonblank_column(1, min_row=4998, max_row=5000)
    ] == ["A4998", "A4999", "A5000"]
    assert [
        cell.coordinate
        for cell in index.iter_nonblank_rectangle(
            min_row=4998,
            max_row=5000,
            min_column=1,
            max_column=1,
        )
    ] == ["A4998", "A4999", "A5000"]


def test_sparse_rectangle_is_row_major_and_applies_one_global_limit() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for coordinate in ("D1", "B1", "C2", "A2", "B3"):
        worksheet[coordinate] = coordinate
    index = WorksheetContentIndex.from_worksheet(worksheet)

    assert [
        cell.coordinate
        for cell in index.iter_nonblank_rectangle(
            min_row=1,
            max_row=3,
            min_column=1,
            max_column=4,
            limit=4,
        )
    ] == ["B1", "D1", "A2", "C2"]


def test_sparse_index_materializes_one_shot_cell_iterable_once() -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    cells = [
        worksheet.cell(1, 3, "C1"),
        worksheet.cell(2, 1, "A2"),
        worksheet.cell(3, 2, "B3"),
    ]
    consumed: list[str] = []

    def cell_stream():
        for cell in cells:
            consumed.append(cell.coordinate)
            yield cell

    index = WorksheetContentIndex.from_cells(worksheet, cell_stream())

    assert consumed == ["C1", "A2", "B3"]
    assert [cell.coordinate for cell in index.iter_nonblank_row(2)] == ["A2"]
    assert [cell.coordinate for cell in index.iter_nonblank_column(2)] == ["B3"]
    assert [
        cell.coordinate
        for cell in index.iter_nonblank_rectangle(
            min_row=1,
            max_row=3,
            min_column=1,
            max_column=3,
        )
    ] == ["C1", "A2", "B3"]


def test_cross_sheet_rule_builds_each_sparse_index_once_per_scan(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Source Data'!B2"
    dashboard["B2"] = "='Source Data'!B3"
    calls: list[str] = []
    original = WorksheetContentIndex.from_worksheet

    def counted(cls: type[WorksheetContentIndex], worksheet: Worksheet) -> WorksheetContentIndex:
        calls.append(worksheet.title)
        return original(worksheet)

    monkeypatch.setattr(WorksheetContentIndex, "from_worksheet", classmethod(counted))

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert scan.findings == []
    assert calls == ["Source Data", "Dashboard"]


def test_cross_sheet_reference_reports_blank_inside_dense_table(tmp_path: Path) -> None:
    workbook = _cross_sheet_book(blank_row=4)
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Source Data'!B4"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert len(scan.findings) == 1
    assert scan.findings[0].evidence.details["reason"] == "blank_inside_dense_table"
    assert scan.findings[0].severity == Severity.WARNING


def test_cross_sheet_reference_accepts_populated_target(tmp_path: Path) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Source Data'!B4"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert scan.findings == []


def test_cross_sheet_reference_does_not_overstate_optional_edge_blank(tmp_path: Path) -> None:
    workbook = _cross_sheet_book(blank_row=8)
    dashboard = workbook["Dashboard"]
    dashboard["B1"] = "='Source Data'!B8"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert scan.findings == []


def test_cross_sheet_reference_reports_obviously_truncated_nonaggregate_range(
    tmp_path: Path,
) -> None:
    workbook = _cross_sheet_book()
    dashboard = workbook["Dashboard"]
    dashboard["A1"] = "Total row count"
    dashboard["B1"] = "=ROWS('Source Data'!B2:B5)"

    scan = _scan(tmp_path, workbook, CrossSheetReferenceValidityRule())

    assert len(scan.findings) == 1
    finding = scan.findings[0]
    assert finding.evidence.details["reason"] == "truncated_cross_sheet_range"
    assert finding.severity == Severity.WARNING
    assert finding.evidence.details["excluded_cells"] == ["B6", "B7", "B8"]


def test_notes_column_reports_isolated_formula(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["ID", "Notes"])
    for row in range(2, 9):
        worksheet.append([row - 1, f"Reviewed note {row}"])
    worksheet["B5"] = "=SUM(A2:A4)"

    scan = _scan(tmp_path, workbook, FormulaInNotesColumnRule())

    assert len(scan.findings) == 1
    assert scan.findings[0].location == "B5"
    assert scan.findings[0].evidence.details["header"] == "Notes"
    assert scan.patches == []


def test_notes_rule_accepts_formula_in_amount_column(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["ID", "Amount"])
    for row in range(2, 9):
        worksheet.append([row - 1, f"=A{row}*10"])

    scan = _scan(tmp_path, workbook, FormulaInNotesColumnRule())

    assert scan.findings == []


def test_notes_rule_does_not_match_noted_date_header(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["ID", "Noted Date"])
    for row in range(2, 9):
        worksheet.append([row - 1, f"2026-08-{row:02d}"])
    worksheet["B5"] = "=TODAY()"

    scan = _scan(tmp_path, workbook, FormulaInNotesColumnRule())

    assert scan.findings == []


@pytest.mark.parametrize(
    ("formula", "kind"),
    [
        ("=A1/A1", "self_division"),
        ("=SUM(A1:A3)-SUM(A1:A3)", "self_subtraction"),
        ("=A1*0", "multiply_by_zero"),
    ],
)
def test_degenerate_formula_rule_reports_exact_top_level_forms(
    tmp_path: Path,
    formula: str,
    kind: str,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = 10
    worksheet["A2"] = 5
    worksheet["B1"] = formula

    scan = _scan(tmp_path, workbook, DegenerateFormulaRule(), name=f"{kind}.xlsx")

    assert len(scan.findings) == 1
    assert scan.findings[0].evidence.details["kind"] == kind
    assert scan.findings[0].patch_ids == []
    assert scan.patches == []


@pytest.mark.parametrize("formula", ["=A1/A2", "=A1-A2", "=A1*1"])
def test_degenerate_formula_rule_accepts_non_degenerate_forms(
    tmp_path: Path,
    formula: str,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = 10
    worksheet["A2"] = 5
    worksheet["B1"] = formula

    scan = _scan(tmp_path, workbook, DegenerateFormulaRule())

    assert scan.findings == []


def test_degenerate_formula_rule_ignores_guarded_inner_identity(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = 10
    worksheet["B1"] = "=IF(A1=0,0,A1/A1)"

    scan = _scan(tmp_path, workbook, DegenerateFormulaRule())

    assert scan.findings == []


def test_format_role_rule_reports_rate_currency_and_amount_percentage(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "Conversion Rate"
    worksheet["B1"] = "=C1/D1"
    worksheet["B1"].number_format = "$#,##0.00"
    worksheet["C1"] = 50
    worksheet["D1"] = 100
    worksheet["A2"] = "Revenue"
    worksheet["B2"] = "=C2*D2"
    worksheet["B2"].number_format = "0.0%"
    worksheet["C2"] = 5
    worksheet["D2"] = 100

    scan = _scan(tmp_path, workbook, FormulaFormatRoleMismatchRule())

    assert [finding.location for finding in scan.findings] == ["B1", "B2"]
    assert all(finding.patch_ids == [] for finding in scan.findings)
    assert scan.patches == []


def test_format_role_rule_accepts_compatible_formats(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "Conversion Rate"
    worksheet["B1"] = "=C1/D1"
    worksheet["B1"].number_format = "0.0%"
    worksheet["C1"] = 50
    worksheet["D1"] = 100
    worksheet["A2"] = "Revenue"
    worksheet["B2"] = "=C2*D2"
    worksheet["B2"].number_format = "$#,##0.00"
    worksheet["C2"] = 5
    worksheet["D2"] = 100

    scan = _scan(tmp_path, workbook, FormulaFormatRoleMismatchRule())

    assert scan.findings == []


@pytest.mark.parametrize("label", ["Hourly Rate", "Exchange Rate", "汇率"])
def test_format_role_rule_skips_ambiguous_currency_rates(tmp_path: Path, label: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = label
    worksheet["B1"] = "=C1/D1"
    worksheet["B1"].number_format = "$#,##0.00"
    worksheet["C1"] = 50
    worksheet["D1"] = 100

    scan = _scan(tmp_path, workbook, FormulaFormatRoleMismatchRule())

    assert scan.findings == []
