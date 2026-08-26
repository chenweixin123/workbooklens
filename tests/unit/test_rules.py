from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree
from openpyxl import Workbook
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side

import workbooklens.rules.builtin as builtin_rules
from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.models import PatchKind, PatchRisk
from workbooklens.rules.builtin import BUILTIN_RULES, TextFormulaInDataRegionRule
from workbooklens.rules.formula_semantics import AggregateRangeCoverageRule
from workbooklens.rules.registry import RuleRegistry
from workbooklens.scanner import ScanResult, scan_workbook


@pytest.fixture(scope="module")
def demo_scan(tmp_path_factory: pytest.TempPathFactory) -> ScanResult:
    directory = tmp_path_factory.mktemp("rules-demo")
    path = directory / "demo.xlsx"
    generate_demo_workbook(path)
    config = {"keys": [{"sheet": "Sales", "range": "A2:A22", "ignore_blank": True}]}
    return scan_workbook(path, config=config)


def _set_formula_cached_error(path: Path, coordinate: str, error: str) -> None:
    temporary = path.with_name(f"{path.stem}-cached-error.xlsx")
    with zipfile.ZipFile(path, "r") as source:
        infos = source.infolist()
        contents = {info.filename: source.read(info.filename) for info in infos}
    part = "xl/worksheets/sheet1.xml"
    root = etree.fromstring(contents[part])
    target = next(
        cell
        for cell in root.iter()
        if etree.QName(cell).localname == "c" and cell.get("r") == coordinate
    )
    target.set("t", "e")
    value = next(
        (child for child in target if etree.QName(child).localname == "v"),
        None,
    )
    if value is None:
        value = etree.SubElement(target, f"{{{etree.QName(target).namespace}}}v")
    value.text = error
    contents[part] = etree.tostring(root, xml_declaration=False, encoding="utf-8")
    with zipfile.ZipFile(temporary, "w") as output:
        for info in infos:
            output.writestr(info, contents[info.filename])
    temporary.replace(path)


def test_demo_exercises_expected_rules_and_generated_patch_kinds(demo_scan: ScanResult) -> None:
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
        "WL016_TEXT_DISPLAY_RISK",
        "WL041_NUMBER_FORMAT_ROLE_CONFLICT",
        "WL042_AGGREGATE_RANGE_COVERAGE",
        "WL046_FORMULA_FORMAT_ROLE_MISMATCH",
    }
    assert {patch.kind for patch in demo_scan.patches} == {
        PatchKind.SET_FORMULA,
        PatchKind.SET_NUMERIC,
        PatchKind.COPY_STYLE,
        PatchKind.CREATE_FORMULA,
        PatchKind.SET_ROW_HEIGHT,
        PatchKind.SET_WRAP_TEXT,
    }
    assert all(
        patch.safe_only_eligible for patch in demo_scan.patches if patch.risk == PatchRisk.SAFE
    )
    assert all(
        not patch.safe and patch.risk == PatchRisk.LAYOUT_REVIEW
        for patch in demo_scan.patches
        if patch.kind in {PatchKind.SET_ROW_HEIGHT, PatchKind.SET_WRAP_TEXT}
    )


def test_formula_pattern_outlier_rule_and_recalculated_proposal(tmp_path: Path) -> None:
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
    assert findings[0].evidence.expected is not None
    assert "aggregate" not in findings[0].explanation.lower()
    patches = [patch for patch in scan.patches if patch.cell == "K2"]
    assert len(patches) == 1
    assert patches[0].kind == PatchKind.SET_FORMULA
    assert patches[0].after == "=K1*2"
    assert patches[0].risk == PatchRisk.FORMULA_DERIVED
    assert not patches[0].safe
    assert patches[0].atomic_group is not None
    assert patches[0].derivation.strategy == "bidirectional_r1c1_consensus"
    assert len(patches[0].derivation.sources) == 2
    assert patches[0].derivation.requires_recalculation


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(B2:T2)",
        "=SUBTOTAL(9,B2:T2)",
        "=AGGREGATE(9,5,B2:T2)",
        "=SUM(B2:T2)+0",
        "=ROUND(SUM(B2:T2),2)",
        "=SUM(B2:T2,N(0))",
        "=SUM(B2:T2)+IFERROR(0,0)",
        "=AVERAGE(B2:T2)",
        "=MIN(B2:T2)",
        "=MAX(B2:T2)",
        "=COUNT(B2:T2)",
        "=COUNTA(B2:T2)",
        "=MEDIAN(B2:T2)",
    ],
)
def test_boundary_aggregate_is_not_treated_as_formula_outlier(tmp_path: Path, formula: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for column in range(2, 21):
        worksheet.cell(1, column, column)
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["U2"] = formula
    path = tmp_path / "boundary-total.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "U2"
        for finding in scan.findings
    )
    assert not any(patch.cell == "U2" for patch in scan.patches)


def test_complete_column_totals_are_not_wl002_but_truncated_totals_remain_wl042(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    budget = workbook.active
    assert budget is not None
    budget.title = "Budget"
    budget["A1"] = "Budget title"
    budget.append([])
    budget.append([])
    budget.append(
        ["Project", "Name", "Department", "Owner", "Budget", "Spent", "Remaining", "Rate"]
    )
    for row in range(5, 25):
        budget.append(
            [
                f"P-{row:03d}",
                f"Project {row}",
                "Ops",
                f"Owner {row}",
                row * 100,
                row * 40,
                f"=E{row}-F{row}",
                f"=IF(E{row}=0,0,F{row}/E{row})",
            ]
        )
    budget.append([])
    budget.append(
        [
            "Total",
            None,
            None,
            None,
            "=SUM(E5:E23)",
            "=SUM(F5:F24)",
            "=SUM(G5:G24)",
            "=AVERAGE(H5:H24)",
        ]
    )

    expenses = workbook.create_sheet("Expenses")
    expenses["A1"] = "Expense title"
    expenses.append([])
    expenses.append([])
    expenses.append(
        ["Claim", "Date", "Project", "Owner", "Category", "Amount", "Rate", "Tax", "Gross"]
    )
    for row in range(5, 35):
        expenses.append(
            [
                f"C-{row:03d}",
                f"2026-08-{((row - 5) % 28) + 1:02d}",
                f"P-{row:03d}",
                f"Owner {row}",
                "Travel",
                row * 10,
                0.06,
                f"=F{row}*G{row}",
                f"=F{row}+H{row}",
            ]
        )
        expenses.cell(row, 7).number_format = "0%"
    expenses.append([])
    expenses.append(
        [
            "Total",
            None,
            None,
            None,
            None,
            "=SUM(F5:F33)",
            None,
            "=SUM(H5:H34)",
            "=SUM(I5:I34)",
        ]
    )

    hours = workbook.create_sheet("Hours")
    hours["A1"] = "Hours title"
    hours.append([])
    hours.append([])
    hours.append(["Date", "Project", "Employee", "Hours", "Rate", "Base", "Factor", "Total"])
    for row in range(5, 33):
        hours.append(
            [
                f"2026-08-{((row - 5) % 28) + 1:02d}",
                f"P-{row:03d}",
                f"Employee {row}",
                8,
                100,
                f"=D{row}*E{row}",
                1.5,
                f"=F{row}*G{row}",
            ]
        )
    hours.append([])
    hours.append(
        [
            "Total",
            None,
            None,
            "=SUM(D5:D31)",
            None,
            "=SUM(F5:F32)",
            None,
            "=SUM(H5:H32)",
        ]
    )

    path = tmp_path / "complete-and-truncated-totals.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(
        path,
        registry=RuleRegistry(
            (
                builtin_rules.FormulaPatternOutlierRule(),
                builtin_rules.BlankInFormulaBandRule(),
                AggregateRangeCoverageRule(),
            )
        ),
    )

    wl002 = {
        (finding.sheet, finding.location)
        for finding in scan.findings
        if finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER"
    }
    assert (
        not {
            ("Budget", "G26"),
            ("Expenses", "H36"),
            ("Expenses", "I36"),
        }
        & wl002
    )
    assert ("Budget", "H26") in wl002

    wl042 = {
        (finding.sheet, finding.location)
        for finding in scan.findings
        if finding.rule_id == "WL042_AGGREGATE_RANGE_COVERAGE"
    }
    assert {
        ("Budget", "E26"),
        ("Expenses", "F36"),
        ("Hours", "D34"),
    } <= wl042
    assert not any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND"
        and finding.sheet == "Expenses"
        and finding.location == "G36"
        for finding in scan.findings
    )


def _scan_horizontal_summary_gap(
    tmp_path: Path,
    *,
    header: str,
    values: list[float],
    number_format: str,
    name: str,
) -> ScanResult:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Summary"
    worksheet.append(["Item", "Amount A", header, "Amount B", "Amount C"])
    for row, value in enumerate(values, start=2):
        worksheet.append([f"R{row:03d}", row * 10, value, row * 20, row * 30])
        worksheet.cell(row, 3).number_format = number_format
    worksheet.append([])
    summary_row = len(values) + 3
    last_body_row = len(values) + 1
    worksheet.append(
        [
            "Total",
            f"=SUM(B2:B{last_body_row})",
            None,
            f"=SUM(D2:D{last_body_row})",
            f"=SUM(E2:E{last_body_row})",
        ]
    )
    path = tmp_path / name
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(
        path,
        registry=RuleRegistry((builtin_rules.BlankInFormulaBandRule(),)),
    )
    assert summary_row == len(values) + 3
    return scan


@pytest.mark.parametrize("header", ["Tax Rate", "Completion Ratio", "税率", "完成率"])
def test_wl003_skips_intentional_rate_gap_in_additive_summary_row(
    tmp_path: Path,
    header: str,
) -> None:
    scan = _scan_horizontal_summary_gap(
        tmp_path,
        header=header,
        values=[0.01, 0.03, 0.06, 0.13, 0.03, 0.06, 0.13, 0.01, 0.03, 0.06],
        number_format="0%",
        name=f"intentional-rate-summary-{len(list(tmp_path.iterdir()))}.xlsx",
    )

    assert not any(finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" for finding in scan.findings)


@pytest.mark.parametrize(
    ("header", "number_format"),
    [
        ("Tax Rate", "General"),
        ("数量", "0%"),
    ],
)
def test_wl003_does_not_skip_summary_gap_with_only_one_rate_signal(
    tmp_path: Path,
    header: str,
    number_format: str,
) -> None:
    scan = _scan_horizontal_summary_gap(
        tmp_path,
        header=header,
        values=[0.01, 0.03, 0.06, 0.13, 0.03, 0.06, 0.13, 0.01, 0.03, 0.06],
        number_format=number_format,
        name=f"weak-rate-summary-{len(list(tmp_path.iterdir()))}.xlsx",
    )

    assert any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "C13"
        for finding in scan.findings
    )


@pytest.mark.parametrize(
    ("header", "values", "number_format"),
    [
        ("Amount", [float(value * 100) for value in range(1, 11)], "$#,##0.00"),
        ("数量", [float(value) for value in range(1, 11)], "0"),
    ],
)
def test_wl003_keeps_missing_additive_summary_for_amount_and_quantity(
    tmp_path: Path,
    header: str,
    values: list[float],
    number_format: str,
) -> None:
    scan = _scan_horizontal_summary_gap(
        tmp_path,
        header=header,
        values=values,
        number_format=number_format,
        name=f"additive-summary-{len(list(tmp_path.iterdir()))}.xlsx",
    )

    assert any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "C13"
        for finding in scan.findings
    )


def test_wl003_keeps_non_summary_horizontal_formula_gap(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["B1"] = 2
    worksheet["C1"] = 3
    worksheet["D1"] = 4
    worksheet["E1"] = 5
    worksheet["B2"] = "=B1*2"
    worksheet["C2"] = None
    worksheet["D2"] = "=D1*2"
    worksheet["E2"] = "=E1*2"
    path = tmp_path / "non-summary-horizontal-gap.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(
        path,
        registry=RuleRegistry((builtin_rules.BlankInFormulaBandRule(),)),
    )
    assert any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "C2"
        for finding in scan.findings
    )


def test_multiple_isolated_formula_outliers_are_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for column in range(2, 22):
        worksheet.cell(1, column, column)
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["K2"] = "=K1*3"
    worksheet["Q2"] = "=Q1+7"
    path = tmp_path / "multiple-outliers.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = {
        finding.location
        for finding in scan.findings
        if finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER"
    }
    assert findings == {"K2", "Q2"}
    assert not any(patch.cell in findings for patch in scan.patches)


def test_data_region_formula_consensus_handles_multiple_anomaly_types(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 102):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B20"] = 999
    worksheet["B30"] = "=C30*3"
    worksheet["B40"] = 888
    worksheet["B50"] = "'=C50*2"
    worksheet["B60"] = None
    worksheet["B70"] = "=C70+7"
    worksheet["B80"] = None
    path = tmp_path / "multi-anomaly-column.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations = {
        rule_id: {finding.location for finding in scan.findings if finding.rule_id == rule_id}
        for rule_id in {
            "WL002_FORMULA_PATTERN_OUTLIER",
            "WL003_BLANK_IN_FORMULA_BAND",
            "WL004_HARDCODED_VALUE_IN_FORMULA_BAND",
            "WL028_TEXT_FORMULA",
        }
    }
    assert locations["WL002_FORMULA_PATTERN_OUTLIER"] == {"B30", "B70"}
    assert locations["WL003_BLANK_IN_FORMULA_BAND"] == {"B60", "B80"}
    assert locations["WL004_HARDCODED_VALUE_IN_FORMULA_BAND"] == {"B20", "B40"}
    assert locations["WL028_TEXT_FORMULA"] == {"B50"}

    formula_patches = {
        patch.cell
        for patch in scan.patches
        if patch.kind in {PatchKind.SET_FORMULA, PatchKind.CREATE_FORMULA}
    }
    assert formula_patches == {"B20", "B30", "B40", "B50", "B70"}
    assert all(
        patch.risk == PatchRisk.FORMULA_DERIVED
        for patch in scan.patches
        if patch.cell in formula_patches
    )


def test_two_anomalies_apply_only_the_uniquely_derived_nonblank_formula(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 102):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B20"] = 999
    worksheet["B40"] = None
    path = tmp_path / "two-anomalies-strong-consensus.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND" and finding.location == "B20"
        for finding in scan.findings
    )
    assert any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "B40"
        for finding in scan.findings
    )
    formula_patches = [
        patch
        for patch in scan.patches
        if patch.cell in {"B20", "B40"}
        and patch.kind in {PatchKind.SET_FORMULA, PatchKind.CREATE_FORMULA}
    ]
    assert [(patch.cell, patch.after) for patch in formula_patches] == [("B20", "=C20*2")]
    assert formula_patches[0].risk == PatchRisk.FORMULA_DERIVED


def test_single_data_region_formula_anomaly_uses_formula_derived_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 52):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B20"] = 999
    path = tmp_path / "single-data-region-anomaly.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    patches = [
        patch
        for patch in scan.patches
        if patch.cell == "B20" and patch.kind == PatchKind.SET_FORMULA
    ]
    assert len(patches) == 1
    assert patches[0].after == "=C20*2"
    assert patches[0].risk == PatchRisk.FORMULA_DERIVED
    assert float(patches[0].confidence) == 0.99
    assert not patches[0].safe_only_eligible


def test_clustered_leading_formula_anomalies_use_template_and_table_boundary(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Input A", "Input B", "Calculated"])
    for row in range(2, 32):
        worksheet.append([f"R{row}", row, row + 1, f"=B{row}*C{row}"])
    worksheet["D2"] = "=B2+C2"
    worksheet["D3"] = "=B4*C3"
    worksheet["D4"] = "=B4-C4"
    path = tmp_path / "clustered-leading-formula-anomalies.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    patches = {
        patch.cell: patch
        for patch in scan.patches
        if patch.kind == PatchKind.SET_FORMULA and patch.cell in {"D2", "D3", "D4"}
    }
    assert {cell: patch.after for cell, patch in patches.items()} == {
        "D2": "=B2*C2",
        "D3": "=B3*C3",
        "D4": "=B4*C4",
    }
    assert all(patch.risk == PatchRisk.FORMULA_DERIVED for patch in patches.values())
    assert all(float(patch.confidence) == 0.99 for patch in patches.values())
    assert all(
        patch.derivation.strategy == "r1c1_template_and_table_boundary"
        for patch in patches.values()
    )
    assert all(len(patch.derivation.sources) == 2 for patch in patches.values())


def test_interior_aggregate_is_reported_but_real_total_is_not(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Quantity", "Price", "Amount"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row, row + 0.5, f"=B{row}*C{row}"])
    worksheet["D10"] = "=SUM(B10:C10)"
    worksheet.append(["Total", None, None, "=SUM(D2:D21)"])
    path = tmp_path / "interior-aggregate.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "D10"
        for finding in scan.findings
    )
    assert not any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "D22"
        for finding in scan.findings
    )
    assert not any(patch.cell in {"D10", "D22"} for patch in scan.patches)


def test_truncated_aggregate_outlier_requests_manual_range_and_label_review(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Quantity", "Price", "Amount"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row, row + 0.5, f"=B{row}*C{row}"])
    worksheet.append(["TOTAL?", None, None, "=SUM(D2:D18)"])
    path = tmp_path / "truncated-aggregate.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "D22"
    )
    assert finding.evidence.expected == {
        "review": "aggregate_range_and_label",
        "detail_formula_replacement_inferred": False,
    }
    assert "aggregate formula" in finding.explanation.lower()
    assert "aggregate range" in finding.expected.lower()
    assert "detail rows" in finding.suggested_action.lower()
    assert finding.patch_ids == []
    aggregate_patch = next(patch for patch in scan.patches if patch.cell == "D22")
    assert aggregate_patch.after == "=SUM(D2:D21)"
    assert aggregate_patch.risk == PatchRisk.FORMULA_DERIVED


def test_provable_formula_errors_are_reported_without_evaluating_branches(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Errors"
    worksheet["A1"] = "=NA()"
    worksheet["A2"] = "=1/0"
    worksheet["A3"] = "=1/-0.0"
    worksheet["A4"] = '=VALUE("abc")'
    worksheet["A5"] = "=#DIV/0!"
    worksheet["B1"] = "=IFERROR(1/0,0)"
    worksheet["B2"] = "=IF(FALSE,1/0,1)"
    worksheet["B3"] = '=IF(TRUE,1,VALUE("abc"))'
    worksheet["B4"] = "=1/C1"
    worksheet["B5"] = '=VALUE("123")'
    worksheet["B6"] = "=UNKNOWN(1/0)"
    worksheet["B7"] = '="#N/A"'
    worksheet["B8"] = "=IFERROR(NA(),0)"
    path = tmp_path / "provable-formula-errors.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = [
        finding for finding in scan.findings if finding.rule_id == "WL034_PROVABLE_FORMULA_ERROR"
    ]
    assert {finding.location for finding in findings} == {"A1", "A2", "A3", "A4", "A5"}
    assert {finding.evidence.details["error"] for finding in findings} == {
        "#N/A",
        "#DIV/0!",
        "#VALUE!",
    }
    assert not any(patch.cell in {"A1", "A2", "A3", "A4", "A5"} for patch in scan.patches)


def test_cached_formula_error_is_reported_as_stale_evidence_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = "=A2+1"
    worksheet["A2"] = 1
    path = tmp_path / "cached-formula-error.xlsx"
    workbook.save(path)
    workbook.close()
    _set_formula_cached_error(path, "A1", "#NUM!")

    scan = scan_workbook(path)
    findings = [
        finding for finding in scan.findings if finding.rule_id == "WL034_PROVABLE_FORMULA_ERROR"
    ]
    assert len(findings) == 1
    finding = findings[0]
    assert finding.location == "A1"
    assert finding.evidence.details == {
        "error": "#NUM!",
        "error_code_proven": True,
        "proof": "cached_formula_error",
        "cached_error": "#NUM!",
    }
    assert float(finding.confidence) == 0.9
    assert "may be stale" in finding.explanation
    assert not any(patch.cell == "A1" for patch in scan.patches)


def test_multiple_hardcodes_with_override_labels_or_styles_are_findings_only(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 102):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B20"] = 999
    worksheet["A20"] = "Exception row"
    worksheet["B40"] = 888
    worksheet["C40"].fill = PatternFill("solid", fgColor="FFFF00")
    path = tmp_path / "multiple-intentional-overrides.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations = {
        finding.location
        for finding in scan.findings
        if finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND"
    }
    assert {"B20", "B40"} <= locations
    assert not any(patch.cell in {"B20", "B40"} for patch in scan.patches)


@pytest.mark.parametrize("hidden_scope", ["rows", "column", "sheet"])
def test_multiple_hardcodes_in_hidden_targets_are_findings_only(
    tmp_path: Path,
    hidden_scope: str,
) -> None:
    workbook = Workbook()
    cover = workbook.active
    assert cover is not None
    cover.title = "Cover"
    worksheet = workbook.create_sheet("Data")
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 52):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B20"] = 999
    worksheet["B40"] = 888
    if hidden_scope == "rows":
        worksheet.row_dimensions[20].hidden = True
        worksheet.row_dimensions[40].hidden = True
    elif hidden_scope == "column":
        worksheet.column_dimensions["B"].hidden = True
    else:
        worksheet.sheet_state = "hidden"
    path = tmp_path / f"multiple-hidden-{hidden_scope}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations = {
        finding.location
        for finding in scan.findings
        if finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND" and finding.sheet == "Data"
    }
    assert {"B20", "B40"} <= locations
    assert not any(patch.sheet == "Data" and patch.cell in {"B20", "B40"} for patch in scan.patches)


def test_large_formula_column_reports_several_adjacent_outliers_below_eighty_percent(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 19):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B10"] = "=C9*2"
    worksheet["B11"] = "=C11*3"
    worksheet["B12"] = "=NA()"
    worksheet["B13"] = "=B13+1"
    path = tmp_path / "several-adjacent-formula-outliers.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations = {
        finding.location
        for finding in scan.findings
        if finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER"
    }
    assert locations == {"B10", "B11", "B12", "B13"}
    assert not any(patch.cell in locations for patch in scan.patches)


def test_blank_separator_before_total_is_never_auto_filled(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Sales"
    worksheet.append(["Record", "Calculated"])
    for row in range(2, 11):
        worksheet.append([f"R{row}", f'=A{row}&"-ok"'])
    worksheet["A12"] = "Total"
    worksheet["B12"] = "=COUNTA(B2:B10)"
    path = tmp_path / "formula-separator-total.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "B11"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B11" for patch in scan.patches)


def test_static_formula_cycles_are_reported_as_sccs_without_patches(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Cycle"
    worksheet["A1"] = "=A1+1"
    worksheet["B1"] = "=C1+1"
    worksheet["C1"] = "=B1+1"
    worksheet["D1"] = "='Other Sheet'!A1"
    worksheet["E1"] = "=1+1"
    other = workbook.create_sheet("Other Sheet")
    other["A1"] = "='Cycle'!D1"
    path = tmp_path / "formula-cycles.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = [
        finding for finding in scan.findings if finding.rule_id == "WL029_CIRCULAR_REFERENCE"
    ]
    assert len(findings) == 3
    assert sorted(finding.evidence.details["component_size"] for finding in findings) == [1, 2, 2]
    peer_sets = {frozenset(finding.evidence.peers) for finding in findings}
    assert frozenset({"Cycle!A1"}) in peer_sets
    assert frozenset({"Cycle!B1", "Cycle!C1"}) in peer_sets
    assert frozenset({"Cycle!D1", "Other Sheet!A1"}) in peer_sets
    circular_cells = {
        peer.rsplit("!", maxsplit=1)[-1] for finding in findings for peer in finding.evidence.peers
    }
    assert not any(patch.cell in circular_cells for patch in scan.patches)


def test_circular_dependency_graph_is_computed_once_per_scan(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 12):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B6"] = "=B6+1"
    path = tmp_path / "dependency-cache.xlsx"
    workbook.save(path)
    workbook.close()

    calls = 0
    original = builtin_rules._static_formula_dependency_graph

    def counted(workbook: Workbook) -> dict[tuple[str, str], set[tuple[str, str]]]:
        nonlocal calls
        calls += 1
        return original(workbook)

    monkeypatch.setattr(builtin_rules, "_static_formula_dependency_graph", counted)
    scan = scan_workbook(path)

    assert calls == 1
    assert any(
        finding.rule_id == "WL029_CIRCULAR_REFERENCE" and finding.location == "B6"
        for finding in scan.findings
    )


def test_large_formula_ranges_do_not_expand_and_still_detect_self_reference(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Large"
    worksheet["A1"] = "=SUM(A:A)"
    for row in range(1, 40):
        worksheet.cell(row, 2, f"=SUM(A:A)+{row}")
    path = tmp_path / "large-formula-ranges.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = [
        finding for finding in scan.findings if finding.rule_id == "WL029_CIRCULAR_REFERENCE"
    ]
    assert len(findings) == 1
    assert findings[0].location == "A1"
    assert findings[0].evidence.peers == ["Large!A1"]


def test_large_formula_range_links_real_formula_dependencies_without_expansion(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Cycle"
    worksheet["A1"] = "=SUM(B:B)"
    worksheet["B1"] = "=A1"
    path = tmp_path / "large-range-formula-cycle.xlsx"
    workbook.save(path)
    workbook.close()

    findings = [
        finding
        for finding in scan_workbook(path).findings
        if finding.rule_id == "WL029_CIRCULAR_REFERENCE"
    ]

    assert len(findings) == 1
    assert findings[0].evidence.details["component_size"] == 2
    assert findings[0].evidence.peers == ["Cycle!A1", "Cycle!B1"]


def test_circular_dependencies_resolve_sheet_names_case_insensitively(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    data = workbook.active
    assert data is not None
    data.title = "Data"
    summary = workbook.create_sheet("Summary")
    data["A1"] = "=summary!A1"
    summary["A1"] = "=DATA!A1"
    path = tmp_path / "case-insensitive-formula-cycle.xlsx"
    workbook.save(path)
    workbook.close()

    findings = [
        finding
        for finding in scan_workbook(path).findings
        if finding.rule_id == "WL029_CIRCULAR_REFERENCE"
    ]

    assert len(findings) == 1
    assert findings[0].evidence.peers == ["Data!A1", "Summary!A1"]


def test_dense_formula_range_dependencies_are_bounded_and_cached(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        builtin_rules,
        "MAX_STATIC_CIRCULAR_DEPENDENCIES_PER_REFERENCE",
        16,
    )
    monkeypatch.setattr(
        builtin_rules,
        "MAX_STATIC_CIRCULAR_GRAPH_NONSELF_EDGES",
        100,
    )
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Dense"
    for row in range(1, 1001):
        worksheet.cell(row, 1, "=SUM(A:A)")

    observed_limits: list[int | None] = []
    original_iterator = builtin_rules.WorksheetContentIndex.iter_nonblank_rectangle

    def recorded_rectangle(self, **kwargs):
        observed_limits.append(kwargs.get("limit"))
        yield from original_iterator(self, **kwargs)

    monkeypatch.setattr(
        builtin_rules.WorksheetContentIndex,
        "iter_nonblank_rectangle",
        recorded_rectangle,
    )

    graph = builtin_rules._static_formula_dependency_graph(workbook)

    assert len(graph) == 1000
    assert all(dependencies == {node} for node, dependencies in graph.items())
    assert observed_limits == [17]


def test_circular_dependency_graph_skips_nonself_batch_when_budget_is_insufficient(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(
        builtin_rules,
        "MAX_STATIC_CIRCULAR_DEPENDENCIES_PER_REFERENCE",
        10,
    )
    monkeypatch.setattr(
        builtin_rules,
        "MAX_STATIC_CIRCULAR_GRAPH_NONSELF_EDGES",
        1,
    )
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Budget"
    worksheet["A1"] = "=SUM(B1:B2)"
    worksheet["A2"] = "=A2"
    worksheet["B1"] = "=1"
    worksheet["B2"] = "=1"

    graph = builtin_rules._static_formula_dependency_graph(workbook)

    assert graph[("Budget", "A1")] == set()
    assert graph[("Budget", "A2")] == {("Budget", "A2")}
    assert (
        sum(
            len({dependency for dependency in dependencies if dependency != node})
            for node, dependencies in graph.items()
        )
        == 0
    )


def test_merged_formula_outlier_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 22):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet["B10"] = "=A10*3"
    worksheet.merge_cells("B10:C10")
    path = tmp_path / "merged-formula-outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


@pytest.mark.parametrize(
    "label",
    [
        "Subtotal",
        "Sub-total",
        "Grand-total",
        "Average",
        "Mean",
        "Minimum",
        "Maximum",
        "Summary",
        "小计",
        "小计金额",
        "合计金额",
        "汇总金额",
        "平均金额",
        "最大金额",
        "最小金额",
        "总金额",
        "累计金额",
        "净额",
        "期末余额",
    ],
)
def test_summary_row_formula_outlier_is_findings_only(tmp_path: Path, label: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A2"] = label
    for column in range(2, 22):
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["K2"] = "=K1*3"
    path = tmp_path / f"summary-formula-outlier-{len(label)}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "K2"
        for finding in scan.findings
    )
    assert not any(patch.cell == "K2" for patch in scan.patches)


def test_hidden_formula_outlier_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 22):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet["B10"] = "=A10*3"
    worksheet.row_dimensions[10].hidden = True
    path = tmp_path / "hidden-formula-outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_formula_band_boundary_never_gets_automatic_replacement(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Band"
    for column in range(2, 21):
        worksheet.cell(1, column, column)
        coordinate = worksheet.cell(2, column).coordinate
        source = worksheet.cell(1, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet["U2"] = "=AVERAGE(Band!B2:T2)"
    path = tmp_path / "explicit-sheet-boundary-summary.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(patch.cell == "U2" for patch in scan.patches)


def test_finding_identity_is_stable_when_evidence_content_changes(tmp_path: Path) -> None:
    first = Workbook()
    first_sheet = first.active
    assert first_sheet is not None
    first_sheet["A1"] = "=#REF!+1"
    first_path = tmp_path / "first.xlsx"
    first.save(first_path)
    first.close()

    second = Workbook()
    second_sheet = second.active
    assert second_sheet is not None
    second_sheet["A1"] = "=#REF!+2"
    second_path = tmp_path / "second.xlsx"
    second.save(second_path)
    second.close()

    first_finding = next(
        finding
        for finding in scan_workbook(first_path).findings
        if finding.rule_id == "WL001_BROKEN_REFERENCE"
    )
    second_finding = next(
        finding
        for finding in scan_workbook(second_path).findings
        if finding.rule_id == "WL001_BROKEN_REFERENCE"
    )
    assert first_finding.id == second_finding.id
    assert first_finding.content_fingerprint != second_finding.content_fingerprint


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


def test_volatile_peer_template_never_generates_formula_derived_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", f"=NOW()+C{row}", row])
    worksheet["B10"] = "=C10*3"
    path = tmp_path / "volatile-peer-template.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_formula_candidate_that_would_create_cycle_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B10"] = "=1"
    worksheet["C10"] = "=B10"
    path = tmp_path / "candidate-cycle.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert any(
        finding.rule_id == "WL002_FORMULA_PATTERN_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_volatile_text_formula_is_never_converted_automatically(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Calculated", "Input"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", f"=C{row}*2", row])
    worksheet["B10"] = "'=NOW()"
    path = tmp_path / "volatile-text-formula.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert any(
        finding.rule_id == "WL028_TEXT_FORMULA" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def _strict_text_formula_book(*, rows: int = 30) -> Workbook:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.title = "Data"
    worksheet.append(["Record", "Input A", "Input B", "Calculated"])
    for row in range(2, rows + 2):
        worksheet.append([f"R{row}", row, row + 1, f"=B{row}*C{row}"])
    worksheet["A2"] = "Special record"
    worksheet["A2"].fill = PatternFill("solid", fgColor="FFF2CC")
    worksheet["D2"] = "'=B2*C2"
    return workbook


def test_exact_same_style_text_formula_uses_unique_peer_and_region_evidence(
    tmp_path: Path,
) -> None:
    workbook = _strict_text_formula_book()
    path = tmp_path / "strict-text-formula.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path, registry=RuleRegistry([TextFormulaInDataRegionRule()]))

    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL028_TEXT_FORMULA" and finding.location == "D2"
    )
    patch = next(patch for patch in scan.patches if patch.cell == "D2")
    assert finding.patch_ids == [patch.id]
    assert patch.after == "=B2*C2"
    assert patch.risk == PatchRisk.FORMULA_DERIVED
    assert patch.derivation.strategy == "r1c1_template_and_table_boundary"
    assert any(
        source.startswith("stored_formula_text_exact_match:") for source in patch.derivation.sources
    )
    assert "target_and_peers_are_visible_and_share_style" in patch.derivation.invariants
    assert patch.derivation.requires_recalculation


@pytest.mark.parametrize("hidden_target", ["row", "column"])
def test_exact_formula_text_in_hidden_target_is_finding_only(
    tmp_path: Path,
    hidden_target: str,
) -> None:
    workbook = _strict_text_formula_book()
    worksheet = workbook["Data"]
    if hidden_target == "row":
        worksheet.row_dimensions[2].hidden = True
    else:
        worksheet.column_dimensions["D"].hidden = True
    path = tmp_path / f"strict-hidden-{hidden_target}-text-formula.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path, registry=RuleRegistry([TextFormulaInDataRegionRule()]))

    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL028_TEXT_FORMULA" and finding.location == "D2"
    )
    assert finding.patch_ids == []
    assert not any(patch.cell == "D2" for patch in scan.patches)


@pytest.mark.parametrize("blocked", ["different_formula", "different_style", "insufficient_peers"])
def test_text_formula_strict_evidence_gate_withholds_patch(
    tmp_path: Path,
    blocked: str,
) -> None:
    workbook = _strict_text_formula_book(rows=7 if blocked == "insufficient_peers" else 30)
    worksheet = workbook["Data"]
    if blocked == "different_formula":
        worksheet["D2"] = "'=B2+C2"
    elif blocked == "different_style":
        worksheet["D2"].fill = PatternFill("solid", fgColor="F4CCCC")
    path = tmp_path / f"strict-text-formula-{blocked}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path, registry=RuleRegistry([TextFormulaInDataRegionRule()]))

    assert any(
        finding.rule_id == "WL028_TEXT_FORMULA" and finding.location == "D2"
        for finding in scan.findings
    )
    assert not any(patch.cell == "D2" for patch in scan.patches)


def test_formula_rule_tokens_ignore_string_literals_but_keep_real_constructs(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = '="#REF! NOW() Table1[Amount] [Book.xlsx]S!A1"'
    worksheet["A2"] = "=#REF!+1"
    worksheet["A3"] = "=NOW()"
    worksheet["A4"] = "='[Book.xlsx]S'!A1"
    path = tmp_path / "formula-tokens.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations_by_rule = {
        rule_id: {finding.location for finding in scan.findings if finding.rule_id == rule_id}
        for rule_id in {
            "WL001_BROKEN_REFERENCE",
            "WL009_EXTERNAL_LINK",
            "WL010_VOLATILE_OR_FRAGILE_FUNCTION",
        }
    }
    assert locations_by_rule["WL001_BROKEN_REFERENCE"] == {"A2"}
    assert locations_by_rule["WL009_EXTERNAL_LINK"] == {"A4"}
    assert locations_by_rule["WL010_VOLATILE_OR_FRAGILE_FUNCTION"] == {"A3"}


def test_numeric_text_excludes_leading_zero_identifiers(demo_scan: ScanResult) -> None:
    locations = {
        finding.location
        for finding in demo_scan.findings
        if finding.rule_id == "WL006_NUMERIC_TEXT"
    }
    assert locations == {"B10"}


@pytest.mark.parametrize(
    "header",
    [
        "Customer ID",
        "SKU",
        "Account Number",
        "Postal Code",
        "Phone Number",
        "Mobile Number",
        "SSN",
        "ISBN",
        "手机号",
        "证件号码",
        "银行卡号",
        "客户编号",
    ],
)
def test_identifier_headers_suppress_numeric_text_patch(tmp_path: Path, header: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append([header, "Amount"])
    for row in range(2, 22):
        worksheet.append([10000 + row, row * 10])
    worksheet["A10"] = "12345"
    path = tmp_path / f"identifier-{len(header)}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
    )
    assert not finding.safe_patch_available
    assert not any(patch.cell == "A10" for patch in scan.patches)


def test_account_balance_still_allows_numeric_measure_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Account Balance", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    path = tmp_path / "account-balance.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        patch.kind == PatchKind.SET_NUMERIC and patch.cell == "A10" for patch in scan.patches
    )


def test_measure_patch_survives_plain_numeric_text_in_identifier_column(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Account ID", "Customer", "Credit Limit"])
    for row in range(2, 22):
        worksheet.append([10000 + row, f"Customer {row}", row * 1000])
    worksheet["A10"] = "12345"
    worksheet["C10"] = "12000"
    path = tmp_path / "identifier-and-measure-numeric-text.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    patches = {(patch.kind, patch.cell) for patch in scan.patches}
    assert (PatchKind.SET_NUMERIC, "A10") not in patches
    assert (PatchKind.SET_NUMERIC, "C10") in patches


def test_unknown_numeric_column_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Business Field", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    path = tmp_path / "unknown-numeric-column.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "A10" for patch in scan.patches)


def test_explicit_text_format_suppresses_numeric_text_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Measure", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    worksheet["A10"].number_format = "@"
    path = tmp_path / "explicit-text.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "A10" for patch in scan.patches)


def test_quote_prefix_suppresses_numeric_and_style_patches(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Measure", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    worksheet["A10"].quotePrefix = True
    path = tmp_path / "quote-prefix.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(
        patch.cell == "A10" and patch.kind in {PatchKind.SET_NUMERIC, PatchKind.COPY_STYLE}
        for patch in scan.patches
    )


def test_protected_worksheet_is_findings_only_for_numeric_text(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Amount", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    worksheet.protection.sheet = True
    path = tmp_path / "protected-numeric.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "A10" for patch in scan.patches)


def test_grouped_numeric_text_is_reported_without_auto_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Amount", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1,200"
    path = tmp_path / "grouped-numeric.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "A10" for patch in scan.patches)


def test_grouped_hidden_nonleading_column_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Context", "Amount", "Reviewer"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", f"C{row}", row * 100, "Chen"])
    worksheet["C10"] = "1200"
    worksheet.column_dimensions.group("B", "D", hidden=True)
    path = tmp_path / "grouped-hidden-nonleading-column.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "C10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "C10" for patch in scan.patches)
    snapshot_sheet = scan.snapshot.sheets[0]
    assert snapshot_sheet.hidden_columns == ["B", "C", "D"]
    assert snapshot_sheet.cells["C10"].column_hidden


@pytest.mark.parametrize("sheet_state", ["hidden", "veryHidden"])
def test_nonvisible_worksheet_is_findings_only_for_numeric_text(
    tmp_path: Path, sheet_state: str
) -> None:
    workbook = Workbook()
    cover = workbook.active
    assert cover is not None
    cover.title = "Cover"
    worksheet = workbook.create_sheet("Data")
    worksheet.append(["Amount", "Context"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"R{row}"])
    worksheet["A10"] = "1200"
    worksheet.sheet_state = sheet_state
    path = tmp_path / f"{sheet_state}-numeric.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL006_NUMERIC_TEXT"
        and finding.sheet == "Data"
        and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(patch.sheet == "Data" for patch in scan.patches)


@pytest.mark.parametrize("sheet_state", ["hidden", "veryHidden"])
def test_nonvisible_worksheet_is_findings_only_for_formula_and_style(
    tmp_path: Path, sheet_state: str
) -> None:
    workbook = Workbook()
    cover = workbook.active
    assert cover is not None
    cover.title = "Cover"
    worksheet = workbook.create_sheet("Data")
    worksheet.append(["Amount", "Calculated"])
    for row in range(2, 22):
        worksheet.append([row * 100, f"=A{row}*2"])
    worksheet["B10"] = 999
    worksheet["A11"].font = Font(bold=True)
    worksheet.sheet_state = sheet_state
    path = tmp_path / f"{sheet_state}-formula-style.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND"
        and finding.sheet == "Data"
        and finding.location == "B10"
        for finding in scan.findings
    )
    assert any(
        finding.rule_id == "WL007_STYLE_OUTLIER"
        and finding.sheet == "Data"
        and finding.location == "A11"
        for finding in scan.findings
    )
    assert not any(patch.sheet == "Data" for patch in scan.patches)


@pytest.mark.parametrize("sheet_state", ["hidden", "veryHidden"])
def test_nonvisible_worksheet_sum_boundary_is_findings_only(
    tmp_path: Path, sheet_state: str
) -> None:
    workbook = Workbook()
    cover = workbook.active
    assert cover is not None
    cover.title = "Cover"
    worksheet = workbook.create_sheet("Data")
    for row in range(2, 10):
        worksheet.cell(row, 1, f"R{row}")
        worksheet.cell(row, 2, row * 100)
    worksheet["A10"] = "Total"
    worksheet["B10"] = "=SUM(B2:B8)"
    worksheet.sheet_state = sheet_state
    path = tmp_path / f"{sheet_state}-sum-boundary.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY" and finding.sheet == "Data"
    )
    assert not finding.safe_patch_available
    assert not any(patch.sheet == "Data" for patch in scan.patches)


def test_protection_only_difference_is_not_a_visual_style_outlier(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Input", "Context"])
    for row in range(2, 22):
        worksheet.append([row, f"R{row}"])
    worksheet["A10"].protection = Protection(locked=False)
    worksheet.protection.sheet = True
    path = tmp_path / "unlocked-style-outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(
        patch.kind == PatchKind.COPY_STYLE and patch.cell == "A10" for patch in scan.patches
    )


def test_style_outlier_with_different_number_format_has_no_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Customer ID", "Context"])
    for row in range(2, 22):
        worksheet.append([10000 + row, f"R{row}"])
    worksheet["A10"].number_format = "000000"
    path = tmp_path / "identifier-number-format-outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "A10"
    )
    assert not finding.safe_patch_available
    assert not any(
        patch.kind == PatchKind.COPY_STYLE and patch.cell == "A10" for patch in scan.patches
    )


@pytest.mark.parametrize("missing_side", ["left", "right"])
def test_single_sided_shared_border_is_visually_equivalent(
    tmp_path: Path, missing_side: str
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Context", "Amount", "Context"])
    side = Side(style="thin", color="336699")
    grid = Border(left=side, right=side, top=side, bottom=side)
    for row in range(2, 22):
        worksheet.append([f"L{row}", row * 100, f"R{row}"])
        for cell in worksheet[row]:
            cell.border = grid
    target = worksheet["B10"]
    target.border = Border(
        left=None if missing_side == "left" else side,
        right=None if missing_side == "right" else side,
        top=side,
        bottom=side,
    )
    path = tmp_path / f"shared-border-{missing_side}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )


@pytest.mark.parametrize(
    ("target_row", "perimeter_side"),
    [(2, "top"), (21, "bottom")],
)
def test_detail_row_perimeter_border_is_not_a_style_outlier(
    tmp_path: Path,
    target_row: int,
    perimeter_side: str,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Context", "Amount"])
    side = Side(style="thin", color="336699")
    perimeter = Side(style="medium", color="336699")
    grid = Border(left=side, right=side, top=side, bottom=side)
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 100])
        for cell in worksheet[row]:
            cell.border = grid
    targets = worksheet[target_row]
    for target in targets:
        target.border = Border(
            left=side,
            right=side,
            top=perimeter if perimeter_side == "top" else side,
            bottom=perimeter if perimeter_side == "bottom" else side,
        )
    path = tmp_path / f"detail-{perimeter_side}-perimeter-border.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER"
        and finding.location in {target.coordinate for target in targets}
        for finding in scan.findings
    )


def test_single_first_detail_row_top_border_without_boundary_consensus_is_reported(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Context", "Amount", "Status"])
    side = Side(style="thin", color="336699")
    grid = Border(left=side, right=side, top=side, bottom=side)
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 100, "Open"])
        for cell in worksheet[row]:
            cell.border = grid
    worksheet["B2"].border = Border(
        left=side,
        right=side,
        top=Side(style="medium", color="FF0000"),
        bottom=side,
    )
    path = tmp_path / "single-first-detail-row-top-border.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "B2"
        for finding in scan.findings
    )
    assert not any(
        finding.rule_id == "WL017_BORDER_EDGE_INCONSISTENCY" and finding.location == "B2"
        for finding in scan.findings
    )


@pytest.mark.parametrize("component", ["font", "fill", "alignment", "number_format", "border"])
def test_material_visual_style_difference_remains_an_outlier(
    tmp_path: Path, component: str
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Context", "Amount"])
    side = Side(style="thin", color="336699")
    grid = Border(left=side, right=side, top=side, bottom=side)
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 100])
        for cell in worksheet[row]:
            cell.border = grid
    target = worksheet["B10"]
    if component == "font":
        target.font = Font(bold=True)
    elif component == "fill":
        target.fill = PatternFill("solid", fgColor="FFF2CC")
    elif component == "alignment":
        target.alignment = Alignment(horizontal="right")
    elif component == "number_format":
        target.number_format = "0.00"
    else:
        target.border = Border(right=side, top=side, bottom=side)
        worksheet["A10"].border = Border(left=side, top=side, bottom=side)
    path = tmp_path / f"material-style-{component}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "B10"
        for finding in scan.findings
    )


def test_last_detail_row_bottom_border_variant_is_not_a_style_outlier(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Month", "Sales"])
    side = Side(style="thin", color="D9E2F3")
    interior = Border(top=side, bottom=side)
    last_row = Border(top=side)
    for row in range(2, 14):
        worksheet.append([f"2026-{row - 1:02d}", row * 100])
        for cell in worksheet[row]:
            cell.border = last_row if row == 13 else interior
    path = tmp_path / "last-row-perimeter-style.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location in {"A13", "B13"}
        for finding in scan.findings
    )


def test_row_local_sum_does_not_hide_a_style_outlier(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "A", "B", "C", "Row total"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row, row + 1, row + 2, f"=SUM(B{row}:D{row})+$A$1"])
    worksheet["C10"].font = Font(bold=True)
    path = tmp_path / "row-local-sum-style-outlier.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "C10"
        for finding in scan.findings
    )


def test_unlabelled_cross_row_sum_still_marks_a_summary_row(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Amount"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 10])
    worksheet.append(["Result", "=SUM(B2:B21)"])
    worksheet["B22"].font = Font(bold=True)
    path = tmp_path / "unlabelled-summary-row.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)

    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "B22"
        for finding in scan.findings
    )


@pytest.mark.parametrize("error_type", [ValueError, IndexError, TokenizerError])
def test_malformed_aggregate_formula_is_handled_conservatively(
    monkeypatch: pytest.MonkeyPatch,
    error_type: type[Exception],
) -> None:
    def reject_formula(_formula: str):
        raise error_type("malformed aggregate formula")

    monkeypatch.setattr(builtin_rules, "Tokenizer", reject_formula)

    assert builtin_rules._aggregate_formula_spans_other_rows("=SUM(A1:A2)", 3)


def test_pivot_button_only_difference_is_not_a_visual_style_outlier(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Amount", "Context"])
    for row in range(2, 22):
        worksheet.append([row, f"R{row}"])
    worksheet["A10"].pivotButton = True
    path = tmp_path / "pivot-button-style.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location == "A10"
        for finding in scan.findings
    )
    assert not any(
        patch.kind == PatchKind.COPY_STYLE and patch.cell == "A10" for patch in scan.patches
    )


def test_multiple_style_outliers_are_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Amount", "Context"])
    for row in range(2, 24):
        worksheet.append([row, f"R{row}"])
    worksheet["A8"].font = Font(bold=True)
    worksheet["A18"].font = Font(italic=True)
    path = tmp_path / "multiple-style-outliers.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    locations = {
        finding.location for finding in scan.findings if finding.rule_id == "WL007_STYLE_OUTLIER"
    }
    assert locations == {"A8", "A18"}
    assert not any(
        patch.kind == PatchKind.COPY_STYLE and patch.cell in locations for patch in scan.patches
    )


@pytest.mark.parametrize("label", ["Grand Total", "Summary", "总金额", "累计金额", "期末余额"])
def test_summary_row_style_outliers_are_not_reported_or_auto_copied(
    tmp_path: Path, label: str
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Amount"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 10])
    worksheet.append([label, "=SUM(B2:B21)"])
    worksheet["A22"].font = Font(bold=True)
    worksheet["B22"].font = Font(bold=True)
    path = tmp_path / "summary-row-style.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert not any(
        finding.rule_id == "WL007_STYLE_OUTLIER" and finding.location in {"A22", "B22"}
        for finding in scan.findings
    )
    assert not any(
        patch.kind == PatchKind.COPY_STYLE and patch.cell in {"A22", "B22"}
        for patch in scan.patches
    )


@pytest.mark.parametrize("label", ["Total", "Summary", "总金额", "累计金额", "期末余额"])
def test_summary_row_numeric_text_is_findings_only(tmp_path: Path, label: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Amount"])
    for row in range(2, 22):
        worksheet.append([f"R{row}", row * 100])
    worksheet["A10"] = label
    worksheet["B10"] = "1200"
    path = tmp_path / f"summary-numeric-{len(label)}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL006_NUMERIC_TEXT" and finding.location == "B10"
    )
    assert not finding.safe_patch_available
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_secondary_row_semantic_override_blocks_all_candidate_patches(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Row type", "Amount", "Calculated"])
    for row in range(2, 22):
        worksheet.append([f"Item {row}", "Regular", row * 100, f"=C{row}*2"])
    worksheet["B10"] = "Manual override"
    worksheet["C10"] = "1200"
    worksheet["C10"].font = Font(bold=True)
    worksheet["D10"] = 999
    path = tmp_path / "secondary-row-semantic-override.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = {(finding.rule_id, finding.location) for finding in scan.findings}
    assert ("WL004_HARDCODED_VALUE_IN_FORMULA_BAND", "D10") in findings
    assert ("WL006_NUMERIC_TEXT", "C10") in findings
    assert ("WL007_STYLE_OUTLIER", "C10") in findings
    assert not any(patch.cell in {"C10", "D10"} for patch in scan.patches)


@pytest.mark.parametrize(
    ("fault", "rule_id"),
    [
        ("formula", "WL002_FORMULA_PATTERN_OUTLIER"),
        ("blank", "WL003_BLANK_IN_FORMULA_BAND"),
    ],
)
def test_secondary_row_semantic_override_blocks_formula_candidate_patches(
    tmp_path: Path, fault: str, rule_id: str
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Row type", "Input", "Calculated"])
    for row in range(2, 22):
        worksheet.append([f"Item {row}", "Regular", row * 100, f"=C{row}*2"])
    worksheet["B10"] = "调整项"
    worksheet["D10"] = "=C10*3" if fault == "formula" else None
    path = tmp_path / f"secondary-row-{fault}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == rule_id and finding.location == "D10"
    )
    assert not finding.safe_patch_available
    assert not any(patch.cell == "D10" for patch in scan.patches)


def test_visually_marked_blank_formula_gap_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Calculated"])
    for row in range(2, 22):
        worksheet.append([f"Item {row}", f'=A{row}&"-done"'])
    worksheet["B10"] = None
    worksheet["B10"].fill = PatternFill("solid", fgColor="FFF2CC")
    path = tmp_path / "highlighted-blank-formula-gap.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "B10"
    )
    assert not finding.safe_patch_available
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_summary_label_outside_inferred_region_suppresses_style_finding_and_patches(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["C1"] = "Item"
    worksheet["D1"] = "Amount"
    for row in range(2, 22):
        worksheet[f"C{row}"] = f"R{row}"
        worksheet[f"D{row}"] = row * 100
    worksheet["A10"] = "Summary"
    worksheet["D10"] = "1200"
    worksheet["D10"].font = Font(bold=True)
    path = tmp_path / "outside-region-summary.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = {(finding.rule_id, finding.location) for finding in scan.findings}
    assert ("WL006_NUMERIC_TEXT", "D10") in findings
    assert ("WL007_STYLE_OUTLIER", "D10") not in findings
    assert not any(patch.cell == "D10" for patch in scan.patches)


def test_intentionally_highlighted_row_blocks_content_and_style_patches(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet.append(["Item", "Amount", "Calculated"])
    for row in range(2, 22):
        worksheet.append([f"Item {row}", row * 100, f"=B{row}*2"])
    worksheet["B10"] = "1200"
    worksheet["C10"] = 999
    highlight = PatternFill("solid", fgColor="FFF2CC")
    for cell in worksheet[10]:
        cell.font = Font(bold=True)
        cell.fill = highlight
    path = tmp_path / "highlighted-override-row.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = {(finding.rule_id, finding.location) for finding in scan.findings}
    assert ("WL004_HARDCODED_VALUE_IN_FORMULA_BAND", "C10") in findings
    assert ("WL006_NUMERIC_TEXT", "B10") in findings
    assert ("WL007_STYLE_OUTLIER", "B10") in findings
    assert not any(patch.cell in {"B10", "C10"} for patch in scan.patches)


def test_freeform_only_row_labels_keep_formula_and_style_findings_review_only(
    tmp_path: Path,
) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    names = [
        "Alice",
        "Bob",
        "Carol",
        "Diego",
        "Eve",
        "Fatima",
        "Grace",
        "Special Case",
        "Hiro",
        "Iris",
        "Jamal",
        "Kai",
        "Lina",
        "Marta",
        "Nora",
        "Omar",
        "Pia",
        "Quinn",
        "Ravi",
        "Sara",
    ]
    worksheet.append(["Name", "Amount", "Calculated"])
    for row, name in enumerate(names, start=2):
        worksheet.append([name, row * 100, f"=B{row}*2"])
    worksheet["B9"].font = Font(bold=True)
    worksheet["C9"] = 999
    path = tmp_path / "freeform-label-special-case.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    findings = {(finding.rule_id, finding.location) for finding in scan.findings}
    assert ("WL004_HARDCODED_VALUE_IN_FORMULA_BAND", "C9") in findings
    assert ("WL007_STYLE_OUTLIER", "B9") in findings
    assert not any(patch.cell in {"B9", "C9"} for patch in scan.patches)


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


def test_sum_boundary_numeric_candidate_is_findings_only(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 10):
        worksheet.cell(row, 1, f"R{row}")
        worksheet.cell(row, 2, row * 100)
    worksheet["A10"] = "Total"
    worksheet["B10"] = "=SUM(B2:B8)"
    path = tmp_path / "numeric-boundary-findings-only.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding for finding in scan.findings if finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY"
    )
    assert finding.location == "B10"
    assert finding.evidence.expected == "=SUM(B2:B9)"
    assert not finding.safe_patch_available
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_sum_boundary_does_not_materialize_missing_adjacent_cell(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 9):
        worksheet.cell(row, 2, row * 100)
    worksheet["B10"] = "=SUM(B2:B8)"
    path = tmp_path / "missing-adjacent-sum-boundary.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(
        path,
        registry=RuleRegistry((builtin_rules.SuspiciousSumBoundaryRule(),)),
    )

    assert not scan.findings


@pytest.mark.parametrize(
    "label",
    [
        "Subtotal",
        "Sub-total",
        "Grand-total",
        "Total adjustment",
        "Average",
        "Mean",
        "Minimum",
        "Maximum",
        "Summary",
        "小计",
        "合计金额",
        "汇总金额",
        "平均金额",
        "最大金额",
        "最小金额",
        "总金额",
        "累计金额",
        "净额",
        "期末余额",
    ],
)
def test_sum_boundary_subtotal_candidate_is_findings_only(tmp_path: Path, label: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 10):
        worksheet.cell(row, 2, row * 100)
    worksheet["A9"] = label
    worksheet["B10"] = "=SUM(B2:B8)"
    path = tmp_path / f"subtotal-boundary-{len(label)}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


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


def test_sum_boundary_reports_adjacent_formula_without_patch(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 9):
        worksheet.cell(row, 1, f"R{row}")
        worksheet.cell(row, 2, row)
    worksheet["A9"] = "R9"
    worksheet["B9"] = "=1+1"
    worksheet["A10"] = "Total"
    worksheet["B10"] = "=SUM(B2:B8)"
    path = tmp_path / "formula-adjacent-sum.xlsx"
    workbook.save(path)
    workbook.close()
    scan = scan_workbook(path)
    finding = next(
        finding for finding in scan.findings if finding.rule_id == "WL005_SUSPICIOUS_SUM_BOUNDARY"
    )
    assert finding.location == "B10"
    assert not finding.safe_patch_available
    assert not any(patch.cell == "B10" for patch in scan.patches)


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


def test_merged_non_anchor_formula_gap_is_never_auto_patched(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(4, 12):
        for column in range(2, 7):
            worksheet.cell(row, column, row + column)
    for column in (2, 3, 5, 6):
        coordinate = worksheet.cell(12, column).coordinate
        source = worksheet.cell(11, column).coordinate
        worksheet[coordinate] = f"={source}*2"
    worksheet.merge_cells("C12:D12")
    path = tmp_path / "merged-non-anchor.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    finding = next(
        finding
        for finding in scan.findings
        if finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "D12"
    )
    assert not finding.safe_patch_available
    assert not any(patch.cell == "D12" for patch in scan.patches)


def test_merged_anchor_formula_gap_is_never_auto_patched(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in (2, 3, 5, 6):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet.merge_cells("B4:C4")
    path = tmp_path / "merged-anchor-gap.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL003_BLANK_IN_FORMULA_BAND" and finding.location == "B4"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B4" for patch in scan.patches)


def test_merged_anchor_literal_is_never_replaced(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 22):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet["B10"] = "Merged note"
    worksheet.merge_cells("B10:C10")
    path = tmp_path / "merged-anchor-literal.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


@pytest.mark.parametrize(
    "label",
    [
        "Subtotal",
        "Sub-total",
        "Grand-total",
        "Average",
        "Mean",
        "Minimum",
        "Maximum",
        "小计",
        "小计金额",
        "合计金额",
        "汇总金额",
        "平均金额",
        "最大金额",
        "最小金额",
    ],
)
def test_summary_row_literal_is_never_replaced(tmp_path: Path, label: str) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 22):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet["A10"] = label
    worksheet["B10"] = 999
    path = tmp_path / f"summary-literal-{len(label)}.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_hidden_literal_is_never_replaced(tmp_path: Path) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    for row in range(2, 22):
        worksheet[f"B{row}"] = f"=A{row}*2"
    worksheet["B10"] = 999
    worksheet.row_dimensions[10].hidden = True
    path = tmp_path / "hidden-literal.xlsx"
    workbook.save(path)
    workbook.close()

    scan = scan_workbook(path)
    assert any(
        finding.rule_id == "WL004_HARDCODED_VALUE_IN_FORMULA_BAND" and finding.location == "B10"
        for finding in scan.findings
    )
    assert not any(patch.cell == "B10" for patch in scan.patches)


def test_builtin_rule_ids_are_stable() -> None:
    legacy_rule_ids = {
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
                "TEXT_DISPLAY_RISK",
                "BORDER_EDGE_INCONSISTENCY",
                "USED_RANGE_INFLATION",
                "IDENTIFIER_SCIENTIFIC_NOTATION",
                "SAVED_VIEW_OFF_CONTENT",
                "WHITESPACE_ONLY_TAIL",
            ],
            start=1,
        )
    }
    assert {rule.rule_id for rule in BUILTIN_RULES} == legacy_rule_ids | {
        "WL028_TEXT_FORMULA",
        "WL029_CIRCULAR_REFERENCE",
        "WL034_PROVABLE_FORMULA_ERROR",
    }
