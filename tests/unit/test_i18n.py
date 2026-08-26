# ruff: noqa: RUF001
from __future__ import annotations

import ast
import json
import re
from pathlib import Path

from typer.testing import CliRunner

from workbooklens.cli import app
from workbooklens.demo.workflow import generate_demo_workbook
from workbooklens.diff import write_diff_report
from workbooklens.exceptions import UsageError
from workbooklens.i18n import (
    BUILTIN_RULE_TITLES,
    assert_catalog_complete,
    assert_error_catalog_complete,
    canonical_text_translation,
    localize_evidence_value,
    localize_exception,
    localize_finding,
    localize_scan_result,
    localized_error,
    new_diagnostic_id,
    normalize_language,
    translate,
)
from workbooklens.models import (
    CellChange,
    Confidence,
    Evidence,
    Finding,
    Severity,
    WorkbookDiff,
)
from workbooklens.reports import write_scan_report
from workbooklens.rules import default_registry
from workbooklens.rules.builtin import BUILTIN_RULES
from workbooklens.rules.data_quality import DATA_QUALITY_RULES
from workbooklens.rules.formula_semantics import FORMULA_SEMANTIC_RULES
from workbooklens.rules.inferred_semantics import INFERRED_SEMANTIC_RULES
from workbooklens.rules.layout_geometry import LAYOUT_GEOMETRY_RULES
from workbooklens.rules.print_quality import PRINT_QUALITY_RULES
from workbooklens.rules.profile_quality import PROFILE_QUALITY_RULES
from workbooklens.rules.relational_semantics import RELATIONAL_SEMANTIC_RULES
from workbooklens.scanner import scan_workbook

runner = CliRunner()


def test_evidence_values_localize_recursively_without_mutating_input() -> None:
    canonical = {
        "proof": "propagated_formula_error",
        "source_proof": [
            {"font_size": 8, "fixed_total_pages": 3},
            ("unknown_key", {"custom": "业务原文"}),
        ],
    }

    chinese = localize_evidence_value(canonical, "zh-CN")
    english = localize_evidence_value(canonical, "en")

    assert chinese == {
        "证明": "传播的公式错误",
        "源证明": [
            {"字号": 8, "固定总页数": 3},
            ("unknown_key", {"custom": "业务原文"}),
        ],
    }
    assert english == {
        "Proof": "Propagated formula error",
        "Source proof": [
            {"Font size": 8, "Fixed total pages": 3},
            ("unknown_key", {"custom": "业务原文"}),
        ],
    }
    assert localize_evidence_value(
        {"observed": ["proof", "formula", "numeric", "hidden", "Chart"]}, "zh-CN"
    ) == {"实际值": ["proof", "formula", "numeric", "hidden", "Chart"]}
    assert localize_evidence_value(
        {"kind": ["formula", "numeric", "hidden", "Chart"]}, "zh-CN"
    ) == {"类型": ["公式", "数值", "隐藏", "图表"]}
    assert localize_evidence_value({"proof": 1, "证明": 2}, "zh-CN") == {
        "proof": 1,
        "证明": 2,
    }
    assert canonical == {
        "proof": "propagated_formula_error",
        "source_proof": [
            {"font_size": 8, "fixed_total_pages": 3},
            ("unknown_key", {"custom": "业务原文"}),
        ],
    }
    assert chinese is not canonical
    assert chinese["源证明"] is not canonical["source_proof"]
    assert english["Source proof"] is not canonical["source_proof"]


def test_catalogs_are_complete_and_locale_normalization_is_bounded() -> None:
    assert_catalog_complete()
    assert_error_catalog_complete()
    assert normalize_language("zh_Hans_CN") == "zh-CN"
    assert normalize_language("en-US") == "en"
    assert normalize_language("fr") == "en"
    assert translate("severity.error", "zh-CN") == "错误"
    assert translate("risk.formula_derived", "zh-CN") == "公式推导"
    assert translate("risk.semantic_review", "zh-CN") == "需确认语义"
    assert translate("patch_kind.normalize_text", "zh-CN") == "规范化数值"
    assert translate("web.results_auto_repair", "zh-CN") == "一键安全修复"
    assert translate("web.profile_label", "zh-CN") == "工作簿 Profile（可选 .yml 或 .yaml）"
    assert translate("web.patch_derivation_evidence", "zh-CN") == "推导证据"
    assert translate("web.patch_candidate_count", "zh-CN") == "候选数量"
    assert translate("web.patch_requires_recalculation", "zh-CN") == "是否需要重算"
    assert translate("validation_status.degraded", "zh-CN") == "已安全降级"


def test_every_builtin_rule_has_exact_bilingual_title_coverage() -> None:
    all_rule_types = (
        *BUILTIN_RULES,
        *DATA_QUALITY_RULES,
        *PROFILE_QUALITY_RULES,
        *INFERRED_SEMANTIC_RULES,
        *FORMULA_SEMANTIC_RULES,
        *LAYOUT_GEOMETRY_RULES,
        *PRINT_QUALITY_RULES,
        *RELATIONAL_SEMANTIC_RULES,
    )
    emitted = {rule.rule_id: rule.title for rule in all_rule_types}
    assert set(emitted) == set(BUILTIN_RULE_TITLES)
    assert {rule.rule_id for rule in default_registry().values()} == set(BUILTIN_RULE_TITLES)
    assert {rule_id: titles[0] for rule_id, titles in BUILTIN_RULE_TITLES.items()} == emitted
    assert all(titles[1] and titles[1] != titles[0] for titles in BUILTIN_RULE_TITLES.values())


def _literal_branches(node: ast.expr) -> list[str]:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return [node.value]
    if isinstance(node, ast.IfExp):
        return [*_literal_branches(node.body), *_literal_branches(node.orelse)]
    return []


def test_all_static_builtin_finding_patch_and_evidence_templates_are_translated() -> None:
    import workbooklens.rules.builtin as builtin
    import workbooklens.rules.data_quality as data_quality
    import workbooklens.rules.formula_semantics as formula_semantics
    import workbooklens.rules.inferred_semantics as inferred_semantics
    import workbooklens.rules.layout_geometry as layout_geometry
    import workbooklens.rules.print_quality as print_quality
    import workbooklens.rules.profile_quality as profile_quality
    import workbooklens.rules.relational_semantics as relational_semantics

    fields = {"description", "explanation", "expected", "suggested_action", "summary"}
    texts: set[str] = set()
    for module in (
        builtin,
        data_quality,
        formula_semantics,
        inferred_semantics,
        layout_geometry,
        print_quality,
        profile_quality,
        relational_semantics,
    ):
        source = Path(module.__file__).read_text(encoding="utf-8")
        tree = ast.parse(source)
        texts.update(
            text
            for node in ast.walk(tree)
            if isinstance(node, ast.Call)
            for keyword in node.keywords
            if keyword.arg in fields
            for text in _literal_branches(keyword.value)
        )
    missing = sorted(text for text in texts if canonical_text_translation(text, "zh-CN") is None)
    assert not missing


def test_all_known_dynamic_builtin_templates_are_translated() -> None:
    samples = (
        "7 of 8 formulas share one signature",
        "A simple total stops one row before a directly adjacent non-hidden numeric peer.",
        "A simple total stops one row before a directly adjacent formula peer.",
        "SUM ends at row 20, while B21 is adjacent",
        "12 peer cells are stored as numbers",
        "One visual style appears in 9 of 10 peer cells",
        "veryHidden sheet contains 3 nonempty cells",
        "Hidden row 8 contains 2 nonempty cells",
        "Hidden column range B:D contains 4 nonempty cells",
        "Configured key value appears 2 times",
        "Values in Data!A2:A20 are unique.",
        "target sheet 'Missing' does not exist",
        "target range 'A0' is invalid",
        "8 peer cells share one validation signature",
        "Increase row 7 height to 42.5 points after review.",
        "Copy the parallel-consensus left border edge after review.",
        "Missing parallel-consensus edge(s): left, right",
        "14 exact blank styled cells and 3 empty row records form a separated tail",
        "6 literal-whitespace cells form an outer tail",
        "A nonblank email value fails conservative structural validation",
        "A percentage field uses a conflicting number-format role",
        "Chart 1 covers 7 populated non-source cells",
        "Image 2 covers 3 populated non-source cells",
        "4 fixed-format numeric values exceed the estimated column width",
        "Explicit row height 48 is 3.20 times the detail-row median",
        "Merged title combines 4 unusual role-specific style components",
        "Body-role component consensus identifies 12 anomalous cells",
        "3 total-row cells use formats inconsistent with their body columns",
        "Confirm whether 'Sales'!B2:B5 should cover 'Sales'!B2:B8.",
        "1 manual row break(s) split the dense leading portion of an inferred table",
    )
    assert all(canonical_text_translation(text, "zh-CN") for text in samples)


def test_demo_findings_and_patches_localize_strictly_without_identity_drift(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    scan = scan_workbook(
        workbook,
        config={"keys": [{"sheet": "Sales", "range": "A2:A22"}]},
    )
    localized = localize_scan_result(scan, "zh-CN", strict=True)

    assert len(localized.findings) == len(scan.findings)
    assert len(localized.patches) == len(scan.patches)
    for canonical, display in zip(scan.findings, localized.findings, strict=True):
        assert display.id == canonical.id
        assert display.content_fingerprint == canonical.content_fingerprint
        assert display.rule_id == canonical.rule_id
        assert display.severity == canonical.severity
        assert display.workbook == canonical.workbook
        assert display.sheet == canonical.sheet
        assert display.location == canonical.location
        assert display.patch_ids == canonical.patch_ids
        assert display.evidence.observed == canonical.evidence.observed
        assert display.evidence.expected == canonical.evidence.expected
        assert display.evidence.peers == canonical.evidence.peers
        assert display.evidence.details == canonical.evidence.details
    for canonical, display in zip(scan.patches, localized.patches, strict=True):
        canonical_payload = canonical.model_dump(mode="json")
        display_payload = display.model_dump(mode="json")
        canonical_payload.pop("description")
        display_payload.pop("description")
        assert display_payload == canonical_payload


def test_unknown_plugin_text_is_preserved_but_explicitly_marked() -> None:
    finding = Finding(
        id="plugin-finding",
        rule_id="PLUGIN001",
        title="Custom check",
        explanation="Plugin explanation",
        severity=Severity.INFO,
        confidence=Confidence(1.0),
        workbook="sample.xlsx",
        evidence=Evidence(summary="Plugin evidence"),
        expected="Plugin expectation",
        suggested_action="Plugin action",
    )
    localized = localize_finding(finding, "zh-CN", strict=True)
    assert localized.title == "第三方插件提供：Custom check"
    assert localized.explanation == finding.explanation


def test_public_errors_are_stable_localized_and_do_not_expose_raw_diagnostics() -> None:
    raw = 'C:\\secret\\book.xls <Objs Version="1.1.0.1">PowerShell CLIXML trace'
    exc = UsageError(raw, error_key="conversion.all_providers_failed")
    public = localize_exception(
        exc,
        "zh-CN",
        operation="conversion",
        diagnostic_id="WL-ABC123",
    )
    assert public.code == "WL-CNV-003"
    assert public.diagnostic_id == "WL-ABC123"
    assert raw not in public.title + public.message + public.suggestion
    assert "PowerShell" not in public.title + public.message + public.suggestion
    assert localized_error("not-a-real-key", "zh-CN").code == "WL-INT-001"
    identifiers = {new_diagnostic_id() for _ in range(4)}
    assert len(identifiers) == 4
    assert all(re.fullmatch(r"WL-[0-9A-F]{12}", identifier) for identifier in identifiers)


def test_chinese_scan_report_localizes_html_but_keeps_json_and_sarif_canonical(
    tmp_path: Path,
) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    scan = scan_workbook(workbook)
    paths = write_scan_report(scan, tmp_path / "report", language="zh-CN")
    html = paths["html"].read_text(encoding="utf-8")
    findings = json.loads(paths["findings"].read_text(encoding="utf-8"))
    sarif = json.loads(paths["sarif"].read_text(encoding="utf-8"))
    assert '<html lang="zh-CN">' in html
    assert "工作表概览" in html
    assert "建议操作" in html
    assert findings["findings"][0]["title"] in {
        titles[0] for titles in BUILTIN_RULE_TITLES.values()
    }
    assert sarif["runs"][0]["tool"]["driver"]["rules"][0]["shortDescription"]["text"] in {
        titles[0] for titles in BUILTIN_RULE_TITLES.values()
    }


def test_chinese_diff_report_localizes_html_but_keeps_json_canonical(tmp_path: Path) -> None:
    diff = WorkbookDiff(
        before_sha256="before",
        after_sha256="after",
        cell_changes=[
            CellChange(
                sheet="Data",
                cell="A1",
                change_type="value",
                before=1,
                after=2,
            )
        ],
    )
    paths = write_diff_report(diff, tmp_path / "diff.html", language="zh-CN")
    html = paths["html"].read_text(encoding="utf-8")
    payload = json.loads(paths["json"].read_text(encoding="utf-8"))
    assert '<html lang="zh-CN">' in html
    assert "工作簿语义差异" in html
    assert "单元格更改" in html
    assert payload["cell_changes"][0]["change_type"] == "value"


def test_cli_language_choice_localizes_human_output_and_html_only(tmp_path: Path) -> None:
    workbook = tmp_path / "demo.xlsx"
    generate_demo_workbook(workbook)
    output = tmp_path / "report"
    result = runner.invoke(
        app,
        ["--language", "zh-CN", "scan", str(workbook), "--out", str(output)],
    )
    assert result.exit_code == 0, result.stdout
    assert "扫描完成" in result.stdout
    assert "工作表概览" in (output / "report.html").read_text(encoding="utf-8")
    payload = json.loads((output / "findings.json").read_text(encoding="utf-8"))
    assert payload["findings"][0]["title"] in {titles[0] for titles in BUILTIN_RULE_TITLES.values()}
