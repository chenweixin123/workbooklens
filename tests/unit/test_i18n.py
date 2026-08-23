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
from workbooklens.rules.builtin import BUILTIN_RULES
from workbooklens.scanner import scan_workbook

runner = CliRunner()


def test_catalogs_are_complete_and_locale_normalization_is_bounded() -> None:
    assert_catalog_complete()
    assert_error_catalog_complete()
    assert normalize_language("zh_Hans_CN") == "zh-CN"
    assert normalize_language("en-US") == "en"
    assert normalize_language("fr") == "en"
    assert translate("severity.error", "zh-CN") == "错误"


def test_every_builtin_rule_has_exact_bilingual_title_coverage() -> None:
    emitted = {rule.rule_id: rule.title for rule in BUILTIN_RULES}
    assert set(emitted) == set(BUILTIN_RULE_TITLES)
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

    source = Path(builtin.__file__).read_text(encoding="utf-8")
    tree = ast.parse(source)
    fields = {"description", "explanation", "expected", "suggested_action", "summary"}
    texts = {
        text
        for node in ast.walk(tree)
        if isinstance(node, ast.Call)
        for keyword in node.keywords
        if keyword.arg in fields
        for text in _literal_branches(keyword.value)
    }
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
