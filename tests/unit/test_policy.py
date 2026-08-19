from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path

import pytest

from workbooklens.exceptions import UsageError
from workbooklens.models import Confidence, Evidence, Finding, Severity
from workbooklens.policy import (
    FindingSuppression,
    apply_finding_policy,
    load_baseline,
    normalize_source_scope,
    source_scope_for_path,
)


def _finding(identifier: str, *, rule: str = "WL001", sheet: str = "Data") -> Finding:
    return Finding(
        id=identifier,
        rule_id=rule,
        title="Example",
        explanation="Example finding",
        severity=Severity.ERROR,
        confidence=Confidence(1.0),
        workbook="book.xlsx",
        sheet=sheet,
        location="A1",
        evidence=Evidence(summary="example", observed="bad", expected="good"),
        expected="good",
        suggested_action="Review",
    )


def test_policy_classifies_baseline_new_suppressed_and_expired() -> None:
    findings = [_finding("old"), _finding("waived", rule="WL010"), _finding("new")]
    suppressions = [
        FindingSuppression(
            id="accepted-volatility",
            reason="Approved template behavior",
            rules=["WL010"],
            sheets=["D*"],
        ),
        FindingSuppression(
            id="expired",
            reason="Temporary legacy waiver",
            finding_ids=["new"],
            expires=date(2025, 1, 1),
        ),
    ]
    policy = apply_finding_policy(
        findings,
        suppressions=suppressions,
        baseline_ids=frozenset({"old"}),
        new_only=True,
        as_of=date(2026, 8, 19),
    )
    assert [finding.id for finding in policy.active_findings] == ["new"]
    assert [finding.id for finding in policy.baseline_findings] == ["old"]
    assert [item.finding.id for item in policy.suppressed_findings] == ["waived"]
    assert policy.expired_suppression_ids == ("expired",)
    assert policy.summary == {
        "total": 3,
        "active": 1,
        "suppressed": 1,
        "baseline_known": 1,
        "new": 1,
    }


def test_sheet_and_location_globs_do_not_match_missing_dimensions() -> None:
    workbook_finding = _finding("workbook-level").model_copy(
        update={"sheet": None, "location": None}
    )
    cell_finding = _finding("cell-level")
    sheet_scope = FindingSuppression(
        id="sheet-scope",
        reason="Only findings attached to a sheet",
        sheets=["*"],
    )
    location_scope = FindingSuppression(
        id="location-scope",
        reason="Only findings attached to a location",
        locations=["*"],
    )

    assert not sheet_scope.matches(workbook_finding, as_of=date(2026, 8, 19))
    assert not location_scope.matches(workbook_finding, as_of=date(2026, 8, 19))
    assert sheet_scope.matches(cell_finding, as_of=date(2026, 8, 19))
    assert location_scope.matches(cell_finding, as_of=date(2026, 8, 19))


def test_baseline_loader_accepts_findings_report_and_rejects_bad_ids(tmp_path: Path) -> None:
    report = tmp_path / "findings.json"
    report.write_text(
        json.dumps(
            {
                "schema_version": 2,
                "source_scope": "books/book.xlsx",
                "findings": [{"id": "finding-a"}],
                "baseline_findings": [{"id": "finding-b"}],
            }
        ),
        encoding="utf-8",
    )
    assert load_baseline(report, expected_source_scope="books/book.xlsx") == frozenset(
        {"finding-a", "finding-b"}
    )
    with pytest.raises(UsageError, match="does not match"):
        load_baseline(report, expected_source_scope="archive/book.xlsx")
    report.write_text(
        json.dumps(
            {
                "finding_ids": ["authoritative"],
                "findings": [{"id": "ignored-new"}],
                "baseline_findings": [{"id": "ignored-old"}],
                "suppressed_findings": [{"finding": {"id": "ignored-suppressed"}}],
            }
        ),
        encoding="utf-8",
    )
    assert load_baseline(report) == frozenset({"authoritative"})
    report.write_text(
        json.dumps(
            {
                "findings": [{"id": "finding-a"}],
                "suppressed_findings": [{"finding": {"id": "not-a-baseline"}}],
            }
        ),
        encoding="utf-8",
    )
    assert load_baseline(report) == frozenset({"finding-a"})

    report.write_text(
        json.dumps(
            {
                "workbooks": {
                    "books/same.xlsx": ["book-a"],
                    "archive/same.xlsx": ["book-b"],
                }
            }
        ),
        encoding="utf-8",
    )
    assert load_baseline(report, expected_source_scope="books/same.xlsx") == frozenset({"book-a"})
    assert load_baseline(report, expected_source_scope="archive/same.xlsx") == frozenset({"book-b"})
    assert load_baseline(report, expected_source_scope="new/same.xlsx") == frozenset()
    report.write_text(json.dumps({"finding_ids": ["", 3]}), encoding="utf-8")
    with pytest.raises(UsageError, match="nonempty strings"):
        load_baseline(report)


@pytest.mark.parametrize("legacy_fields", [{}, {"source": ""}, {"source": 3}])
def test_scoped_baseline_rejects_findings_report_without_usable_source(
    tmp_path: Path, legacy_fields: dict[str, object]
) -> None:
    report = tmp_path / "findings.json"
    report.write_text(
        json.dumps({**legacy_fields, "findings": [{"id": "finding-a"}]}),
        encoding="utf-8",
    )

    with pytest.raises(UsageError, match="finding_ids"):
        load_baseline(report, expected_source_scope="books/book.xlsx")


def test_scoped_baseline_accepts_legacy_report_with_matching_source(tmp_path: Path) -> None:
    report = tmp_path / "findings.json"
    report.write_text(
        json.dumps({"source": "book.xlsx", "findings": [{"id": "finding-a"}]}),
        encoding="utf-8",
    )

    assert load_baseline(report, expected_source_scope="books/book.xlsx") == frozenset(
        {"finding-a"}
    )


@pytest.mark.parametrize("value", ["", "../book.xlsx", "/book.xlsx", "a/./book.xlsx"])
def test_source_scope_rejects_nonportable_paths(value: str) -> None:
    with pytest.raises(UsageError, match="Source scope"):
        normalize_source_scope(value)


def test_external_source_scopes_distinguish_same_basenames_without_leaking_paths(
    tmp_path: Path,
) -> None:
    base = tmp_path / "workspace"
    first = tmp_path / "one" / "same.xlsx"
    second = tmp_path / "two" / "same.xlsx"
    first_scope = source_scope_for_path(first, base=base)
    second_scope = source_scope_for_path(second, base=base)

    assert first_scope != second_scope
    assert first_scope.endswith("/same.xlsx")
    assert str(tmp_path).replace("\\", "/") not in first_scope


def test_external_source_scope_uses_platform_path_case_semantics(tmp_path: Path) -> None:
    base = tmp_path / "workspace"
    upper = tmp_path / "CaseParent" / "same.xlsx"
    lower = tmp_path / "caseparent" / "same.xlsx"

    upper_scope = source_scope_for_path(upper, base=base)
    lower_scope = source_scope_for_path(lower, base=base)

    if os.path.normcase("CaseParent") == os.path.normcase("caseparent"):
        assert upper_scope == lower_scope
    else:
        assert upper_scope != lower_scope
