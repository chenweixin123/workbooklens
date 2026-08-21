from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from workbooklens.models import (
    Evidence,
    Finding,
    PatchKind,
    PatchOperation,
    PatchPrecondition,
    PatchRisk,
    Severity,
)
from workbooklens.rules import RuleContext, RuleRegistry, RuleResult, WorkbookRule
from workbooklens.scanner import scan_workbook
from workbooklens.snapshot import cell_fingerprint


class _PatchRule(WorkbookRule):
    title = "test"

    def __init__(self, rule_id: str, *, description: str, emit_patch: bool = True) -> None:
        self.rule_id = rule_id
        self.description = description
        self.emit_patch = emit_patch

    def run(self, context: RuleContext) -> RuleResult:
        worksheet = context.workbook.active
        cell = worksheet["A1"]
        patch = PatchOperation(
            id="patch-shared-identity",
            kind=PatchKind.SET_NUMERIC,
            sheet=worksheet.title,
            cell="A1",
            before=1,
            after=2,
            confidence=0.99,
            safe=True,
            risk=PatchRisk.SAFE,
            description=self.description,
            precondition=PatchPrecondition(
                cell_fingerprint=cell_fingerprint(cell),
                expected_value=1,
                expected_style_id=cell.style_id,
            ),
        )
        finding = Finding(
            id=f"finding-{self.rule_id}",
            rule_id=self.rule_id,
            title=self.title,
            explanation="test",
            severity=Severity.WARNING,
            confidence=0.99,
            workbook=context.path.name,
            sheet=worksheet.title,
            location="A1",
            evidence=Evidence(summary="test"),
            expected="test",
            suggested_action="test",
            safe_patch_available=True,
            patch_ids=[patch.id],
        )
        return RuleResult(findings=[finding], patches=[patch] if self.emit_patch else [])


def _source(tmp_path: Path) -> Path:
    path = tmp_path / "scanner-normalization.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    worksheet["A1"] = 1
    workbook.save(path)
    workbook.close()
    return path


def test_identical_duplicate_patch_payload_is_deduplicated(tmp_path: Path) -> None:
    source = _source(tmp_path)
    registry = RuleRegistry(
        [
            _PatchRule("TEST_A", description="same"),
            _PatchRule("TEST_B", description="same"),
        ]
    )

    scan = scan_workbook(source, registry=registry)

    assert [patch.id for patch in scan.patches] == ["patch-shared-identity"]
    assert all(finding.patch_ids == ["patch-shared-identity"] for finding in scan.findings)


def test_conflicting_duplicate_patch_payload_is_rejected(tmp_path: Path) -> None:
    source = _source(tmp_path)
    registry = RuleRegistry(
        [
            _PatchRule("TEST_A", description="first"),
            _PatchRule("TEST_B", description="second"),
        ]
    )

    with pytest.raises(RuntimeError, match="Conflicting patch identity"):
        scan_workbook(source, registry=registry)


def test_finding_cannot_reference_a_missing_patch(tmp_path: Path) -> None:
    source = _source(tmp_path)
    registry = RuleRegistry([_PatchRule("TEST_A", description="missing", emit_patch=False)])

    with pytest.raises(RuntimeError, match="references missing patches"):
        scan_workbook(source, registry=registry)
