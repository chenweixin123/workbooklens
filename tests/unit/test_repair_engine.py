from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any

import pytest

import workbooklens.repair.engine as engine
from workbooklens.exceptions import PatchValidationError, UsageError
from workbooklens.models import (
    Evidence,
    Finding,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
    Severity,
)


def _layout_plan() -> PatchPlan:
    patches = [
        PatchOperation(
            id=patch_id,
            kind=PatchKind.SET_COLUMN_WIDTH,
            sheet="Sheet1",
            cell=cell,
            before=8.43,
            after=16.0,
            confidence=0.99,
            safe=False,
            risk=PatchRisk.LAYOUT_REVIEW,
            description="test layout patch",
            precondition=PatchPrecondition(cell_fingerprint=f"fingerprint-{patch_id}"),
            atomic_group="layout-group",
        )
        for patch_id, cell in (("p1", "A1"), ("p2", "B1"))
    ]
    return PatchPlan(
        tool_version="2.1.0",
        source_name="input.xlsx",
        source_sha256="abc123",
        patches=patches,
    )


def test_apply_resolves_atomic_group_before_calling_low_level_patcher(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = _layout_plan()
    scan = SimpleNamespace(findings=[])
    captured: dict[str, Any] = {}

    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        captured.update(kwargs)
        selected_ids = kwargs["selected_ids"]
        selected = [patch for patch in plan.patches if patch.id in selected_ids]
        low_level = SimpleNamespace(
            source_sha256="source-hash",
            output_sha256="output-hash",
            changes=[],
            formula_changed=False,
        )
        return low_level, selected

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        tmp_path / "output.xlsx",
        selected_ids=["p1"],
        accept_layout_risk=True,
    )

    assert captured["selected_ids"] == {"p1", "p2"}
    assert captured["safe_only"] is False
    assert captured["accept_layout_risk"] is True
    assert result.applied_patch_ids == ["p1", "p2"]


def test_apply_rejects_selection_before_scanning(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = _layout_plan()
    scanned = False

    def fake_scan(*args: Any, **kwargs: Any) -> None:
        nonlocal scanned
        scanned = True

    monkeypatch.setattr(engine, "scan_workbook", fake_scan)

    with pytest.raises(UsageError, match="Unknown patch IDs"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            tmp_path / "output.xlsx",
            selected_ids=["missing"],
        )

    assert not scanned


def test_apply_rejects_targeted_finding_that_changes_identity_but_remains(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = _layout_plan()

    def finding(identifier: str) -> Finding:
        return Finding(
            id=identifier,
            rule_id="WL016_TEXT_DISPLAY_RISK",
            title="display risk",
            explanation="test",
            severity=Severity.WARNING,
            confidence=0.99,
            workbook="input.xlsx",
            sheet="Sheet1",
            location="A1",
            evidence=Evidence(summary="test"),
            expected="visible text",
            suggested_action="test",
            safe_patch_available=False,
            patch_ids=["p1", "p2"] if identifier == "before-id" else [],
        )

    scans = iter(
        [
            SimpleNamespace(findings=[finding("before-id")]),
            SimpleNamespace(findings=[finding("after-id")]),
        ]
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        output = args[2]
        output.write_bytes(b"candidate")
        low_level = SimpleNamespace(
            source_sha256="source-hash",
            output_sha256="output-hash",
            changes=[],
            formula_changed=False,
        )
        return low_level, plan.patches

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="did not resolve targeted findings"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=["p1"],
            accept_layout_risk=True,
        )

    assert not output.exists()
