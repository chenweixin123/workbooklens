from __future__ import annotations

import json
from pathlib import Path

import pytest

from workbooklens.exceptions import UsageError
from workbooklens.models import (
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
)
from workbooklens.repair.planning import load_patch_plan, resolve_patch_selection


def _patch(
    patch_id: str,
    *,
    kind: PatchKind = PatchKind.SET_FORMULA,
    safe: bool = True,
    risk: PatchRisk = PatchRisk.SAFE,
    confidence: float = 0.99,
    atomic_group: str | None = None,
) -> PatchOperation:
    return PatchOperation(
        id=patch_id,
        kind=kind,
        sheet="Sheet1",
        cell=f"A{patch_id.removeprefix('p') or '1'}",
        before=None,
        after="=1+1",
        confidence=confidence,
        safe=safe,
        risk=risk,
        description="test patch",
        precondition=PatchPrecondition(cell_fingerprint=f"fingerprint-{patch_id}"),
        atomic_group=atomic_group,
    )


def _plan(*patches: PatchOperation) -> PatchPlan:
    return PatchPlan(
        tool_version="2.1.0",
        source_name="input.xlsx",
        source_sha256="abc123",
        patches=list(patches),
    )


def _layout_patch(patch_id: str, *, atomic_group: str | None = None) -> PatchOperation:
    return _patch(
        patch_id,
        kind=PatchKind.SET_COLUMN_WIDTH,
        safe=False,
        risk=PatchRisk.LAYOUT_REVIEW,
        atomic_group=atomic_group,
    )


def test_explicit_selection_expands_the_entire_atomic_group() -> None:
    plan = _plan(
        _patch("p1", atomic_group="g1"),
        _patch("p2", atomic_group="g1"),
        _patch("p3"),
    )

    assert resolve_patch_selection(plan, ["p1"]) == {"p1", "p2"}


def test_explicit_selection_expands_multiple_groups_and_keeps_ungrouped_patches() -> None:
    plan = _plan(
        _patch("p1", atomic_group="g1"),
        _patch("p2", atomic_group="g1"),
        _patch("p3", atomic_group="g2"),
        _patch("p4", atomic_group="g2"),
        _patch("p5"),
    )

    assert resolve_patch_selection(plan, ["p1", "p3", "p5"]) == {
        "p1",
        "p2",
        "p3",
        "p4",
        "p5",
    }


def test_safe_only_excludes_layout_review_patches() -> None:
    plan = _plan(_patch("p1"), _layout_patch("p2"))

    assert resolve_patch_selection(plan, safe_only=True) == {"p1"}


def test_safe_only_skips_a_whole_atomic_group_with_a_review_member() -> None:
    plan = _plan(
        _patch("p1", atomic_group="mixed"),
        _layout_patch("p2", atomic_group="mixed"),
        _patch("p3"),
    )

    assert resolve_patch_selection(plan, safe_only=True) == {"p3"}


def test_layout_review_requires_explicit_acceptance_and_then_expands_group() -> None:
    plan = _plan(
        _layout_patch("p1", atomic_group="layout"),
        _layout_patch("p2", atomic_group="layout"),
    )

    with pytest.raises(UsageError, match="--accept-layout-risk"):
        resolve_patch_selection(plan, ["p1"])

    assert resolve_patch_selection(
        plan,
        ["p1"],
        accept_layout_risk=True,
    ) == {"p1", "p2"}


def test_layout_acceptance_does_not_enable_other_unsafe_or_low_confidence_patches() -> None:
    unsafe_plan = _plan(_patch("p1", safe=False))
    with pytest.raises(UsageError, match="refuses unsafe"):
        resolve_patch_selection(unsafe_plan, ["p1"], accept_layout_risk=True)

    low_confidence_plan = _plan(_patch("p2", confidence=0.94))
    with pytest.raises(UsageError, match=r"0\.95 confidence"):
        resolve_patch_selection(low_confidence_plan, ["p2"], accept_layout_risk=True)


def test_selection_rejects_duplicate_unknown_empty_and_mutually_exclusive_inputs() -> None:
    plan = _plan(_patch("p1"))

    with pytest.raises(UsageError, match="Duplicate --patch-id"):
        resolve_patch_selection(plan, ["p1", "p1"])
    with pytest.raises(UsageError, match="Unknown patch IDs"):
        resolve_patch_selection(plan, ["missing"])
    with pytest.raises(UsageError, match="Select at least one"):
        resolve_patch_selection(plan)
    with pytest.raises(UsageError, match="mutually exclusive"):
        resolve_patch_selection(plan, ["p1"], safe_only=True)
    with pytest.raises(UsageError, match="mutually exclusive"):
        resolve_patch_selection(plan, safe_only=True, accept_layout_risk=True)


def test_selection_rejects_duplicate_plan_ids() -> None:
    plan = _plan(_patch("p1"), _patch("p2"))
    plan.patches[1].id = "p1"

    with pytest.raises(UsageError, match="duplicate patch IDs"):
        resolve_patch_selection(plan, safe_only=True)


def test_safe_only_rejects_an_empty_eligible_selection() -> None:
    plan = _plan(_layout_patch("p1"))

    with pytest.raises(UsageError, match="No eligible patches"):
        resolve_patch_selection(plan, safe_only=True)


def test_load_patch_plan_clearly_rejects_schema_one(tmp_path: Path) -> None:
    path = tmp_path / "schema-one.json"
    path.write_text(json.dumps({"schema_version": 1}), encoding="utf-8")

    with pytest.raises(UsageError, match="schema version 2"):
        load_patch_plan(path)
