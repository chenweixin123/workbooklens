from __future__ import annotations

import pytest

from workbooklens.exceptions import UsageError
from workbooklens.models import (
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
)
from workbooklens.repair.ooxml_patch import _select_patches


def _patch(
    patch_id: str,
    *,
    kind: PatchKind = PatchKind.SET_FORMULA,
    safe: bool = True,
    risk: PatchRisk = PatchRisk.SAFE,
    confidence: float = 0.99,
    atomic_group: str | None = None,
    after: object | None = None,
) -> PatchOperation:
    patch_after = after
    if patch_after is None:
        patch_after = 16.0 if kind == PatchKind.SET_COLUMN_WIDTH else "=1+1"
    return PatchOperation(
        id=patch_id,
        kind=kind,
        sheet="Sheet1",
        cell=f"A{patch_id.removeprefix('p')}",
        after=patch_after,
        confidence=confidence,
        safe=safe,
        risk=risk,
        description="selection test patch",
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


def test_low_level_explicit_selection_rejects_incomplete_atomic_group() -> None:
    plan = _plan(
        _patch("p1", atomic_group="g1"),
        _patch("p2", atomic_group="g1"),
    )

    with pytest.raises(UsageError, match="Atomic patch groups must be selected in full"):
        _select_patches(plan, {"p1"}, safe_only=False)


def test_low_level_layout_acceptance_still_rejects_low_confidence() -> None:
    plan = _plan(
        _patch(
            "p1",
            kind=PatchKind.SET_COLUMN_WIDTH,
            safe=False,
            risk=PatchRisk.LAYOUT_REVIEW,
            confidence=0.94,
        )
    )

    with pytest.raises(UsageError, match="authorized repair risk boundary"):
        _select_patches(
            plan,
            {"p1"},
            safe_only=False,
            accept_layout_risk=True,
        )
