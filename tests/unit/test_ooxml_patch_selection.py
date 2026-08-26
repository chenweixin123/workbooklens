from __future__ import annotations

from pathlib import Path

import pytest

from workbooklens.exceptions import PatchValidationError, UsageError
from workbooklens.models import (
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
)
from workbooklens.repair.ooxml_patch import (
    _patch_application_sort_key,
    _select_patches,
    _validated_temporary_path,
)


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


def test_patch_application_orders_normalization_before_formula_before_layout() -> None:
    numeric = _patch("p1", kind=PatchKind.SET_NUMERIC, after=1)
    number_format = PatchOperation(
        id="p2",
        kind=PatchKind.COPY_NUMBER_FORMAT,
        sheet="Sheet1",
        cell="B1",
        before="General",
        after="#,##0",
        source_cell="B2",
        confidence=0.99,
        safe=True,
        risk=PatchRisk.SAFE,
        description="copy a peer number format",
        precondition=PatchPrecondition(cell_fingerprint="fingerprint-p2"),
    )
    formula = _patch("p3", kind=PatchKind.SET_FORMULA)
    layout = _patch(
        "p4",
        kind=PatchKind.SET_COLUMN_WIDTH,
        safe=False,
        risk=PatchRisk.LAYOUT_REVIEW,
        after=18.0,
    )

    ordered = sorted(
        [layout, formula, number_format, numeric],
        key=_patch_application_sort_key,
    )

    assert {patch.kind for patch in ordered[:2]} == {
        PatchKind.SET_NUMERIC,
        PatchKind.COPY_NUMBER_FORMAT,
    }
    assert ordered[2].kind == PatchKind.SET_FORMULA
    assert ordered[3].kind == PatchKind.SET_COLUMN_WIDTH


@pytest.mark.parametrize("enforce_safety", [True, False])
def test_low_level_selection_never_applies_migrated_v2_formula(
    enforce_safety: bool,
) -> None:
    payload = _plan(_patch("p1")).model_dump(mode="json")
    payload["schema_version"] = 2
    payload["patches"][0].pop("derivation")
    plan = PatchPlan.model_validate(payload)

    with pytest.raises(UsageError, match="Legacy schema v2 formula patches are unverified"):
        _select_patches(
            plan,
            {"p1"},
            safe_only=False,
            enforce_safety=enforce_safety,
            accept_semantic_risk=True,
        )


def test_ooxml_temporary_path_must_be_direct_child(tmp_path: Path) -> None:
    intended = tmp_path / "intended"
    escaped = tmp_path / "escaped"
    intended.mkdir()
    escaped.mkdir()

    with pytest.raises(PatchValidationError, match="escaped its intended parent directory"):
        _validated_temporary_path(escaped / "candidate.xlsx", intended)
