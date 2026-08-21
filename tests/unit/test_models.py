from __future__ import annotations

import pytest
from pydantic import ValidationError

from workbooklens.models import (
    Confidence,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
    Severity,
    SheetSnapshot,
)


@pytest.mark.parametrize("value", [0.0, 0.25, 1.0])
def test_confidence_accepts_closed_unit_interval(value: float) -> None:
    assert float(Confidence(value)) == value


@pytest.mark.parametrize("value", [-0.01, 1.01])
def test_confidence_rejects_values_outside_unit_interval(value: float) -> None:
    with pytest.raises(ValidationError):
        Confidence(value)


def test_severity_values_are_stable() -> None:
    assert [value.value for value in Severity] == ["info", "warning", "error", "critical"]


def _patch(
    *,
    patch_id: str = "patch-1",
    kind: PatchKind = PatchKind.SET_FORMULA,
    safe: bool = True,
    risk: PatchRisk | None = None,
    atomic_group: str | None = None,
) -> PatchOperation:
    payload: dict[str, object] = {
        "id": patch_id,
        "kind": kind,
        "sheet": "Sheet1",
        "cell": "A1",
        "before": None,
        "after": "=1+1",
        "confidence": 0.99,
        "safe": safe,
        "description": "test patch",
        "precondition": PatchPrecondition(
            cell_fingerprint="cell-fingerprint",
            layout_fingerprint="layout-fingerprint",
        ),
        "atomic_group": atomic_group,
    }
    if risk is not None:
        payload["risk"] = risk
    return PatchOperation.model_validate(payload)


def test_non_layout_patch_defaults_to_safe_risk() -> None:
    patch = _patch()
    assert patch.risk == PatchRisk.SAFE
    assert patch.safe_only_eligible


def test_layout_patch_infers_review_risk_but_must_not_be_marked_safe() -> None:
    patch = _patch(kind=PatchKind.SET_COLUMN_WIDTH, safe=False)
    assert patch.risk == PatchRisk.LAYOUT_REVIEW
    assert not patch.safe_only_eligible

    with pytest.raises(ValidationError, match="safe=false"):
        _patch(kind=PatchKind.SET_COLUMN_WIDTH, safe=True)


def test_layout_patch_cannot_claim_safe_risk() -> None:
    with pytest.raises(ValidationError, match="require risk='layout_review'"):
        _patch(
            kind=PatchKind.COPY_BORDER,
            safe=False,
            risk=PatchRisk.SAFE,
        )


def test_patch_plan_schema_two_round_trip_preserves_layout_metadata() -> None:
    patch = _patch(
        kind=PatchKind.SET_ROW_HEIGHT,
        safe=False,
        risk=PatchRisk.LAYOUT_REVIEW,
        atomic_group="layout-group-1",
    )
    plan = PatchPlan(
        tool_version="2.1.0",
        source_name="input.xlsx",
        source_sha256="abc123",
        patches=[patch],
    )

    restored = PatchPlan.model_validate_json(plan.model_dump_json())

    assert restored.schema_version == 2
    assert restored.patches[0].atomic_group == "layout-group-1"
    assert restored.patches[0].risk == PatchRisk.LAYOUT_REVIEW
    assert restored.patches[0].precondition.layout_fingerprint == "layout-fingerprint"


def test_legacy_sheet_snapshot_payload_loads_with_layout_defaults() -> None:
    snapshot = SheetSnapshot.model_validate(
        {
            "name": "Sheet1",
            "index": 0,
            "state": "visible",
            "max_row": 1,
            "max_column": 1,
        }
    )

    assert snapshot.declared_dimension is None
    assert snapshot.content_dimension is None
    assert snapshot.row_heights == {}
    assert snapshot.column_widths == {}
    assert snapshot.view_top_left_cell is None
    assert snapshot.view_zoom_scale is None
