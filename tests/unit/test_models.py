from __future__ import annotations

import pytest
from pydantic import ValidationError

from workbooklens.models import (
    Confidence,
    PatchDerivation,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchResult,
    PatchRisk,
    RecalculationProvider,
    Severity,
    SheetSnapshot,
    ValidationStatus,
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
    prerequisite_patch_ids: list[str] | None = None,
    confidence: float = 0.99,
    derivation: PatchDerivation | None = None,
) -> PatchOperation:
    payload: dict[str, object] = {
        "id": patch_id,
        "kind": kind,
        "sheet": "Sheet1",
        "cell": "A1",
        "before": None,
        "after": "=1+1",
        "confidence": confidence,
        "safe": safe,
        "description": "test patch",
        "precondition": PatchPrecondition(
            cell_fingerprint="cell-fingerprint",
            layout_fingerprint="layout-fingerprint",
        ),
        "atomic_group": atomic_group,
        "prerequisite_patch_ids": prerequisite_patch_ids or [],
    }
    if risk is not None:
        payload["risk"] = risk
    if derivation is not None:
        payload["derivation"] = derivation
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


def test_patch_plan_schema_three_round_trip_preserves_layout_metadata() -> None:
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

    assert restored.schema_version == 3
    assert restored.patches[0].atomic_group == "layout-group-1"
    assert restored.patches[0].risk == PatchRisk.LAYOUT_REVIEW
    assert restored.patches[0].precondition.layout_fingerprint == "layout-fingerprint"


def test_patch_plan_migrates_schema_two_in_memory_and_only_serializes_three() -> None:
    patch_payload = _patch(kind=PatchKind.SET_NUMERIC).model_dump(mode="json")
    patch_payload.pop("derivation")

    plan = PatchPlan.model_validate(
        {
            "schema_version": 2,
            "tool_version": "2.3.0",
            "source_name": "legacy.xlsx",
            "source_sha256": "abc123",
            "patches": [patch_payload],
        }
    )

    assert plan.schema_version == 3
    assert plan.patches[0].derivation.strategy == "deterministic_rule"
    assert plan.model_dump(mode="json")["schema_version"] == 3


def test_formula_derived_patch_requires_unique_independent_recalculated_evidence() -> None:
    patch = _patch(
        safe=False,
        risk=PatchRisk.FORMULA_DERIVED,
        derivation=PatchDerivation(
            strategy="bidirectional_r1c1_consensus",
            sources=["formula_above:Sheet1!A1", "formula_below:Sheet1!A3"],
            candidate_count=1,
            invariants=["same_formula_region", "no_external_reference"],
            requires_recalculation=True,
        ),
    )

    assert patch.risk == PatchRisk.FORMULA_DERIVED
    assert patch.derivation.source_families == ("formula_above", "formula_below")
    assert not patch.safe_only_eligible


@pytest.mark.parametrize(
    ("confidence", "sources", "candidate_count", "requires_recalculation", "message"),
    [
        (
            0.98,
            ["formula_above:Sheet1!A1", "formula_below:Sheet1!A3"],
            1,
            True,
            "confidence >= 0.99",
        ),
        (0.99, ["formula_above:Sheet1!A1"], 1, True, "two independent evidence"),
        (
            0.99,
            ["formula_above:Sheet1!A1", "formula_below:Sheet1!A3"],
            2,
            True,
            "exactly one candidate",
        ),
        (
            0.99,
            ["formula_above:Sheet1!A1", "formula_below:Sheet1!A3"],
            1,
            False,
            "require recalculation",
        ),
        (
            0.99,
            ["peer_formula:Sheet1!A1", "peer_formula:Sheet1!A3"],
            1,
            True,
            "independent evidence families",
        ),
        (
            0.99,
            ["formula_above", "formula_below"],
            1,
            True,
            "structured family:detail",
        ),
        (
            0.99,
            ["formula_above::", "formula_below:___"],
            1,
            True,
            "structured family:detail",
        ),
    ],
)
def test_formula_derived_patch_rejects_incomplete_proof(
    confidence: float,
    sources: list[str],
    candidate_count: int,
    requires_recalculation: bool,
    message: str,
) -> None:
    with pytest.raises(ValidationError, match=message):
        _patch(
            safe=False,
            risk=PatchRisk.FORMULA_DERIVED,
            confidence=confidence,
            derivation=PatchDerivation(
                strategy="test",
                sources=sources,
                candidate_count=candidate_count,
                invariants=["same_region"],
                requires_recalculation=requires_recalculation,
            ),
        )


def test_formula_evidence_detail_allows_extra_colons_and_survives_json_roundtrip() -> None:
    patch = _patch(
        safe=False,
        risk=PatchRisk.FORMULA_DERIVED,
        derivation=PatchDerivation(
            strategy="independent_structured_evidence",
            sources=["formula_above::extra", "formula_below:Sheet1!A3:peer"],
            candidate_count=1,
            invariants=["same_region"],
            requires_recalculation=True,
        ),
    )

    restored = PatchOperation.model_validate_json(patch.model_dump_json())

    assert restored.derivation.source_families == ("formula_above", "formula_below")
    assert restored.derivation.sources == [
        "formula_above::extra",
        "formula_below:Sheet1!A3:peer",
    ]


def test_semantic_review_patch_can_never_claim_safe_status() -> None:
    with pytest.raises(ValidationError, match="semantic_review patches require safe=false"):
        _patch(safe=True, risk=PatchRisk.SEMANTIC_REVIEW)


def test_patch_dependencies_reject_self_reference_duplicates_missing_and_cycles() -> None:
    with pytest.raises(ValidationError, match="cannot depend on itself"):
        _patch(prerequisite_patch_ids=["patch-1"])
    with pytest.raises(ValidationError, match="must be distinct"):
        _patch(prerequisite_patch_ids=["patch-2", "patch-2"])

    with pytest.raises(ValidationError, match="missing prerequisite"):
        PatchPlan(
            tool_version="2.4.0",
            source_name="input.xlsx",
            source_sha256="abc",
            patches=[_patch(prerequisite_patch_ids=["missing"])],
        )

    with pytest.raises(ValidationError, match="contains a cycle"):
        PatchPlan(
            tool_version="2.4.0",
            source_name="input.xlsx",
            source_sha256="abc",
            patches=[
                _patch(patch_id="patch-1", prerequisite_patch_ids=["patch-2"]),
                _patch(patch_id="patch-2", prerequisite_patch_ids=["patch-1"]),
            ],
        )


def test_patch_result_v3_fields_have_backward_compatible_defaults() -> None:
    result = PatchResult(
        source_sha256="source",
        output_sha256="output",
        output_path="output.xlsx",
        applied_patch_ids=[],
        package_changes=[],
    )

    assert result.recalculation_provider == RecalculationProvider.NONE
    assert result.formula_errors_before == []
    assert result.formula_errors_after == []
    assert result.downgraded_patch_ids == []
    assert result.skipped_patch_ids == []
    assert result.validation_status == ValidationStatus.NOT_RUN
    assert not result.rollback_performed


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
