from __future__ import annotations

import json
from pathlib import Path

import pytest

from workbooklens.exceptions import UsageError
from workbooklens.models import (
    PatchDerivation,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
)
from workbooklens.repair.planning import (
    load_patch_plan,
    resolve_patch_selection,
    write_patch_plan,
)


def _patch(
    patch_id: str,
    *,
    kind: PatchKind = PatchKind.SET_FORMULA,
    safe: bool = True,
    risk: PatchRisk = PatchRisk.SAFE,
    confidence: float = 0.99,
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
    derivation: PatchDerivation | None = None,
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
        prerequisite_patch_ids=prerequisite_patch_ids or [],
        derivation=derivation or PatchDerivation(),
    )


def _plan(*patches: PatchOperation) -> PatchPlan:
    return PatchPlan(
        tool_version="2.1.0",
        source_name="input.xlsx",
        source_sha256="abc123",
        patches=list(patches),
    )


def _layout_patch(
    patch_id: str,
    *,
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
) -> PatchOperation:
    return _patch(
        patch_id,
        kind=PatchKind.SET_COLUMN_WIDTH,
        safe=False,
        risk=PatchRisk.LAYOUT_REVIEW,
        atomic_group=atomic_group,
        prerequisite_patch_ids=prerequisite_patch_ids,
    )


def _formula_derived_patch(
    patch_id: str,
    *,
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
) -> PatchOperation:
    return _patch(
        patch_id,
        safe=False,
        risk=PatchRisk.FORMULA_DERIVED,
        atomic_group=atomic_group,
        prerequisite_patch_ids=prerequisite_patch_ids,
        derivation=PatchDerivation(
            strategy="bidirectional_r1c1_consensus",
            sources=["formula_above:Sheet1!A1", "formula_below:Sheet1!A3"],
            candidate_count=1,
            invariants=["same_formula_region"],
            requires_recalculation=True,
        ),
    )


def _semantic_patch(
    patch_id: str,
    *,
    confidence: float = 0.99,
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
) -> PatchOperation:
    return _patch(
        patch_id,
        kind=PatchKind.SET_NUMERIC,
        safe=False,
        risk=PatchRisk.SEMANTIC_REVIEW,
        confidence=confidence,
        atomic_group=atomic_group,
        prerequisite_patch_ids=prerequisite_patch_ids,
    )


def test_explicit_selection_expands_the_entire_atomic_group() -> None:
    plan = _plan(
        _patch("p1", atomic_group="g1"),
        _patch("p2", atomic_group="g1"),
        _patch("p3"),
    )

    assert resolve_patch_selection(plan, ["p1"]) == {"p1", "p2"}


def test_formula_selection_recursively_expands_safe_prerequisites_and_their_group() -> None:
    plan = _plan(
        _patch("p1", kind=PatchKind.SET_NUMERIC, atomic_group="normalize"),
        _patch("p2", kind=PatchKind.NORMALIZE_TEXT, atomic_group="normalize"),
        _formula_derived_patch("p3", prerequisite_patch_ids=["p1"]),
    )

    assert resolve_patch_selection(
        plan,
        ["p3"],
        accept_formula_derived=True,
    ) == {"p1", "p2", "p3"}
    assert resolve_patch_selection(plan, auto_repair=True) == {"p1", "p2", "p3"}


def test_auto_repair_fails_closed_when_formula_prerequisite_needs_review() -> None:
    plan = _plan(
        _semantic_patch("p1"),
        _formula_derived_patch("p2", prerequisite_patch_ids=["p1"]),
    )

    with pytest.raises(UsageError, match="must be explicitly selected"):
        resolve_patch_selection(plan, auto_repair=True)


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


def test_auto_repair_selects_safe_and_formula_derived_but_not_review_risks() -> None:
    plan = _plan(
        _patch("p1"),
        _formula_derived_patch("p2"),
        _layout_patch("p3"),
        _semantic_patch("p4"),
    )

    assert resolve_patch_selection(plan, auto_repair=True) == {"p1", "p2"}


def test_auto_repair_preserves_atomic_groups_fail_closed() -> None:
    plan = _plan(
        _patch("p1", atomic_group="eligible"),
        _formula_derived_patch("p2", atomic_group="eligible"),
        _patch("p3", atomic_group="review"),
        _semantic_patch("p4", atomic_group="review"),
    )

    assert resolve_patch_selection(plan, auto_repair=True) == {"p1", "p2"}


def test_auto_repair_adds_only_explicitly_authorized_semantic_candidates() -> None:
    plan = _plan(
        _patch("safe"),
        _formula_derived_patch("formula"),
        _semantic_patch("semantic"),
        _layout_patch("layout"),
    )

    assert resolve_patch_selection(
        plan,
        ["semantic"],
        auto_repair=True,
        accept_semantic_risk=True,
    ) == {"safe", "formula", "semantic"}


def test_auto_repair_semantic_consent_without_selection_changes_nothing() -> None:
    plan = _plan(_patch("safe"), _formula_derived_patch("formula"), _semantic_patch("semantic"))

    assert resolve_patch_selection(
        plan,
        auto_repair=True,
        accept_semantic_risk=True,
    ) == {"safe", "formula"}


def test_auto_repair_semantic_extras_fail_closed() -> None:
    plan = _plan(
        _patch("safe"),
        _formula_derived_patch("formula"),
        _semantic_patch("semantic"),
        _layout_patch("layout"),
    )

    with pytest.raises(UsageError, match="--accept-semantic-risk"):
        resolve_patch_selection(plan, ["semantic"], auto_repair=True)
    with pytest.raises(UsageError, match="only explicitly selected semantic-review"):
        resolve_patch_selection(
            plan,
            ["layout"],
            auto_repair=True,
            accept_semantic_risk=True,
        )
    with pytest.raises(UsageError, match="only explicitly selected semantic-review"):
        resolve_patch_selection(
            plan,
            ["safe"],
            auto_repair=True,
            accept_semantic_risk=True,
        )


def test_auto_repair_requires_every_semantic_atomic_member_to_be_explicit() -> None:
    plan = _plan(
        _semantic_patch("semantic-a", atomic_group="semantic-group"),
        _semantic_patch("semantic-b", atomic_group="semantic-group"),
        _patch("safe"),
    )

    with pytest.raises(UsageError, match="Every semantic-review patch"):
        resolve_patch_selection(
            plan,
            ["semantic-a"],
            auto_repair=True,
            accept_semantic_risk=True,
        )
    assert resolve_patch_selection(
        plan,
        ["semantic-a", "semantic-b"],
        auto_repair=True,
        accept_semantic_risk=True,
    ) == {"safe", "semantic-a", "semantic-b"}


def test_auto_repair_revalidates_every_expanded_semantic_atomic_member() -> None:
    plan = _plan(
        _semantic_patch("semantic", atomic_group="mixed"),
        _patch("unsafe-safe-risk", safe=False, atomic_group="mixed"),
    )

    with pytest.raises(UsageError, match="safe-only eligible"):
        resolve_patch_selection(
            plan,
            ["semantic"],
            auto_repair=True,
            accept_semantic_risk=True,
        )


def test_auto_repair_rejects_low_confidence_safe_semantic_prerequisite() -> None:
    plan = _plan(
        _patch("low-safe", confidence=0.94),
        _semantic_patch("semantic", prerequisite_patch_ids=["low-safe"]),
    )

    with pytest.raises(UsageError, match="safe-only eligible"):
        resolve_patch_selection(
            plan,
            ["semantic"],
            auto_repair=True,
            accept_semantic_risk=True,
        )


def test_auto_repair_rejects_layout_semantic_prerequisite() -> None:
    plan = _plan(
        _layout_patch("layout"),
        _semantic_patch("semantic", prerequisite_patch_ids=["layout"]),
    )

    with pytest.raises(UsageError, match="layout-review or unsupported unsafe"):
        resolve_patch_selection(
            plan,
            ["semantic"],
            auto_repair=True,
            accept_semantic_risk=True,
        )


def test_auto_repair_accepts_verified_formula_semantic_prerequisite() -> None:
    plan = _plan(
        _formula_derived_patch("formula"),
        _semantic_patch("semantic", prerequisite_patch_ids=["formula"]),
    )

    assert resolve_patch_selection(
        plan,
        ["semantic"],
        auto_repair=True,
        accept_semantic_risk=True,
    ) == {"formula", "semantic"}


def test_auto_repair_accepts_safe_only_eligible_semantic_prerequisite() -> None:
    plan = _plan(
        _patch("safe"),
        _semantic_patch("semantic", prerequisite_patch_ids=["safe"]),
    )

    assert resolve_patch_selection(
        plan,
        ["semantic"],
        auto_repair=True,
        accept_semantic_risk=True,
    ) == {"safe", "semantic"}


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


def test_formula_derived_explicit_selection_requires_recalculation_authorization() -> None:
    plan = _plan(_formula_derived_patch("p1"))

    with pytest.raises(UsageError, match="recalculation authorization"):
        resolve_patch_selection(plan, ["p1"])

    assert resolve_patch_selection(
        plan,
        ["p1"],
        accept_formula_derived=True,
    ) == {"p1"}


def test_semantic_review_requires_explicit_acceptance() -> None:
    plan = _plan(_semantic_patch("p1"))

    with pytest.raises(UsageError, match="--accept-semantic-risk"):
        resolve_patch_selection(plan, ["p1"])

    assert resolve_patch_selection(
        plan,
        ["p1"],
        accept_semantic_risk=True,
    ) == {"p1"}


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
    with pytest.raises(UsageError, match="only explicitly selected semantic-review"):
        resolve_patch_selection(plan, ["p1"], auto_repair=True)
    with pytest.raises(UsageError, match="never accepts layout"):
        resolve_patch_selection(plan, auto_repair=True, accept_layout_risk=True)


def test_selection_rejects_duplicate_plan_ids() -> None:
    plan = _plan(_patch("p1"), _patch("p2"))
    plan.patches[1].id = "p1"

    with pytest.raises(UsageError, match="duplicate patch IDs"):
        resolve_patch_selection(plan, safe_only=True)


def test_safe_only_rejects_an_empty_eligible_selection() -> None:
    plan = _plan(_layout_patch("p1"))

    with pytest.raises(UsageError, match="No eligible patches"):
        resolve_patch_selection(plan, safe_only=True)


def test_load_patch_plan_migrates_v2_and_write_only_emits_v3(tmp_path: Path) -> None:
    legacy_path = tmp_path / "schema-two.json"
    legacy_payload = _plan(_patch("p1")).model_dump(mode="json")
    legacy_payload["schema_version"] = 2
    legacy_payload["patches"][0].pop("derivation")
    legacy_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    plan = load_patch_plan(legacy_path)
    output_path = tmp_path / "schema-three.json"
    write_patch_plan(output_path, plan)

    assert plan.schema_version == 3
    assert json.loads(output_path.read_text(encoding="utf-8"))["schema_version"] == 3


def test_v2_formula_patch_migrates_to_unverified_review_candidate(tmp_path: Path) -> None:
    legacy_path = tmp_path / "schema-two-formula.json"
    legacy_payload = _plan(_patch("formula")).model_dump(mode="json")
    legacy_payload["schema_version"] = 2
    legacy_payload["patches"][0].pop("derivation")
    legacy_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    plan = load_patch_plan(legacy_path)
    patch = plan.patches[0]

    assert patch.risk == PatchRisk.SEMANTIC_REVIEW
    assert not patch.safe
    assert patch.derivation.strategy == "legacy_schema_v2_unverified_formula"
    assert patch.derivation.sources == ["legacy_plan:schema_v2"]
    assert patch.derivation.requires_recalculation


@pytest.mark.parametrize(
    ("kwargs", "message"),
    [
        ({"safe_only": True}, "No eligible patches"),
        ({"auto_repair": True}, "No eligible patches"),
        (
            {"selected_ids": ["formula"], "accept_semantic_risk": True},
            "Legacy schema v2 formula patches are unverified",
        ),
    ],
)
def test_v2_formula_patch_can_never_be_selected(
    tmp_path: Path,
    kwargs: dict[str, object],
    message: str,
) -> None:
    legacy_path = tmp_path / "schema-two-formula.json"
    legacy_payload = _plan(_patch("formula")).model_dump(mode="json")
    legacy_payload["schema_version"] = 2
    legacy_payload["patches"][0].pop("derivation")
    legacy_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    with pytest.raises(UsageError, match=message):
        resolve_patch_selection(load_patch_plan(legacy_path), **kwargs)


def test_v2_formula_atomic_group_cannot_enter_through_safe_sibling(tmp_path: Path) -> None:
    legacy_path = tmp_path / "schema-two-atomic.json"
    legacy_payload = _plan(
        _patch("safe", kind=PatchKind.SET_NUMERIC, atomic_group="mixed"),
        _patch("formula", atomic_group="mixed"),
    ).model_dump(mode="json")
    legacy_payload["schema_version"] = 2
    for patch in legacy_payload["patches"]:
        patch.pop("derivation")
    legacy_path.write_text(json.dumps(legacy_payload), encoding="utf-8")

    plan = load_patch_plan(legacy_path)

    with pytest.raises(UsageError, match="No eligible patches"):
        resolve_patch_selection(plan, safe_only=True)
    with pytest.raises(UsageError, match="Legacy schema v2 formula patches are unverified"):
        resolve_patch_selection(plan, ["safe"], accept_semantic_risk=True)


def test_load_patch_plan_clearly_rejects_schema_one(tmp_path: Path) -> None:
    path = tmp_path / "schema-one.json"
    path.write_text(json.dumps({"schema_version": 1}), encoding="utf-8")

    with pytest.raises(UsageError, match="schema version 2 or 3"):
        load_patch_plan(path)
