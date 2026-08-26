from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
from typing import Any, Literal

import pytest

import workbooklens.repair.engine as engine
import workbooklens.repair.recalculation as recalculation
from workbooklens import conversion
from workbooklens.exceptions import PatchValidationError, StalePlanError, UsageError
from workbooklens.models import (
    CellSnapshot,
    Evidence,
    Finding,
    PatchDerivation,
    PatchKind,
    PatchOperation,
    PatchPlan,
    PatchPrecondition,
    PatchRisk,
    RecalculationProvider,
    Severity,
    SheetSnapshot,
    ValidationStatus,
    WorkbookSnapshot,
)
from workbooklens.repair.recalculation import RecalculationValidation
from workbooklens.utils import sha256_file


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


def _safe_patch(
    patch_id: str = "safe",
    *,
    cell: str = "A2",
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
) -> PatchOperation:
    return PatchOperation(
        id=patch_id,
        kind=PatchKind.NORMALIZE_TEXT,
        sheet="Sheet1",
        cell=cell,
        before=" value ",
        after="value",
        confidence=1.0,
        safe=True,
        description="lossless text normalization",
        precondition=PatchPrecondition(cell_fingerprint=f"fingerprint-{patch_id}"),
        atomic_group=atomic_group,
        prerequisite_patch_ids=prerequisite_patch_ids or [],
    )


def _formula_patch(
    patch_id: str = "formula",
    *,
    cell: str = "B2",
    atomic_group: str | None = None,
    prerequisite_patch_ids: list[str] | None = None,
) -> PatchOperation:
    return PatchOperation(
        id=patch_id,
        kind=PatchKind.SET_FORMULA,
        sheet="Sheet1",
        cell=cell,
        before="=A2/0",
        after="=A2",
        confidence=0.99,
        safe=False,
        risk=PatchRisk.FORMULA_DERIVED,
        description="unique formula reconstruction",
        precondition=PatchPrecondition(
            cell_fingerprint=f"fingerprint-{patch_id}",
            expected_formula="=A2/0",
        ),
        atomic_group=atomic_group,
        prerequisite_patch_ids=prerequisite_patch_ids or [],
        derivation=PatchDerivation(
            strategy="bidirectional_r1c1_consensus",
            sources=["formula_above:Sheet1!B1", "formula_below:Sheet1!B3"],
            candidate_count=1,
            invariants=["same_formula_band", "ordinary_a1_references"],
            requires_recalculation=True,
        ),
    )


def _semantic_numeric_patch() -> PatchOperation:
    return PatchOperation(
        id="semantic",
        kind=PatchKind.SET_NUMERIC,
        sheet="Sheet1",
        cell="A2",
        before="八万九千",
        after=89000,
        confidence=0.99,
        safe=False,
        risk=PatchRisk.SEMANTIC_REVIEW,
        description="reviewed Chinese numeric normalization",
        precondition=PatchPrecondition(
            cell_fingerprint="fingerprint-semantic",
            expected_value="八万九千",
        ),
        derivation=PatchDerivation(
            strategy="explicit_profile_chinese_number",
            sources=["profile_numeric_role", "unique_chinese_number_parse"],
            candidate_count=1,
            invariants=["source_precondition", "column_role_preserved"],
            requires_recalculation=True,
        ),
    )


def _formula_snapshot() -> WorkbookSnapshot:
    return WorkbookSnapshot(
        source_name="input.xlsx",
        source_sha256="source-hash",
        format="xlsx",
        sheets=[
            SheetSnapshot(
                name="Sheet1",
                index=0,
                state="visible",
                max_row=2,
                max_column=3,
                cells={
                    "A2": CellSnapshot(coordinate="A2", value=5, data_type="n"),
                    "B2": CellSnapshot(
                        coordinate="B2",
                        formula="=A2/0",
                        data_type="f",
                    ),
                    "C2": CellSnapshot(
                        coordinate="C2",
                        formula="=B2+1",
                        data_type="f",
                    ),
                },
            )
        ],
    )


def _dependency_snapshot(
    sheet_names: list[str],
    formula: str,
) -> WorkbookSnapshot:
    sheets = [
        SheetSnapshot(
            name=name,
            index=index,
            state="visible",
            max_row=1,
            max_column=1,
            cells={"A1": CellSnapshot(coordinate="A1", value=index + 1, data_type="n")},
        )
        for index, name in enumerate(sheet_names)
    ]
    sheets.append(
        SheetSnapshot(
            name="Calc",
            index=len(sheets),
            state="visible",
            max_row=1,
            max_column=2,
            cells={"B1": CellSnapshot(coordinate="B1", formula=formula, data_type="f")},
        )
    )
    return WorkbookSnapshot(
        source_name="input.xlsx",
        source_sha256="source-hash",
        format="xlsx",
        sheets=sheets,
    )


def _causal_snapshot(formulas: dict[str, str]) -> WorkbookSnapshot:
    cells = {
        coordinate: CellSnapshot(coordinate=coordinate, formula=formula, data_type="f")
        for coordinate, formula in formulas.items()
    }
    return WorkbookSnapshot(
        source_name="input.xlsx",
        source_sha256="source-hash",
        format="xlsx",
        sheets=[
            SheetSnapshot(
                name="Sheet1",
                index=0,
                state="visible",
                max_row=100,
                max_column=100,
                cells=cells,
            )
        ],
    )


def _formula_group_for(*patches: PatchOperation) -> engine._FormulaRepairGroup:
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=list(patches),
    )
    groups = engine._formula_repair_groups(plan, {patch.id for patch in patches})
    assert len(groups) == 1
    return groups[0]


def _scan(
    *,
    findings: list[Finding] | None = None,
    patches: list[PatchOperation] | None = None,
    snapshot: WorkbookSnapshot | None = None,
) -> SimpleNamespace:
    return SimpleNamespace(
        findings=findings or [],
        patches=patches or [],
        snapshot=snapshot or _formula_snapshot(),
    )


def test_formula_dependency_exact_unicode_sheet_names_do_not_collide() -> None:
    snapshot = _dependency_snapshot(
        ["Straße", "STRASSE"],
        "='Straße'!A1+'STRASSE'!A1",
    )

    assert engine._opaque_formula_dependencies(snapshot, {("Straße", "A1")}) == []
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("Straße", "A1")},
    )
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("STRASSE", "A1")},
    )


def test_formula_dependency_exact_case_distinct_sheet_names_take_priority() -> None:
    snapshot = _dependency_snapshot(
        ["SheetA", "sheeta"],
        "=SheetA!A1+sheeta!A1",
    )

    assert engine._opaque_formula_dependencies(snapshot, {("SheetA", "A1")}) == []
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("SheetA", "A1")},
    )
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("sheeta", "A1")},
    )


def test_formula_dependency_unique_case_compatible_sheet_name_resolves() -> None:
    snapshot = _dependency_snapshot(["SheetA"], "=sheeta!A1")

    assert engine._opaque_formula_dependencies(snapshot, {("SHEETA", "A1")}) == []
    assert engine._formula_dependency_closure(snapshot, {("SHEETA", "A1")}) == {
        ("SheetA", "A1"),
        ("Calc", "B1"),
    }


def test_formula_dependency_ambiguous_compatible_sheet_name_fails_closed() -> None:
    snapshot = _dependency_snapshot(["SheetA", "sheeta"], "=SHEETA!A1")

    opaque = engine._opaque_formula_dependencies(snapshot, {("SheetA", "A1")})

    assert any("ambiguous worksheet reference" in item for item in opaque)
    with pytest.raises(engine._OpaqueFormulaDependencyError, match="opaque references"):
        engine._formula_dependency_closure(snapshot, {("SheetA", "A1")})


def test_formula_dependency_unknown_sheet_name_fails_closed() -> None:
    snapshot = _dependency_snapshot(["Data"], "=Missing!A1")

    opaque = engine._opaque_formula_dependencies(snapshot, {("Data", "A1")})

    assert any("unavailable worksheet 'Missing'" in item for item in opaque)
    with pytest.raises(engine._OpaqueFormulaDependencyError, match="opaque references"):
        engine._formula_dependency_closure(snapshot, {("Data", "A1")})


def test_formula_dependency_supports_quoted_and_unquoted_sheet_references() -> None:
    snapshot = _dependency_snapshot(
        ["Data One", "Data"],
        "='Data One'!A1+Data!A1",
    )

    assert engine._opaque_formula_dependencies(snapshot, {("Data One", "A1")}) == []
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("Data One", "A1")},
    )
    assert ("Calc", "B1") in engine._formula_dependency_closure(
        snapshot,
        {("Data", "A1")},
    )


@pytest.mark.parametrize("formula", ["=A1(", "=SUM(A1:A2", "=A1+", "=A1)"])
def test_formula_dependency_malformed_formula_fails_closed(formula: str) -> None:
    snapshot = _dependency_snapshot(["Data"], formula)

    opaque = engine._opaque_formula_dependencies(snapshot, {("Data", "A1")})

    assert any(
        "malformed formula structure" in item or "formula tokenizer rejected expression" in item
        for item in opaque
    )
    with pytest.raises(engine._OpaqueFormulaDependencyError, match="opaque references"):
        engine._formula_dependency_closure(snapshot, {("Data", "A1")})


def test_formula_group_retains_target_and_downstream_errors_from_shared_old_root() -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    before = _causal_snapshot({"A2": "=1/0", "B2": "=A2+0", "C2": "=A2+B2"})
    after = _causal_snapshot({"A2": "=1/0", "B2": "=A2", "C2": "=A2+B2"})

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        before,
        after,
        [],
        [],
        ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"],
        ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"],
        {("Sheet1", "B2")},
    )

    assert validation.failures == {}
    assert any("Sheet1!A2:#DIV/0!" in message for message in validation.retained_messages)


def test_formula_group_retains_b8_b9_target_errors_from_distinct_shared_old_roots() -> None:
    first = _formula_patch("formula-b8", cell="B8", atomic_group="dashboard-group")
    second = _formula_patch("formula-b9", cell="B9", atomic_group="dashboard-group")
    group = _formula_group_for(first, second)
    before = _causal_snapshot({"H12": "=1/0", "H16": "=1/0", "B8": "=H12+0", "B9": "=H16+0"})
    after = _causal_snapshot({"H12": "=1/0", "H16": "=1/0", "B8": "=H12", "B9": "=H16"})
    errors = [
        "Sheet1!H12:#VALUE!",
        "Sheet1!H16:#VALUE!",
        "Sheet1!B8:#VALUE!",
        "Sheet1!B9:#VALUE!",
    ]

    validation = engine._validate_formula_groups(
        [group],
        {first.id, second.id},
        before,
        after,
        [],
        [],
        errors,
        errors,
        {("Sheet1", "B8"), ("Sheet1", "B9")},
    )

    assert validation.failures == {}
    proof = " ".join(validation.retained_messages)
    assert "Sheet1!B8:#VALUE! <= Sheet1!H12:#VALUE!" in proof
    assert "Sheet1!B9:#VALUE! <= Sheet1!H16:#VALUE!" in proof


@pytest.mark.parametrize("function", ["SUM", "AVERAGE"])
def test_formula_group_retains_error_through_transparent_aggregate(function: str) -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    snapshot = _causal_snapshot({"A2": "=1/0", "B2": "=1", "C2": f"={function}(A2:B2)"})
    errors = ["Sheet1!A2:#DIV/0!", "Sheet1!C2:#DIV/0!"]

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        snapshot,
        snapshot,
        [],
        [],
        errors,
        errors,
        {("Sheet1", "B2")},
    )

    assert validation.failures == {}
    assert any("Sheet1!A2:#DIV/0!" in message for message in validation.retained_messages)


@pytest.mark.parametrize(
    "formula",
    [
        "=IF(FALSE,A2,B2/0)",
        "=IFERROR(A2,B2/0)",
    ],
)
def test_formula_group_rejects_nontransparent_conditional_error_path(formula: str) -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    snapshot = _causal_snapshot({"A2": "=1/0", "B2": "=1", "C2": formula})
    errors = ["Sheet1!A2:#DIV/0!", "Sheet1!C2:#DIV/0!"]

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        snapshot,
        snapshot,
        [],
        [],
        errors,
        errors,
        {("Sheet1", "B2")},
    )

    assert group.label in validation.failures
    assert "Sheet1!C2:#DIV/0!" in validation.failures[group.label][1]


def test_formula_group_rejects_root_with_different_error_code() -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    snapshot = _causal_snapshot({"A2": "=NA()", "B2": "=1", "C2": "=A2+B2"})
    errors = ["Sheet1!A2:#N/A", "Sheet1!C2:#DIV/0!"]

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        snapshot,
        snapshot,
        [],
        [],
        errors,
        errors,
        {("Sheet1", "B2")},
    )

    assert group.label in validation.failures
    assert "Sheet1!C2:#DIV/0!" in validation.failures[group.label][1]


@pytest.mark.parametrize(
    ("after_formulas", "modified_nodes", "after_errors"),
    [
        (
            {"A2": "=1/0", "B2": "=1/0", "C2": "=B2+1"},
            {("Sheet1", "B2")},
            ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"],
        ),
        (
            {"A2": "=1/0", "D2": "=A2", "B2": "=D2", "C2": "=B2+1"},
            {("Sheet1", "B2"), ("Sheet1", "D2")},
            ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"],
        ),
        (
            {"A2": "=1/0", "B2": "=A2", "C2": "=B2+1"},
            {("Sheet1", "A2"), ("Sheet1", "B2")},
            ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"],
        ),
        (
            {"A2": "=1/0", "B2": "=A2", "C2": "=B2+1"},
            {("Sheet1", "B2")},
            ["Sheet1!A2:#DIV/0!", "Sheet1!B2:#VALUE!", "Sheet1!C2:#DIV/0!"],
        ),
    ],
)
def test_formula_group_rejects_unproven_or_changed_residual_errors(
    after_formulas: dict[str, str],
    modified_nodes: set[tuple[str, str]],
    after_errors: list[str],
) -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    before = _causal_snapshot({"A2": "=1/0", "D2": "=A2", "B2": "=A2+0", "C2": "=B2+1"})
    before_errors = [
        "Sheet1!A2:#DIV/0!",
        "Sheet1!B2:#DIV/0!",
        "Sheet1!C2:#DIV/0!",
    ]

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        before,
        _causal_snapshot(after_formulas),
        [],
        [],
        before_errors,
        after_errors,
        modified_nodes,
    )

    assert group.label in validation.failures


def test_formula_group_rejects_when_target_finding_remains() -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    before_finding = _finding("before", Severity.ERROR, patch_ids=[patch.id], location="B2")
    after_finding = _finding("after", Severity.ERROR, patch_ids=[patch.id], location="B2")
    snapshot = _causal_snapshot({"B2": "=1"})

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        snapshot,
        snapshot,
        [before_finding],
        [after_finding],
        [],
        [],
        {("Sheet1", "B2")},
    )

    assert group.label in validation.failures
    assert any(
        "finding:WL_TEST:Sheet1!B2" in reason for reason in validation.failures[group.label][1]
    )


def test_formula_group_rejects_opaque_candidate_formula() -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        _causal_snapshot({"B2": "=A2"}),
        _causal_snapshot({"B2": '=INDIRECT("A2")'}),
        [],
        [],
        [],
        [],
        {("Sheet1", "B2")},
    )

    assert any("dependency-proof" in reason for reason in validation.failures[group.label][1])


def test_formula_group_dependency_budget_exhaustion_fails_closed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    patch = _formula_patch(cell="B2", atomic_group="formula-group")
    group = _formula_group_for(patch)
    snapshot = _causal_snapshot({"B2": "=A2", "C2": "=SUM(B2:B3)"})
    monkeypatch.setattr(engine, "_MAX_DEPENDENCY_RANGE_CHECKS", 0)

    validation = engine._validate_formula_groups(
        [group],
        {patch.id},
        snapshot,
        snapshot,
        [],
        [],
        [],
        [],
        {("Sheet1", "B2")},
    )

    assert any("range-check budget" in reason for reason in validation.failures[group.label][1])


def _fake_low_level(candidate: Path, *, formula_changed: bool = True) -> SimpleNamespace:
    return SimpleNamespace(
        source_sha256="source-hash",
        output_sha256=sha256_file(candidate),
        changes=[],
        formula_changed=formula_changed,
    )


def _finding(
    identifier: str,
    severity: Severity,
    *,
    patch_ids: list[str] | None = None,
    location: str = "A2",
) -> Finding:
    return Finding(
        id=identifier,
        rule_id="WL_TEST",
        title="test finding",
        explanation="test",
        severity=severity,
        confidence=0.99,
        workbook="input.xlsx",
        sheet="Sheet1",
        location=location,
        evidence=Evidence(summary="test"),
        expected="test",
        suggested_action="test",
        patch_ids=patch_ids or [],
    )


def test_grouped_finding_subset_is_remaining_not_new() -> None:
    before = _finding("before-group", Severity.WARNING, location="F15,F22")
    after = _finding("after-single", Severity.WARNING, location="F15")

    changes = engine._compare_findings([before], [after])

    assert changes.resolved_ids == set()
    assert changes.remaining_ids == {after.id}
    assert changes.new_ids == set()
    assert changes.dangerous_findings == []


def test_grouped_finding_merge_is_remaining_not_new() -> None:
    before_first = _finding("before-f15", Severity.WARNING, location="F15")
    before_second = _finding("before-f22", Severity.WARNING, location="F22")
    after = _finding("after-group", Severity.WARNING, location="F15,F22")

    changes = engine._compare_findings([before_first, before_second], [after])

    assert changes.resolved_ids == set()
    assert changes.remaining_ids == {after.id}
    assert changes.new_ids == set()


def test_grouped_finding_expansion_preserves_true_new_member() -> None:
    before = _finding("before-group", Severity.WARNING, location="F15,F22")
    after = _finding("after-group", Severity.WARNING, location="F15,F30")

    changes = engine._compare_findings([before], [after])

    assert changes.resolved_ids == set()
    assert changes.remaining_ids == set()
    assert changes.new_ids == {after.id}


def test_grouped_finding_subset_severity_escalation_is_dangerous() -> None:
    before = _finding("before-group", Severity.WARNING, location="F15,F22")
    after = _finding("after-single", Severity.ERROR, location="F15")

    changes = engine._compare_findings([before], [after])

    assert changes.dangerous_findings == [after]


def test_grouped_finding_uses_highest_predecessor_severity_per_cell() -> None:
    before_warning = _finding("before-warning", Severity.WARNING, location="F15,F22")
    before_error = _finding("before-error", Severity.ERROR, location="F15")
    after = _finding("after-error", Severity.ERROR, location="F15")

    changes = engine._compare_findings([before_warning, before_error], [after])

    assert changes.dangerous_findings == []


def test_non_grouped_location_falls_back_to_exact_finding_id() -> None:
    before = _finding("before-id", Severity.WARNING, location="A1")
    after = _finding("after-id", Severity.WARNING, location="A1")

    changes = engine._compare_findings([before], [after])

    assert changes.resolved_ids == {before.id}
    assert changes.remaining_ids == set()
    assert changes.new_ids == {after.id}


def test_apply_resolves_atomic_group_before_calling_low_level_patcher(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = _layout_plan()
    scan = SimpleNamespace(findings=[], patches=[])
    captured: dict[str, Any] = {}

    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        captured.update(kwargs)
        captured["candidate"] = args[2]
        assert args[2] != tmp_path / "output.xlsx"
        assert not (tmp_path / "output.xlsx").exists()
        selected_ids = kwargs["selected_ids"]
        selected = [patch for patch in plan.patches if patch.id in selected_ids]
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2], formula_changed=False), selected

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
    assert (tmp_path / "output.xlsx").read_bytes() == b"candidate"
    assert not captured["candidate"].exists()


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


def test_apply_rejects_partially_remaining_grouped_target(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = _layout_plan()
    before = _finding(
        "before-group",
        Severity.WARNING,
        location="F15,F22",
        patch_ids=["p1", "p2"],
    )
    after = _finding("after-single", Severity.WARNING, location="F15")
    scans = iter(
        [
            _scan(findings=[before]),
            _scan(findings=[after]),
        ]
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2], formula_changed=False), plan.patches

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


def test_auto_repair_without_provider_applies_safe_and_degrades_formula(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    safe_patch = _safe_patch()
    formula_patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[safe_patch, formula_patch],
    )
    scans = iter([_scan(), _scan()])
    captured: dict[str, Any] = {}
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("untrusted workbooks must not start provider discovery"),
    )

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        captured.update(kwargs)
        args[2].write_bytes(b"normalized")
        selected = [patch for patch in plan.patches if patch.id in kwargs["selected_ids"]]
        return _fake_low_level(args[2], formula_changed=False), selected

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: pytest.fail("a downgraded formula must not be recalculated"),
    )

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        tmp_path / "output.xlsx",
        auto_repair=True,
        recalc_provider="auto",
    )

    assert captured["selected_ids"] == {safe_patch.id}
    assert result.applied_patch_ids == [safe_patch.id]
    assert result.downgraded_patch_ids == [formula_patch.id]
    assert result.skipped_patch_ids == [formula_patch.id]
    assert result.validation_status == ValidationStatus.DEGRADED
    assert result.recalculation_provider == RecalculationProvider.NONE
    assert any("not explicitly trusted" in message for message in result.validation_messages)
    assert not any(
        "recalculation was disabled" in message for message in result.validation_messages
    )


@pytest.mark.parametrize(
    ("recalc_provider", "trust_workbook", "expected_reason", "unexpected_reason"),
    [
        ("none", False, "recalculation was disabled", "not explicitly trusted"),
        ("auto", True, "no requested local recalculation provider", "not explicitly trusted"),
    ],
)
def test_auto_repair_reports_the_direct_recalculation_downgrade_reason(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    recalc_provider: Literal["auto", "none"],
    trust_workbook: bool,
    expected_reason: str,
    unexpected_reason: str,
) -> None:
    safe_patch = _safe_patch()
    formula_patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[safe_patch, formula_patch],
    )
    scans = iter([_scan(), _scan()])
    provider_preferences: list[str] = []
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def no_providers(preference: str) -> tuple[object, ...]:
        provider_preferences.append(preference)
        return ()

    monkeypatch.setattr(engine, "candidate_providers", no_providers)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"normalized")
        selected = [patch for patch in plan.patches if patch.id in kwargs["selected_ids"]]
        return _fake_low_level(args[2], formula_changed=False), selected

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: pytest.fail("a downgraded formula must not be recalculated"),
    )

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        tmp_path / "output.xlsx",
        auto_repair=True,
        recalc_provider=recalc_provider,
        trust_workbook_for_recalculation=trust_workbook,
    )

    assert any(expected_reason in message for message in result.validation_messages)
    assert not any(unexpected_reason in message for message in result.validation_messages)
    assert provider_preferences == ([] if recalc_provider == "none" else ["auto"])


def test_auto_repair_without_provider_skips_entire_formula_atomic_group(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source workbook bytes")
    formula_patch = _formula_patch(atomic_group="formula-and-format")
    format_patch = _safe_patch(
        patch_id="format",
        cell="B2",
        atomic_group="formula-and-format",
    )
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256=sha256_file(source),
        patches=[formula_patch, format_patch],
    )
    scan = _scan()
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("untrusted workbooks must not start provider discovery"),
    )
    monkeypatch.setattr(
        engine,
        "patch_ooxml_package",
        lambda *args, **kwargs: pytest.fail("a fully downgraded group must not be applied"),
    )
    output = tmp_path / "output.xlsx"

    result = engine.apply_patch_plan(
        source,
        plan,
        output,
        auto_repair=True,
        recalc_provider="auto",
    )

    assert output.read_bytes() == source.read_bytes()
    assert result.applied_patch_ids == []
    assert result.downgraded_patch_ids == [formula_patch.id]
    assert result.skipped_patch_ids == sorted([formula_patch.id, format_patch.id])
    assert result.validation_status == ValidationStatus.DEGRADED


def test_explicit_formula_repair_rejects_missing_installed_provider(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: ())
    output = tmp_path / "output.xlsx"

    with pytest.raises(UsageError, match="no requested provider is installed") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
            recalc_provider="auto",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()
    assert caught.value.patch_result.validation_status == ValidationStatus.FAILED
    assert caught.value.patch_result.recalculation_provider == RecalculationProvider.NONE
    assert caught.value.patch_result.rollback_performed is True
    assert caught.value.patch_result.failure_report_path is not None


def test_semantic_repair_requiring_recalculation_is_rejected_without_provider(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _semantic_numeric_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("untrusted workbooks must not start provider discovery"),
    )
    output = tmp_path / "output.xlsx"

    with pytest.raises(UsageError, match="require a local recalculation provider"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
            accept_semantic_risk=True,
            recalc_provider="auto",
        )

    assert not output.exists()


def test_formula_repair_records_provider_and_clears_dependency_errors(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    scans = iter([_scan(), _scan()])
    captured: dict[str, Any] = {}
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def fake_candidates(preference: str) -> tuple[conversion.ConversionProvider, ...]:
        captured["discovery_preference"] = preference
        return (provider,)

    monkeypatch.setattr(engine, "candidate_providers", fake_candidates)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2]), [patch]

    def fake_validate(*args: Any, **kwargs: Any) -> RecalculationValidation:
        captured.update(kwargs)
        return RecalculationValidation(
            provider,
            ("Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"),
            (),
        )

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(engine, "validate_formula_recalculation", fake_validate)

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        tmp_path / "output.xlsx",
        selected_ids=[patch.id],
        recalc_provider="excel",
        trust_workbook_for_recalculation=True,
    )

    assert captured["preference"] == "excel"
    assert captured["discovery_preference"] == "excel"
    assert captured["expected_source_sha256"] == plan.source_sha256
    assert captured["expected_candidate_sha256"] == sha256_file(tmp_path / "output.xlsx")
    assert result.recalculation_provider == RecalculationProvider.EXCEL
    assert result.formula_errors_before == ["Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"]
    assert result.formula_errors_after == []
    assert result.validation_status == ValidationStatus.PASSED


def test_formula_recalculation_entry_pins_each_round_input_hashes(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source workbook bytes")
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256=sha256_file(source),
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    scans = iter([_scan(), _scan()])
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))
    candidate_hashes: list[str] = []

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        candidate = args[2]
        candidate.write_bytes(b"candidate round input")
        low_level = _fake_low_level(candidate)
        candidate_hashes.append(low_level.output_sha256)
        return low_level, [patch]

    captured: dict[str, Any] = {}

    def fake_validate(*args: Any, **kwargs: Any) -> RecalculationValidation:
        captured.update(kwargs)
        assert sha256_file(args[0]) == kwargs["expected_source_sha256"]
        assert sha256_file(args[1]) == kwargs["expected_candidate_sha256"]
        return RecalculationValidation(provider, (), ())

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(engine, "validate_formula_recalculation", fake_validate)

    engine.apply_patch_plan(
        source,
        plan,
        tmp_path / "output.xlsx",
        selected_ids=[patch.id],
        recalc_provider="excel",
        trust_workbook_for_recalculation=True,
    )

    assert captured["expected_source_sha256"] == plan.source_sha256
    assert captured["expected_candidate_sha256"] == candidate_hashes[-1]


@pytest.mark.parametrize("mutation_target", ["source", "candidate"])
def test_formula_recalculation_input_replacement_fails_closed_before_provider(
    mutation_target: Literal["source", "candidate"],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source workbook bytes")
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256=sha256_file(source),
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))
    monkeypatch.setattr(recalculation, "candidate_providers", lambda _preference: (provider,))
    monkeypatch.setattr(
        recalculation,
        "_run_provider",
        lambda *args, **kwargs: pytest.fail("replaced inputs must fail before Excel starts"),
    )

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        candidate = args[2]
        candidate.write_bytes(b"candidate round input")
        return _fake_low_level(candidate), [patch]

    def replace_then_validate(*args: Any, **kwargs: Any) -> RecalculationValidation:
        target = args[0] if mutation_target == "source" else args[1]
        target.write_bytes(b"replacement after patch generation")
        return recalculation.validate_formula_recalculation(*args, **kwargs)

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(engine, "validate_formula_recalculation", replace_then_validate)
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="SHA-256 changed") as caught:
        engine.apply_patch_plan(
            source,
            plan,
            output,
            selected_ids=[patch.id],
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()
    assert caught.value.patch_result.validation_status == ValidationStatus.FAILED
    assert caught.value.patch_result.rollback_performed


def test_formula_repair_rolls_back_when_downstream_dependency_error_remains(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2]), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: RecalculationValidation(
            provider,
            ("Sheet1!B2:#DIV/0!", "Sheet1!C2:#DIV/0!"),
            ("Sheet1!C2:#DIV/0!",),
        ),
    )
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="target dependency chain"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()


def test_formula_repair_rolls_back_when_recalculation_rejects_new_error(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2]), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)

    def reject_recalculation(*args: Any, **kwargs: Any) -> RecalculationValidation:
        error = PatchValidationError("Formula recalculation introduced new errors")
        error.__dict__["recalculation_provider"] = "excel"
        raise error

    monkeypatch.setattr(engine, "validate_formula_recalculation", reject_recalculation)
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="introduced new errors") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()
    assert caught.value.patch_result.recalculation_provider == RecalculationProvider.EXCEL
    assert caught.value.patch_result.failure_report_path is not None
    report_text = Path(caught.value.patch_result.failure_report_path).read_text(encoding="utf-8")
    assert "Recalculation provider / 重算提供者: excel" in report_text


def test_formula_repair_rolls_back_when_same_action_is_proposed_again(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    scans = iter([_scan(), _scan(patches=[patch])])
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2]), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: RecalculationValidation(provider, (), ()),
    )
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="not idempotent"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()


@pytest.mark.parametrize("same_path", [False, True])
def test_degraded_copy_never_overwrites_existing_or_source(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    same_path: bool,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source workbook bytes")
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256="source-hash",
        patches=[patch],
    )
    scan = _scan()
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: ())
    output = source if same_path else tmp_path / "output.xlsx"
    if not same_path:
        output.write_bytes(b"existing output")

    with pytest.raises(UsageError, match=r"source workbook is never overwritten|already exists"):
        engine.apply_patch_plan(
            source,
            plan,
            output,
            auto_repair=True,
            recalc_provider="auto",
        )

    assert source.read_bytes() == b"source workbook bytes"
    if not same_path:
        assert output.read_bytes() == b"existing output"


def test_degraded_copy_rejects_source_changed_after_canonical_scan(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"changed workbook bytes")
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256="0" * 64,
        patches=[patch],
    )
    scan = _scan()
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: ())
    output = tmp_path / "output.xlsx"

    with pytest.raises(StalePlanError, match="does not match workbook hash"):
        engine.apply_patch_plan(
            source,
            plan,
            output,
            auto_repair=True,
            recalc_provider="auto",
        )

    assert not output.exists()


def test_explicit_formula_repair_requires_workbook_trust_before_provider_discovery(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("provider discovery must remain disabled"),
    )

    with pytest.raises(UsageError, match="recalculation authorization"):
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            tmp_path / "output.xlsx",
            selected_ids=[patch.id],
            recalc_provider="excel",
        )


def test_formula_auto_repair_degrades_when_dependency_graph_is_opaque(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    safe_patch = _safe_patch()
    formula_patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[safe_patch, formula_patch],
    )
    snapshot = _formula_snapshot()
    snapshot.sheets[0].max_column = 4
    snapshot.sheets[0].cells["D2"] = CellSnapshot(
        coordinate="D2",
        formula='=INDIRECT("B2")',
        data_type="f",
    )
    scans = iter(
        [
            _scan(snapshot=snapshot),
            _scan(snapshot=snapshot),
            _scan(snapshot=snapshot),
        ]
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("opaque formula groups must degrade before discovery"),
    )

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        assert kwargs["selected_ids"] == {safe_patch.id}
        args[2].write_bytes(b"normalized")
        return _fake_low_level(args[2], formula_changed=False), [safe_patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: pytest.fail("opaque formula groups must not be recalculated"),
    )

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        tmp_path / "output.xlsx",
        auto_repair=True,
        recalc_provider="excel",
        trust_workbook_for_recalculation=True,
    )

    assert result.applied_patch_ids == [safe_patch.id]
    assert result.downgraded_patch_ids == [formula_patch.id]
    assert result.validation_status == ValidationStatus.DEGRADED
    assert any("opaque references" in message for message in result.validation_messages)


def test_explicit_formula_repair_rejects_opaque_dependency_graph(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    formula_patch = _formula_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[formula_patch],
    )
    snapshot = _formula_snapshot()
    snapshot.sheets[0].max_column = 4
    snapshot.sheets[0].cells["D2"] = CellSnapshot(
        coordinate="D2",
        formula="=NamedBudget+1",
        data_type="f",
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan(snapshot=snapshot))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("opaque formulas must be rejected before discovery"),
    )

    output = tmp_path / "output.xlsx"
    with pytest.raises(PatchValidationError, match="opaque references") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[formula_patch.id],
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()
    failure = caught.value.patch_result
    assert failure.validation_status == ValidationStatus.FAILED
    assert failure.rollback_performed
    assert failure.failure_report_path is not None
    report_text = Path(failure.failure_report_path).read_text(encoding="utf-8")
    assert "validation failure" in report_text
    assert "修复验证失败" in report_text


def test_same_finding_id_escalating_to_error_rolls_back_and_writes_bilingual_report(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _safe_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    scans = iter(
        [
            _scan(findings=[_finding("stable-id", Severity.WARNING)]),
            _scan(findings=[_finding("stable-id", Severity.ERROR)]),
        ]
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2], formula_changed=False), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="severity-escalated") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
        )

    assert not output.exists()
    failure = caught.value.patch_result
    assert failure.validation_status == ValidationStatus.FAILED
    assert failure.rollback_performed is True
    assert failure.failure_report_path is not None
    report = Path(failure.failure_report_path)
    assert report.exists()
    report_text = report.read_text(encoding="utf-8")
    assert "validation failure" in report_text
    assert "修复验证失败" in report_text


def test_failure_report_does_not_claim_cleanup_when_rollback_is_incomplete(
    tmp_path: Path,
) -> None:
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[_safe_patch()],
    )
    output = tmp_path / "output.xlsx"
    result = engine._failure_result(
        plan=plan,
        output=output,
        before_scan=None,
        attempted_ids={"safe"},
        downgraded_ids=set(),
        skipped_ids=set(),
        provider=RecalculationProvider.NONE,
        formula_errors_before=(),
        formula_errors_after=(),
        reason="cleanup test",
        rollback_performed=False,
    )

    report = engine._write_failure_report(
        output,
        result,
        attempted_ids={"safe"},
        reason="cleanup test",
    )

    assert report is not None
    report_text = report.read_text(encoding="utf-8")
    assert "cleanup was incomplete" in report_text
    assert "未完全清理" in report_text
    assert "output may still exist" in report_text
    assert "输出可能仍然存在" in report_text
    assert "intended output was not published" not in report_text
    assert "private candidate removed" not in report_text


def test_private_candidate_workspace_cannot_escape_output_parent(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    output_parent = tmp_path / "output"
    escaped = tmp_path / "escaped"
    output_parent.mkdir()
    escaped.mkdir()
    monkeypatch.setattr(engine.tempfile, "mkdtemp", lambda **_kwargs: str(escaped))

    with pytest.raises(PatchValidationError, match="escaped its intended parent directory"):
        engine._create_private_candidate(output_parent / "result.xlsx")


def test_private_candidate_names_have_bounded_output_independent_overhead(
    tmp_path: Path,
) -> None:
    output = tmp_path / ("x" * 180 + ".xlsx")

    workspace, candidate = engine._create_private_candidate(output)
    try:
        assert workspace.parent == tmp_path
        assert workspace.name.startswith(".wl-r-")
        assert output.stem not in workspace.name
        assert candidate.parent == workspace
        assert candidate.name.startswith("c-")
        assert len(str(candidate)) - len(str(tmp_path)) < 40
    finally:
        assert engine._cleanup_private_candidate(workspace, candidate)


def test_workspace_escape_fails_before_apply_and_writes_bilingual_report(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _safe_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    escaped = tmp_path / "escaped"
    escaped.mkdir()
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    monkeypatch.setattr(engine.tempfile, "mkdtemp", lambda **_kwargs: str(escaped))

    with pytest.raises(
        PatchValidationError, match="escaped its intended parent directory"
    ) as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            tmp_path / "output" / "result.xlsx",
            selected_ids=[patch.id],
        )

    failure = caught.value.patch_result
    assert failure.rollback_performed
    assert failure.failure_report_path is not None
    report_text = Path(failure.failure_report_path).read_text(encoding="utf-8")
    assert "cleanup completed" in report_text
    assert "清理已完成" in report_text


def test_failure_report_rejects_escaped_mkstemp_without_writing_it(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[_safe_patch()],
    )
    output_parent = tmp_path / "output"
    escaped_parent = tmp_path / "escaped"
    output_parent.mkdir()
    escaped_parent.mkdir()
    escaped_report = escaped_parent / "report.txt"
    escaped_report.write_text("sentinel", encoding="utf-8")
    descriptor = engine.os.open(escaped_report, engine.os.O_RDWR)
    monkeypatch.setattr(
        engine.tempfile,
        "mkstemp",
        lambda **_kwargs: (descriptor, str(escaped_report)),
    )
    result = engine._failure_result(
        plan=plan,
        output=output_parent / "result.xlsx",
        before_scan=None,
        attempted_ids={"safe"},
        downgraded_ids=set(),
        skipped_ids=set(),
        provider=RecalculationProvider.NONE,
        formula_errors_before=(),
        formula_errors_after=(),
        reason="escape test",
        rollback_performed=True,
    )

    report = engine._write_failure_report(
        output_parent / "result.xlsx",
        result,
        attempted_ids={"safe"},
        reason="escape test",
    )

    assert report is None
    assert escaped_report.read_text(encoding="utf-8") == "sentinel"


def test_link_created_before_publication_error_is_removed_and_reported_as_rolled_back(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _safe_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    scans = iter([_scan(), _scan()])
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2], formula_changed=False), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    real_link = engine.os.link

    def link_then_fail(source: Path, destination: Path) -> None:
        real_link(source, destination)
        raise OSError("fault after link")

    monkeypatch.setattr(engine.os, "link", link_then_fail)
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="atomic no-overwrite") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
        )

    assert not output.exists()
    assert caught.value.patch_result.rollback_performed


def test_concurrent_output_created_during_validation_is_never_deleted(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    patch = _safe_patch()
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[patch],
    )
    output = tmp_path / "output.xlsx"
    scan_count = 0

    def fake_scan(*args: Any, **kwargs: Any) -> SimpleNamespace:
        nonlocal scan_count
        scan_count += 1
        if scan_count == 2:
            output.write_bytes(b"concurrent owner")
        return _scan()

    monkeypatch.setattr(engine, "scan_workbook", fake_scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2], formula_changed=False), [patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)

    with pytest.raises(UsageError, match="already exists") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            selected_ids=[patch.id],
        )

    assert output.read_bytes() == b"concurrent owner"
    assert caught.value.patch_result.validation_status == ValidationStatus.FAILED
    assert caught.value.patch_result.rollback_performed is True


def test_formula_only_degradation_uses_canonical_patch_authority_not_plan_dump(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    source = tmp_path / "input.xlsx"
    source.write_bytes(b"source workbook bytes")
    patch = _formula_patch()
    canonical = PatchPlan(
        tool_version="2.4.0",
        source_name=source.name,
        source_sha256=sha256_file(source),
        patches=[patch],
    )
    submitted = canonical.model_copy(deep=True)
    submitted.finding_ids = ["presentation-only-difference"]
    scan = _scan()
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: scan)
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: canonical)
    monkeypatch.setattr(
        engine,
        "candidate_providers",
        lambda _preference: pytest.fail("untrusted workbooks must not start provider discovery"),
    )

    result = engine.apply_patch_plan(
        source,
        submitted,
        tmp_path / "output.xlsx",
        auto_repair=True,
    )

    assert result.validation_status == ValidationStatus.DEGRADED
    assert (tmp_path / "output.xlsx").read_bytes() == source.read_bytes()


def test_auto_repair_rebuilds_from_source_after_formula_group_downgrade(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    prerequisite = _safe_patch("normalize", cell="A2")
    rejected_formula = _formula_patch(
        "rejected-formula",
        cell="B2",
        atomic_group="rejected-group",
        prerequisite_patch_ids=[prerequisite.id],
    )
    rejected_companion = _safe_patch(
        "rejected-companion",
        cell="C3",
        atomic_group="rejected-group",
    )
    retained_formula = _formula_patch(
        "retained-formula",
        cell="D2",
        atomic_group="retained-group",
        prerequisite_patch_ids=[prerequisite.id],
    )
    reverse_dependent = _formula_patch(
        "reverse-dependent",
        cell="H2",
        atomic_group="dependent-group",
        prerequisite_patch_ids=[rejected_formula.id],
    )
    patches = [
        prerequisite,
        rejected_formula,
        rejected_companion,
        retained_formula,
        reverse_dependent,
    ]
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=patches,
    )
    snapshot = WorkbookSnapshot(
        source_name="input.xlsx",
        source_sha256="source-hash",
        format="xlsx",
        sheets=[
            SheetSnapshot(
                name="Sheet1",
                index=0,
                state="visible",
                max_row=3,
                max_column=9,
                cells={
                    "A2": CellSnapshot(coordinate="A2", value=5, data_type="n"),
                    "B2": CellSnapshot(coordinate="B2", formula="=A2/0", data_type="f"),
                    "C2": CellSnapshot(coordinate="C2", formula="=B2+1", data_type="f"),
                    "C3": CellSnapshot(coordinate="C3", value=" value ", data_type="s"),
                    "D2": CellSnapshot(coordinate="D2", formula="=A2/0", data_type="f"),
                    "E2": CellSnapshot(coordinate="E2", formula="=D2+1", data_type="f"),
                    "H2": CellSnapshot(coordinate="H2", formula="=A2/0", data_type="f"),
                    "I2": CellSnapshot(coordinate="I2", formula="=H2+1", data_type="f"),
                },
            )
        ],
    )
    scans = iter(
        [
            _scan(snapshot=snapshot),
            _scan(snapshot=snapshot),
            _scan(snapshot=snapshot),
        ]
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))
    selected_rounds: list[set[str]] = []
    candidate_paths: list[Path] = []
    source_paths: list[Path] = []

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        source_paths.append(args[0])
        candidate = args[2]
        candidate_paths.append(candidate)
        if len(candidate_paths) == 2:
            assert candidate_paths[0] != candidate
            assert not candidate_paths[0].exists()
        selected_ids = set(kwargs["selected_ids"])
        selected_rounds.append(selected_ids)
        candidate.write_text(",".join(sorted(selected_ids)), encoding="utf-8")
        selected = [patch for patch in patches if patch.id in selected_ids]
        return _fake_low_level(candidate), selected

    preferences: list[str] = []

    def fake_validate(*args: Any, **kwargs: Any) -> RecalculationValidation:
        preferences.append(kwargs["preference"])
        if len(preferences) == 1:
            return RecalculationValidation(
                provider,
                ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
                ("Sheet1!C2:#VALUE!",),
            )
        return RecalculationValidation(
            provider,
            ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
            (),
        )

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(engine, "validate_formula_recalculation", fake_validate)
    source = tmp_path / "input.xlsx"
    output = tmp_path / "output.xlsx"

    result = engine.apply_patch_plan(
        source,
        plan,
        output,
        auto_repair=True,
        recalc_provider="auto",
        trust_workbook_for_recalculation=True,
    )

    assert selected_rounds == [
        {patch.id for patch in patches},
        {prerequisite.id, retained_formula.id},
    ]
    assert source_paths == [source.resolve(), source.resolve()]
    assert preferences == ["auto", "excel"]
    assert result.applied_patch_ids == [prerequisite.id, retained_formula.id]
    assert result.downgraded_patch_ids == sorted([rejected_formula.id, reverse_dependent.id])
    assert result.skipped_patch_ids == sorted(
        [rejected_formula.id, rejected_companion.id, reverse_dependent.id]
    )
    assert result.validation_status == ValidationStatus.DEGRADED
    assert any("recalculation round 1" in message for message in result.validation_messages)
    assert any("重算第 1 轮" in message for message in result.validation_messages)
    assert output.read_text(encoding="utf-8") == ",".join(
        sorted([prerequisite.id, retained_formula.id])
    )


def test_auto_repair_recalculates_and_publishes_safe_subset_when_all_formula_groups_fail(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    safe_patch = _safe_patch("safe", cell="A2")
    formula_patch = _formula_patch("formula", cell="B2", atomic_group="formula-group")
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[safe_patch, formula_patch],
    )
    scan = _scan()
    scans = iter([scan, scan, scan])
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: next(scans))
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))
    selected_rounds: list[set[str]] = []

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        selected_ids = set(kwargs["selected_ids"])
        selected_rounds.append(selected_ids)
        args[2].write_bytes(b"safe subset" if selected_ids == {safe_patch.id} else b"full")
        selected = [patch for patch in plan.patches if patch.id in selected_ids]
        return _fake_low_level(args[2], formula_changed=formula_patch.id in selected_ids), selected

    validations = iter(
        [
            RecalculationValidation(
                provider,
                ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
                ("Sheet1!C2:#VALUE!",),
            ),
            RecalculationValidation(
                provider,
                ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
                ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
            ),
        ]
    )
    preferences: list[str] = []

    def fake_validate(*args: Any, **kwargs: Any) -> RecalculationValidation:
        preferences.append(kwargs["preference"])
        return next(validations)

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(engine, "validate_formula_recalculation", fake_validate)
    output = tmp_path / "output.xlsx"

    result = engine.apply_patch_plan(
        tmp_path / "input.xlsx",
        plan,
        output,
        auto_repair=True,
        recalc_provider="auto",
        trust_workbook_for_recalculation=True,
    )

    assert selected_rounds == [{safe_patch.id, formula_patch.id}, {safe_patch.id}]
    assert preferences == ["auto", "excel"]
    assert result.applied_patch_ids == [safe_patch.id]
    assert result.downgraded_patch_ids == [formula_patch.id]
    assert result.skipped_patch_ids == [formula_patch.id]
    assert result.formula_errors_after == ["Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"]
    assert result.recalculation_provider == RecalculationProvider.EXCEL
    assert result.validation_status == ValidationStatus.DEGRADED
    assert output.read_bytes() == b"safe subset"


def test_auto_repair_formula_group_downgrade_fails_closed_when_no_progress_is_possible(
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
) -> None:
    formula_patch = _formula_patch("formula", atomic_group="formula-group")
    plan = PatchPlan(
        tool_version="2.4.0",
        source_name="input.xlsx",
        source_sha256="source-hash",
        patches=[formula_patch],
    )
    monkeypatch.setattr(engine, "scan_workbook", lambda *args, **kwargs: _scan())
    monkeypatch.setattr(engine, "build_patch_plan", lambda _scan_result: plan)
    provider = conversion.ConversionProvider("excel", "Microsoft Excel", Path("excel.exe"))
    monkeypatch.setattr(engine, "candidate_providers", lambda _preference: (provider,))

    def fake_patch(*args: Any, **kwargs: Any) -> tuple[SimpleNamespace, list[PatchOperation]]:
        args[2].write_bytes(b"candidate")
        return _fake_low_level(args[2]), [formula_patch]

    monkeypatch.setattr(engine, "patch_ooxml_package", fake_patch)
    monkeypatch.setattr(
        engine,
        "validate_formula_recalculation",
        lambda *args, **kwargs: RecalculationValidation(
            provider,
            ("Sheet1!B2:#VALUE!", "Sheet1!C2:#VALUE!"),
            ("Sheet1!C2:#VALUE!",),
        ),
    )
    monkeypatch.setattr(engine, "_atomic_dependent_removals", lambda *args, **kwargs: set())
    output = tmp_path / "output.xlsx"

    with pytest.raises(PatchValidationError, match="made no progress") as caught:
        engine.apply_patch_plan(
            tmp_path / "input.xlsx",
            plan,
            output,
            auto_repair=True,
            recalc_provider="excel",
            trust_workbook_for_recalculation=True,
        )

    assert not output.exists()
    assert caught.value.patch_result.validation_status == ValidationStatus.FAILED
    assert caught.value.patch_result.rollback_performed
