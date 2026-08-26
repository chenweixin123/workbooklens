"""Creation and loading of source-bound JSON patch manifests."""

from __future__ import annotations

import json
from collections.abc import Collection
from pathlib import Path

from pydantic import ValidationError

from workbooklens import __version__
from workbooklens.exceptions import UsageError
from workbooklens.models import FORMULA_PATCH_KINDS, PatchOperation, PatchPlan, PatchRisk
from workbooklens.scanner import ScanResult
from workbooklens.utils import write_json


def build_patch_plan(scan: ScanResult) -> PatchPlan:
    """Create a deterministic plan containing every reviewable proposed patch."""

    return PatchPlan(
        tool_version=__version__,
        source_name=scan.inspection.path.name,
        source_sha256=scan.snapshot.source_sha256,
        patches=scan.patches,
        finding_ids=[finding.id for finding in scan.findings if finding.patch_ids],
        findings=[finding for finding in scan.findings if finding.patch_ids],
    )


def write_patch_plan(path: Path, plan: PatchPlan) -> None:
    """Write a plan using the public JSON schema representation."""

    write_json(path, plan.model_dump(mode="json"))


def load_patch_plan(path: Path, max_bytes: int = 10 * 1024 * 1024) -> PatchPlan:
    """Read and strictly validate a patch plan supplied by a user."""

    try:
        encoded = path.read_bytes()
    except OSError as exc:
        raise UsageError(f"Unable to read patch plan {path}: {exc}") from exc
    if len(encoded) > max_bytes:
        raise UsageError(f"Patch plan exceeds the {max_bytes}-byte limit")
    try:
        payload = json.loads(encoded)
    except (json.JSONDecodeError, UnicodeDecodeError) as exc:
        raise UsageError(f"Patch plan is not valid UTF-8 JSON: {exc}") from exc
    if not isinstance(payload, dict) or payload.get("schema_version") not in {2, 3}:
        raise UsageError(
            "Patch plan must use schema version 2 or 3; regenerate it with WorkbookLens"
        )
    try:
        return PatchPlan.model_validate(payload)
    except ValidationError as exc:
        raise UsageError(f"Patch plan does not match schema version 3: {exc}") from exc


def _is_unverified_legacy_formula(patch: PatchOperation) -> bool:
    """Return whether a migrated v2 formula lacks the proof required by schema v3."""

    return (
        patch.kind in FORMULA_PATCH_KINDS
        and patch.derivation.strategy == "legacy_schema_v2_unverified_formula"
    )


def _reject_unverified_legacy_formulas(
    by_id: dict[str, PatchOperation],
    resolved: Collection[str],
) -> None:
    legacy = sorted(
        patch_id for patch_id in resolved if _is_unverified_legacy_formula(by_id[patch_id])
    )
    if legacy:
        raise UsageError(
            "Legacy schema v2 formula patches are unverified and cannot be applied; "
            "regenerate the plan with WorkbookLens schema v3: " + ", ".join(legacy)
        )


def resolve_patch_selection(
    plan: PatchPlan,
    selected_ids: Collection[str] | None = None,
    safe_only: bool = False,
    *,
    accept_layout_risk: bool = False,
    accept_formula_derived: bool = False,
    accept_semantic_risk: bool = False,
    auto_repair: bool = False,
) -> set[str]:
    """Resolve a fail-closed selection while preserving atomic repair groups."""

    by_id = {patch.id: patch for patch in plan.patches}
    if len(by_id) != len(plan.patches):
        raise UsageError("Patch plan contains duplicate patch IDs")
    requested_ids = list(selected_ids or ())
    if len(requested_ids) != len(set(requested_ids)):
        raise UsageError("Duplicate --patch-id selections are not allowed")
    requested = set(requested_ids)
    if safe_only and requested:
        raise UsageError("--safe-only and --patch-id are mutually exclusive")
    if auto_repair and safe_only:
        raise UsageError("--auto-repair and --safe-only are mutually exclusive")
    if safe_only and accept_layout_risk:
        raise UsageError("--safe-only and --accept-layout-risk are mutually exclusive")
    if safe_only and (accept_formula_derived or accept_semantic_risk):
        raise UsageError("--safe-only cannot be combined with repair-risk authorization")
    if auto_repair and accept_layout_risk:
        raise UsageError("--auto-repair never accepts layout repair risk")

    groups: dict[str, set[str]] = {}
    for patch in plan.patches:
        if patch.atomic_group is not None:
            groups.setdefault(patch.atomic_group, set()).add(patch.id)

    def expand_required(seed: set[str]) -> set[str]:
        resolved = set(seed)
        pending = list(seed)
        while pending:
            patch_id = pending.pop()
            patch = by_id[patch_id]
            required = set(patch.prerequisite_patch_ids)
            if patch.atomic_group is not None:
                required.update(groups[patch.atomic_group])
            unknown_required = required - by_id.keys()
            if unknown_required:
                raise UsageError(
                    "Patch plan references unknown prerequisite patch IDs: "
                    + ", ".join(sorted(unknown_required))
                )
            for required_id in required - resolved:
                resolved.add(required_id)
                pending.append(required_id)
        return resolved

    if safe_only or auto_repair:
        if auto_repair:
            unknown = requested - by_id.keys()
            if unknown:
                raise UsageError("Unknown patch IDs: " + ", ".join(sorted(unknown)))
            non_semantic = sorted(
                patch_id
                for patch_id in requested
                if by_id[patch_id].risk != PatchRisk.SEMANTIC_REVIEW
            )
            if non_semantic:
                raise UsageError(
                    "--auto-repair accepts only explicitly selected semantic-review patch IDs; "
                    "safe and formula-derived patches are selected automatically: "
                    + ", ".join(non_semantic)
                )
            if requested and not accept_semantic_risk:
                raise UsageError(
                    "Semantic-review patches require explicit --accept-semantic-risk: "
                    + ", ".join(sorted(requested))
                )
            below_threshold = sorted(
                patch_id for patch_id in requested if float(by_id[patch_id].confidence) < 0.95
            )
            if below_threshold:
                raise UsageError(
                    "WorkbookLens refuses patches below the 0.95 confidence threshold: "
                    + ", ".join(below_threshold)
                )

        def automatically_eligible(patch_id: str) -> bool:
            patch = by_id[patch_id]
            return not _is_unverified_legacy_formula(patch) and (
                patch.safe_only_eligible
                or (auto_repair and patch.risk == PatchRisk.FORMULA_DERIVED)
            )

        def validate_auto_boundary(patch_ids: set[str]) -> None:
            invalid_safe = sorted(
                patch_id
                for patch_id in patch_ids
                if by_id[patch_id].risk == PatchRisk.SAFE and not by_id[patch_id].safe_only_eligible
            )
            if invalid_safe:
                raise UsageError(
                    "Automatic repair requires every safe-risk dependency to be "
                    "safe-only eligible: " + ", ".join(invalid_safe)
                )
            invalid_formula = sorted(
                patch_id
                for patch_id in patch_ids
                if by_id[patch_id].risk == PatchRisk.FORMULA_DERIVED
                and not automatically_eligible(patch_id)
            )
            if invalid_formula:
                raise UsageError(
                    "Automatic repair accepts only verified formula-derived patches: "
                    + ", ".join(invalid_formula)
                )
            implicit_semantic = sorted(
                patch_id
                for patch_id in patch_ids
                if by_id[patch_id].risk == PatchRisk.SEMANTIC_REVIEW and patch_id not in requested
            )
            if implicit_semantic:
                raise UsageError(
                    "Every semantic-review patch in an automatic repair dependency chain must "
                    "be explicitly selected: " + ", ".join(implicit_semantic)
                )
            low_confidence_semantic = sorted(
                patch_id
                for patch_id in patch_ids
                if by_id[patch_id].risk == PatchRisk.SEMANTIC_REVIEW
                and float(by_id[patch_id].confidence) < 0.95
            )
            if low_confidence_semantic:
                raise UsageError(
                    "WorkbookLens refuses patches below the 0.95 confidence threshold: "
                    + ", ".join(low_confidence_semantic)
                )
            forbidden = sorted(
                patch_id
                for patch_id in patch_ids
                if by_id[patch_id].risk
                not in {
                    PatchRisk.SAFE,
                    PatchRisk.FORMULA_DERIVED,
                    PatchRisk.SEMANTIC_REVIEW,
                }
            )
            if forbidden:
                raise UsageError(
                    "Automatic repair forbids layout-review or unsupported unsafe dependencies: "
                    + ", ".join(forbidden)
                )
            if (
                any(by_id[patch_id].risk == PatchRisk.SEMANTIC_REVIEW for patch_id in patch_ids)
                and not accept_semantic_risk
            ):
                raise UsageError(
                    "Semantic-review patches require explicit --accept-semantic-risk: "
                    + ", ".join(
                        sorted(
                            patch_id
                            for patch_id in patch_ids
                            if by_id[patch_id].risk == PatchRisk.SEMANTIC_REVIEW
                        )
                    )
                )

        resolved = {
            patch.id
            for patch in plan.patches
            if patch.atomic_group is None and automatically_eligible(patch.id)
        }
        for member_ids in groups.values():
            if all(automatically_eligible(patch_id) for patch_id in member_ids):
                resolved.update(member_ids)
        resolved = expand_required(resolved)
        if safe_only:
            ineligible_prerequisites = sorted(
                patch_id for patch_id in resolved if not automatically_eligible(patch_id)
            )
            if ineligible_prerequisites:
                raise UsageError(
                    "Automatic repair dependencies fall outside the safe repair boundary: "
                    + ", ".join(ineligible_prerequisites)
                )
        if auto_repair and requested:
            semantic_resolved = expand_required(requested)
            resolved.update(semantic_resolved)
        if auto_repair:
            validate_auto_boundary(resolved)
    else:
        if not requested:
            raise UsageError(
                "Select at least one --patch-id or pass --safe-only",
                error_key="repair.selection_required",
            )
        unknown = requested - by_id.keys()
        if unknown:
            raise UsageError("Unknown patch IDs: " + ", ".join(sorted(unknown)))
        resolved = set(requested)
        resolved = expand_required(resolved)

        _reject_unverified_legacy_formulas(by_id, resolved)

        below_threshold = sorted(
            patch_id for patch_id in resolved if float(by_id[patch_id].confidence) < 0.95
        )
        if below_threshold:
            raise UsageError(
                "WorkbookLens refuses patches below the 0.95 confidence threshold: "
                + ", ".join(below_threshold)
            )
        layout_review = sorted(
            patch_id for patch_id in resolved if by_id[patch_id].risk == PatchRisk.LAYOUT_REVIEW
        )
        if layout_review and not accept_layout_risk:
            raise UsageError(
                "Layout-review patches require explicit --accept-layout-risk: "
                + ", ".join(layout_review)
            )
        formula_derived = sorted(
            patch_id for patch_id in resolved if by_id[patch_id].risk == PatchRisk.FORMULA_DERIVED
        )
        if formula_derived and not accept_formula_derived:
            raise UsageError(
                "Formula-derived patches require recalculation authorization: "
                + ", ".join(formula_derived)
            )
        semantic_review = sorted(
            patch_id for patch_id in resolved if by_id[patch_id].risk == PatchRisk.SEMANTIC_REVIEW
        )
        if semantic_review and not accept_semantic_risk:
            raise UsageError(
                "Semantic-review patches require explicit --accept-semantic-risk: "
                + ", ".join(semantic_review)
            )
        unsafe = sorted(
            patch_id
            for patch_id in resolved
            if not by_id[patch_id].safe
            and by_id[patch_id].risk
            not in {
                PatchRisk.LAYOUT_REVIEW,
                PatchRisk.FORMULA_DERIVED,
                PatchRisk.SEMANTIC_REVIEW,
            }
        )
        if unsafe:
            raise UsageError("WorkbookLens refuses unsafe patches: " + ", ".join(unsafe))

    if not resolved:
        raise UsageError("No eligible patches were selected")
    _reject_unverified_legacy_formulas(by_id, resolved)
    return resolved
