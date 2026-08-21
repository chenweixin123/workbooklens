"""Creation and loading of source-bound JSON patch manifests."""

from __future__ import annotations

import json
from collections.abc import Collection
from pathlib import Path

from pydantic import ValidationError

from workbooklens import __version__
from workbooklens.exceptions import UsageError
from workbooklens.models import PatchPlan, PatchRisk
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
    if not isinstance(payload, dict) or payload.get("schema_version") != 2:
        raise UsageError("Patch plan must use schema version 2; regenerate it with WorkbookLens")
    try:
        return PatchPlan.model_validate(payload)
    except ValidationError as exc:
        raise UsageError(f"Patch plan does not match schema version 2: {exc}") from exc


def resolve_patch_selection(
    plan: PatchPlan,
    selected_ids: Collection[str] | None = None,
    safe_only: bool = False,
    *,
    accept_layout_risk: bool = False,
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
    if safe_only and accept_layout_risk:
        raise UsageError("--safe-only and --accept-layout-risk are mutually exclusive")

    groups: dict[str, set[str]] = {}
    for patch in plan.patches:
        if patch.atomic_group is not None:
            groups.setdefault(patch.atomic_group, set()).add(patch.id)

    if safe_only:
        resolved = {
            patch.id
            for patch in plan.patches
            if patch.atomic_group is None and patch.safe_only_eligible
        }
        for member_ids in groups.values():
            members = [by_id[patch_id] for patch_id in member_ids]
            if all(patch.safe_only_eligible for patch in members):
                resolved.update(member_ids)
    else:
        if not requested:
            raise UsageError("Select at least one --patch-id or pass --safe-only")
        unknown = requested - by_id.keys()
        if unknown:
            raise UsageError("Unknown patch IDs: " + ", ".join(sorted(unknown)))
        resolved = set(requested)
        for patch_id in tuple(resolved):
            group = by_id[patch_id].atomic_group
            if group is not None:
                resolved.update(groups[group])

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
        unsafe = sorted(
            patch_id
            for patch_id in resolved
            if not by_id[patch_id].safe and by_id[patch_id].risk != PatchRisk.LAYOUT_REVIEW
        )
        if unsafe:
            raise UsageError("WorkbookLens refuses unsafe patches: " + ", ".join(unsafe))

    if not resolved:
        raise UsageError("No eligible patches were selected")
    return resolved
