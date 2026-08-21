"""Deterministic scan orchestration and finding de-duplication."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from workbooklens.models import SEVERITY_RANK, Finding, PatchKind, PatchOperation, WorkbookSnapshot
from workbooklens.ooxml.formula_ranges import find_unsupported_formula_ranges
from workbooklens.ooxml.safety import PackageInspection, PackageLimits, inspect_package
from workbooklens.regions import infer_data_regions, infer_formula_bands
from workbooklens.rules import RuleContext, RuleRegistry, default_registry
from workbooklens.snapshot import load_for_analysis, snapshot_from_workbook
from workbooklens.utils import sha256_file


@dataclass(slots=True)
class ScanResult:
    """Internal aggregate consumed by reports, plans, tests, and the web UI."""

    inspection: PackageInspection
    snapshot: WorkbookSnapshot
    findings: list[Finding]
    patches: list[PatchOperation]


def _normalize_column_width_patches(
    findings: list[Finding],
    patches: list[PatchOperation],
) -> tuple[list[Finding], list[PatchOperation]]:
    """Collapse monotonic same-column width requests to their sufficient maximum."""

    groups: dict[tuple[str, str], list[PatchOperation]] = {}
    for patch in patches:
        if patch.kind != PatchKind.SET_COLUMN_WIDTH or not isinstance(patch.after, dict):
            continue
        column = patch.after.get("column")
        width = patch.after.get("width")
        if (
            not isinstance(column, str)
            or isinstance(width, bool)
            or not isinstance(width, (int, float))
        ):
            continue
        groups.setdefault((patch.sheet, column.upper()), []).append(patch)
    replacement: dict[str, str] = {}
    removed: set[str] = set()
    for group in groups.values():
        if len(group) < 2:
            continue
        if any(patch.atomic_group is not None for patch in group):
            continue
        chosen = sorted(
            group,
            key=lambda patch: (-float(patch.after["width"]), patch.id),
        )[0]
        for patch in group:
            replacement[patch.id] = chosen.id
            if patch.id != chosen.id:
                removed.add(patch.id)
    normalized_patches = [patch for patch in patches if patch.id not in removed]
    by_id = {patch.id: patch for patch in normalized_patches}
    normalized_findings: list[Finding] = []
    for finding in findings:
        identifiers: list[str] = []
        for patch_id in finding.patch_ids:
            normalized = replacement.get(patch_id, patch_id)
            if normalized not in identifiers:
                identifiers.append(normalized)
        missing = [patch_id for patch_id in identifiers if patch_id not in by_id]
        if missing:
            raise RuntimeError(
                f"Finding {finding.id} references missing patches: {', '.join(sorted(missing))}"
            )
        normalized_findings.append(
            finding.model_copy(
                update={
                    "patch_ids": identifiers,
                    "safe_patch_available": any(
                        by_id[patch_id].safe_only_eligible for patch_id in identifiers
                    ),
                }
            )
        )
    return normalized_findings, normalized_patches


def scan_workbook(
    path: Path,
    *,
    config: dict[str, Any] | None = None,
    limits: PackageLimits | None = None,
    registry: RuleRegistry | None = None,
) -> ScanResult:
    """Inspect and analyze one workbook without saving or evaluating it."""

    inspection = inspect_package(path, limits)
    source_sha256 = sha256_file(inspection.path)
    workbook = load_for_analysis(inspection.path, limits)
    try:
        snapshot = snapshot_from_workbook(workbook, inspection.path, source_sha256)
        data_regions = {
            worksheet.title: infer_data_regions(worksheet) for worksheet in workbook.worksheets
        }
        formula_bands = {
            worksheet.title: infer_formula_bands(worksheet) for worksheet in workbook.worksheets
        }
        unsupported_formula_ranges = find_unsupported_formula_ranges(
            inspection.path,
            limits,
        )
        context = RuleContext(
            path=inspection.path,
            workbook=workbook,
            snapshot=snapshot,
            config=config or {},
            data_regions=data_regions,
            formula_bands=formula_bands,
            unsupported_formula_ranges=unsupported_formula_ranges,
        )
        finding_by_id: dict[str, Finding] = {}
        patch_by_id: dict[str, PatchOperation] = {}
        active_registry = registry or default_registry()
        for rule in active_registry.values():
            rule_result = rule.run(context)
            for finding in rule_result.findings:
                if finding.id in finding_by_id:
                    raise RuntimeError(
                        "Duplicate finding identity generated for "
                        f"{finding.rule_id} at {finding.sheet or 'Workbook'}!"
                        f"{finding.location or ''}"
                    )
                finding_by_id[finding.id] = finding
            for patch in rule_result.patches:
                existing = patch_by_id.get(patch.id)
                if existing is not None and existing.model_dump(mode="json") != patch.model_dump(
                    mode="json"
                ):
                    raise RuntimeError(f"Conflicting patch identity generated: {patch.id}")
                patch_by_id[patch.id] = patch
            context.prior_patches = tuple(patch_by_id.values())
        if inspection.format == "xlsm":
            finding_by_id = {
                finding_id: finding.model_copy(
                    update={"safe_patch_available": False, "patch_ids": []}
                )
                for finding_id, finding in finding_by_id.items()
            }
            patch_by_id.clear()
        normalized_findings, normalized_patches = _normalize_column_width_patches(
            list(finding_by_id.values()),
            list(patch_by_id.values()),
        )
        findings = sorted(
            normalized_findings,
            key=lambda finding: (
                -SEVERITY_RANK[finding.severity],
                finding.rule_id,
                finding.sheet or "",
                finding.location or "",
                finding.id,
            ),
        )
        patches = sorted(
            normalized_patches, key=lambda patch: (patch.sheet, patch.cell, patch.kind, patch.id)
        )
        return ScanResult(
            inspection=inspection,
            snapshot=snapshot,
            findings=findings,
            patches=patches,
        )
    finally:
        workbook.close()
