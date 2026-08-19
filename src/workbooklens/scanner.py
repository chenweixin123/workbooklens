"""Deterministic scan orchestration and finding de-duplication."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from workbooklens.models import SEVERITY_RANK, Finding, PatchOperation, WorkbookSnapshot
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
                patch_by_id[patch.id] = patch
        if inspection.format == "xlsm":
            finding_by_id = {
                finding_id: finding.model_copy(
                    update={"safe_patch_available": False, "patch_ids": []}
                )
                for finding_id, finding in finding_by_id.items()
            }
            patch_by_id.clear()
        findings = sorted(
            finding_by_id.values(),
            key=lambda finding: (
                -SEVERITY_RANK[finding.severity],
                finding.rule_id,
                finding.sheet or "",
                finding.location or "",
                finding.id,
            ),
        )
        patches = sorted(
            patch_by_id.values(), key=lambda patch: (patch.sheet, patch.cell, patch.kind, patch.id)
        )
        return ScanResult(
            inspection=inspection,
            snapshot=snapshot,
            findings=findings,
            patches=patches,
        )
    finally:
        workbook.close()
