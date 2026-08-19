"""End-to-end patch application with rescanning and fail-closed cleanup."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from workbooklens.exceptions import PatchValidationError
from workbooklens.models import PatchPlan, PatchResult, Severity
from workbooklens.ooxml.safety import PackageLimits
from workbooklens.repair.ooxml_patch import patch_ooxml_package
from workbooklens.repair.planning import build_patch_plan
from workbooklens.scanner import scan_workbook


def apply_patch_plan(
    source: Path,
    plan: PatchPlan,
    output: Path,
    *,
    selected_ids: set[str] | None = None,
    safe_only: bool = False,
    config: dict[str, Any] | None = None,
    limits: PackageLimits | None = None,
) -> PatchResult:
    """Apply, reopen, rescan, compare findings, and delete invalid output."""

    before_scan = scan_workbook(source, config=config, limits=limits)
    canonical_plan = build_patch_plan(before_scan)
    low_level, selected = patch_ooxml_package(
        source,
        plan,
        output,
        selected_ids=selected_ids,
        safe_only=safe_only,
        limits=limits,
        canonical_plan=canonical_plan,
    )
    try:
        after_scan = scan_workbook(output, config=config, limits=limits)
        before_ids = {finding.id for finding in before_scan.findings}
        after_ids = {finding.id for finding in after_scan.findings}
        new_findings = [finding for finding in after_scan.findings if finding.id not in before_ids]
        dangerous_new = [
            finding
            for finding in new_findings
            if finding.severity in {Severity.ERROR, Severity.CRITICAL}
        ]
        if dangerous_new:
            identifiers = ", ".join(finding.id for finding in dangerous_new)
            raise PatchValidationError(f"Repair introduced new error-level findings: {identifiers}")
        messages = [
            "Output reopened through the secure OOXML reader and openpyxl read-only mode.",
            "Every untouched ZIP part has byte-identical uncompressed content.",
        ]
        if low_level.formula_changed:
            messages.append(
                "Formula cached values were removed and full recalculation is requested on next open; "
                "WorkbookLens did not calculate formulas."
            )
        return PatchResult(
            source_sha256=low_level.source_sha256,
            output_sha256=low_level.output_sha256,
            output_path=str(output.resolve()),
            applied_patch_ids=[patch.id for patch in selected],
            package_changes=low_level.changes,
            resolved_finding_ids=sorted(before_ids - after_ids),
            remaining_finding_ids=sorted(before_ids & after_ids),
            new_finding_ids=sorted(after_ids - before_ids),
            validation_messages=messages,
        )
    except BaseException:
        output.unlink(missing_ok=True)
        raise
