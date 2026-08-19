"""Scan report serialization with no network-loaded assets."""

from __future__ import annotations

from collections import Counter
from dataclasses import replace
from pathlib import Path
from typing import Any

from jinja2 import Environment, PackageLoader, select_autoescape
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.utils.exceptions import CellCoordinatesException

from workbooklens import __version__
from workbooklens.models import Finding, Severity
from workbooklens.policy import FindingPolicyResult, apply_finding_policy, source_scope_for_path
from workbooklens.scanner import ScanResult
from workbooklens.utils import atomic_write_bytes, write_json


def _weighted_finding_score(findings: list[Finding] | tuple[Finding, ...]) -> int:
    """Return an additive severity score; it is not a probability or percentage."""

    weights = {
        Severity.CRITICAL: 35,
        Severity.ERROR: 15,
        Severity.WARNING: 5,
        Severity.INFO: 1,
    }
    return sum(weights[finding.severity] for finding in findings)


def _sarif_level(severity: Severity) -> str:
    return {
        Severity.CRITICAL: "error",
        Severity.ERROR: "error",
        Severity.WARNING: "warning",
        Severity.INFO: "note",
    }[severity]


def _sarif_location(finding: Finding, source_name: str) -> dict[str, Any]:
    physical: dict[str, Any] = {"artifactLocation": {"uri": source_name}}
    if finding.location and finding.sheet:
        try:
            column_letters, row = coordinate_from_string(finding.location)
            physical["region"] = {
                "startLine": row,
                "startColumn": column_index_from_string(column_letters),
            }
        except (CellCoordinatesException, ValueError, TypeError):
            pass
    location: dict[str, Any] = {"physicalLocation": physical}
    if finding.sheet:
        logical_name = finding.sheet
        if finding.location:
            logical_name += f"!{finding.location}"
        location["logicalLocations"] = [{"name": logical_name, "kind": "spreadsheetLocation"}]
    return location


def build_sarif(scan: ScanResult, *, source_uri: str | None = None) -> dict[str, Any]:
    """Build a GitHub code-scanning-compatible SARIF 2.1.0 document."""

    logical_source = source_uri or source_scope_for_path(scan.inspection.path)
    rule_examples: dict[str, Finding] = {}
    for finding in scan.findings:
        rule_examples.setdefault(finding.rule_id, finding)
    rules = [
        {
            "id": rule_id,
            "name": rule_id,
            "shortDescription": {"text": example.title},
            "fullDescription": {"text": example.explanation},
            "help": {"text": example.suggested_action},
            "properties": {"defaultSeverity": example.severity.value},
        }
        for rule_id, example in sorted(rule_examples.items())
    ]
    results = [
        {
            "ruleId": finding.rule_id,
            "level": _sarif_level(finding.severity),
            "message": {"text": f"{finding.title}: {finding.explanation}"},
            "locations": [_sarif_location(finding, logical_source)],
            "fingerprints": {"workbooklensFindingId": finding.id},
            "properties": {
                "contentFingerprint": finding.content_fingerprint,
                "confidence": float(finding.confidence),
                "sheet": finding.sheet,
                "cellOrRange": finding.location,
                "safePatchAvailable": finding.safe_patch_available,
            },
        }
        for finding in scan.findings
    ]
    return {
        "$schema": "https://json.schemastore.org/sarif-2.1.0.json",
        "version": "2.1.0",
        "runs": [
            {
                "tool": {
                    "driver": {
                        "name": "WorkbookLens",
                        "semanticVersion": __version__,
                        "informationUri": "https://github.com/chenweixin123/workbooklens",
                        "rules": rules,
                    }
                },
                "automationDetails": {"id": f"workbooklens/{logical_source}"},
                "results": results,
            }
        ],
    }


def write_scan_report(
    scan: ScanResult,
    output_directory: Path,
    *,
    policy: FindingPolicyResult | None = None,
) -> dict[str, Path]:
    """Write HTML, findings JSON, snapshot JSON, and SARIF into one directory."""

    active_policy = policy or apply_finding_policy(scan.findings)
    logical_source = active_policy.source_scope or source_scope_for_path(scan.inspection.path)
    active_scan = replace(scan, findings=list(active_policy.active_findings))
    output_directory.mkdir(parents=True, exist_ok=True)
    findings_path = output_directory / "findings.json"
    snapshot_path = output_directory / "snapshot.json"
    sarif_path = output_directory / "results.sarif"
    html_path = output_directory / "report.html"
    findings_payload = {
        "schema_version": 2,
        "tool_version": __version__,
        "source": scan.inspection.path.name,
        "source_scope": logical_source,
        "source_sha256": scan.snapshot.source_sha256,
        "weighted_finding_score": _weighted_finding_score(active_policy.active_findings),
        "summary": active_policy.summary,
        "baseline": {
            "path": active_policy.baseline_path,
            "new_only": active_policy.new_only,
        },
        "expired_suppression_ids": list(active_policy.expired_suppression_ids),
        "findings": [finding.model_dump(mode="json") for finding in active_policy.active_findings],
        "suppressed_findings": [item.model_dump() for item in active_policy.suppressed_findings],
        "baseline_findings": [
            finding.model_dump(mode="json") for finding in active_policy.baseline_findings
        ]
        if active_policy.new_only
        else [],
    }
    write_json(findings_path, findings_payload)
    write_json(snapshot_path, scan.snapshot.model_dump(mode="json"))
    write_json(sarif_path, build_sarif(active_scan, source_uri=logical_source))

    severity_counts = Counter(finding.severity.value for finding in active_policy.active_findings)
    rule_counts = Counter(finding.rule_id for finding in active_policy.active_findings)
    patch_map = {patch.id: patch for patch in scan.patches}
    environment = Environment(
        loader=PackageLoader("workbooklens.reports", "templates"),
        autoescape=select_autoescape(("html", "xml")),
    )
    template = environment.get_template("scan.html.j2")
    html = template.render(
        source=logical_source,
        source_sha256=scan.snapshot.source_sha256,
        weighted_finding_score=_weighted_finding_score(active_policy.active_findings),
        findings=active_policy.active_findings,
        suppressed_findings=active_policy.suppressed_findings,
        baseline_findings=active_policy.baseline_findings if active_policy.new_only else (),
        policy_summary=active_policy.summary,
        baseline_path=active_policy.baseline_path,
        new_only=active_policy.new_only,
        expired_suppression_ids=active_policy.expired_suppression_ids,
        patch_map=patch_map,
        sheets=scan.snapshot.sheets,
        severity_counts=severity_counts,
        rule_counts=sorted(rule_counts.items()),
        tool_version=__version__,
    )
    atomic_write_bytes(html_path, html.encode("utf-8"))
    return {
        "html": html_path,
        "findings": findings_path,
        "snapshot": snapshot_path,
        "sarif": sarif_path,
    }
