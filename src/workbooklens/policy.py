"""Finding baselines and explicit, reviewable suppression policy."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date
from fnmatch import fnmatchcase
from hashlib import sha256
from pathlib import Path, PurePosixPath
from typing import Any

from pydantic import BaseModel, ConfigDict, Field, model_validator

from workbooklens.exceptions import UsageError
from workbooklens.models import Finding


class FindingSuppression(BaseModel):
    """A documented waiver scoped to stable finding attributes."""

    model_config = ConfigDict(extra="forbid")

    id: str = Field(min_length=1)
    reason: str = Field(min_length=3)
    finding_ids: list[str] = Field(default_factory=list)
    rules: list[str] = Field(default_factory=list)
    sheets: list[str] = Field(default_factory=list)
    locations: list[str] = Field(default_factory=list)
    expires: date | None = None

    @model_validator(mode="after")
    def require_scope(self) -> FindingSuppression:
        if not any((self.finding_ids, self.rules, self.sheets, self.locations)):
            raise ValueError(
                "suppression must select at least one finding_id, rule, sheet, or location"
            )
        return self

    def is_expired(self, *, as_of: date) -> bool:
        return self.expires is not None and self.expires < as_of

    def matches(self, finding: Finding, *, as_of: date) -> bool:
        """Match every configured dimension, with shell-style globs for sheet/location."""

        if self.is_expired(as_of=as_of):
            return False
        if self.finding_ids and finding.id not in self.finding_ids:
            return False
        if self.rules and finding.rule_id not in self.rules:
            return False
        if self.sheets and (
            finding.sheet is None
            or not any(fnmatchcase(finding.sheet, pattern) for pattern in self.sheets)
        ):
            return False
        if self.locations:
            return finding.location is not None and any(
                fnmatchcase(finding.location, pattern) for pattern in self.locations
            )
        return True


@dataclass(frozen=True, slots=True)
class SuppressedFinding:
    """Finding plus the waiver that removed it from active policy gates."""

    finding: Finding
    suppression_id: str
    reason: str
    expires: date | None

    def model_dump(self) -> dict[str, Any]:
        return {
            "finding": self.finding.model_dump(mode="json"),
            "suppression": {
                "id": self.suppression_id,
                "reason": self.reason,
                "expires": self.expires.isoformat() if self.expires else None,
            },
        }


@dataclass(frozen=True, slots=True)
class FindingPolicyResult:
    """Classification used consistently by reports, SARIF, and CLI exit gates."""

    all_findings: tuple[Finding, ...]
    active_findings: tuple[Finding, ...]
    suppressed_findings: tuple[SuppressedFinding, ...]
    baseline_findings: tuple[Finding, ...]
    new_findings: tuple[Finding, ...]
    baseline_ids: frozenset[str]
    baseline_path: str | None
    source_scope: str | None
    new_only: bool
    expired_suppression_ids: tuple[str, ...]

    @property
    def summary(self) -> dict[str, int]:
        return {
            "total": len(self.all_findings),
            "active": len(self.active_findings),
            "suppressed": len(self.suppressed_findings),
            "baseline_known": len(self.baseline_findings),
            "new": len(self.new_findings),
        }


def normalize_source_scope(value: str) -> str:
    """Validate a portable repository-style logical workbook path."""

    normalized = value.replace("\\", "/").strip()
    raw_parts = normalized.split("/")
    scope = PurePosixPath(normalized)
    if (
        not normalized
        or scope.is_absolute()
        or any(part in {"", ".", ".."} for part in raw_parts)
        or ":" in raw_parts[0]
    ):
        raise UsageError("Source scope must be a nonempty relative POSIX-style path")
    return scope.as_posix()


def source_scope_for_path(
    path: Path,
    *,
    base: Path | None = None,
    explicit: str | None = None,
) -> str:
    """Return a stable logical path without storing an absolute local filesystem path."""

    if explicit is not None:
        return normalize_source_scope(explicit)
    resolved = path.expanduser().resolve()
    root = (base or Path.cwd()).expanduser().resolve()
    try:
        relative = resolved.relative_to(root)
    except ValueError:
        normalized_parent = os.path.normcase(str(resolved.parent))
        parent_token = sha256(normalized_parent.encode("utf-8")).hexdigest()[:12]
        return normalize_source_scope(f"external/{parent_token}/{resolved.name}")
    return normalize_source_scope(relative.as_posix())


def _finding_ids(items: Any, label: str) -> list[str | None]:
    if not isinstance(items, list):
        raise UsageError(f"Baseline '{label}' must be a JSON array")
    return [item.get("id") if isinstance(item, dict) else None for item in items]


def _validate_report_scope(raw: dict[str, Any], expected_source_scope: str | None) -> None:
    if expected_source_scope is None:
        return
    expected = normalize_source_scope(expected_source_scope)
    report_scope = raw.get("source_scope")
    if report_scope is not None:
        if not isinstance(report_scope, str):
            raise UsageError("Baseline 'source_scope' must be a string")
        if normalize_source_scope(report_scope) != expected:
            raise UsageError(f"Baseline source scope {report_scope!r} does not match {expected!r}")
        return
    legacy_source = raw.get("source")
    if not isinstance(legacy_source, str) or not legacy_source.strip():
        raise UsageError(
            "Baseline findings report must contain a nonempty 'source_scope' or legacy "
            "'source'; use 'finding_ids' for an intentional global baseline"
        )
    legacy_name = PurePosixPath(legacy_source.replace("\\", "/")).name
    if legacy_name != PurePosixPath(expected).name:
        raise UsageError(f"Baseline source {legacy_source!r} does not match {expected!r}")


def load_baseline(
    path: Path,
    max_bytes: int = 20 * 1024 * 1024,
    *,
    expected_source_scope: str | None = None,
) -> frozenset[str]:
    """Load IDs from a WorkbookLens findings report or a minimal baseline JSON file."""

    try:
        data = path.read_bytes()
    except OSError as exc:
        raise UsageError(f"Unable to read finding baseline {path}: {exc}") from exc
    if len(data) > max_bytes:
        raise UsageError(f"Finding baseline exceeds the {max_bytes}-byte limit")
    try:
        raw = json.loads(data.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise UsageError(f"Finding baseline must be valid UTF-8 JSON: {exc}") from exc

    identifiers: Any
    if isinstance(raw, list):
        identifiers = raw
    elif isinstance(raw, dict) and "finding_ids" in raw:
        identifiers = raw["finding_ids"]
    elif isinstance(raw, dict) and "workbooks" in raw:
        if expected_source_scope is None:
            raise UsageError("A workbook baseline manifest requires an expected source scope")
        workbooks = raw["workbooks"]
        if not isinstance(workbooks, dict):
            raise UsageError("Baseline 'workbooks' must be a JSON object")
        normalized_workbooks: dict[str, Any] = {}
        for scope, scoped_ids in workbooks.items():
            if not isinstance(scope, str):
                raise UsageError("Baseline workbook scopes must be strings")
            normalized_scope = normalize_source_scope(scope)
            if normalized_scope in normalized_workbooks:
                raise UsageError(f"Duplicate baseline workbook scope: {normalized_scope}")
            normalized_workbooks[normalized_scope] = scoped_ids
        identifiers = normalized_workbooks.get(normalize_source_scope(expected_source_scope), [])
    elif isinstance(raw, dict) and "findings" in raw:
        _validate_report_scope(raw, expected_source_scope)
        finding_groups = [("findings", raw["findings"])]
        if "baseline_findings" in raw:
            finding_groups.append(("baseline_findings", raw["baseline_findings"]))
        identifiers = []
        for label, findings in finding_groups:
            identifiers.extend(_finding_ids(findings, label))
    else:
        raise UsageError(
            "Baseline must be an ID array, contain 'finding_ids' or 'workbooks', or be a findings.json report"
        )
    if not isinstance(identifiers, list) or any(
        not isinstance(identifier, str) or not identifier for identifier in identifiers
    ):
        raise UsageError("Baseline finding IDs must be nonempty strings")
    return frozenset(identifiers)


def apply_finding_policy(
    findings: list[Finding],
    *,
    suppressions: list[FindingSuppression] | None = None,
    baseline_ids: frozenset[str] | None = None,
    baseline_path: Path | None = None,
    source_scope: str | None = None,
    new_only: bool = False,
    as_of: date | None = None,
) -> FindingPolicyResult:
    """Classify findings without mutating the scanner's deterministic raw result."""

    effective_date = as_of or date.today()
    configured_suppressions = suppressions or []
    known_ids = baseline_ids or frozenset()
    expired = tuple(
        suppression.id
        for suppression in configured_suppressions
        if suppression.is_expired(as_of=effective_date)
    )
    suppressed: list[SuppressedFinding] = []
    unsuppressed: list[Finding] = []
    for finding in findings:
        waiver = next(
            (
                suppression
                for suppression in configured_suppressions
                if suppression.matches(finding, as_of=effective_date)
            ),
            None,
        )
        if waiver is None:
            unsuppressed.append(finding)
            continue
        suppressed.append(
            SuppressedFinding(
                finding=finding,
                suppression_id=waiver.id,
                reason=waiver.reason,
                expires=waiver.expires,
            )
        )

    baseline_findings = [finding for finding in unsuppressed if finding.id in known_ids]
    new_findings = [finding for finding in unsuppressed if finding.id not in known_ids]
    active = new_findings if new_only else unsuppressed
    return FindingPolicyResult(
        all_findings=tuple(findings),
        active_findings=tuple(active),
        suppressed_findings=tuple(suppressed),
        baseline_findings=tuple(baseline_findings),
        new_findings=tuple(new_findings),
        baseline_ids=known_ids,
        baseline_path=source_scope_for_path(baseline_path) if baseline_path else None,
        source_scope=normalize_source_scope(source_scope) if source_scope else None,
        new_only=new_only,
        expired_suppression_ids=expired,
    )


__all__ = [
    "FindingPolicyResult",
    "FindingSuppression",
    "SuppressedFinding",
    "apply_finding_policy",
    "load_baseline",
    "normalize_source_scope",
    "source_scope_for_path",
]
