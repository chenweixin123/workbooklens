"""Typed public domain models used by WorkbookLens APIs and reports."""

from __future__ import annotations

from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, RootModel, field_validator


class StrictModel(BaseModel):
    """Base class for stable, strict, forward-compatible public models."""

    model_config = ConfigDict(extra="forbid", frozen=False, validate_assignment=True)


class Severity(StrEnum):
    """Finding severity ordered from informational to release-blocking."""

    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    CRITICAL = "critical"


SEVERITY_RANK: dict[Severity, int] = {
    Severity.INFO: 0,
    Severity.WARNING: 1,
    Severity.ERROR: 2,
    Severity.CRITICAL: 3,
}


class Confidence(RootModel[float]):
    """Validated confidence score in the inclusive interval [0, 1]."""

    @field_validator("root")
    @classmethod
    def validate_range(cls, value: float) -> float:
        if not 0.0 <= value <= 1.0:
            raise ValueError("confidence must be between 0 and 1")
        return value

    def __float__(self) -> float:
        return self.root


class CellSnapshot(StrictModel):
    """Serializable semantic state of one relevant workbook cell."""

    coordinate: str
    value: Any = None
    formula: str | None = None
    data_type: str
    style_id: int = 0
    style_fingerprint: str = ""
    number_format: str = "General"
    row_hidden: bool = False
    column_hidden: bool = False


class Region(StrictModel):
    """A conservatively inferred rectangular or one-dimensional workbook region."""

    sheet: str
    min_row: int
    max_row: int
    min_column: int
    max_column: int
    kind: Literal["data", "formula_row", "formula_column", "style"]
    confidence: Confidence


class SheetSnapshot(StrictModel):
    """Serializable sheet structure and all semantically relevant populated cells."""

    name: str
    index: int
    state: Literal["visible", "hidden", "veryHidden"]
    max_row: int
    max_column: int
    cells: dict[str, CellSnapshot] = Field(default_factory=dict)
    merged_ranges: list[str] = Field(default_factory=list)
    hidden_rows: list[int] = Field(default_factory=list)
    hidden_columns: list[str] = Field(default_factory=list)
    data_validations: list[str] = Field(default_factory=list)


class WorkbookSnapshot(StrictModel):
    """Deterministic semantic snapshot of a workbook without formula execution."""

    source_name: str
    source_sha256: str
    format: Literal["xlsx", "xlsm"]
    sheets: list[SheetSnapshot]
    defined_names: dict[str, str] = Field(default_factory=dict)
    calculation_mode: str | None = None


class Evidence(StrictModel):
    """Observed facts and peer context that justify a finding."""

    summary: str
    observed: Any = None
    expected: Any = None
    peers: list[str] = Field(default_factory=list)
    details: dict[str, Any] = Field(default_factory=dict)


class PatchKind(StrEnum):
    """Declarative operations supported by the preservation-oriented patch engine."""

    SET_FORMULA = "set_formula"
    SET_NUMERIC = "set_numeric"
    COPY_STYLE = "copy_style"
    EXTEND_SUM = "extend_sum"
    CREATE_FORMULA = "create_formula"


class PatchPrecondition(StrictModel):
    """Expected source cell state used to reject stale or mismatched plans."""

    cell_fingerprint: str
    expected_value: Any = None
    expected_formula: str | None = None
    expected_style_id: int | None = None


class PatchOperation(StrictModel):
    """Reviewable, serializable workbook patch operation."""

    id: str
    kind: PatchKind
    sheet: str
    cell: str
    before: Any = None
    after: Any = None
    source_cell: str | None = None
    confidence: Confidence
    safe: bool
    description: str
    precondition: PatchPrecondition


class Finding(StrictModel):
    """One deterministic workbook quality finding with auditable evidence."""

    id: str
    content_fingerprint: str = ""
    rule_id: str
    title: str
    explanation: str
    severity: Severity
    confidence: Confidence
    workbook: str
    sheet: str | None = None
    location: str | None = None
    evidence: Evidence
    expected: str
    suggested_action: str
    safe_patch_available: bool = False
    patch_ids: list[str] = Field(default_factory=list)


class PatchPlan(StrictModel):
    """A source-bound collection of proposed patches for explicit review."""

    schema_version: Literal[1] = 1
    tool_version: str
    source_name: str
    source_sha256: str
    patches: list[PatchOperation]
    finding_ids: list[str] = Field(default_factory=list)
    findings: list[Finding] = Field(default_factory=list)


class PackageChange(StrictModel):
    """Hash-level record of a changed OOXML package part."""

    part: str
    action: Literal["modified", "added", "removed"]
    before_sha256: str | None = None
    after_sha256: str | None = None


class PatchResult(StrictModel):
    """Validated result and package manifest for an apply operation."""

    source_sha256: str
    output_sha256: str
    output_path: str
    applied_patch_ids: list[str]
    package_changes: list[PackageChange]
    resolved_finding_ids: list[str] = Field(default_factory=list)
    remaining_finding_ids: list[str] = Field(default_factory=list)
    new_finding_ids: list[str] = Field(default_factory=list)
    validation_messages: list[str] = Field(default_factory=list)


class CellChange(StrictModel):
    """One semantic cell-level difference."""

    sheet: str
    cell: str
    change_type: Literal["value", "formula", "style", "number_format"]
    before: Any = None
    after: Any = None
    importance: Severity = Severity.WARNING
    before_signature: str | None = None
    after_signature: str | None = None


class StructuralChange(StrictModel):
    """One workbook- or sheet-structure difference."""

    change_type: str
    subject: str
    before: Any = None
    after: Any = None
    importance: Severity = Severity.WARNING


class WorkbookDiff(StrictModel):
    """Semantic workbook comparison independent of ZIP serialization details."""

    before_sha256: str
    after_sha256: str
    cell_changes: list[CellChange] = Field(default_factory=list)
    structural_changes: list[StructuralChange] = Field(default_factory=list)


class WorkbookAssertion(StrictModel):
    """A supported YAML assertion in normalized model form."""

    id: str
    type: Literal[
        "no_findings",
        "unique",
        "equals",
        "allowed_values",
        "nonblank",
        "numeric_bounds",
    ]
    sheet: str | None = None
    range: str | None = None
    rules: list[str] = Field(default_factory=list)
    ignore_blank: bool = True
    left: str | None = None
    right: str | None = None
    tolerance: float = 0.0
    values: list[Any] = Field(default_factory=list)
    minimum: float | None = None
    maximum: float | None = None


class AssertionResult(StrictModel):
    """Pass/fail result with concrete evidence for one workbook assertion."""

    assertion_id: str
    passed: bool
    message: str
    observed: Any = None
    expected: Any = None


__all__ = [
    "SEVERITY_RANK",
    "AssertionResult",
    "CellChange",
    "CellSnapshot",
    "Confidence",
    "Evidence",
    "Finding",
    "PackageChange",
    "PatchKind",
    "PatchOperation",
    "PatchPlan",
    "PatchPrecondition",
    "PatchResult",
    "Region",
    "Severity",
    "SheetSnapshot",
    "StructuralChange",
    "WorkbookAssertion",
    "WorkbookDiff",
    "WorkbookSnapshot",
]
