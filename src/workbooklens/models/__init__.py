"""Typed public domain models used by WorkbookLens APIs and reports."""

from __future__ import annotations

import re
from enum import StrEnum
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, RootModel, field_validator, model_validator


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
    declared_dimension: str | None = None
    content_dimension: str | None = None
    row_heights: dict[str, float] = Field(default_factory=dict)
    column_widths: dict[str, float] = Field(default_factory=dict)
    view_top_left_cell: str | None = None
    view_zoom_scale: int | None = None
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
    NORMALIZE_TEXT = "normalize_text"
    COPY_NUMBER_FORMAT = "copy_number_format"
    COPY_STYLE = "copy_style"
    EXTEND_SUM = "extend_sum"
    CREATE_FORMULA = "create_formula"
    SET_COLUMN_WIDTH = "set_column_width"
    SET_ROW_HEIGHT = "set_row_height"
    SET_WRAP_TEXT = "set_wrap_text"
    SET_SHRINK_TO_FIT = "set_shrink_to_fit"
    SET_TEXT = "set_text"
    SET_SHEET_VIEW = "set_sheet_view"
    COPY_BORDER = "copy_border"
    CLEAR_FORMATTING_TAIL = "clear_formatting_tail"
    REMOVE_WHITESPACE_TAIL_CELLS = "remove_whitespace_tail_cells"


class PatchRisk(StrEnum):
    """Review boundary for repair selection, independent of rule confidence."""

    SAFE = "safe"
    LAYOUT_REVIEW = "layout_review"
    FORMULA_DERIVED = "formula_derived"
    SEMANTIC_REVIEW = "semantic_review"


class RecalculationProvider(StrEnum):
    """Local application used to recalculate isolated validation copies."""

    NONE = "none"
    EXCEL = "excel"
    LIBREOFFICE = "libreoffice"


class ValidationStatus(StrEnum):
    """Overall validation outcome recorded for an apply operation."""

    NOT_RUN = "not_run"
    PASSED = "passed"
    DEGRADED = "degraded"
    FAILED = "failed"


LAYOUT_REVIEW_PATCH_KINDS = frozenset(
    {
        PatchKind.SET_COLUMN_WIDTH,
        PatchKind.SET_ROW_HEIGHT,
        PatchKind.SET_WRAP_TEXT,
        PatchKind.SET_SHRINK_TO_FIT,
        PatchKind.SET_TEXT,
        PatchKind.SET_SHEET_VIEW,
        PatchKind.COPY_BORDER,
        PatchKind.CLEAR_FORMATTING_TAIL,
        PatchKind.REMOVE_WHITESPACE_TAIL_CELLS,
    }
)

FORMULA_PATCH_KINDS = frozenset(
    {
        PatchKind.SET_FORMULA,
        PatchKind.EXTEND_SUM,
        PatchKind.CREATE_FORMULA,
    }
)


class PatchDerivation(StrictModel):
    """Machine-readable proof obligations behind one proposed patch."""

    strategy: str = "deterministic_rule"
    sources: list[str] = Field(default_factory=lambda: ["finding_evidence"], min_length=1)
    candidate_count: int = Field(default=1, ge=1)
    invariants: list[str] = Field(
        default_factory=lambda: ["source_precondition"],
        min_length=1,
    )
    requires_recalculation: bool = False

    @field_validator("strategy")
    @classmethod
    def require_nonblank_strategy(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("derivation strategy must not be blank")
        return normalized

    @field_validator("sources", "invariants")
    @classmethod
    def require_distinct_nonblank_entries(cls, value: list[str]) -> list[str]:
        normalized = [entry.strip() for entry in value]
        if any(not entry for entry in normalized):
            raise ValueError("derivation evidence entries must not be blank")
        folded = [entry.casefold() for entry in normalized]
        if len(folded) != len(set(folded)):
            raise ValueError("derivation evidence entries must be distinct")
        return normalized

    @property
    def source_families(self) -> tuple[str, ...]:
        """Return validated evidence-family prefixes, or an empty tuple if unstructured."""

        families: list[str] = []
        for source in self.sources:
            family, separator, detail = source.partition(":")
            if (
                not separator
                or not detail.strip()
                or not any(character.isalnum() for character in detail)
                or re.fullmatch(r"[a-z][a-z0-9_]*", family) is None
            ):
                return ()
            families.append(family)
        return tuple(families)


class PatchPrecondition(StrictModel):
    """Expected source cell state used to reject stale or mismatched plans."""

    cell_fingerprint: str
    expected_value: Any = None
    expected_formula: str | None = None
    expected_style_id: int | None = None
    layout_fingerprint: str | None = None


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
    risk: PatchRisk = PatchRisk.SAFE
    description: str
    precondition: PatchPrecondition
    atomic_group: str | None = None
    prerequisite_patch_ids: list[str] = Field(default_factory=list)
    derivation: PatchDerivation = Field(default_factory=PatchDerivation)

    @model_validator(mode="before")
    @classmethod
    def infer_layout_review_risk(cls, value: Any) -> Any:
        """Keep legacy constructors working while never inferring layout safety from confidence."""

        if not isinstance(value, dict) or "risk" in value:
            return value
        kind = value.get("kind")
        try:
            if isinstance(kind, PatchKind):
                patch_kind = kind
            elif isinstance(kind, str):
                patch_kind = PatchKind(kind)
            else:
                return value
        except ValueError:
            return value
        if patch_kind not in LAYOUT_REVIEW_PATCH_KINDS:
            return value
        normalized = dict(value)
        normalized["risk"] = PatchRisk.LAYOUT_REVIEW
        return normalized

    @model_validator(mode="after")
    def require_risk_invariants(self) -> PatchOperation:
        """Prevent callers from bypassing a patch risk boundary."""

        if len(self.prerequisite_patch_ids) != len(set(self.prerequisite_patch_ids)):
            raise ValueError("prerequisite patch IDs must be distinct")
        if self.id in self.prerequisite_patch_ids:
            raise ValueError("a patch cannot depend on itself")
        if self.kind in LAYOUT_REVIEW_PATCH_KINDS and self.risk != PatchRisk.LAYOUT_REVIEW:
            raise ValueError(f"{self.kind.value} patches require risk='layout_review'")
        if self.risk == PatchRisk.LAYOUT_REVIEW and self.safe:
            raise ValueError("layout_review patches require safe=false")
        if self.risk == PatchRisk.SEMANTIC_REVIEW and self.safe:
            raise ValueError("semantic_review patches require safe=false")
        if self.kind == PatchKind.COPY_NUMBER_FORMAT:
            if not self.source_cell:
                raise ValueError("copy_number_format patches require source_cell")
            if not isinstance(self.before, str) or not isinstance(self.after, str):
                raise ValueError("copy_number_format patches require string formats")
        if self.risk == PatchRisk.FORMULA_DERIVED:
            if self.kind not in FORMULA_PATCH_KINDS:
                raise ValueError("formula_derived risk requires a formula patch kind")
            if self.safe:
                raise ValueError("formula_derived patches require safe=false before validation")
            if float(self.confidence) < 0.99:
                raise ValueError("formula_derived patches require confidence >= 0.99")
            if self.derivation.candidate_count != 1:
                raise ValueError("formula_derived patches require exactly one candidate")
            if len(self.derivation.sources) < 2:
                raise ValueError(
                    "formula_derived patches require at least two independent evidence sources"
                )
            source_families = self.derivation.source_families
            if len(source_families) != len(self.derivation.sources):
                raise ValueError(
                    "formula_derived evidence sources require structured family:detail labels"
                )
            if len(set(source_families)) < 2:
                raise ValueError(
                    "formula_derived patches require at least two independent evidence families"
                )
            if not self.derivation.requires_recalculation:
                raise ValueError("formula_derived patches require recalculation validation")
        return self

    @property
    def safe_only_eligible(self) -> bool:
        """Return whether ``--safe-only`` may select this operation."""

        return self.safe and self.risk == PatchRisk.SAFE and float(self.confidence) >= 0.95


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

    schema_version: Literal[3] = 3
    tool_version: str
    source_name: str
    source_sha256: str
    patches: list[PatchOperation]
    finding_ids: list[str] = Field(default_factory=list)
    findings: list[Finding] = Field(default_factory=list)

    @model_validator(mode="before")
    @classmethod
    def migrate_schema_two(cls, value: Any) -> Any:
        """Read legacy v2 plans as v3 models without ever serializing v2 again."""

        if not isinstance(value, dict) or value.get("schema_version", 3) != 2:
            return value
        migrated = dict(value)
        migrated["schema_version"] = 3
        patches = value.get("patches")
        if isinstance(patches, list):
            migrated_patches: list[Any] = []
            formula_kinds = {member.value for member in FORMULA_PATCH_KINDS}
            for patch in patches:
                if not isinstance(patch, dict):
                    migrated_patches.append(patch)
                    continue
                migrated_patch = dict(patch)
                kind = migrated_patch.get("kind")
                kind_value = kind.value if isinstance(kind, PatchKind) else kind
                if kind_value in formula_kinds:
                    migrated_patch.update(
                        safe=False,
                        risk=PatchRisk.SEMANTIC_REVIEW,
                        derivation={
                            "strategy": "legacy_schema_v2_unverified_formula",
                            "sources": ["legacy_plan:schema_v2"],
                            "candidate_count": 1,
                            "invariants": ["manual_review_required"],
                            "requires_recalculation": True,
                        },
                    )
                migrated_patches.append(migrated_patch)
            migrated["patches"] = migrated_patches
        return migrated

    @model_validator(mode="after")
    def validate_patch_dependencies(self) -> PatchPlan:
        """Reject missing or cyclic patch dependencies before any selection is resolved."""

        by_id = {patch.id: patch for patch in self.patches}
        if len(by_id) != len(self.patches):
            raise ValueError("patch plan contains duplicate patch IDs")
        for patch in self.patches:
            missing = sorted(set(patch.prerequisite_patch_ids) - by_id.keys())
            if missing:
                raise ValueError(
                    f"patch {patch.id} references missing prerequisite patches: "
                    + ", ".join(missing)
                )

        visiting: set[str] = set()
        visited: set[str] = set()

        def visit(patch_id: str) -> None:
            if patch_id in visiting:
                raise ValueError("patch prerequisite graph contains a cycle")
            if patch_id in visited:
                return
            visiting.add(patch_id)
            for prerequisite_id in by_id[patch_id].prerequisite_patch_ids:
                visit(prerequisite_id)
            visiting.remove(patch_id)
            visited.add(patch_id)

        for patch_id in by_id:
            visit(patch_id)
        return self


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
    recalculation_provider: RecalculationProvider = RecalculationProvider.NONE
    formula_errors_before: list[str] = Field(default_factory=list)
    formula_errors_after: list[str] = Field(default_factory=list)
    downgraded_patch_ids: list[str] = Field(default_factory=list)
    skipped_patch_ids: list[str] = Field(default_factory=list)
    validation_status: ValidationStatus = ValidationStatus.NOT_RUN
    rollback_performed: bool = False
    failure_report_path: str | None = None


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
    "PatchDerivation",
    "PatchKind",
    "PatchOperation",
    "PatchPlan",
    "PatchPrecondition",
    "PatchResult",
    "PatchRisk",
    "RecalculationProvider",
    "Region",
    "Severity",
    "SheetSnapshot",
    "StructuralChange",
    "ValidationStatus",
    "WorkbookAssertion",
    "WorkbookDiff",
    "WorkbookSnapshot",
]
