"""Public interfaces for first-party and third-party workbook rules."""

from __future__ import annotations

from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.cell_range import CellRange

from workbooklens.models import Finding, PatchOperation, Region, WorkbookSnapshot


@dataclass(slots=True)
class RuleContext:
    """Read-only semantic inputs shared by deterministic rules."""

    path: Path
    workbook: Workbook
    snapshot: WorkbookSnapshot
    config: dict[str, Any]
    data_regions: dict[str, list[Region]]
    formula_bands: dict[str, list[Region]]
    unsupported_formula_ranges: dict[str, tuple[CellRange, ...]] = field(default_factory=dict)


@dataclass(slots=True)
class RuleResult:
    """Findings plus any declarative patches proposed by a rule."""

    findings: list[Finding] = field(default_factory=list)
    patches: list[PatchOperation] = field(default_factory=list)

    def extend(self, other: RuleResult) -> None:
        """Append another result while preserving rule execution order."""

        self.findings.extend(other.findings)
        self.patches.extend(other.patches)


class WorkbookRule(ABC):
    """Plugin interface: implement ``run`` and register an instance."""

    rule_id: str
    title: str

    @abstractmethod
    def run(self, context: RuleContext) -> RuleResult:
        """Analyze a workbook without mutating or evaluating it."""
