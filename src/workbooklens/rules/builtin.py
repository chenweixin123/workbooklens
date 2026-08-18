"""The fifteen deterministic WorkbookLens v0.1 rules."""

from __future__ import annotations

import math
import re
from collections import Counter, defaultdict
from collections.abc import Iterable, Sequence
from decimal import Decimal, InvalidOperation
from typing import Any, cast

from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import (
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.formulas import (
    UnsupportedFormulaError,
    analyze_formula,
    normalize_formula,
    translate_formula,
)
from workbooklens.models import (
    Confidence,
    Evidence,
    Finding,
    PatchKind,
    PatchOperation,
    PatchPrecondition,
    Region,
    Severity,
)
from workbooklens.rules.base import RuleContext, RuleResult, WorkbookRule
from workbooklens.snapshot import cell_fingerprint
from workbooklens.utils import stable_id

EXCEL_ERRORS = {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"}
SIMPLE_SUM_RE = re.compile(
    r"^=SUM\((?P<sheet>(?:'(?:[^']|'')+'|[^'!]+)!)?"
    r"(?P<col1>\$?[A-Z]{1,3})(?P<start>\$?\d+):"
    r"(?P<col2>\$?[A-Z]{1,3})(?P<end>\$?\d+)\)$",
    re.IGNORECASE,
)


def _confidence(value: float) -> Confidence:
    return Confidence(max(0.0, min(1.0, value)))


def _make_finding(
    *,
    context: RuleContext,
    rule_id: str,
    title: str,
    explanation: str,
    severity: Severity,
    confidence: float,
    sheet: str | None,
    location: str | None,
    evidence: Evidence,
    expected: str,
    suggested_action: str,
    patches: Sequence[PatchOperation] = (),
) -> Finding:
    identifier = stable_id("finding", rule_id, sheet, location, evidence.model_dump(mode="json"))
    return Finding(
        id=identifier,
        rule_id=rule_id,
        title=title,
        explanation=explanation,
        severity=severity,
        confidence=_confidence(confidence),
        workbook=context.path.name,
        sheet=sheet,
        location=location,
        evidence=evidence,
        expected=expected,
        suggested_action=suggested_action,
        safe_patch_available=any(patch.safe for patch in patches),
        patch_ids=[patch.id for patch in patches],
    )


def _make_patch(
    *,
    kind: PatchKind,
    worksheet: Worksheet,
    cell: Cell,
    before: Any,
    after: Any,
    confidence: float,
    description: str,
    source_cell: str | None = None,
) -> PatchOperation:
    safe = confidence >= 0.95
    expected_formula = cell.value if cell.data_type == "f" and isinstance(cell.value, str) else None
    patch_id = stable_id(
        "patch", kind.value, worksheet.title, cell.coordinate, before, after, source_cell
    )
    return PatchOperation(
        id=patch_id,
        kind=kind,
        sheet=worksheet.title,
        cell=cell.coordinate,
        before=before,
        after=after,
        source_cell=source_cell,
        confidence=_confidence(confidence),
        safe=safe,
        description=description,
        precondition=PatchPrecondition(
            cell_fingerprint=cell_fingerprint(cell),
            expected_value=None if cell.data_type == "f" else cell.value,
            expected_formula=expected_formula,
            expected_style_id=cell.style_id,
        ),
    )


def _cell(worksheet: Worksheet, row: int, column: int) -> Cell:
    return cast(Cell, worksheet.cell(row, column))


def _band_cells(worksheet: Worksheet, band: Region) -> list[Cell]:
    if band.kind == "formula_row":
        return [
            _cell(worksheet, band.min_row, column)
            for column in range(band.min_column, band.max_column + 1)
        ]
    return [_cell(worksheet, row, band.min_column) for row in range(band.min_row, band.max_row + 1)]


def _formula_signature(cell: Cell) -> str | None:
    if cell.data_type != "f" or not isinstance(cell.value, str):
        return None
    try:
        features = analyze_formula(cell.value)
        if features.external_references or features.unsupported_reason:
            return None
        return normalize_formula(cell.value, cell.coordinate)
    except (ValueError, UnsupportedFormulaError):
        return None


def _in_unsupported_formula_range(context: RuleContext, sheet: str, coordinate: str) -> bool:
    return any(
        coordinate in formula_range
        for formula_range in context.unsupported_formula_ranges.get(sheet, ())
    )


def _band_has_unsupported_formula(context: RuleContext, sheet: str, cells: Sequence[Cell]) -> bool:
    return any(_in_unsupported_formula_range(context, sheet, cell.coordinate) for cell in cells)


def _translated_consensus(
    target: Cell, peers: Iterable[Cell], required_signature: str | None = None
) -> tuple[str, str] | None:
    translated: list[tuple[str, str]] = []
    for peer in peers:
        if peer.data_type != "f" or not isinstance(peer.value, str):
            continue
        signature = _formula_signature(peer)
        if signature is None or (
            required_signature is not None and signature != required_signature
        ):
            continue
        try:
            value = translate_formula(peer.value, peer.coordinate, target.coordinate)
        except UnsupportedFormulaError:
            continue
        translated.append((value, peer.coordinate))
    if len(translated) < 2:
        return None
    counts = Counter(value for value, _ in translated)
    formula, count = counts.most_common(1)[0]
    if count < 2 or len(counts) != 1:
        return None
    source = min(source for value, source in translated if value == formula)
    return formula, source


class BrokenReferenceRule(WorkbookRule):
    rule_id = "WL001_BROKEN_REFERENCE"
    title = "Broken formula reference"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in worksheet._cells.values():
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                if "#REF!" not in cell.value.upper():
                    continue
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="The formula contains an explicit #REF! token and cannot resolve as written.",
                        severity=Severity.ERROR,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(summary="Formula contains #REF!", observed=cell.value),
                        expected="Every formula reference resolves to an existing cell or range.",
                        suggested_action="Review the deleted or moved source range; no automatic guess was made.",
                    )
                )
        return result


class FormulaPatternOutlierRule(WorkbookRule):
    rule_id = "WL002_FORMULA_PATTERN_OUTLIER"
    title = "Formula pattern outlier"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for band in context.formula_bands[worksheet.title]:
                cells = _band_cells(worksheet, band)
                if _band_has_unsupported_formula(context, worksheet.title, cells):
                    continue
                formula_cells = [cell for cell in cells if cell.data_type == "f"]
                signature_by_cell = {
                    cell.coordinate: _formula_signature(cell) for cell in formula_cells
                }
                if any(signature is None for signature in signature_by_cell.values()):
                    continue
                signatures = [
                    cast(str, signature_by_cell[cell.coordinate]) for cell in formula_cells
                ]
                if len(signatures) < 4:
                    continue
                counts = Counter(signatures)
                consensus, consensus_count = counts.most_common(1)[0]
                ratio = consensus_count / len(signatures)
                outliers = [
                    cell
                    for cell in formula_cells
                    if signature_by_cell[cell.coordinate] != consensus
                ]
                if ratio < 0.8 or len(outliers) != 1:
                    continue
                cell = outliers[0]
                if (worksheet.title, cell.coordinate) in seen:
                    continue
                seen.add((worksheet.title, cell.coordinate))
                proposal = _translated_consensus(cell, formula_cells, consensus)
                patches: list[PatchOperation] = []
                if proposal is not None and ratio >= 0.95:
                    formula, source = proposal
                    patch = _make_patch(
                        kind=PatchKind.SET_FORMULA,
                        worksheet=worksheet,
                        cell=cell,
                        before=cell.value,
                        after=formula,
                        confidence=ratio,
                        source_cell=source,
                        description="Replace the one-off formula with the exact translated peer consensus.",
                    )
                    patches.append(patch)
                    result.patches.append(patch)
                peers = [
                    item.coordinate
                    for item in formula_cells
                    if signature_by_cell[item.coordinate] == consensus
                ]
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="One formula has a different relative-reference signature from a strong band consensus.",
                        severity=Severity.WARNING,
                        confidence=ratio,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary=f"{consensus_count} of {len(signatures)} formulas share one signature",
                            observed=cell.value,
                            expected=consensus,
                            peers=peers[:12],
                        ),
                        expected="Copied formulas in this band have the same structural signature.",
                        suggested_action="Compare the cell with the listed peers and review any proposed formula.",
                        patches=patches,
                    )
                )
        return result


class BlankInFormulaBandRule(WorkbookRule):
    rule_id = "WL003_BLANK_IN_FORMULA_BAND"
    title = "Blank interrupts formula band"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for band in context.formula_bands[worksheet.title]:
                cells = _band_cells(worksheet, band)
                if _band_has_unsupported_formula(context, worksheet.title, cells):
                    continue
                blanks = [cell for cell in cells[1:-1] if cell.value is None]
                formula_cells = [cell for cell in cells if cell.data_type == "f"]
                raw_signatures = [_formula_signature(cell) for cell in formula_cells]
                if any(signature is None for signature in raw_signatures):
                    continue
                signatures = [cast(str, signature) for signature in raw_signatures]
                if len(blanks) != 1 or len(signatures) < 3:
                    continue
                consensus, count = Counter(signatures).most_common(1)[0]
                if count / len(signatures) < 0.9:
                    continue
                cell = blanks[0]
                if (worksheet.title, cell.coordinate) in seen:
                    continue
                seen.add((worksheet.title, cell.coordinate))
                proposal = _translated_consensus(cell, formula_cells, consensus)
                if proposal is None:
                    continue
                formula, source = proposal
                patch = _make_patch(
                    kind=PatchKind.CREATE_FORMULA,
                    worksheet=worksheet,
                    cell=cell,
                    before=None,
                    after=formula,
                    confidence=0.99,
                    source_cell=source,
                    description="Create the missing cell with the exact translated formula agreed by peers.",
                )
                result.patches.append(patch)
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A single blank lies between formulas whose translations agree exactly at this cell.",
                        severity=Severity.ERROR,
                        confidence=0.99,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary="Independent neighboring formulas translate to the same expression",
                            observed=None,
                            expected=formula,
                            peers=[peer.coordinate for peer in formula_cells[:12]],
                        ),
                        expected="The contiguous formula band has no unexplained blank.",
                        suggested_action="Review and select the proposed translated formula.",
                        patches=[patch],
                    )
                )
        return result


class HardcodedValueInFormulaBandRule(WorkbookRule):
    rule_id = "WL004_HARDCODED_VALUE_IN_FORMULA_BAND"
    title = "Hardcoded value interrupts formula band"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for band in context.formula_bands[worksheet.title]:
                cells = _band_cells(worksheet, band)
                if _band_has_unsupported_formula(context, worksheet.title, cells):
                    continue
                literals = [
                    cell for cell in cells[1:-1] if cell.value is not None and cell.data_type != "f"
                ]
                formula_cells = [cell for cell in cells if cell.data_type == "f"]
                raw_signatures = [_formula_signature(cell) for cell in formula_cells]
                if any(signature is None for signature in raw_signatures):
                    continue
                signatures = [cast(str, signature) for signature in raw_signatures]
                if len(literals) != 1 or len(signatures) < 3:
                    continue
                consensus, count = Counter(signatures).most_common(1)[0]
                ratio = count / len(signatures)
                if ratio < 0.9:
                    continue
                cell = literals[0]
                if (worksheet.title, cell.coordinate) in seen:
                    continue
                seen.add((worksheet.title, cell.coordinate))
                proposal = _translated_consensus(cell, formula_cells, consensus)
                if proposal is None:
                    continue
                formula, source = proposal
                confidence = 0.99 if ratio >= 0.95 else 0.94
                patches: list[PatchOperation] = []
                if confidence >= 0.95:
                    patch = _make_patch(
                        kind=PatchKind.SET_FORMULA,
                        worksheet=worksheet,
                        cell=cell,
                        before=cell.value,
                        after=formula,
                        confidence=confidence,
                        source_cell=source,
                        description="Replace the isolated literal with the exact translated peer formula.",
                    )
                    patches.append(patch)
                    result.patches.append(patch)
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A literal value replaces one cell in an otherwise consistent copied-formula band.",
                        severity=Severity.ERROR,
                        confidence=confidence,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary="Peer formulas translate to one exact replacement",
                            observed=cell.value,
                            expected=formula,
                            peers=[peer.coordinate for peer in formula_cells[:12]],
                        ),
                        expected="The formula band follows its consensus structure.",
                        suggested_action="Confirm the literal is not an intentional override before selecting the patch.",
                        patches=patches,
                    )
                )
        return result


class SuspiciousSumBoundaryRule(WorkbookRule):
    rule_id = "WL005_SUSPICIOUS_SUM_BOUNDARY"
    title = "Suspicious SUM boundary"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in worksheet._cells.values():
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                match = SIMPLE_SUM_RE.fullmatch(cell.value)
                if match is None:
                    continue
                if match.group("sheet") or _in_unsupported_formula_range(
                    context, worksheet.title, cell.coordinate
                ):
                    continue
                first_column = match.group("col1").replace("$", "").upper()
                second_column = match.group("col2").replace("$", "").upper()
                if first_column != second_column or first_column != cell.column_letter:
                    continue
                start_row = int(match.group("start").replace("$", ""))
                end_row = int(match.group("end").replace("$", ""))
                candidate_row = end_row + 1
                if candidate_row != cell.row - 1 or start_row >= end_row:
                    continue
                candidate = _cell(worksheet, candidate_row, cell.column)
                if isinstance(candidate.value, bool) or not isinstance(
                    candidate.value, (int, float)
                ):
                    continue
                if worksheet.row_dimensions[candidate_row].hidden:
                    continue
                if any(candidate.coordinate in merged for merged in worksheet.merged_cells.ranges):
                    continue
                raw_end = match.group("end")
                replacement_end = (
                    "$" + str(candidate_row) if raw_end.startswith("$") else str(candidate_row)
                )
                formula = (
                    cell.value[: match.start("end")]
                    + replacement_end
                    + cell.value[match.end("end") :]
                )
                patch = _make_patch(
                    kind=PatchKind.EXTEND_SUM,
                    worksheet=worksheet,
                    cell=cell,
                    before=cell.value,
                    after=formula,
                    confidence=0.96,
                    source_cell=candidate.coordinate,
                    description="Extend the simple SUM through the directly adjacent numeric peer.",
                )
                result.patches.append(patch)
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A simple total stops one row before a directly adjacent non-hidden numeric peer.",
                        severity=Severity.WARNING,
                        confidence=0.96,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary=f"SUM ends at row {end_row}, while {candidate.coordinate} is adjacent",
                            observed=cell.value,
                            expected=formula,
                            peers=[candidate.coordinate],
                        ),
                        expected="A simple contiguous total includes its directly adjacent peer row.",
                        suggested_action="Check for intentional exclusions or subtotals, then select the patch if appropriate.",
                        patches=[patch],
                    )
                )
        return result


def _numeric_text(value: Any) -> int | float | None:
    if not isinstance(value, str):
        return None
    stripped = value.strip()
    if not stripped or re.match(r"^[+-]?0\d+", stripped):
        return None
    digits = re.sub(r"[^0-9]", "", stripped)
    if len(digits) > 15 or re.search(r"[()\-/]", stripped):
        return None
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", stripped):
        return None
    try:
        number = Decimal(stripped)
    except InvalidOperation:
        return None
    if not number.is_finite():
        return None
    if number == number.to_integral_value():
        return int(number)
    result = float(number)
    return result if math.isfinite(result) else None


class NumericTextRule(WorkbookRule):
    rule_id = "WL006_NUMERIC_TEXT"
    title = "Numeric text in numeric region"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for region in context.data_regions[worksheet.title]:
                for column in range(region.min_column, region.max_column + 1):
                    cells = [
                        _cell(worksheet, row, column)
                        for row in range(region.min_row + 1, region.max_row + 1)
                        if _cell(worksheet, row, column).value is not None
                    ]
                    numeric_count = sum(
                        isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool)
                        for cell in cells
                    )
                    if numeric_count < 3:
                        continue
                    for cell in cells:
                        converted = _numeric_text(cell.value)
                        if converted is None or (worksheet.title, cell.coordinate) in seen:
                            continue
                        ratio = numeric_count / len(cells)
                        if ratio < 0.75:
                            continue
                        seen.add((worksheet.title, cell.coordinate))
                        confidence = 0.97 if ratio >= 0.85 else 0.9
                        patches: list[PatchOperation] = []
                        if confidence >= 0.95:
                            patch = _make_patch(
                                kind=PatchKind.SET_NUMERIC,
                                worksheet=worksheet,
                                cell=cell,
                                before=cell.value,
                                after=converted,
                                confidence=confidence,
                                description="Convert an unambiguous numeric string to an OOXML numeric value.",
                            )
                            patches.append(patch)
                            result.patches.append(patch)
                        result.findings.append(
                            _make_finding(
                                context=context,
                                rule_id=self.rule_id,
                                title=self.title,
                                explanation="A plain numeric string appears in a column dominated by numeric values.",
                                severity=Severity.WARNING,
                                confidence=confidence,
                                sheet=worksheet.title,
                                location=cell.coordinate,
                                evidence=Evidence(
                                    summary=f"{numeric_count} peer cells are stored as numbers",
                                    observed=cell.value,
                                    expected=converted,
                                    peers=[
                                        item.coordinate
                                        for item in cells
                                        if isinstance(item.value, (int, float))
                                    ][:12],
                                ),
                                expected="Numeric measures use numeric cell storage, while identifiers remain text.",
                                suggested_action="Confirm the value is a measure rather than an identifier.",
                                patches=patches,
                            )
                        )
        return result


class StyleOutlierRule(WorkbookRule):
    rule_id = "WL007_STYLE_OUTLIER"
    title = "Style outlier in homogeneous region"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for region in context.data_regions[worksheet.title]:
                for column in range(region.min_column, region.max_column + 1):
                    cells = [
                        _cell(worksheet, row, column)
                        for row in range(region.min_row + 1, region.max_row + 1)
                        if _cell(worksheet, row, column).value is not None
                    ]
                    if len(cells) < 5:
                        continue
                    counts = Counter(cell.style_id for cell in cells)
                    style_id, count = counts.most_common(1)[0]
                    ratio = count / len(cells)
                    outliers = [cell for cell in cells if cell.style_id != style_id]
                    if ratio < 0.8 or len(outliers) != 1:
                        continue
                    cell = outliers[0]
                    if (worksheet.title, cell.coordinate) in seen:
                        continue
                    seen.add((worksheet.title, cell.coordinate))
                    sources = [
                        peer
                        for peer in cells
                        if peer.style_id == style_id
                        and not _in_unsupported_formula_range(
                            context, worksheet.title, peer.coordinate
                        )
                    ]
                    source = (
                        min(
                            sources,
                            key=lambda peer: (
                                abs(peer.row - cell.row),
                                peer.coordinate,
                            ),
                        )
                        if sources
                        else None
                    )
                    confidence = ratio
                    patches: list[PatchOperation] = []
                    if (
                        confidence >= 0.95
                        and source is not None
                        and not _in_unsupported_formula_range(
                            context, worksheet.title, cell.coordinate
                        )
                    ):
                        patch = _make_patch(
                            kind=PatchKind.COPY_STYLE,
                            worksheet=worksheet,
                            cell=cell,
                            before=cell.style_id,
                            after=style_id,
                            confidence=confidence,
                            source_cell=source.coordinate,
                            description="Copy the existing consensus style ID from the nearest peer.",
                        )
                        patches.append(patch)
                        result.patches.append(patch)
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation="One populated cell uses a different existing style ID from its column peers.",
                            severity=Severity.INFO,
                            confidence=confidence,
                            sheet=worksheet.title,
                            location=cell.coordinate,
                            evidence=Evidence(
                                summary=f"Style {style_id} appears in {count} of {len(cells)} peer cells",
                                observed=cell.style_id,
                                expected=style_id,
                                peers=[
                                    peer.coordinate for peer in cells if peer.style_id == style_id
                                ][:12],
                                details={"number_format": cell.number_format},
                            ),
                            expected="A homogeneous measure column uses its consensus style.",
                            suggested_action="Check whether the visual distinction is intentional.",
                            patches=patches,
                        )
                    )
        return result


class HiddenNonemptyDataRule(WorkbookRule):
    rule_id = "WL008_HIDDEN_NONEMPTY_DATA"
    title = "Hidden nonempty data"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            nonempty = [cell for cell in worksheet._cells.values() if cell.value is not None]
            if worksheet.sheet_state != "visible" and nonempty:
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A hidden worksheet contains data or formulas that may affect interpretation.",
                        severity=Severity.WARNING,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=None,
                        evidence=Evidence(
                            summary=f"{worksheet.sheet_state} sheet contains {len(nonempty)} nonempty cells"
                        ),
                        expected="Hidden content is reviewed and documented.",
                        suggested_action="Inspect the hidden sheet manually; WorkbookLens never unhides it automatically.",
                    )
                )
            for row, row_dimension in sorted(worksheet.row_dimensions.items()):
                cells = [cell for cell in nonempty if cell.row == row]
                if not row_dimension.hidden or not cells:
                    continue
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A hidden row contains values or formulas.",
                        severity=Severity.WARNING
                        if any(cell.data_type == "f" for cell in cells)
                        else Severity.INFO,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=f"{row}:{row}",
                        evidence=Evidence(
                            summary=f"Hidden row {row} contains {len(cells)} nonempty cells",
                            peers=[cell.coordinate for cell in cells[:12]],
                        ),
                        expected="Hidden rows with consequential content are intentionally documented.",
                        suggested_action="Review the row manually; no automatic unhide is offered.",
                    )
                )
            for _column, column_dimension in sorted(worksheet.column_dimensions.items()):
                min_column = column_dimension.min
                max_column = column_dimension.max
                if min_column is None or max_column is None:
                    continue
                cells = [cell for cell in nonempty if min_column <= cell.column <= max_column]
                if not column_dimension.hidden or not cells:
                    continue
                location = f"{get_column_letter(min_column)}:{get_column_letter(max_column)}"
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A hidden column contains values or formulas.",
                        severity=Severity.WARNING
                        if any(cell.data_type == "f" for cell in cells)
                        else Severity.INFO,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=location,
                        evidence=Evidence(
                            summary=f"Hidden column range {location} contains {len(cells)} nonempty cells",
                            peers=[cell.coordinate for cell in cells[:12]],
                        ),
                        expected="Hidden columns with consequential content are intentionally documented.",
                        suggested_action="Review the column manually; no automatic unhide is offered.",
                    )
                )
        return result


class ExternalLinkRule(WorkbookRule):
    rule_id = "WL009_EXTERNAL_LINK"
    title = "External workbook link"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in worksheet._cells.values():
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                features = analyze_formula(cell.value)
                if not features.external_references:
                    continue
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="The formula depends on another workbook; WorkbookLens does not fetch it.",
                        severity=Severity.WARNING,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary="Formula contains external workbook reference",
                            observed=cell.value,
                            details={"references": list(features.external_references)},
                        ),
                        expected="External dependencies are explicit, available, and reviewed.",
                        suggested_action="Verify the linked workbook and consider replacing fragile dependencies.",
                    )
                )
        for name, target in context.snapshot.defined_names.items():
            if re.search(r"\[[^\]]+\][^!]*!", target):
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A defined name refers to another workbook.",
                        severity=Severity.WARNING,
                        confidence=1.0,
                        sheet=None,
                        location=name,
                        evidence=Evidence(
                            summary="Defined name contains external reference", observed=target
                        ),
                        expected="Defined-name dependencies remain local or are explicitly reviewed.",
                        suggested_action="Review the external target; WorkbookLens never opens it.",
                    )
                )
        return result


class VolatileOrFragileFunctionRule(WorkbookRule):
    rule_id = "WL010_VOLATILE_OR_FRAGILE_FUNCTION"
    title = "Volatile or fragile formula construct"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in worksheet._cells.values():
                if cell.data_type != "f" or not isinstance(cell.value, str):
                    continue
                features = analyze_formula(cell.value)
                if not features.volatile_functions and not features.has_whole_column_reference:
                    continue
                constructs = list(features.volatile_functions)
                if features.has_whole_column_reference:
                    constructs.append("whole-column reference")
                severity = (
                    Severity.WARNING
                    if any(
                        item in {"OFFSET", "INDIRECT", "whole-column reference"}
                        for item in constructs
                    )
                    else Severity.INFO
                )
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="The formula uses constructs that can recalculate frequently or resist static tracing.",
                        severity=severity,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(summary=", ".join(constructs), observed=cell.value),
                        expected="Performance-sensitive and auditable models avoid unnecessary fragile constructs.",
                        suggested_action="Review whether a bounded direct reference can express the same intent.",
                    )
                )
        return result


class ErrorCellRule(WorkbookRule):
    rule_id = "WL011_ERROR_CELL"
    title = "Stored Excel error value"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in worksheet._cells.values():
                value = cell.value
                if cell.data_type != "e" and not (
                    isinstance(value, str) and value.upper() in EXCEL_ERRORS
                ):
                    continue
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="The cell stores a recognized Excel error value.",
                        severity=Severity.ERROR,
                        confidence=1.0,
                        sheet=worksheet.title,
                        location=cell.coordinate,
                        evidence=Evidence(summary="Stored error cell", observed=value),
                        expected="Calculated or imported values do not contain Excel error tokens.",
                        suggested_action="Trace the producing formula or upstream data; no value is fabricated.",
                    )
                )
        return result


def _configured_key_identity(value: Any) -> tuple[str, Any]:
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)):
        return ("number", value)
    try:
        hash(value)
        normalized = value
    except TypeError:
        normalized = repr(value)
    return (type(value).__qualname__, normalized)


class DuplicateConfiguredKeyRule(WorkbookRule):
    rule_id = "WL012_DUPLICATE_CONFIGURED_KEY"
    title = "Duplicate configured key"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        keys = context.config.get("keys", [])
        if not isinstance(keys, list):
            return result
        for specification in keys:
            if not isinstance(specification, dict):
                continue
            sheet_name = specification.get("sheet")
            range_text = specification.get("range")
            if not isinstance(sheet_name, str) or not isinstance(range_text, str):
                continue
            if sheet_name not in context.workbook.sheetnames:
                continue
            worksheet = context.workbook[sheet_name]
            try:
                min_col, min_row, max_col, max_row = range_boundaries(range_text)
            except ValueError:
                continue
            if None in {min_col, min_row, max_col, max_row}:
                continue
            min_col = cast(int, min_col)
            min_row = cast(int, min_row)
            max_col = cast(int, max_col)
            max_row = cast(int, max_row)
            if (
                min_col != max_col
                or min_col < 1
                or min_row < 1
                or max_col > 16_384
                or max_row > 1_048_576
                or max_row - min_row + 1 > 100_000
            ):
                continue
            by_value: dict[tuple[str, Any], list[str]] = defaultdict(list)
            observed_values: dict[tuple[str, Any], Any] = {}
            for row in range(min_row, max_row + 1):
                cell = _cell(worksheet, row, min_col)
                if cell.value is None and specification.get("ignore_blank", True):
                    continue
                key = _configured_key_identity(cell.value)
                observed_values.setdefault(key, cell.value)
                by_value[key].append(cell.coordinate)
            for key, locations in by_value.items():
                if len(locations) < 2:
                    continue
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation="A value repeats in a column explicitly configured as a unique key.",
                        severity=Severity.ERROR,
                        confidence=1.0,
                        sheet=sheet_name,
                        location=",".join(locations),
                        evidence=Evidence(
                            summary=f"Configured key value appears {len(locations)} times",
                            observed=observed_values[key],
                            peers=locations,
                        ),
                        expected=f"Values in {sheet_name}!{range_text} are unique.",
                        suggested_action="Resolve the duplicate records or revise the explicit key configuration.",
                    )
                )
        return result


class BrokenDefinedNameRule(WorkbookRule):
    rule_id = "WL013_BROKEN_DEFINED_NAME"
    title = "Broken defined name"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for defined_name in context.workbook.defined_names.values():
            target = defined_name.attr_text or ""
            reason = None
            if "#REF!" in target.upper():
                reason = "target contains #REF!"
            elif defined_name.type == "RANGE":
                try:
                    destinations = list(defined_name.destinations)
                except (TypeError, ValueError, AttributeError):
                    destinations = []
                    reason = "range target could not be parsed"
                if not destinations and reason is None:
                    reason = "range target has no resolvable destination"
                for sheet_name, range_text in destinations:
                    if sheet_name not in context.workbook.sheetnames:
                        reason = f"target sheet {sheet_name!r} does not exist"
                        break
                    try:
                        range_boundaries(range_text)
                    except ValueError:
                        reason = f"target range {range_text!r} is invalid"
                        break
            if reason is None:
                continue
            result.findings.append(
                _make_finding(
                    context=context,
                    rule_id=self.rule_id,
                    title=self.title,
                    explanation="A workbook defined name cannot resolve to an existing valid range.",
                    severity=Severity.ERROR,
                    confidence=1.0,
                    sheet=None,
                    location=defined_name.name,
                    evidence=Evidence(summary=reason, observed=target),
                    expected="Defined names resolve to valid local sheets and ranges.",
                    suggested_action="Repair or remove the name in Excel after confirming downstream usage.",
                )
            )
        return result


def _ranges_intersect(left: CellRange, right: Region) -> bool:
    return not (
        left.max_row < right.min_row
        or left.min_row > right.max_row
        or left.max_col < right.min_column
        or left.min_col > right.max_column
    )


class MergedCellInDataRegionRule(WorkbookRule):
    rule_id = "WL014_MERGED_CELL_IN_DATA_REGION"
    title = "Merged cells intersect a data region"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for merged in worksheet.merged_cells.ranges:
                for region in context.data_regions[worksheet.title]:
                    if not _ranges_intersect(merged, region):
                        continue
                    if merged.max_row == merged.min_row == region.min_row:
                        # A single merged header row is common and not itself table corruption.
                        continue
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation="A merge intersects the body of a dense table-like region.",
                            severity=Severity.WARNING,
                            confidence=0.9,
                            sheet=worksheet.title,
                            location=str(merged),
                            evidence=Evidence(
                                summary="Merged range overlaps inferred data body",
                                observed=str(merged),
                                expected={
                                    "min_row": region.min_row,
                                    "max_row": region.max_row,
                                    "min_column": region.min_column,
                                    "max_column": region.max_column,
                                },
                            ),
                            expected="Table-like data bodies use one logical value per cell.",
                            suggested_action="Review downstream sort/filter behavior; no automatic unmerge is offered.",
                        )
                    )
                    break
        return result


def _validation_signature(worksheet: Worksheet, cell: Cell) -> str | None:
    if worksheet.data_validations is None:
        return None
    signatures: list[str] = []
    for validation in worksheet.data_validations.dataValidation:
        for cell_range in validation.ranges.ranges:
            if (
                cell_range.min_row <= cell.row <= cell_range.max_row
                and cell_range.min_col <= cell.column <= cell_range.max_col
            ):
                signatures.append(
                    "|".join(
                        [
                            validation.type or "",
                            validation.operator or "",
                            validation.formula1 or "",
                            validation.formula2 or "",
                            str(bool(validation.allow_blank)),
                        ]
                    )
                )
    return ";".join(sorted(signatures)) or None


class InconsistentDataValidationRule(WorkbookRule):
    rule_id = "WL015_INCONSISTENT_DATA_VALIDATION"
    title = "Inconsistent data validation"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        seen: set[tuple[str, str]] = set()
        for worksheet in context.workbook.worksheets:
            for region in context.data_regions[worksheet.title]:
                for column in range(region.min_column, region.max_column + 1):
                    cells = [
                        _cell(worksheet, row, column)
                        for row in range(region.min_row + 1, region.max_row + 1)
                        if _cell(worksheet, row, column).value is not None
                    ]
                    if len(cells) < 4:
                        continue
                    signatures = [_validation_signature(worksheet, cell) for cell in cells]
                    nonempty = [signature for signature in signatures if signature is not None]
                    if len(nonempty) < 3:
                        continue
                    consensus, count = Counter(nonempty).most_common(1)[0]
                    if count / len(cells) < 0.75:
                        continue
                    outliers = [
                        cell
                        for cell, signature in zip(cells, signatures, strict=True)
                        if signature != consensus
                    ]
                    if len(outliers) != 1:
                        continue
                    cell = outliers[0]
                    if (worksheet.title, cell.coordinate) in seen:
                        continue
                    seen.add((worksheet.title, cell.coordinate))
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation="One populated input cell lacks or differs from the validation used by its peers.",
                            severity=Severity.WARNING,
                            confidence=count / len(cells),
                            sheet=worksheet.title,
                            location=cell.coordinate,
                            evidence=Evidence(
                                summary=f"{count} peer cells share one validation signature",
                                observed=_validation_signature(worksheet, cell),
                                expected=consensus,
                                peers=[
                                    peer.coordinate
                                    for peer in cells
                                    if _validation_signature(worksheet, peer) == consensus
                                ][:12],
                            ),
                            expected="Cells in a homogeneous input column share validation constraints.",
                            suggested_action="Review and restore the intended validation rule manually.",
                        )
                    )
        return result


BUILTIN_RULES: tuple[type[WorkbookRule], ...] = (
    BrokenReferenceRule,
    FormulaPatternOutlierRule,
    BlankInFormulaBandRule,
    HardcodedValueInFormulaBandRule,
    SuspiciousSumBoundaryRule,
    NumericTextRule,
    StyleOutlierRule,
    HiddenNonemptyDataRule,
    ExternalLinkRule,
    VolatileOrFragileFunctionRule,
    ErrorCellRule,
    DuplicateConfiguredKeyRule,
    BrokenDefinedNameRule,
    MergedCellInDataRegionRule,
    InconsistentDataValidationRule,
)
