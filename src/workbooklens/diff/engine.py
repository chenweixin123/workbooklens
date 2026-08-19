"""Semantic comparison of workbook values, formulas, styles, and structure."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from jinja2 import Environment, PackageLoader, select_autoescape

from workbooklens.formulas import UnsupportedFormulaError, normalize_formula
from workbooklens.models import (
    CellChange,
    CellSnapshot,
    Severity,
    SheetSnapshot,
    StructuralChange,
    WorkbookDiff,
    WorkbookSnapshot,
)
from workbooklens.snapshot import create_snapshot
from workbooklens.utils import atomic_write_bytes, sha256_bytes, stable_json_bytes, write_json


def _sheet_content_fingerprint(sheet: SheetSnapshot) -> str:
    payload = sheet.model_dump(mode="json", exclude={"name", "index", "state"})
    for cell in payload["cells"].values():
        # ``style_id`` is only meaningful within one workbook. Fresh snapshots carry a
        # workbook-independent fingerprint; retain IDs only for legacy snapshots.
        if cell.get("style_fingerprint"):
            cell.pop("style_id", None)
    return sha256_bytes(stable_json_bytes(payload))


def _renamed_sheets(before: WorkbookSnapshot, after: WorkbookSnapshot) -> dict[str, str]:
    before_names = {sheet.name for sheet in before.sheets}
    after_names = {sheet.name for sheet in after.sheets}
    removed = [sheet for sheet in before.sheets if sheet.name not in after_names]
    added = [sheet for sheet in after.sheets if sheet.name not in before_names]
    removed_by_hash: dict[str, list[SheetSnapshot]] = {}
    for sheet in removed:
        removed_by_hash.setdefault(_sheet_content_fingerprint(sheet), []).append(sheet)
    added_by_hash: dict[str, list[SheetSnapshot]] = {}
    for sheet in added:
        added_by_hash.setdefault(_sheet_content_fingerprint(sheet), []).append(sheet)
    renamed: dict[str, str] = {}
    for fingerprint in sorted(removed_by_hash.keys() & added_by_hash.keys()):
        old_matches = removed_by_hash[fingerprint]
        new_matches = added_by_hash[fingerprint]
        if len(old_matches) == 1 and len(new_matches) == 1:
            renamed[old_matches[0].name] = new_matches[0].name
    return renamed


def _signature(formula: str | None, coordinate: str) -> str | None:
    if formula is None:
        return None
    try:
        return normalize_formula(formula, coordinate)
    except (ValueError, UnsupportedFormulaError):
        return None


def _typed_equal(before: Any, after: Any) -> bool:
    """Compare cell values without collapsing booleans into integers or other coercions."""

    return type(before) is type(after) and before == after


def _style_fingerprint(cell: CellSnapshot | None) -> str:
    if cell is None:
        return "default"
    if cell.style_fingerprint:
        return cell.style_fingerprint
    # Backward compatibility for deserialized v0.1 snapshots without semantic fingerprints.
    return "default" if cell.style_id == 0 else f"legacy-style-id:{cell.style_id}"


def _compare_cell(
    sheet_name: str,
    coordinate: str,
    before: CellSnapshot | None,
    after: CellSnapshot | None,
) -> list[CellChange]:
    changes: list[CellChange] = []
    before_formula = before.formula if before else None
    after_formula = after.formula if after else None
    if before_formula != after_formula:
        changes.append(
            CellChange(
                sheet=sheet_name,
                cell=coordinate,
                change_type="formula",
                before=before_formula,
                after=after_formula,
                importance=Severity.ERROR,
                before_signature=_signature(before_formula, coordinate),
                after_signature=_signature(after_formula, coordinate),
            )
        )
    before_value = before.value if before and before_formula is None else None
    after_value = after.value if after and after_formula is None else None
    if not _typed_equal(before_value, after_value):
        changes.append(
            CellChange(
                sheet=sheet_name,
                cell=coordinate,
                change_type="value",
                before=before_value,
                after=after_value,
                importance=Severity.ERROR,
            )
        )
    before_style = _style_fingerprint(before)
    after_style = _style_fingerprint(after)
    if before_style != after_style:
        changes.append(
            CellChange(
                sheet=sheet_name,
                cell=coordinate,
                change_type="style",
                before=before_style,
                after=after_style,
                importance=Severity.INFO,
            )
        )
    before_format = before.number_format if before else "General"
    after_format = after.number_format if after else "General"
    if before_format != after_format:
        changes.append(
            CellChange(
                sheet=sheet_name,
                cell=coordinate,
                change_type="number_format",
                before=before_format,
                after=after_format,
                importance=Severity.WARNING,
            )
        )
    return changes


def _list_change(
    changes: list[StructuralChange],
    change_type: str,
    subject_prefix: str,
    before: list[Any],
    after: list[Any],
    importance: Severity = Severity.WARNING,
) -> None:
    before_set = set(before)
    after_set = set(after)
    for item in sorted(before_set - after_set, key=str):
        changes.append(
            StructuralChange(
                change_type=f"{change_type}_removed",
                subject=f"{subject_prefix}{item}",
                before=item,
                importance=importance,
            )
        )
    for item in sorted(after_set - before_set, key=str):
        changes.append(
            StructuralChange(
                change_type=f"{change_type}_added",
                subject=f"{subject_prefix}{item}",
                after=item,
                importance=importance,
            )
        )


def compare_snapshots(before: WorkbookSnapshot, after: WorkbookSnapshot) -> WorkbookDiff:
    """Compare two already-extracted snapshots deterministically."""

    structural: list[StructuralChange] = []
    cell_changes: list[CellChange] = []
    before_map = {sheet.name: sheet for sheet in before.sheets}
    after_map = {sheet.name: sheet for sheet in after.sheets}
    renamed = _renamed_sheets(before, after)
    renamed_targets = set(renamed.values())
    for old_name, new_name in sorted(renamed.items()):
        structural.append(
            StructuralChange(
                change_type="sheet_renamed",
                subject=old_name,
                before=old_name,
                after=new_name,
                importance=Severity.WARNING,
            )
        )
    for name in sorted(before_map.keys() - after_map.keys() - renamed.keys()):
        structural.append(
            StructuralChange(
                change_type="sheet_removed",
                subject=name,
                before=name,
                importance=Severity.ERROR,
            )
        )
    for name in sorted(after_map.keys() - before_map.keys() - renamed_targets):
        structural.append(
            StructuralChange(
                change_type="sheet_added",
                subject=name,
                after=name,
                importance=Severity.WARNING,
            )
        )

    pairs: list[tuple[SheetSnapshot, SheetSnapshot]] = []
    for name in sorted(before_map.keys() & after_map.keys()):
        pairs.append((before_map[name], after_map[name]))
    for old_name, new_name in sorted(renamed.items()):
        pairs.append((before_map[old_name], after_map[new_name]))
    for before_sheet, after_sheet in pairs:
        display_name = after_sheet.name
        if before_sheet.index != after_sheet.index:
            structural.append(
                StructuralChange(
                    change_type="sheet_reordered",
                    subject=display_name,
                    before=before_sheet.index,
                    after=after_sheet.index,
                    importance=Severity.INFO,
                )
            )
        if before_sheet.state != after_sheet.state:
            structural.append(
                StructuralChange(
                    change_type="sheet_visibility",
                    subject=display_name,
                    before=before_sheet.state,
                    after=after_sheet.state,
                    importance=Severity.WARNING,
                )
            )
        _list_change(
            structural,
            "merged_range",
            f"{display_name}!",
            before_sheet.merged_ranges,
            after_sheet.merged_ranges,
        )
        _list_change(
            structural,
            "hidden_row",
            f"{display_name}!row ",
            before_sheet.hidden_rows,
            after_sheet.hidden_rows,
            Severity.INFO,
        )
        _list_change(
            structural,
            "hidden_column",
            f"{display_name}!column ",
            before_sheet.hidden_columns,
            after_sheet.hidden_columns,
            Severity.INFO,
        )
        _list_change(
            structural,
            "data_validation",
            f"{display_name}!",
            before_sheet.data_validations,
            after_sheet.data_validations,
        )
        for coordinate in sorted(before_sheet.cells.keys() | after_sheet.cells.keys()):
            cell_changes.extend(
                _compare_cell(
                    display_name,
                    coordinate,
                    before_sheet.cells.get(coordinate),
                    after_sheet.cells.get(coordinate),
                )
            )

    all_names = before.defined_names.keys() | after.defined_names.keys()
    for name in sorted(all_names):
        before_target = before.defined_names.get(name)
        after_target = after.defined_names.get(name)
        if before_target != after_target:
            structural.append(
                StructuralChange(
                    change_type="defined_name",
                    subject=name,
                    before=before_target,
                    after=after_target,
                    importance=Severity.WARNING,
                )
            )
    if before.calculation_mode != after.calculation_mode:
        structural.append(
            StructuralChange(
                change_type="calculation_mode",
                subject="workbook",
                before=before.calculation_mode,
                after=after.calculation_mode,
                importance=Severity.WARNING,
            )
        )
    return WorkbookDiff(
        before_sha256=before.source_sha256,
        after_sha256=after.source_sha256,
        cell_changes=sorted(
            cell_changes,
            key=lambda item: (item.sheet, item.cell, item.change_type),
        ),
        structural_changes=sorted(
            structural,
            key=lambda item: (item.change_type, item.subject),
        ),
    )


def compare_workbooks(before_path: Path, after_path: Path) -> WorkbookDiff:
    """Securely snapshot and compare two workbook files."""

    return compare_snapshots(create_snapshot(before_path), create_snapshot(after_path))


def write_diff_report(diff: WorkbookDiff, output_html: Path) -> dict[str, Path]:
    """Write JSON beside a self-contained, filterable HTML semantic diff."""

    if output_html.suffix.lower() != ".html":
        raise ValueError("semantic diff output must have an .html extension")
    output_html.parent.mkdir(parents=True, exist_ok=True)
    json_path = output_html.with_suffix(".json")
    write_json(json_path, diff.model_dump(mode="json"))
    environment = Environment(
        loader=PackageLoader("workbooklens.diff", "templates"),
        autoescape=select_autoescape(("html", "xml")),
    )
    html = environment.get_template("diff.html.j2").render(diff=diff)
    atomic_write_bytes(output_html, html.encode("utf-8"))
    return {"html": output_html, "json": json_path}
