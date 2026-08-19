"""Workbook-to-model snapshot extraction without formula or macro execution."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path
from typing import Any, Literal, cast

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.models import CellSnapshot, SheetSnapshot, WorkbookSnapshot
from workbooklens.ooxml.safety import PackageLimits, inspect_package
from workbooklens.utils import sha256_bytes, sha256_file, stable_json_bytes


def _formula_payload(cell: Cell) -> str | None:
    if cell.data_type != "f":
        return None
    if isinstance(cell.value, str):
        return cell.value
    if isinstance(cell.value, ArrayFormula):
        return f"<array ref={cell.value.ref!r} text={cell.value.text!r}>"
    if isinstance(cell.value, DataTableFormula):
        attributes = ",".join(
            f"{name}={getattr(cell.value, name)!r}"
            for name in ("ref", "ca", "dt2D", "dtr", "r1", "r2", "del1", "del2")
        )
        return f"<dataTable {attributes}>"
    return f"<unsupported formula object {type(cell.value).__qualname__}>"


def cell_semantic_payload(cell: Cell) -> dict[str, Any]:
    """Return the source-bound fields used by patch preconditions."""

    formula = _formula_payload(cell)
    value = None if formula is not None else cell.value
    return {
        "coordinate": cell.coordinate,
        "value": value,
        "formula": formula,
        "data_type": cell.data_type,
        "style_id": cell.style_id,
        "number_format": cell.number_format,
    }


def cell_fingerprint(cell: Cell) -> str:
    """Hash the semantic state relevant to every supported patch operation."""

    return sha256_bytes(stable_json_bytes(cell_semantic_payload(cell)))


def _xml_component_payload(component: Any) -> dict[str, Any]:
    """Return a deterministic, workbook-local-ID-free style component tree."""

    element = component.to_tree()

    def normalize(node: Any) -> dict[str, Any]:
        return {
            "tag": str(node.tag),
            "attributes": dict(
                sorted((str(key), str(value)) for key, value in node.attrib.items())
            ),
            "text": node.text or "",
            "children": [normalize(child) for child in node],
        }

    return normalize(element)


def _style_payload(cell: Cell) -> dict[str, Any]:
    """Describe effective non-number-format styling without using ``style_id``."""

    return {
        "font": _xml_component_payload(cell.font),
        "fill": _xml_component_payload(cell.fill),
        "border": _xml_component_payload(cell.border),
        "alignment": _xml_component_payload(cell.alignment),
        "protection": _xml_component_payload(cell.protection),
        "quote_prefix": bool(cell.quotePrefix),
        "pivot_button": bool(cell.pivotButton),
    }


@lru_cache(maxsize=1)
def _default_style_payload() -> dict[str, Any]:
    workbook = Workbook()
    try:
        worksheet = workbook.active
        if worksheet is None:  # pragma: no cover - openpyxl always creates an active sheet
            raise RuntimeError("new workbook has no active worksheet")
        return _style_payload(worksheet["A1"])
    finally:
        workbook.close()


def style_fingerprint(cell: Cell) -> str:
    """Hash effective visual/protection styling independently of workbook style-table IDs."""

    payload = _style_payload(cell)
    if payload == _default_style_payload():
        return "default"
    return sha256_bytes(stable_json_bytes(payload))


def _snapshot_cell(cell: Cell, worksheet: Worksheet) -> CellSnapshot:
    payload = cell_semantic_payload(cell)
    column_letter = cell.column_letter
    return CellSnapshot(
        **payload,
        style_fingerprint=style_fingerprint(cell),
        row_hidden=bool(worksheet.row_dimensions[cell.row].hidden),
        column_hidden=bool(worksheet.column_dimensions[column_letter].hidden),
    )


def _defined_names(workbook: Workbook) -> dict[str, str]:
    names: dict[str, str] = {}
    for name in workbook.defined_names.values():
        key = name.name
        if name.localSheetId is not None:
            key = f"{key}@sheet:{name.localSheetId}"
        names[key] = name.attr_text or ""
    return dict(sorted(names.items()))


def snapshot_from_workbook(workbook: Workbook, path: Path, source_sha256: str) -> WorkbookSnapshot:
    """Build a deterministic snapshot from an already-open, non-read-only workbook."""

    sheets: list[SheetSnapshot] = []
    for index, worksheet in enumerate(workbook.worksheets):
        cells: dict[str, CellSnapshot] = {}
        # Iterating the sparse cell store avoids trusting a maliciously huge dimension range.
        sparse_cells = [cell for cell in worksheet._cells.values() if isinstance(cell, Cell)]
        for cell in sorted(
            sparse_cells,
            key=lambda item: (item.row, item.column),
        ):
            if cell.value is not None or cell.has_style:
                cells[cell.coordinate] = _snapshot_cell(cell, worksheet)
        validations: list[str] = []
        if worksheet.data_validations is not None:
            for validation in worksheet.data_validations.dataValidation:
                validations.append(
                    "|".join(
                        [
                            str(validation.sqref),
                            validation.type or "",
                            validation.operator or "",
                            validation.formula1 or "",
                            validation.formula2 or "",
                        ]
                    )
                )
        sheets.append(
            SheetSnapshot(
                name=worksheet.title,
                index=index,
                state=worksheet.sheet_state,
                max_row=max((cell.row for cell in sparse_cells), default=0),
                max_column=max((cell.column for cell in sparse_cells), default=0),
                cells=cells,
                merged_ranges=sorted(str(item) for item in worksheet.merged_cells.ranges),
                hidden_rows=sorted(
                    index
                    for index, dimension in worksheet.row_dimensions.items()
                    if dimension.hidden
                ),
                hidden_columns=sorted(
                    key
                    for key, dimension in worksheet.column_dimensions.items()
                    if dimension.hidden
                ),
                data_validations=sorted(validations),
            )
        )
    calculation_mode = None
    if workbook.calculation is not None:
        calculation_mode = workbook.calculation.calcMode
    workbook_format = cast(Literal["xlsx", "xlsm"], path.suffix.lower().lstrip("."))
    return WorkbookSnapshot(
        source_name=path.name,
        source_sha256=source_sha256,
        format=workbook_format,
        sheets=sheets,
        defined_names=_defined_names(workbook),
        calculation_mode=calculation_mode,
    )


def load_for_analysis(path: Path, limits: PackageLimits | None = None) -> Workbook:
    """Safety-check and open a workbook without evaluating formulas or external content."""

    inspection = inspect_package(path, limits)
    return load_workbook(
        inspection.path,
        read_only=False,
        data_only=False,
        keep_links=False,
        keep_vba=False,
    )


def create_snapshot(path: Path, limits: PackageLimits | None = None) -> WorkbookSnapshot:
    """Inspect, open, snapshot, and close one workbook."""

    resolved = path.expanduser().resolve()
    workbook = load_for_analysis(resolved, limits)
    try:
        return snapshot_from_workbook(workbook, resolved, sha256_file(resolved))
    finally:
        workbook.close()
