"""Extract formula constructs that must never participate in automatic repair."""

from __future__ import annotations

import posixpath
import zipfile
from pathlib import Path
from typing import cast

from lxml import etree
from openpyxl.worksheet.cell_range import CellRange

from workbooklens.exceptions import UnsafeWorkbookError
from workbooklens.formulas import analyze_formula
from workbooklens.ooxml.safety import PackageLimits, parse_xml_part

UNSUPPORTED_FORMULA_TYPES = {"shared", "array", "dataTable"}


def _bounded_part(archive: zipfile.ZipFile, part: str, limits: PackageLimits) -> bytes:
    try:
        info = archive.getinfo(part)
    except KeyError as exc:
        raise UnsafeWorkbookError(f"Worksheet package part is missing: {part!r}") from exc
    if info.file_size > limits.max_xml_bytes:
        raise UnsafeWorkbookError(
            f"Package part {part!r} exceeds the {limits.max_xml_bytes}-byte XML limit"
        )
    with archive.open(info, "r") as handle:
        data = handle.read(limits.max_xml_bytes + 1)
    if len(data) > limits.max_xml_bytes:
        raise UnsafeWorkbookError(f"Package part {part!r} expanded beyond its XML limit")
    return data


def _sheet_parts(archive: zipfile.ZipFile, limits: PackageLimits) -> dict[str, str]:
    workbook_part = "xl/workbook.xml"
    relationships_part = "xl/_rels/workbook.xml.rels"
    workbook_root = parse_xml_part(_bounded_part(archive, workbook_part, limits), workbook_part)
    relationships_root = parse_xml_part(
        _bounded_part(archive, relationships_part, limits), relationships_part
    )
    relationships: dict[str, str] = {}
    for relationship in relationships_root:
        if etree.QName(relationship).localname != "Relationship":
            continue
        relationship_id = relationship.get("Id")
        target = relationship.get("Target")
        relationship_type = relationship.get("Type", "")
        if not relationship_id or not target or not relationship_type.endswith("/worksheet"):
            continue
        candidate = target.lstrip("/") if target.startswith("/") else posixpath.join("xl", target)
        normalized = posixpath.normpath(candidate)
        if normalized.startswith("../") or normalized in {"", ".", ".."}:
            raise UnsafeWorkbookError(
                f"Worksheet relationship target escapes the package: {target!r}"
            )
        relationships[relationship_id] = normalized
    result: dict[str, str] = {}
    for sheet in workbook_root.iter():
        if etree.QName(sheet).localname != "sheet":
            continue
        name = sheet.get("name")
        relationship_id = cast(
            str | None,
            next(
                (
                    value
                    for attribute, value in sheet.attrib.items()
                    if etree.QName(attribute).localname == "id"
                ),
                None,
            ),
        )
        if not name or not relationship_id or relationship_id not in relationships:
            raise UnsafeWorkbookError("Workbook contains a malformed worksheet relationship")
        result[name] = relationships[relationship_id]
    return result


def find_unsupported_formula_ranges(
    path: Path,
    limits: PackageLimits | None = None,
) -> dict[str, tuple[CellRange, ...]]:
    """Return raw shared/array/data-table/dynamic formula ranges by worksheet."""

    active_limits = limits or PackageLimits()
    result: dict[str, tuple[CellRange, ...]] = {}
    try:
        with zipfile.ZipFile(path, "r") as archive:
            sheet_parts = _sheet_parts(archive, active_limits)
            for sheet_name, part in sheet_parts.items():
                root = parse_xml_part(_bounded_part(archive, part, active_limits), part)
                ranges: dict[str, CellRange] = {}
                for cell in root.iter():
                    if etree.QName(cell).localname != "c":
                        continue
                    coordinate = cell.get("r")
                    formula = next(
                        (child for child in cell if etree.QName(child).localname == "f"),
                        None,
                    )
                    if coordinate is None or formula is None:
                        continue
                    formula_type = formula.get("t")
                    reference = formula.get("ref")
                    text = "=" + (formula.text or "")
                    features = analyze_formula(text)
                    if not (
                        formula_type in UNSUPPORTED_FORMULA_TYPES
                        or reference
                        or features.external_references
                        or features.unsupported_reason
                    ):
                        continue
                    range_text = reference or coordinate
                    try:
                        formula_range = CellRange(range_text)
                    except ValueError as exc:
                        raise UnsafeWorkbookError(
                            f"Malformed formula range {range_text!r} in {part!r}"
                        ) from exc
                    ranges[str(formula_range)] = formula_range
                result[sheet_name] = tuple(ranges[key] for key in sorted(ranges))
    except (zipfile.BadZipFile, OSError, RuntimeError) as exc:
        raise UnsafeWorkbookError(
            f"Unable to inspect worksheet formula metadata safely: {exc}"
        ) from exc
    return result
