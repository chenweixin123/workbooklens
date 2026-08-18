"""Sparse, bounded region inference that avoids trusting worksheet dimensions."""

from __future__ import annotations

from collections import deque

from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.models import Confidence, Region


def _nonempty_coordinates(worksheet: Worksheet) -> set[tuple[int, int]]:
    return {
        (cell.row, cell.column)
        for cell in worksheet._cells.values()
        if isinstance(cell, Cell) and cell.value is not None
    }


def infer_data_regions(worksheet: Worksheet) -> list[Region]:
    """Infer dense connected components, excluding sparse decorative cell islands."""

    remaining = _nonempty_coordinates(worksheet)
    regions: list[Region] = []
    while remaining:
        seed = min(remaining)
        queue = deque([seed])
        component: set[tuple[int, int]] = set()
        remaining.remove(seed)
        while queue:
            current = queue.popleft()
            component.add(current)
            row, column = current
            for neighbor in (
                (row - 1, column),
                (row + 1, column),
                (row, column - 1),
                (row, column + 1),
            ):
                if neighbor in remaining:
                    remaining.remove(neighbor)
                    queue.append(neighbor)
        if len(component) < 6:
            continue
        rows = [item[0] for item in component]
        columns = [item[1] for item in component]
        min_row, max_row = min(rows), max(rows)
        min_column, max_column = min(columns), max(columns)
        height = max_row - min_row + 1
        width = max_column - min_column + 1
        if height < 2 or width < 2:
            continue
        density = len(component) / (height * width)
        if density < 0.5:
            continue
        regions.append(
            Region(
                sheet=worksheet.title,
                min_row=min_row,
                max_row=max_row,
                min_column=min_column,
                max_column=max_column,
                kind="data",
                confidence=Confidence(max(0.5, min(1.0, density))),
            )
        )
    return sorted(regions, key=lambda item: (item.min_row, item.min_column))


def _runs(values: list[int], allowed_gaps: int = 1) -> list[tuple[int, int, int]]:
    if not values:
        return []
    results: list[tuple[int, int, int]] = []
    start = previous = values[0]
    count = 1
    for value in values[1:]:
        if value - previous <= allowed_gaps + 1:
            count += 1
            previous = value
            continue
        results.append((start, previous, count))
        start = previous = value
        count = 1
    results.append((start, previous, count))
    return results


def infer_formula_bands(worksheet: Worksheet) -> list[Region]:
    """Find row/column formula runs with at most one intervening anomaly."""

    formula_cells = [
        cell
        for cell in worksheet._cells.values()
        if isinstance(cell, Cell) and cell.data_type == "f"
    ]
    by_row: dict[int, list[int]] = {}
    by_column: dict[int, list[int]] = {}
    for cell in formula_cells:
        row = cell.row
        column = cell.column
        by_row.setdefault(row, []).append(column)
        by_column.setdefault(column, []).append(row)
    bands: list[Region] = []
    for row, columns in by_row.items():
        for start, end, count in _runs(sorted(columns)):
            span = end - start + 1
            if count >= 3 and span - count <= 1:
                bands.append(
                    Region(
                        sheet=worksheet.title,
                        min_row=row,
                        max_row=row,
                        min_column=start,
                        max_column=end,
                        kind="formula_row",
                        confidence=Confidence(count / span),
                    )
                )
    for column, rows in by_column.items():
        for start, end, count in _runs(sorted(rows)):
            span = end - start + 1
            if count >= 3 and span - count <= 1:
                bands.append(
                    Region(
                        sheet=worksheet.title,
                        min_row=start,
                        max_row=end,
                        min_column=column,
                        max_column=column,
                        kind="formula_column",
                        confidence=Confidence(count / span),
                    )
                )
    return sorted(
        bands,
        key=lambda item: (item.kind, item.min_row, item.min_column, item.max_row, item.max_column),
    )
