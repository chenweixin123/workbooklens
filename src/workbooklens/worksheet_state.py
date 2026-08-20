"""Range-aware worksheet visibility helpers."""

from __future__ import annotations

from collections.abc import Iterator

from openpyxl.utils.cell import column_index_from_string, get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.exceptions import UnsafeWorkbookError

MAX_EXCEL_COLUMN = 16_384


def _validate_span(min_column: int, max_column: int) -> tuple[int, int]:
    if 1 <= min_column <= max_column <= MAX_EXCEL_COLUMN:
        return min_column, max_column
    raise UnsafeWorkbookError(
        f"Hidden column span {min_column}:{max_column} is outside Excel's valid A:XFD range"
    )


def _column_span(dimension: ColumnDimension) -> tuple[int, int] | None:
    min_column = dimension.min
    max_column = dimension.max
    if isinstance(min_column, int) and isinstance(max_column, int):
        return _validate_span(min_column, max_column)
    index = dimension.index
    if not isinstance(index, str):
        return None
    try:
        column = column_index_from_string(index)
    except ValueError:
        raise UnsafeWorkbookError(f"Invalid hidden column dimension index {index!r}") from None
    return _validate_span(column, column)


def hidden_column_spans(worksheet: Worksheet) -> Iterator[tuple[int, int]]:
    """Yield every hidden column span, including grouped dimensions."""

    for dimension in worksheet.column_dimensions.values():
        if not dimension.hidden:
            continue
        span = _column_span(dimension)
        if span is not None:
            yield span


def is_column_hidden(worksheet: Worksheet, column: int) -> bool:
    """Return whether a column falls inside any hidden dimension span."""

    return any(
        min_column <= column <= max_column
        for min_column, max_column in hidden_column_spans(worksheet)
    )


def hidden_column_labels(worksheet: Worksheet) -> list[str]:
    """Expand hidden dimension spans to deterministic column labels."""

    columns = {
        get_column_letter(column)
        for min_column, max_column in hidden_column_spans(worksheet)
        for column in range(min_column, max_column + 1)
    }
    return sorted(columns, key=column_index_from_string)
