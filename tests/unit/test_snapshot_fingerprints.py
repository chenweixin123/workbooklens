from __future__ import annotations

from copy import copy

from openpyxl import Workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.utils.indexed_list import IndexedList

from workbooklens.snapshot import cell_fingerprint


def _styled_formula_workbook() -> tuple[Workbook, object]:
    workbook = Workbook()
    worksheet = workbook.active
    assert worksheet is not None
    cell = worksheet["B2"]
    cell.value = "=ROUND(1.25*2,2)"
    cell.number_format = '"$"#,##0.00'
    cell.font = Font(name="Arial", bold=True)
    cell.border = Border()
    return workbook, cell


def test_cell_fingerprint_ignores_workbook_local_style_table_index() -> None:
    workbook, cell = _styled_formula_workbook()
    try:
        before_style_id = cell.style_id
        before = cell_fingerprint(cell)

        styles = list(workbook._cell_styles)
        dummy_style = copy(styles[0])
        dummy_style.fontId = 999
        workbook._cell_styles = IndexedList([dummy_style, *styles])

        assert cell.style_id != before_style_id
        assert cell_fingerprint(cell) == before
    finally:
        workbook.close()


def test_cell_fingerprint_still_detects_effective_border_change() -> None:
    workbook, cell = _styled_formula_workbook()
    try:
        before = cell_fingerprint(cell)
        cell.border = Border(top=Side(style="dashed"))
        assert cell_fingerprint(cell) != before
    finally:
        workbook.close()
