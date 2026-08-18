"""Generated workbook fixtures shared by unit, integration, and security tests."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest
from openpyxl import Workbook


@pytest.fixture
def workbook_factory(tmp_path: Path) -> Callable[[str], tuple[Workbook, Path]]:
    """Return a fresh workbook and deterministic destination in the test directory."""

    def factory(name: str = "workbook.xlsx") -> tuple[Workbook, Path]:
        workbook = Workbook()
        return workbook, tmp_path / name

    return factory
