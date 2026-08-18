from __future__ import annotations

import pytest
from pydantic import ValidationError

from workbooklens.models import Confidence, Severity


@pytest.mark.parametrize("value", [0.0, 0.25, 1.0])
def test_confidence_accepts_closed_unit_interval(value: float) -> None:
    assert float(Confidence(value)) == value


@pytest.mark.parametrize("value", [-0.01, 1.01])
def test_confidence_rejects_values_outside_unit_interval(value: float) -> None:
    with pytest.raises(ValidationError):
        Confidence(value)


def test_severity_values_are_stable() -> None:
    assert [value.value for value in Severity] == ["info", "warning", "error", "critical"]
