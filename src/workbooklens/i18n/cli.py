"""Localization helpers for canonical assertion messages shown by the CLI."""

from __future__ import annotations

import re

from workbooklens.i18n.catalog import normalize_language, translate
from workbooklens.i18n.rules import severity_label


def localize_assertion_message(message: str, language: str | None = None) -> str:
    """Translate known evaluator messages without changing assertion JSON."""

    if normalize_language(language) == "en":
        return message
    match = re.fullmatch(
        r"Observed (?P<observed>\d+) (?P<severity>[a-z]+) findings; maximum is (?P<maximum>\d+)",
        message,
    )
    if match:
        values = match.groupdict()
        return translate(
            "assertion.threshold",
            language,
            observed=values["observed"],
            severity=severity_label(values["severity"], language),
            maximum=values["maximum"],
        )
    patterns = (
        (r"Matched (?P<count>\d+) prohibited findings", "assertion.prohibited"),
        (r"Found (?P<count>\d+) duplicate value groups", "assertion.duplicates"),
        (r"Found (?P<count>\d+) values outside the allowed domain", "assertion.allowed"),
        (r"Found (?P<count>\d+) blank cells", "assertion.blank"),
        (r"Found (?P<count>\d+) nonnumeric or out-of-bounds cells", "assertion.numeric"),
        (r"Compared (?P<left>.+) with (?P<right>.+)", "assertion.compared"),
    )
    for pattern, key in patterns:
        match = re.fullmatch(pattern, message)
        if match:
            return translate(key, language, **match.groupdict())
    if message.startswith("Assertion could not be evaluated safely:"):
        return translate("assertion.unsafe", language)
    return message


__all__ = ["localize_assertion_message"]
