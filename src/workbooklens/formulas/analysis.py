"""Conservative A1 formula analysis that never evaluates workbook expressions."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast

from openpyxl.formula import Tokenizer
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

CELL_RE = re.compile(r"^(?P<col_abs>\$?)(?P<col>[A-Z]{1,3})(?P<row_abs>\$?)(?P<row>\d{1,7})$")
SHEET_PREFIX_RE = re.compile(r"^(?P<sheet>(?:'(?:[^']|'')+'|[^'!]+)!)?(?P<body>.+)$")
EXTERNAL_RE = re.compile(r"\[[^\]]+\][^!]*!")
WHOLE_COLUMN_RE = re.compile(r"(?<![A-Z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Z0-9_])", re.I)
STRUCTURED_REFERENCE_RE = re.compile(r"(?:\b[A-Z_][A-Z0-9_.]*\s*)?\[[^\]]+\]", re.I)
ADVANCED_FORMULA_RE = re.compile(
    r"(?:(?:_xlfn|_xlws)\.)*(?:FILTER|UNIQUE|SORT|SORTBY|SEQUENCE|RANDARRAY|LAMBDA)\s*\(",
    re.IGNORECASE,
)
VOLATILE_FUNCTIONS = (
    "OFFSET",
    "INDIRECT",
    "NOW",
    "TODAY",
    "RAND",
    "RANDBETWEEN",
)


class UnsupportedFormulaError(ValueError):
    """Raised when translation would require guessing about an advanced formula syntax."""


@dataclass(frozen=True, slots=True)
class FormulaFeatures:
    """Non-executing feature summary used by rules and reports."""

    references: tuple[str, ...]
    external_references: tuple[str, ...]
    volatile_functions: tuple[str, ...]
    has_whole_column_reference: bool
    unsupported_reason: str | None = None


def _split_sheet(token: str) -> tuple[str, str]:
    match = SHEET_PREFIX_RE.match(token)
    if match is None:
        return "", token
    return match.group("sheet") or "", match.group("body")


def _r1c1_cell(cell_ref: str, origin: str) -> str | None:
    match = CELL_RE.fullmatch(cell_ref.upper())
    if match is None:
        return None
    origin_col_letters, origin_row = coordinate_from_string(origin.upper())
    origin_col = column_index_from_string(origin_col_letters)
    col = column_index_from_string(match.group("col"))
    row = int(match.group("row"))
    column_component = f"C{col}" if match.group("col_abs") else f"C[{col - origin_col}]"
    row_component = f"R{row}" if match.group("row_abs") else f"R[{row - origin_row}]"
    return row_component + column_component


def _normalize_range_token(token: str, origin: str) -> str:
    sheet, body = _split_sheet(token)
    if EXTERNAL_RE.search(token):
        return f"EXTERNAL({token.upper()})"
    if STRUCTURED_REFERENCE_RE.search(body):
        return f"UNSUPPORTED({token.upper()})"
    components = body.split(":")
    if len(components) > 2:
        return f"NAME({token.upper()})"
    normalized: list[str] = []
    for component in components:
        cell = _r1c1_cell(component, origin)
        if cell is None:
            return f"NAME({token.upper()})"
        normalized.append(cell)
    normalized_sheet = sheet.upper().replace("''", "'")
    return normalized_sheet + ":".join(normalized)


def normalize_formula(formula: str, origin: str) -> str:
    """Return an R1C1-like signature relative to ``origin`` without evaluating it."""

    if not isinstance(formula, str) or not formula.startswith("="):
        raise ValueError("formula must be a string beginning with '='")
    try:
        tokens = Tokenizer(formula).items
    except (ValueError, IndexError) as exc:
        raise UnsupportedFormulaError(f"formula tokenizer rejected the expression: {exc}") from exc
    result: list[str] = []
    for token in tokens:
        if token.type == "WHITE-SPACE":
            continue
        value = token.value
        if token.type == "OPERAND" and token.subtype == "RANGE":
            value = _normalize_range_token(value, origin)
        elif token.type == "FUNC":
            value = value.upper()
        result.append(value)
    return "".join(result)


def _range_operands(formula: str) -> list[str]:
    try:
        tokens = Tokenizer(formula).items
    except (ValueError, IndexError):
        return []
    return [token.value for token in tokens if token.type == "OPERAND" and token.subtype == "RANGE"]


def analyze_formula(formula: str) -> FormulaFeatures:
    """Extract references and risky constructs while making no compatibility claims."""

    operands = _range_operands(formula)
    external = tuple(sorted({item for item in operands if EXTERNAL_RE.search(item)}))
    upper = formula.upper()
    volatile = tuple(
        function
        for function in VOLATILE_FUNCTIONS
        if re.search(rf"(?<![A-Z0-9_])(?:(?:_XLFN|_XLWS)\.)*{function}\s*\(", upper)
    )
    unsupported_reason = None
    if STRUCTURED_REFERENCE_RE.search(formula) and not external:
        unsupported_reason = "structured reference"
    elif ADVANCED_FORMULA_RE.search(formula) or "#" in formula:
        unsupported_reason = "dynamic or advanced formula"
    return FormulaFeatures(
        references=tuple(operands),
        external_references=external,
        volatile_functions=volatile,
        has_whole_column_reference=bool(WHOLE_COLUMN_RE.search(formula)),
        unsupported_reason=unsupported_reason,
    )


def translate_formula(formula: str, origin: str, target: str) -> str:
    """Translate ordinary A1 references, rejecting external and structured references."""

    features = analyze_formula(formula)
    if features.external_references:
        raise UnsupportedFormulaError("external workbook references are never translated")
    if features.unsupported_reason:
        raise UnsupportedFormulaError(f"cannot safely translate {features.unsupported_reason}")
    try:
        translated = Translator(formula, origin=origin).translate_formula(target)
    except (TranslatorError, ValueError, TypeError) as exc:
        raise UnsupportedFormulaError(f"formula translation failed: {exc}") from exc
    if not translated.startswith("="):
        raise UnsupportedFormulaError("translated result is not an Excel formula")
    return cast(str, translated)
