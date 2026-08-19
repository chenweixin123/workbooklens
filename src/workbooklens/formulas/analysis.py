"""Conservative A1 formula analysis that never evaluates workbook expressions."""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import cast

from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError
from openpyxl.formula.translate import Translator, TranslatorError
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string

CELL_RE = re.compile(r"^(?P<col_abs>\$?)(?P<col>[A-Z]{1,3})(?P<row_abs>\$?)(?P<row>\d{1,7})$")
SHEET_PREFIX_RE = re.compile(r"^(?P<sheet>(?:'(?:[^']|'')+'|[^'!]+)!)?(?P<body>.+)$")
EXTERNAL_RE = re.compile(r"\[[^\]]+\][^!]*!")
WHOLE_COLUMN_RE = re.compile(r"(?<![A-Z0-9_])\$?[A-Z]{1,3}:\$?[A-Z]{1,3}(?![A-Z0-9_])", re.I)
STRUCTURED_REFERENCE_RE = re.compile(r"(?:\b[A-Z_][A-Z0-9_.]*\s*)?\[[^\]]+\]", re.I)
ADVANCED_FUNCTIONS = frozenset(
    {"FILTER", "UNIQUE", "SORT", "SORTBY", "SEQUENCE", "RANDARRAY", "LAMBDA"}
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
    broken_references: tuple[str, ...]
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
    except (ValueError, IndexError, TokenizerError) as exc:
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


def _tokens(formula: str) -> tuple[Token, ...] | None:
    try:
        return tuple(Tokenizer(formula).items)
    except (ValueError, IndexError, TokenizerError):
        return None


def _range_operands(tokens: tuple[Token, ...]) -> list[str]:
    return [token.value for token in tokens if token.type == "OPERAND" and token.subtype == "RANGE"]


def _function_name(token: Token) -> str | None:
    if token.type != "FUNC" or token.subtype != "OPEN":
        return None
    value = cast(str, token.value)[:-1].upper()
    return value.rsplit(".", maxsplit=1)[-1]


def _without_string_literals(formula: str) -> str:
    """Blank Excel double-quoted strings for conservative tokenizer-failure fallback checks."""

    output: list[str] = []
    index = 0
    inside = False
    while index < len(formula):
        character = formula[index]
        if character == '"':
            if inside and index + 1 < len(formula) and formula[index + 1] == '"':
                output.extend((" ", " "))
                index += 2
                continue
            inside = not inside
            output.append(" ")
        else:
            output.append(" " if inside else character)
        index += 1
    return "".join(output)


def analyze_formula(formula: str) -> FormulaFeatures:
    """Extract references and risky constructs while making no compatibility claims."""

    tokens = _tokens(formula)
    if tokens is None:
        # openpyxl currently rejects some valid newer syntax (for example spill references).
        # Remove string literals before conservative fallback matching so display text is inert.
        searchable = _without_string_literals(formula)
        upper = searchable.upper()
        external = tuple(sorted(set(EXTERNAL_RE.findall(searchable))))
        broken = ("#REF!",) if "#REF!" in upper else ()
        volatile = tuple(
            function
            for function in VOLATILE_FUNCTIONS
            if re.search(
                rf"(?<![A-Z0-9_])(?:(?:_XLFN|_XLWS)\.)*{function}\s*\(",
                upper,
            )
        )
        structured = bool(STRUCTURED_REFERENCE_RE.search(searchable)) and not external
        advanced = (
            any(
                re.search(
                    rf"(?<![A-Z0-9_])(?:(?:_XLFN|_XLWS)\.)*{function}\s*\(",
                    upper,
                )
                for function in ADVANCED_FUNCTIONS
            )
            or "#" in searchable
        )
        unsupported_reason = (
            "structured reference"
            if structured
            else "dynamic or advanced formula"
            if advanced
            else None
        )
        return FormulaFeatures(
            references=(),
            external_references=external,
            broken_references=broken,
            volatile_functions=volatile,
            has_whole_column_reference=bool(WHOLE_COLUMN_RE.search(searchable)),
            unsupported_reason=unsupported_reason,
        )

    operands = _range_operands(tokens)
    external = tuple(sorted({item for item in operands if EXTERNAL_RE.search(item)}))
    broken = tuple(
        sorted(
            {
                token.value
                for token in tokens
                if token.type == "OPERAND"
                and token.subtype in {"ERROR", "RANGE"}
                and "#REF!" in token.value.upper()
            }
        )
    )
    functions = {_function_name(token) for token in tokens}
    functions.discard(None)
    volatile = tuple(function for function in VOLATILE_FUNCTIONS if function in functions)
    structured = any(
        STRUCTURED_REFERENCE_RE.search(operand) and not EXTERNAL_RE.search(operand)
        for operand in operands
    )
    advanced = bool(functions & ADVANCED_FUNCTIONS) or any(
        "#" in token.value
        for token in tokens
        if token.type == "OPERAND" and token.subtype != "TEXT"
    )
    unsupported_reason = None
    if structured:
        unsupported_reason = "structured reference"
    elif advanced:
        unsupported_reason = "dynamic or advanced formula"
    return FormulaFeatures(
        references=tuple(operands),
        external_references=external,
        broken_references=broken,
        volatile_functions=volatile,
        has_whole_column_reference=any(WHOLE_COLUMN_RE.search(item) for item in operands),
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
