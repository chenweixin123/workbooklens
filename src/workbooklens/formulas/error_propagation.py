"""Conservative static propagation of unconditional Excel formula errors."""

from __future__ import annotations

import re
from collections import defaultdict, deque
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any, cast

from openpyxl.cell.cell import Cell
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import TokenizerError
from openpyxl.workbook.workbook import Workbook

from workbooklens.formulas.ir import FormulaReference, WorksheetContentIndex, parse_formula_ir

EXCEL_ERRORS = frozenset({"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"})
ERROR_PROPAGATING_AGGREGATES = frozenset({"AVERAGE", "MAX", "MIN", "SUM"})
ERROR_PROPAGATING_INFIX_OPERATORS = frozenset(
    {"+", "-", "*", "/", "^", "&", "=", "<>", "<", ">", "<=", ">="}
)
ERROR_PROPAGATING_PREFIX_OPERATORS = frozenset({"+", "-"})
ERROR_PROPAGATING_POSTFIX_OPERATORS = frozenset({"%"})
PURE_OPERATOR_OPERANDS = frozenset({"LOGICAL", "NUMBER", "TEXT"})
MAX_DEPENDENCIES_PER_REFERENCE = 4096
MAX_GRAPH_EDGES = 250_000
MAX_PURE_OPERATOR_TOKENS = 4096
MAX_PURE_OPERATOR_NESTING = 128
EXPLICIT_REF_FORMULA_RE = re.compile(r"^\s*=\s*#REF!\s*$", re.IGNORECASE)

FormulaNode = tuple[str, str]
DirectProof = Callable[[Any], tuple[str, str] | None]


@dataclass(frozen=True, slots=True)
class StaticFormulaErrorProof:
    """One unconditional formula-error proof and its immediate propagation witness."""

    error: str | None
    proof: str
    source: FormulaNode | None = None
    function: str | None = None
    source_proof: str | None = None
    reportable: bool = True


@dataclass(frozen=True, slots=True)
class _PropagationCandidate:
    references: tuple[FormulaReference, ...]
    function: str


def _raw_tokens(formula: str) -> tuple[Any, ...] | None:
    try:
        return tuple(Tokenizer(formula).items)
    except (IndexError, TokenizerError, ValueError):
        return None


def _tokens(formula: str) -> tuple[Any, ...] | None:
    tokens = _raw_tokens(formula)
    if tokens is None:
        return None
    return tuple(token for token in tokens if token.type != "WHITE-SPACE")


def _reference_token_key(value: Any) -> str:
    return str(value).replace("$", "").casefold()


def _pure_operator_references(
    formula: str,
    references: tuple[FormulaReference, ...],
) -> tuple[FormulaReference, ...] | None:
    """Recognize a bounded expression made only from error-propagating operators."""

    if not references or any(not reference.is_single_cell for reference in references):
        return None
    reference_lookup = {_reference_token_key(reference.raw): reference for reference in references}
    if len(reference_lookup) != len(references):
        return None
    tokens = _raw_tokens(formula)
    if (
        tokens is None
        or not tokens
        or len(tokens) > MAX_PURE_OPERATOR_TOKENS
        or any(token.type == "WHITE-SPACE" for token in tokens)
    ):
        # Excel whitespace can be the range-intersection operator, so this intentionally
        # rejects even cosmetically spaced expressions instead of guessing its role.
        return None

    expected_reference_keys = set(reference_lookup)
    seen_reference_keys: set[str] = set()
    expect_operand = True
    parenthesis_depth = 0
    previous_was_postfix = False
    for token in tokens:
        if token.type == "OPERAND":
            if not expect_operand:
                return None
            if token.subtype == "RANGE":
                raw = str(token.value)
                if raw.startswith("@") or "#" in raw or ":" in raw:
                    return None
                key = _reference_token_key(raw)
                if key not in reference_lookup:
                    # Reject names, structured references, and any range token that
                    # Formula IR did not resolve to the same ordinary A1 cell.
                    return None
                seen_reference_keys.add(key)
            elif token.subtype not in PURE_OPERATOR_OPERANDS:
                return None
            expect_operand = False
            previous_was_postfix = False
            continue
        if token.type == "PAREN" and token.subtype == "OPEN":
            if not expect_operand or parenthesis_depth >= MAX_PURE_OPERATOR_NESTING:
                return None
            parenthesis_depth += 1
            previous_was_postfix = False
            continue
        if token.type == "PAREN" and token.subtype == "CLOSE":
            if expect_operand or parenthesis_depth <= 0:
                return None
            parenthesis_depth -= 1
            previous_was_postfix = False
            continue
        if token.type == "OPERATOR-PREFIX":
            if not expect_operand or token.value not in ERROR_PROPAGATING_PREFIX_OPERATORS:
                return None
            previous_was_postfix = False
            continue
        if token.type == "OPERATOR-POSTFIX":
            if (
                expect_operand
                or previous_was_postfix
                or token.value not in ERROR_PROPAGATING_POSTFIX_OPERATORS
            ):
                return None
            previous_was_postfix = True
            continue
        if token.type == "OPERATOR-INFIX":
            if expect_operand or token.value not in ERROR_PROPAGATING_INFIX_OPERATORS:
                return None
            expect_operand = True
            previous_was_postfix = False
            continue
        return None

    if expect_operand or parenthesis_depth != 0 or seen_reference_keys != expected_reference_keys:
        return None
    return references


def _propagation_candidate(
    formula: str,
    sheet: str,
    coordinate: str,
) -> _PropagationCandidate | None:
    """Recognize only complete supported formulas that must propagate Excel errors."""

    try:
        ir = parse_formula_ir(formula, sheet, coordinate)
    except ValueError:
        return None
    if ir.features.external_references or ir.features.unsupported_reason or not ir.references:
        return None
    tokens = _tokens(formula)
    if tokens is None:
        return None
    if (
        len(ir.references) == 1
        and len(tokens) == 1
        and tokens[0].type == "OPERAND"
        and tokens[0].subtype == "RANGE"
        and ":" not in str(tokens[0].value)
        and ir.references[0].is_single_cell
    ):
        return _PropagationCandidate(
            references=(ir.references[0],),
            function="DIRECT_REFERENCE",
        )
    if len(ir.references) == 1 and len(tokens) == 3:
        opening, operand, closing = tokens
        if (
            opening.type == "FUNC"
            and opening.subtype == "OPEN"
            and operand.type == "OPERAND"
            and operand.subtype == "RANGE"
            and closing.type == "FUNC"
            and closing.subtype == "CLOSE"
        ):
            function = str(opening.value)[:-1].upper().rsplit(".", maxsplit=1)[-1]
            if function in ERROR_PROPAGATING_AGGREGATES:
                return _PropagationCandidate(
                    references=(ir.references[0],),
                    function=function,
                )
    operator_references = _pure_operator_references(formula, ir.references)
    if operator_references is None:
        return None
    return _PropagationCandidate(
        references=operator_references,
        function="PURE_OPERATOR_EXPRESSION",
    )


def find_static_formula_errors(
    workbook: Workbook,
    direct_formula_error: DirectProof,
    *,
    max_dependencies_per_reference: int = MAX_DEPENDENCIES_PER_REFERENCE,
    max_graph_edges: int = MAX_GRAPH_EDGES,
) -> dict[FormulaNode, StaticFormulaErrorProof]:
    """Return formula errors proved from exact error-propagating dependency chains.

    Cached formula results are deliberately absent from the seed set. The graph contains
    only formulas whose complete expression is a direct cell reference, one exact
    ``SUM``/``AVERAGE``/``MIN``/``MAX`` reference, or a strictly bounded expression of
    unconditional error-propagating operators. Conditional/error-handling functions,
    range algebra, dynamic arrays, and Excel aggregates that can ignore errors never
    inherit a proof.
    """

    formulas: dict[FormulaNode, Cell] = {}
    candidate_cells: dict[str, list[Cell]] = {}
    nodes_by_position: dict[str, dict[tuple[int, int], FormulaNode]] = {}
    positions: dict[FormulaNode, tuple[int, int, int]] = {}
    proofs: dict[FormulaNode, StaticFormulaErrorProof] = {}
    sheet_order = {worksheet.title: index for index, worksheet in enumerate(workbook.worksheets)}

    for worksheet in workbook.worksheets:
        retained: list[Cell] = []
        sheet_nodes: dict[tuple[int, int], FormulaNode] = {}
        for cell in worksheet._cells.values():
            if not isinstance(cell, Cell):
                continue
            node = (worksheet.title, cell.coordinate)
            if cell.data_type == "f" and isinstance(cell.value, str):
                formulas[node] = cell
                retained.append(cell)
                sheet_nodes[(cell.row, cell.column)] = node
                positions[node] = (sheet_order[worksheet.title], cell.row, cell.column)
                direct = direct_formula_error(cell.value)
                if direct is not None:
                    error, proof = direct
                    proofs[node] = StaticFormulaErrorProof(error=error, proof=proof)
                elif EXPLICIT_REF_FORMULA_RE.fullmatch(cell.value):
                    # Keep a broken-reference formula in the propagation graph,
                    # while WL001 remains the sole finding for the source cell.
                    proofs[node] = StaticFormulaErrorProof(
                        error="#REF!",
                        proof="explicit_broken_reference_token",
                        reportable=False,
                    )
            elif cell.data_type == "e" and str(cell.value).upper() in EXCEL_ERRORS:
                retained.append(cell)
                sheet_nodes[(cell.row, cell.column)] = node
                positions[node] = (sheet_order[worksheet.title], cell.row, cell.column)
                proofs[node] = StaticFormulaErrorProof(
                    error=str(cell.value).upper(),
                    proof="literal_error_value",
                )
        candidate_cells[worksheet.title] = retained
        nodes_by_position[worksheet.title] = sheet_nodes

    indexes = {
        worksheet.title: WorksheetContentIndex.from_cells(
            worksheet,
            candidate_cells[worksheet.title],
        )
        for worksheet in workbook.worksheets
    }
    worksheets = {worksheet.title.casefold(): worksheet.title for worksheet in workbook.worksheets}
    reverse: dict[FormulaNode, list[tuple[FormulaNode, str]]] = defaultdict(list)
    edge_count = 0

    for node, cell in sorted(formulas.items(), key=lambda item: positions[item[0]]):
        candidate = _propagation_candidate(cast(str, cell.value), node[0], node[1])
        if candidate is None:
            continue
        if len(candidate.references) > max_dependencies_per_reference:
            continue
        if candidate.function == "PURE_OPERATOR_EXPRESSION" and any(
            reference.sheet.casefold() == node[0].casefold()
            and reference.min_row == cell.row
            and reference.min_column == cell.column
            for reference in candidate.references
        ):
            # Circular-calculation settings make self-referential expressions unsafe to
            # classify through ordinary error propagation, even with another error source.
            continue
        dependencies: set[FormulaNode] = set()
        oversized_reference = False
        unresolved_reference = False
        for reference in candidate.references:
            target_sheet = worksheets.get(reference.sheet.casefold())
            if target_sheet is None:
                unresolved_reference = True
                break
            hits = tuple(
                indexes[target_sheet].iter_nonblank_rectangle(
                    min_row=reference.min_row,
                    max_row=reference.max_row,
                    min_column=reference.min_column,
                    max_column=reference.max_column,
                    limit=max_dependencies_per_reference + 1,
                )
            )
            if len(hits) > max_dependencies_per_reference:
                oversized_reference = True
                break
            target_nodes = nodes_by_position[target_sheet]
            dependencies.update(
                target_nodes[(hit.row, hit.column)]
                for hit in hits
                if (hit.row, hit.column) in target_nodes
            )
        if oversized_reference or unresolved_reference:
            continue
        dependencies.discard(node)
        if edge_count + len(dependencies) > max_graph_edges:
            continue
        for source in sorted(dependencies, key=positions.__getitem__):
            reverse[source].append((node, candidate.function))
        edge_count += len(dependencies)

    queue = deque(sorted(proofs, key=positions.__getitem__))
    while queue:
        source = queue.popleft()
        source_proof = proofs[source]
        for dependent, function in sorted(
            reverse.get(source, ()),
            key=lambda item: positions[item[0]],
        ):
            if dependent in proofs:
                continue
            proofs[dependent] = StaticFormulaErrorProof(
                error=(None if function == "PURE_OPERATOR_EXPRESSION" else source_proof.error),
                proof="propagated_formula_error",
                source=source,
                function=function,
                source_proof=source_proof.proof,
            )
            queue.append(dependent)

    return {node: proof for node, proof in proofs.items() if node in formulas}


__all__ = ["FormulaNode", "StaticFormulaErrorProof", "find_static_formula_errors"]
