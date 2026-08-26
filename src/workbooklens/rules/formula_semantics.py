"""Conservative, report-only formula semantics rules."""

from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from collections.abc import Mapping, Sequence
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Literal, cast

from openpyxl.cell.cell import Cell
from openpyxl.formula import Tokenizer
from openpyxl.formula.tokenizer import Token, TokenizerError
from openpyxl.utils.cell import get_column_letter, quote_sheetname
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.formulas import UnsupportedFormulaError, analyze_formula, normalize_formula
from workbooklens.formulas.ir import (
    FormulaIR,
    FormulaReference,
    inspect_reference_target,
    parse_formula_ir,
    worksheet_content_index,
)
from workbooklens.models import (
    Confidence,
    Evidence,
    Finding,
    PatchDerivation,
    PatchKind,
    PatchOperation,
    PatchPrecondition,
    PatchRisk,
    Region,
    Severity,
)
from workbooklens.rules.base import RuleContext, RuleResult, WorkbookRule
from workbooklens.snapshot import cell_fingerprint
from workbooklens.utils import stable_id
from workbooklens.worksheet_state import is_column_hidden

SUMMARY_RE = re.compile(
    r"(?<![a-z0-9])(?:grand[\s-]+total|subtotal|total|average|avg|count|summary|sum)"
    r"(?![a-z0-9])|(?:合计|总计|小计|汇总|平均|总数|计数)",
    re.IGNORECASE,
)
NOTE_TOKENS = {
    "comment",
    "comments",
    "description",
    "descriptions",
    "instruction",
    "instructions",
    "memo",
    "memos",
    "note",
    "notes",
    "remark",
    "remarks",
}
NOTE_MARKERS = ("备注", "说明", "注释", "意见", "描述", "附注")
RATE_TOKENS = {
    "churn",
    "completion",
    "conversion",
    "discount",
    "growth",
    "interest",
    "margin",
    "percent",
    "percentage",
    "ratio",
    "retention",
    "share",
    "success",
    "tax",
}
RATE_QUALIFIERS = RATE_TOKENS - {"percent", "percentage", "ratio"}
RATE_MARKERS = (
    "百分比",
    "比例",
    "占比",
    "增长率",
    "转换率",
    "转化率",
    "执行率",
    "完成率",
    "成功率",
    "利润率",
    "税率",
    "折扣率",
)
AMOUNT_TOKENS = {
    "amount",
    "budget",
    "cost",
    "expense",
    "income",
    "payroll",
    "price",
    "profit",
    "revenue",
    "salary",
    "sales",
    "turnover",
    "wage",
}
AMOUNT_MARKERS = (
    "金额",
    "预算",
    "成本",
    "费用",
    "收入",
    "营收",
    "销售额",
    "单价",
    "价格",
    "工资",
    "薪资",
)
AGGREGATE_METRIC_TOKENS = {
    "amount",
    "budget",
    "cost",
    "expense",
    "gross",
    "income",
    "inventory",
    "payroll",
    "profit",
    "reimbursement",
    "revenue",
    "salary",
    "sales",
    "turnover",
    "value",
}
AGGREGATE_METRIC_MARKERS = (
    "金额",
    "预算",
    "成本",
    "支出",
    "收入",
    "营收",
    "报销",
    "库存",
    "工资",
    "薪资",
    "销售额",
    "价值",
)
WINDOW_SCOPE_TOKENS = {
    "bottom",
    "first",
    "last",
    "period",
    "q1",
    "q2",
    "q3",
    "q4",
    "recent",
    "rolling",
    "top",
    "window",
}
WINDOW_SCOPE_MARKERS = (
    "最近",
    "近几",
    "滚动",
    "第一季度",
    "第二季度",
    "第三季度",
    "第四季度",
    "前几",
    "后几",
)
CURRENCY_MARKERS = ("$", "€", "£", "¥", "￥", "usd", "eur", "gbp", "cny", "rmb")
QUANTITY_TOKENS = {"count", "qty", "quantities", "quantity", "units"}
QUANTITY_MARKERS = ("件数", "数量", "数目", "个数", "套数", "台数")
NUMERIC_AGGREGATE_KINDS = frozenset({"formula", "number"})
CONDITIONAL_CRITERIA_KINDS = frozenset({"date", "formula", "logical", "number", "text"})
NUMERIC_COMPARISON_RE = re.compile(
    r"^(?P<operator><=|>=|<>|=|<|>)?\s*"
    r"(?P<value>[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)$"
)
SIMPLE_COLUMN_RANGE_RE = re.compile(
    r"^(?P<sheet>(?:'(?:[^']|'')+'|[^'!]+)!)?"
    r"(?P<column1>\$?[A-Z]{1,3})(?P<row1>\$?\d+):"
    r"(?P<column2>\$?[A-Z]{1,3})(?P<row2>\$?\d+)$",
    re.IGNORECASE,
)


def _confidence(value: float) -> Confidence:
    return Confidence(max(0.0, min(1.0, value)))


def _make_finding(
    *,
    context: RuleContext,
    rule_id: str,
    title: str,
    explanation: str,
    severity: Severity,
    confidence: float,
    worksheet: Worksheet,
    location: str,
    evidence: Evidence,
    expected: str,
    suggested_action: str,
    discriminator: Any = None,
    patches: Sequence[PatchOperation] = (),
) -> Finding:
    return Finding(
        id=stable_id("finding", rule_id, worksheet.title, location, discriminator),
        content_fingerprint=stable_id(
            "finding-content", evidence.model_dump(mode="json"), length=24
        ),
        rule_id=rule_id,
        title=title,
        explanation=explanation,
        severity=severity,
        confidence=_confidence(confidence),
        workbook=context.path.name,
        sheet=worksheet.title,
        location=location,
        evidence=evidence,
        expected=expected,
        suggested_action=suggested_action,
        safe_patch_available=any(patch.safe_only_eligible for patch in patches),
        patch_ids=[patch.id for patch in patches],
    )


def _make_formula_patch(
    *,
    worksheet: Worksheet,
    cell: Cell,
    formula: str,
    derivation: PatchDerivation,
    description: str,
    prerequisite_patch_ids: Sequence[str] = (),
) -> PatchOperation:
    return PatchOperation(
        id=stable_id(
            "patch",
            PatchKind.SET_FORMULA.value,
            worksheet.title,
            cell.coordinate,
            cell.value,
            formula,
        ),
        kind=PatchKind.SET_FORMULA,
        sheet=worksheet.title,
        cell=cell.coordinate,
        before=cell.value,
        after=formula,
        confidence=_confidence(0.99),
        safe=False,
        risk=PatchRisk.FORMULA_DERIVED,
        description=description,
        precondition=PatchPrecondition(
            cell_fingerprint=cell_fingerprint(cell),
            expected_formula=str(cell.value),
            expected_style_id=cell.style_id,
        ),
        atomic_group=stable_id("atomic", "formula-derived", worksheet.title, cell.coordinate),
        prerequisite_patch_ids=list(prerequisite_patch_ids),
        derivation=derivation,
    )


def _formula_cells(context: RuleContext, worksheet: Worksheet) -> tuple[Cell, ...]:
    cache_key = "formula-semantics-formula-cells-v1"
    cached = context.analysis_cache.get(cache_key)
    if not isinstance(cached, dict):
        cached = {}
        context.analysis_cache[cache_key] = cached
    formulas = cached.get(worksheet.title)
    if isinstance(formulas, tuple) and all(isinstance(cell, Cell) for cell in formulas):
        return formulas
    formulas = tuple(
        sorted(
            (
                cell
                for cell in worksheet._cells.values()
                if isinstance(cell, Cell) and cell.data_type == "f" and isinstance(cell.value, str)
            ),
            key=lambda cell: (cell.row, cell.column),
        )
    )
    cached[worksheet.title] = formulas
    return formulas


def _formula_ir(context: RuleContext, worksheet: Worksheet, cell: Cell) -> FormulaIR:
    cache_key = "formula-semantics-ir-v1"
    cache = context.analysis_cache.setdefault(cache_key, {})
    if not isinstance(cache, dict):
        cache = {}
        context.analysis_cache[cache_key] = cache
    key = (worksheet.title, cell.coordinate, cell.value)
    cached = cache.get(key)
    if isinstance(cached, FormulaIR):
        return cached
    parsed = parse_formula_ir(str(cell.value), worksheet.title, cell.coordinate)
    cache[key] = parsed
    return parsed


def _worksheet_by_name(context: RuleContext, name: str) -> Worksheet | None:
    cache_key = "formula-semantics-worksheets-by-name-v1"
    cached = context.analysis_cache.get(cache_key)
    if not isinstance(cached, dict):
        cached = {sheet.title.casefold(): sheet for sheet in context.workbook.worksheets}
        context.analysis_cache[cache_key] = cached
    return cast(Worksheet | None, cached.get(name.casefold()))


def _is_nonblank(value: Any) -> bool:
    return value is not None and not (isinstance(value, str) and not value.strip())


def _cell_at(worksheet: Worksheet, row: int, column: int) -> Cell | None:
    cell = worksheet._cells.get((row, column))
    return cell if isinstance(cell, Cell) else None


def _normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return " ".join(unicodedata.normalize("NFKC", value).strip().casefold().split())


def _text_tokens(value: str) -> set[str]:
    return set(re.findall(r"[a-z0-9]+", value))


def _formula_tokens(value: Any) -> tuple[Token, ...] | None:
    if not isinstance(value, str) or not value.startswith("="):
        return None
    try:
        return tuple(token for token in Tokenizer(value).items if token.type != "WHITE-SPACE")
    except (IndexError, TokenizerError, ValueError):
        return None


def _token_function_name(token: Token) -> str | None:
    if token.type != "FUNC" or token.subtype != "OPEN":
        return None
    return str(token.value)[:-1].upper().rsplit(".", maxsplit=1)[-1]


def _simple_function_arguments(
    tokens: tuple[Token, ...] | None,
    functions: frozenset[str],
) -> tuple[str, tuple[tuple[Token, ...], ...]] | None:
    if not tokens or len(tokens) < 3:
        return None
    function = _token_function_name(tokens[0])
    if function not in functions or tokens[-1].type != "FUNC" or tokens[-1].subtype != "CLOSE":
        return None
    arguments: list[tuple[Token, ...]] = []
    current: list[Token] = []
    depth = 0
    for token in tokens[1:-1]:
        if token.type in {"FUNC", "PAREN"} and token.subtype == "OPEN":
            depth += 1
        elif token.type in {"FUNC", "PAREN"} and token.subtype == "CLOSE":
            depth -= 1
            if depth < 0:
                return None
        if token.type == "SEP" and token.subtype == "ARG" and depth == 0:
            if not current:
                return None
            arguments.append(tuple(current))
            current = []
        else:
            current.append(token)
    if depth != 0 or not current:
        return None
    arguments.append(tuple(current))
    return function, tuple(arguments)


def _reference_argument(
    ir: FormulaIR,
    argument: tuple[Token, ...],
) -> FormulaReference | None:
    if len(argument) != 1 or argument[0].type != "OPERAND" or argument[0].subtype != "RANGE":
        return None
    raw = str(argument[0].value).replace("$", "").strip().casefold()
    return next(
        (
            reference
            for reference in ir.references
            if reference.raw.replace("$", "").strip().casefold() == raw
        ),
        None,
    )


def _conditional_aggregate_ranges(
    ir: FormulaIR,
) -> tuple[str, tuple[tuple[str, FormulaReference], ...]] | None:
    parsed = _simple_function_arguments(
        _formula_tokens(ir.formula),
        frozenset({"SUMIF", "SUMIFS"}),
    )
    if parsed is None:
        return None
    function, arguments = parsed
    ranges: list[tuple[str, FormulaReference]] = []
    if function == "SUMIF":
        if len(arguments) not in {2, 3}:
            return None
        criteria_range = _reference_argument(ir, arguments[0])
        sum_range = _reference_argument(ir, arguments[2]) if len(arguments) == 3 else criteria_range
        if criteria_range is None or sum_range is None:
            return None
        ranges = [
            ("criteria_range", criteria_range),
            ("sum_range", sum_range),
        ]
    else:
        if len(arguments) < 3 or (len(arguments) - 1) % 2:
            return None
        sum_range = _reference_argument(ir, arguments[0])
        if sum_range is None:
            return None
        ranges.append(("sum_range", sum_range))
        for index in range(1, len(arguments), 2):
            criteria_range = _reference_argument(ir, arguments[index])
            if criteria_range is None or not arguments[index + 1]:
                return None
            ranges.append((f"criteria_range_{(index + 1) // 2}", criteria_range))
    references = [reference for _, reference in ranges]
    if any(
        not reference.is_cross_sheet or not reference.is_single_column for reference in references
    ):
        return None
    if len({reference.sheet.casefold() for reference in references}) != 1:
        return None
    if len({(reference.min_row, reference.max_row) for reference in references}) != 1:
        return None
    return function, tuple(ranges)


def _top_level_division_denominator(
    tokens: tuple[Token, ...] | None,
) -> tuple[Token, ...] | None:
    if not tokens:
        return None
    depth = 0
    operators: list[tuple[int, Token]] = []
    for index, token in enumerate(tokens):
        if token.type in {"FUNC", "PAREN"} and token.subtype == "OPEN":
            depth += 1
        elif token.type in {"FUNC", "PAREN"} and token.subtype == "CLOSE":
            depth -= 1
            if depth < 0:
                return None
        elif depth == 0 and token.type == "OPERATOR-INFIX":
            operators.append((index, token))
    if depth != 0 or len(operators) != 1 or operators[0][1].value != "/":
        return None
    index = operators[0][0]
    if index == 0 or index + 1 >= len(tokens):
        return None
    return tokens[index + 1 :]


def _numeric_countif_criterion(
    argument: tuple[Token, ...],
) -> tuple[str, Decimal] | None:
    if len(argument) != 1 or argument[0].type != "OPERAND":
        return None
    token = argument[0]
    if token.subtype == "TEXT":
        raw = str(token.value)
        if len(raw) < 2 or raw[0] != '"' or raw[-1] != '"':
            return None
        raw = raw[1:-1].replace('""', '"')
    elif token.subtype == "NUMBER":
        raw = str(token.value)
    else:
        return None
    match = NUMERIC_COMPARISON_RE.fullmatch(raw.strip())
    if match is None:
        return None
    return match.group("operator") or "=", Decimal(match.group("value"))


def _comparison_matches(value: Decimal, operator: str, threshold: Decimal) -> bool:
    if operator == "=":
        return value == threshold
    if operator == "<>":
        return value != threshold
    if operator == ">":
        return value > threshold
    if operator == ">=":
        return value >= threshold
    if operator == "<":
        return value < threshold
    return value <= threshold


def _regions_for(context: RuleContext, worksheet: Worksheet) -> tuple[Region, ...]:
    return tuple(context.data_regions.get(worksheet.title, ()))


def _region_for_cell(
    context: RuleContext,
    worksheet: Worksheet,
    row: int,
    column: int,
) -> Region | None:
    candidates = [
        region
        for region in _regions_for(context, worksheet)
        if region.min_row <= row <= region.max_row
        and region.min_column <= column <= region.max_column
    ]
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda region: (
            (region.max_row - region.min_row + 1) * (region.max_column - region.min_column + 1),
            region.min_row,
            region.min_column,
        ),
    )


def _region_for_reference(
    context: RuleContext,
    worksheet: Worksheet,
    reference: FormulaReference,
) -> Region | None:
    candidates = [
        region
        for region in _regions_for(context, worksheet)
        if region.min_column <= reference.min_column <= reference.max_column <= region.max_column
        and region.min_row <= reference.min_row <= reference.max_row <= region.max_row
        and region.max_column > region.min_column
    ]
    if not candidates:
        return None
    return min(
        candidates,
        key=lambda region: (
            (region.max_row - region.min_row + 1) * (region.max_column - region.min_column + 1),
            region.min_row,
            region.min_column,
        ),
    )


def _row_has_summary_label(context: RuleContext, worksheet: Worksheet, row: int) -> bool:
    cache_key = "formula-semantics-summary-row-labels-v1"
    cached = context.analysis_cache.setdefault(cache_key, {})
    if not isinstance(cached, dict):
        cached = {}
        context.analysis_cache[cache_key] = cached
    key = (worksheet.title, row)
    if key in cached:
        return cast(bool, cached[key])
    result = any(
        cell.row == row
        and cell.data_type != "f"
        and isinstance(cell.value, str)
        and SUMMARY_RE.search(_normalize_text(cell.value))
        for cell in worksheet_content_index(context, worksheet).iter_nonblank_row(row)
    )
    cached[key] = result
    return result


def _summary_label_cells(context: RuleContext, worksheet: Worksheet, row: int) -> tuple[Cell, ...]:
    return tuple(
        cell
        for cell in worksheet_content_index(context, worksheet).iter_nonblank_row(row)
        if cell.data_type != "f"
        and isinstance(cell.value, str)
        and SUMMARY_RE.search(_normalize_text(cell.value))
    )


def _nearby_explicit_aggregate_label(
    worksheet: Worksheet,
    cell: Cell,
    function: str,
) -> Cell | None:
    """Return a tight upper-left label that explicitly names this aggregate role."""

    markers: dict[str, tuple[re.Pattern[str], tuple[str, ...]]] = {
        "SUM": (
            re.compile(
                r"(?<![a-z0-9])(?:grand[\s-]+total|subtotal|total|sum)(?![a-z0-9])",
                re.IGNORECASE,
            ),
            ("合计", "总计", "小计", "汇总"),
        ),
        "AVERAGE": (
            re.compile(r"(?<![a-z0-9])(?:average|avg)(?![a-z0-9])", re.IGNORECASE),
            ("平均",),
        ),
        "COUNT": (
            re.compile(r"(?<![a-z0-9])count(?![a-z0-9])", re.IGNORECASE),
            ("计数", "总数"),
        ),
        "MIN": (re.compile(r"(?<![a-z0-9])min(?:imum)?(?![a-z0-9])", re.IGNORECASE), ("最小",)),
        "MAX": (re.compile(r"(?<![a-z0-9])max(?:imum)?(?![a-z0-9])", re.IGNORECASE), ("最大",)),
    }
    marker = markers.get(function.upper())
    if marker is None:
        return None
    pattern, localized_markers = marker
    for row in range(max(1, cell.row - 2), cell.row):
        for column in range(max(1, cell.column - 1), cell.column + 1):
            candidate = _cell_at(worksheet, row, column)
            if (
                candidate is None
                or candidate.data_type == "f"
                or not isinstance(candidate.value, str)
            ):
                continue
            normalized = _normalize_text(candidate.value)
            if pattern.search(normalized) or any(item in normalized for item in localized_markers):
                return candidate
    return None


def _row_has_unscoped_metric_label(context: RuleContext, worksheet: Worksheet, row: int) -> bool:
    cache_key = "formula-semantics-unscoped-metric-row-labels-v1"
    cached = context.analysis_cache.setdefault(cache_key, {})
    if not isinstance(cached, dict):
        cached = {}
        context.analysis_cache[cache_key] = cached
    key = (worksheet.title, row)
    if key in cached:
        return cast(bool, cached[key])
    result = False
    for cell in worksheet_content_index(context, worksheet).iter_nonblank_row(row):
        if cell.data_type == "f" or not isinstance(cell.value, str):
            continue
        normalized = _normalize_text(cell.value)
        tokens = _text_tokens(normalized)
        if tokens & WINDOW_SCOPE_TOKENS or any(
            marker in normalized for marker in WINDOW_SCOPE_MARKERS
        ):
            continue
        if tokens & RATE_TOKENS or any(marker in normalized for marker in RATE_MARKERS):
            continue
        if tokens & AGGREGATE_METRIC_TOKENS or any(
            marker in normalized for marker in AGGREGATE_METRIC_MARKERS
        ):
            result = True
            break
    cached[key] = result
    return result


def _row_is_summary(context: RuleContext, worksheet: Worksheet, region: Region, row: int) -> bool:
    return any(
        cell.data_type != "f"
        and isinstance(cell.value, str)
        and SUMMARY_RE.search(_normalize_text(cell.value))
        for cell in worksheet_content_index(context, worksheet).iter_nonblank_row(
            row,
            min_column=region.min_column,
            max_column=region.max_column,
        )
    )


def _literal_value_kind(value: Any, data_type: str | None = None) -> str | None:
    if not _is_nonblank(value):
        return None
    if data_type == "f":
        return "formula"
    if isinstance(value, bool):
        return "logical"
    if isinstance(value, (int, float, Decimal)):
        return "number"
    if isinstance(value, (date, datetime)):
        return "date"
    if isinstance(value, str):
        return "text"
    return "other"


def _safe_normalization_patches(
    context: RuleContext,
) -> dict[tuple[str, str], PatchOperation]:
    """Return canonical safe numeric normalizations without mutating the workbook."""

    cache_key = "formula-semantics-safe-normalizations-v1"
    cached = context.analysis_cache.get(cache_key)
    if isinstance(cached, dict):
        return cast(dict[tuple[str, str], PatchOperation], cached)
    from workbooklens.rules.profile_quality import LosslessValueNormalizationRule

    proposals = LosslessValueNormalizationRule().run(context)
    patches = {
        (patch.sheet.casefold(), patch.cell.upper()): patch
        for patch in proposals.patches
        if patch.kind == PatchKind.SET_NUMERIC and patch.safe_only_eligible
    }
    context.analysis_cache[cache_key] = patches
    return patches


def _value_kind(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell | None,
) -> str | None:
    if cell is None or not _is_nonblank(cell.value):
        return None
    actual = _literal_value_kind(cell.value, cell.data_type)
    if actual != "text":
        return actual
    patch = _safe_normalization_patches(context).get(
        (worksheet.title.casefold(), cell.coordinate.upper())
    )
    if patch is None or patch.precondition.cell_fingerprint != cell_fingerprint(cell):
        return actual
    return _literal_value_kind(patch.after)


def _normalization_prerequisite_id(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell | None,
) -> str | None:
    if cell is None or _literal_value_kind(cell.value, cell.data_type) != "text":
        return None
    patch = _safe_normalization_patches(context).get(
        (worksheet.title.casefold(), cell.coordinate.upper())
    )
    if (
        patch is None
        or patch.precondition.cell_fingerprint != cell_fingerprint(cell)
        or _literal_value_kind(patch.after) != "number"
    ):
        return None
    return patch.id


def _append_patch_with_prerequisites(
    result: RuleResult,
    context: RuleContext,
    patch: PatchOperation,
) -> None:
    from workbooklens.rules.profile_quality import LosslessValueNormalizationRule

    proposals = LosslessValueNormalizationRule().run(context)
    available = [*context.prior_patches, *proposals.patches]
    by_id = {candidate.id: candidate for candidate in available}
    for prerequisite_id in patch.prerequisite_patch_ids:
        prerequisite = by_id.get(prerequisite_id)
        if prerequisite is None:
            raise RuntimeError(f"Formula patch {patch.id} references an unavailable prerequisite")
        required = (
            [
                candidate
                for candidate in available
                if candidate.atomic_group == prerequisite.atomic_group
            ]
            if prerequisite.atomic_group is not None
            else [prerequisite]
        )
        existing_ids = {candidate.id for candidate in result.patches}
        result.patches.extend(
            candidate for candidate in required if candidate.id not in existing_ids
        )
    result.patches.append(patch)


def _dominant_reference_kind(
    context: RuleContext,
    worksheet: Worksheet,
    reference: FormulaReference,
    *,
    allowed_kinds: frozenset[str] = NUMERIC_AGGREGATE_KINDS,
) -> tuple[str, float] | None:
    kinds = [
        kind
        for cell in worksheet_content_index(context, worksheet).iter_nonblank_rectangle(
            min_row=reference.min_row,
            max_row=reference.max_row,
            min_column=reference.min_column,
            max_column=reference.max_column,
        )
        if (kind := _value_kind(context, worksheet, cell)) is not None
    ]
    if len(kinds) < 3:
        return None
    kind, count = Counter(kinds).most_common(1)[0]
    ratio = count / len(kinds)
    if kind not in allowed_kinds or ratio < 0.75:
        return None
    return kind, ratio


def _contiguous_extension(
    context: RuleContext,
    origin_worksheet: Worksheet,
    origin_cell: Cell,
    reference: FormulaReference,
    *,
    allowed_kinds: frozenset[str] = NUMERIC_AGGREGATE_KINDS,
) -> tuple[tuple[str, ...], str, float] | None:
    if not reference.is_single_column:
        return None
    target_worksheet = _worksheet_by_name(context, reference.sheet)
    if target_worksheet is None:
        return None
    region = _region_for_reference(context, target_worksheet, reference)
    if region is None:
        return None
    dominant = _dominant_reference_kind(
        context,
        target_worksheet,
        reference,
        allowed_kinds=allowed_kinds,
    )
    if dominant is None:
        return None
    kind, ratio = dominant
    column = reference.min_column

    def candidate(row: int) -> Cell | None:
        if not region.min_row < row <= region.max_row or _row_is_summary(
            context, target_worksheet, region, row
        ):
            return None
        cell = _cell_at(target_worksheet, row, column)
        if (
            target_worksheet.title.casefold() == origin_worksheet.title.casefold()
            and row == origin_cell.row
            and column == origin_cell.column
        ):
            return None
        return cell if _value_kind(context, target_worksheet, cell) == kind else None

    prefix: list[str] = []
    for row in range(reference.min_row - 1, region.min_row, -1):
        cell = candidate(row)
        if cell is None:
            break
        prefix.append(cell.coordinate)
        if len(prefix) >= 256:
            break
    prefix.reverse()
    suffix: list[str] = []
    for row in range(reference.max_row + 1, region.max_row + 1):
        cell = candidate(row)
        if cell is None:
            break
        suffix.append(cell.coordinate)
        if len(suffix) >= 256:
            break
    extension = tuple(prefix + suffix)
    return (extension, kind, ratio) if extension else None


def _reference_key(reference: FormulaReference) -> tuple[str, int, int, int, int]:
    return (
        reference.sheet.casefold(),
        reference.min_row,
        reference.max_row,
        reference.min_column,
        reference.max_column,
    )


def _expanded_reference(reference: FormulaReference, extension: tuple[str, ...]) -> str:
    rows = [reference.min_row, reference.max_row]
    for coordinate in extension:
        cell_row = int(re.search(r"\d+$", coordinate).group())  # type: ignore[union-attr]
        rows.append(cell_row)
    match = SIMPLE_COLUMN_RANGE_RE.fullmatch(reference.raw)
    column = get_column_letter(reference.min_column)
    first_column = column
    second_column = column
    first_row_prefix = ""
    second_row_prefix = ""
    if match is not None:
        first_column = match.group("column1").replace(
            get_column_letter(reference.min_column),
            column,
        )
        second_column = match.group("column2").replace(
            get_column_letter(reference.max_column),
            column,
        )
        first_row_prefix = "$" if match.group("row1").startswith("$") else ""
        second_row_prefix = "$" if match.group("row2").startswith("$") else ""
    local = (
        f"{first_column}{first_row_prefix}{min(rows)}:{second_column}{second_row_prefix}{max(rows)}"
    )
    return f"{quote_sheetname(reference.sheet)}!{local}" if reference.explicit_sheet else local


def _formula_target_is_hidden(worksheet: Worksheet, cell: Cell) -> bool:
    row_dimension = worksheet.row_dimensions.get(cell.row)
    return bool(
        worksheet.sheet_state != "visible"
        or (row_dimension is not None and row_dimension.hidden)
        or is_column_hidden(worksheet, cell.column)
    )


def _formula_patch_target_is_safe(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
    candidate: str,
) -> bool:
    if (
        any(cell.coordinate in merged for merged in worksheet.merged_cells.ranges)
        or worksheet.protection.sheet
        or _formula_target_is_hidden(worksheet, cell)
        or any(
            cell.coordinate in formula_range
            for formula_range in context.unsupported_formula_ranges.get(worksheet.title, ())
        )
    ):
        return False
    current = analyze_formula(str(cell.value))
    proposed = analyze_formula(candidate)
    for features in (current, proposed):
        if (
            features.external_references
            or features.broken_references
            or features.volatile_functions
            or features.unsupported_reason
            or features.has_whole_column_reference
        ):
            return False
    from workbooklens.rules.builtin import (
        _circular_formula_components,
        _formula_candidate_creates_cycle,
    )

    if any(
        (worksheet.title, cell.coordinate) in component
        for component in _circular_formula_components(context)
    ):
        return False
    return not _formula_candidate_creates_cycle(context, worksheet, cell, candidate)


def _replace_formula_references(
    formula: str,
    replacements: Mapping[str, str],
) -> str | None:
    tokens = _formula_tokens(formula)
    if tokens is None or not replacements:
        return None
    remaining = set(replacements)
    pieces: list[str] = []
    for token in tokens:
        value = str(token.value)
        if token.type == "OPERAND" and token.subtype == "RANGE" and value in replacements:
            value = replacements[value]
            remaining.discard(str(token.value))
        pieces.append(value)
    if remaining:
        return None
    candidate = "=" + "".join(pieces)
    return candidate if candidate != formula else None


def _complete_reference_body(
    context: RuleContext,
    origin_worksheet: Worksheet,
    origin_cell: Cell,
    reference: FormulaReference,
    extension: tuple[str, ...],
    *,
    allowed_kinds: frozenset[str],
) -> tuple[Region, tuple[int, int], tuple[str, ...]] | None:
    target_worksheet = _worksheet_by_name(context, reference.sheet)
    if target_worksheet is None:
        return None
    region = _region_for_reference(context, target_worksheet, reference)
    if region is None or float(region.confidence) < 0.9:
        return None
    dominant = _dominant_reference_kind(
        context,
        target_worksheet,
        reference,
        allowed_kinds=allowed_kinds,
    )
    if dominant is None:
        return None
    kind, ratio = dominant
    if ratio < 0.9:
        return None
    rows = [reference.min_row, reference.max_row]
    rows.extend(
        int(match.group())
        for coordinate in extension
        if (match := re.search(r"\d+$", coordinate)) is not None
    )
    expected_min, expected_max = min(rows), max(rows)
    body_rows = []
    prerequisite_patch_ids: set[str] = set()
    for row in range(region.min_row + 1, region.max_row + 1):
        if _row_is_summary(context, target_worksheet, region, row):
            continue
        if (
            target_worksheet.title.casefold() == origin_worksheet.title.casefold()
            and row == origin_cell.row
            and reference.min_column == origin_cell.column
        ):
            continue
        body_cell = _cell_at(target_worksheet, row, reference.min_column)
        if _value_kind(context, target_worksheet, body_cell) != kind:
            return None
        prerequisite_id = _normalization_prerequisite_id(
            context,
            target_worksheet,
            body_cell,
        )
        if prerequisite_id is not None:
            prerequisite_patch_ids.add(prerequisite_id)
        body_rows.append(row)
    if (
        not body_rows
        or body_rows != list(range(body_rows[0], body_rows[-1] + 1))
        or (expected_min, expected_max) != (body_rows[0], body_rows[-1])
    ):
        return None
    return region, (expected_min, expected_max), tuple(sorted(prerequisite_patch_ids))


def _ordinary_formula_signature(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
    formula: str,
) -> str | None:
    if any(
        cell.coordinate in formula_range
        for formula_range in context.unsupported_formula_ranges.get(worksheet.title, ())
    ):
        return None
    from workbooklens.rules.builtin import _formula_is_ordinary_derived_candidate

    if not _formula_is_ordinary_derived_candidate(context, worksheet, formula):
        return None
    try:
        return normalize_formula(formula, cell.coordinate)
    except (ValueError, UnsupportedFormulaError):
        return None


def _prior_formula_repairs(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
) -> tuple[PatchOperation, ...]:
    repairs = []
    for patch in context.prior_patches:
        if (
            patch.sheet.casefold() != worksheet.title.casefold()
            or patch.cell.upper() != cell.coordinate.upper()
            or patch.risk != PatchRisk.FORMULA_DERIVED
            or patch.kind != PatchKind.SET_FORMULA
            or patch.precondition.cell_fingerprint != cell_fingerprint(cell)
            or patch.derivation.candidate_count != 1
            or not patch.derivation.requires_recalculation
            or not isinstance(patch.after, str)
            or not patch.after.startswith("=")
        ):
            continue
        repairs.append(patch)
    return tuple(sorted(repairs, key=lambda patch: patch.id))


def _complete_formula_reference_body(
    context: RuleContext,
    origin_worksheet: Worksheet,
    origin_cell: Cell,
    reference: FormulaReference,
    extension: tuple[str, ...],
) -> tuple[Region, tuple[int, int], tuple[str, ...]] | None:
    """Prove a full formula-column body after unique prior formula repairs."""

    target_worksheet = _worksheet_by_name(context, reference.sheet)
    if target_worksheet is None or not reference.is_single_column:
        return None
    region = _region_for_reference(context, target_worksheet, reference)
    if region is None or float(region.confidence) < 0.95:
        return None
    same_sheet_target = (
        target_worksheet.title.casefold() == origin_worksheet.title.casefold()
        and reference.min_column == origin_cell.column
    )
    body_rows: list[int] = []
    for row in range(region.min_row + 1, region.max_row + 1):
        if _row_is_summary(context, target_worksheet, region, row):
            if same_sheet_target and row == origin_cell.row:
                continue
            return None
        body_rows.append(row)
    if not body_rows or body_rows != list(range(body_rows[0], body_rows[-1] + 1)):
        return None
    first_body_row = body_rows[0]
    last_body_row = body_rows[-1]
    extension_rows = tuple(
        int(match.group())
        for coordinate in extension
        if (match := re.search(r"\d+$", coordinate)) is not None
    )
    if (
        reference.min_row != first_body_row
        or reference.max_row >= last_body_row
        or extension_rows != tuple(range(reference.max_row + 1, last_body_row + 1))
    ):
        return None
    body_cells: list[Cell] = []
    current_signatures: dict[str, str] = {}
    repair_signatures: dict[str, list[tuple[PatchOperation, str]]] = {}
    for row in body_rows:
        body_cell = _cell_at(target_worksheet, row, reference.min_column)
        if body_cell is None or not _is_nonblank(body_cell.value):
            return None
        body_cells.append(body_cell)
        if body_cell.data_type == "f" and isinstance(body_cell.value, str):
            signature = _ordinary_formula_signature(
                context,
                target_worksheet,
                body_cell,
                body_cell.value,
            )
            if signature is None:
                return None
            current_signatures[body_cell.coordinate] = signature
        candidates = []
        for patch in _prior_formula_repairs(context, target_worksheet, body_cell):
            signature = _ordinary_formula_signature(
                context,
                target_worksheet,
                body_cell,
                cast(str, patch.after),
            )
            if signature is not None:
                candidates.append((patch, signature))
        repair_signatures[body_cell.coordinate] = candidates

    formula_ratio = len(current_signatures) / len(body_cells)
    signature_counts = Counter(current_signatures.values())
    if formula_ratio < 0.9 or not signature_counts:
        return None
    ranked = signature_counts.most_common()
    dominant_signature, dominant_count = ranked[0]
    if (
        dominant_count < 8
        or dominant_count / len(current_signatures) < 0.9
        or (len(ranked) > 1 and ranked[1][1] == dominant_count)
    ):
        return None

    prerequisite_patch_ids: set[str] = set()
    for body_cell in body_cells:
        if current_signatures.get(body_cell.coordinate) == dominant_signature:
            continue
        matching_repairs = {
            patch.id: patch
            for patch, signature in repair_signatures[body_cell.coordinate]
            if signature == dominant_signature
        }
        if len(matching_repairs) != 1:
            return None
        prerequisite_patch_ids.update(matching_repairs)
    return (
        region,
        (first_body_row, last_body_row),
        tuple(sorted(prerequisite_patch_ids)),
    )


def _aggregate_label_source(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
    function: str,
    *,
    cross_sheet: bool,
) -> str | None:
    summary_cells = _summary_label_cells(context, worksheet, cell.row)
    function = function.upper()

    def explicitly_matches(value: str) -> bool:
        normalized = _normalize_text(value)
        if function == "AVERAGE":
            return bool(
                re.search(r"(?<![a-z0-9])(?:average|avg|mean)(?![a-z0-9])", normalized)
                or "平均" in normalized
                or "均值" in normalized
            )
        return bool(
            re.search(
                r"(?<![a-z0-9])(?:grand[\s-]+total|subtotal|total|sum)(?![a-z0-9])",
                normalized,
            )
            or any(marker in normalized for marker in ("合计", "总计", "小计", "汇总"))
        )

    matching_summary = next(
        (
            label
            for label in summary_cells
            if isinstance(label.value, str) and explicitly_matches(label.value)
        ),
        None,
    )
    if matching_summary is not None:
        return f"aggregate_label:{worksheet.title}!{matching_summary.coordinate}:{function}"
    if not cross_sheet or not _row_has_unscoped_metric_label(context, worksheet, cell.row):
        return None
    for label in worksheet_content_index(context, worksheet).iter_nonblank_row(cell.row):
        if label.data_type == "f" or not isinstance(label.value, str):
            continue
        normalized = _normalize_text(label.value)
        tokens = _text_tokens(normalized)
        if tokens & WINDOW_SCOPE_TOKENS or any(
            marker in normalized for marker in WINDOW_SCOPE_MARKERS
        ):
            continue
        if tokens & RATE_TOKENS or any(marker in normalized for marker in RATE_MARKERS):
            continue
        if function == "AVERAGE" and not explicitly_matches(label.value):
            continue
        if tokens & AGGREGATE_METRIC_TOKENS or any(
            marker in normalized for marker in AGGREGATE_METRIC_MARKERS
        ):
            return f"aggregate_metric_label:{worksheet.title}!{label.coordinate}:{function}"
    return None


def _range_derivation(
    *,
    strategy: str,
    label_source: str,
    source_sheet: str,
    region: Region,
    synchronized_ranges: bool,
    prerequisite_patch_ids: Sequence[str] = (),
    formula_repair_patch_ids: Sequence[str] = (),
) -> PatchDerivation:
    boundary = (
        f"{get_column_letter(region.min_column)}{region.min_row}:"
        f"{get_column_letter(region.max_column)}{region.max_row}"
    )
    invariants = [
        "unique_complete_table_boundary",
        "ordinary_bounded_a1_references",
        "no_external_dynamic_volatile_or_broken_reference",
        "target_not_merged_hidden_or_protected",
        "no_static_cycle",
    ]
    if synchronized_ranges:
        invariants.append("criteria_and_sum_ranges_share_identical_row_bounds")
    sources = [label_source, f"inferred_table_boundary:{source_sheet}!{boundary}"]
    if prerequisite_patch_ids:
        sources.append(
            "safe_normalization_prerequisite:" + ",".join(sorted(prerequisite_patch_ids))
        )
        invariants.append("lossless_normalization_prerequisites_selected")
    if formula_repair_patch_ids:
        sources.append("formula_repair_prerequisite:" + ",".join(sorted(formula_repair_patch_ids)))
        invariants.append("formula_repairs_selected_before_range_extension")
    return PatchDerivation(
        strategy=strategy,
        sources=sources,
        candidate_count=1,
        invariants=invariants,
        requires_recalculation=True,
    )


def _conditional_aggregate_extensions(
    context: RuleContext,
    origin_worksheet: Worksheet,
    origin_cell: Cell,
    ranges: tuple[tuple[str, FormulaReference], ...],
) -> (
    tuple[
        tuple[tuple[str, FormulaReference, tuple[str, ...], str, float], ...],
        tuple[int, ...],
    ]
    | None
):
    region_keys: set[tuple[int, int, int, int]] = set()
    for _, reference in ranges:
        target_worksheet = _worksheet_by_name(context, reference.sheet)
        if target_worksheet is None:
            return None
        region = _region_for_reference(context, target_worksheet, reference)
        if region is None:
            return None
        region_keys.add(
            (
                region.min_row,
                region.max_row,
                region.min_column,
                region.max_column,
            )
        )
    if len(region_keys) != 1:
        return None

    results: list[tuple[str, FormulaReference, tuple[str, ...], str, float]] = []
    row_extensions: set[tuple[int, ...]] = set()
    for role, reference in ranges:
        allowed_kinds = (
            NUMERIC_AGGREGATE_KINDS if role == "sum_range" else CONDITIONAL_CRITERIA_KINDS
        )
        extension_data = _contiguous_extension(
            context,
            origin_worksheet,
            origin_cell,
            reference,
            allowed_kinds=allowed_kinds,
        )
        if extension_data is None:
            return None
        extension, kind, ratio = extension_data
        rows = tuple(
            int(re.search(r"\d+$", coordinate).group())  # type: ignore[union-attr]
            for coordinate in extension
        )
        row_extensions.add(rows)
        results.append((role, reference, extension, kind, ratio))
    if len(row_extensions) != 1:
        return None
    return tuple(results), next(iter(row_extensions))


def _provable_zero_countif_denominator(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
    ir: FormulaIR,
) -> dict[str, Any] | None:
    denominator = _top_level_division_denominator(_formula_tokens(cell.value))
    parsed = _simple_function_arguments(denominator, frozenset({"COUNTIF"}))
    if parsed is None:
        return None
    _, arguments = parsed
    if len(arguments) != 2:
        return None
    reference = _reference_argument(ir, arguments[0])
    criterion = _numeric_countif_criterion(arguments[1])
    if (
        reference is None
        or criterion is None
        or not reference.is_single_column
        or reference.area > 4096
    ):
        return None
    target_worksheet = _worksheet_by_name(context, reference.sheet)
    if target_worksheet is None:
        return None
    values: list[Decimal] = []
    for row in range(reference.min_row, reference.max_row + 1):
        source = _cell_at(target_worksheet, row, reference.min_column)
        if (
            source is None
            or source.data_type == "f"
            or isinstance(source.value, bool)
            or not isinstance(source.value, (int, float, Decimal))
        ):
            return None
        value = Decimal(str(source.value))
        if not value.is_finite():
            return None
        values.append(value)
    operator, threshold = criterion
    if not values or any(_comparison_matches(value, operator, threshold) for value in values):
        return None
    return {
        "proof_level": "proven_static_values",
        "kind": "provable_zero_countif_denominator",
        "denominator_function": "COUNTIF",
        "reference": reference.raw,
        "comparison": f"{operator}{threshold}",
        "literal_value_count": len(values),
        "observed_min": str(min(values)),
        "observed_max": str(max(values)),
        "automatic_formula_change": False,
    }


def _blank_inside_dense_region(
    context: RuleContext,
    worksheet: Worksheet,
    reference: FormulaReference,
) -> bool:
    if not reference.is_single_cell:
        return False
    region = _region_for_cell(
        context,
        worksheet,
        reference.min_row,
        reference.min_column,
    )
    if region is None or reference.min_row <= region.min_row:
        return False
    above = _cell_at(worksheet, reference.min_row - 1, reference.min_column)
    below = _cell_at(worksheet, reference.min_row + 1, reference.min_column)
    populated_in_row = sum(
        _is_nonblank(candidate.value)
        for column in range(region.min_column, region.max_column + 1)
        if (candidate := _cell_at(worksheet, reference.min_row, column)) is not None
    )
    return (
        _is_nonblank(above.value if above is not None else None)
        and _is_nonblank(below.value if below is not None else None)
        and populated_in_row >= max(2, math.ceil((region.max_column - region.min_column + 1) * 0.5))
    )


def _looks_like_notes_header(value: Any) -> bool:
    normalized = _normalize_text(value)
    return bool(
        normalized
        and (
            _text_tokens(normalized) & NOTE_TOKENS
            or any(marker in normalized for marker in NOTE_MARKERS)
        )
    )


def _column_note_profile(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
) -> tuple[str, list[str], float] | None:
    region = _region_for_cell(context, worksheet, cell.row, cell.column)
    if region is None or cell.row <= region.min_row:
        return None
    header = _cell_at(worksheet, region.min_row, cell.column)
    if header is None or not _looks_like_notes_header(header.value):
        return None
    body = [
        candidate
        for row in range(region.min_row + 1, region.max_row + 1)
        if not _row_is_summary(context, worksheet, region, row)
        and (candidate := _cell_at(worksheet, row, cell.column)) is not None
        and _is_nonblank(candidate.value)
    ]
    text_peers = [
        candidate
        for candidate in body
        if candidate.data_type != "f" and isinstance(candidate.value, str)
    ]
    formula_count = sum(candidate.data_type == "f" for candidate in body)
    literal_count = sum(candidate.data_type != "f" for candidate in body)
    text_ratio = len(text_peers) / literal_count if literal_count else 0.0
    if (
        len(text_peers) < 3
        or text_ratio < 0.75
        or formula_count > max(1, math.floor(len(body) * 0.2))
    ):
        return None
    return str(header.value), [peer.coordinate for peer in text_peers[:12]], text_ratio


def _semantic_role(value: Any) -> Literal["rate", "amount"] | None:
    normalized = _normalize_text(value)
    if not normalized:
        return None
    words = _text_tokens(normalized)
    rate = bool(
        words & RATE_TOKENS
        or ("rate" in words and words & RATE_QUALIFIERS)
        or any(marker in normalized for marker in RATE_MARKERS)
    )
    if "exchange rate" in normalized or "汇率" in normalized:
        rate = False
    amount = bool(words & AMOUNT_TOKENS or any(marker in normalized for marker in AMOUNT_MARKERS))
    if rate == amount:
        return None
    return "rate" if rate else "amount"


def _aggregate_input_role(value: Any) -> Literal["rate", "amount", "quantity"] | None:
    semantic = _semantic_role(value)
    if semantic is not None:
        return semantic
    normalized = _normalize_text(value)
    if not normalized:
        return None
    words = _text_tokens(normalized)
    quantity = bool(
        words & QUANTITY_TOKENS or any(marker in normalized for marker in QUANTITY_MARKERS)
    )
    return "quantity" if quantity else None


def _heterogeneous_aggregate_column_roles(
    context: RuleContext,
    worksheet: Worksheet,
    reference: FormulaReference,
) -> list[dict[str, Any]] | None:
    """Return fully known, incompatible source-column roles for one aggregate range."""

    if reference.is_single_column:
        return None
    region = _region_for_reference(context, worksheet, reference)
    if region is None:
        return None
    columns: list[dict[str, Any]] = []
    for column in range(reference.min_column, reference.max_column + 1):
        header = _cell_at(worksheet, region.min_row, column)
        if header is None or not isinstance(header.value, str):
            return None
        role = _aggregate_input_role(header.value)
        if role is None:
            return None
        columns.append({"column": get_column_letter(column), "header": header.value, "role": role})
    return columns if len({item["role"] for item in columns}) >= 2 else None


def _format_role(number_format: str) -> Literal["percentage", "currency"] | None:
    normalized = number_format.casefold()
    percentage = "%" in number_format
    currency = any(marker in normalized for marker in CURRENCY_MARKERS)
    if percentage == currency:
        return None
    return "percentage" if percentage else "currency"


def _result_label(
    context: RuleContext,
    worksheet: Worksheet,
    cell: Cell,
) -> tuple[str, Literal["rate", "amount"]] | None:
    candidates: list[Any] = []
    for column in range(cell.column - 1, max(0, cell.column - 4), -1):
        left = _cell_at(worksheet, cell.row, column)
        if left is not None and isinstance(left.value, str) and left.data_type != "f":
            candidates.append(left.value)
    region = _region_for_cell(context, worksheet, cell.row, cell.column)
    if region is not None and cell.row > region.min_row:
        header = _cell_at(worksheet, region.min_row, cell.column)
        if header is not None and isinstance(header.value, str):
            candidates.append(header.value)
    for candidate in candidates:
        role = _semantic_role(candidate)
        if role is not None:
            return str(candidate), role
    return None


class AggregateRangeCoverageRule(WorkbookRule):
    rule_id = "WL042_AGGREGATE_RANGE_COVERAGE"
    title = "Aggregate range excludes contiguous table rows"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            formula_cells = _formula_cells(context, worksheet)
            formulas_by_column_lists: dict[int, list[Cell]] = {}
            for formula_cell in formula_cells:
                formulas_by_column_lists.setdefault(formula_cell.column, []).append(formula_cell)
            formulas_by_column: dict[int, tuple[Cell, ...]] = {
                column: tuple(cells) for column, cells in formulas_by_column_lists.items()
            }
            proven_by_column: dict[int, dict[str, Cell]] = {}
            for cell in formula_cells:
                ir = _formula_ir(context, worksheet, cell)
                for aggregate in ir.aggregate_references:
                    has_scope_label = _row_has_summary_label(context, worksheet, cell.row)
                    if aggregate.reference.is_cross_sheet:
                        has_scope_label = has_scope_label or _row_has_unscoped_metric_label(
                            context, worksheet, cell.row
                        )
                    if not has_scope_label:
                        continue
                    extension_data = _contiguous_extension(
                        context, worksheet, cell, aggregate.reference
                    )
                    if extension_data is None:
                        continue
                    extension, kind, ratio = extension_data
                    expected_reference = _expanded_reference(aggregate.reference, extension)
                    patches: list[PatchOperation] = []
                    label_source = _aggregate_label_source(
                        context,
                        worksheet,
                        cell,
                        aggregate.function,
                        cross_sheet=aggregate.reference.is_cross_sheet,
                    )
                    complete_body = _complete_reference_body(
                        context,
                        worksheet,
                        cell,
                        aggregate.reference,
                        extension,
                        allowed_kinds=NUMERIC_AGGREGATE_KINDS,
                    )
                    formula_repair_patch_ids: tuple[str, ...] = ()
                    if kind == "formula":
                        complete_body = _complete_formula_reference_body(
                            context,
                            worksheet,
                            cell,
                            aggregate.reference,
                            extension,
                        )
                        if complete_body is not None:
                            source_region, bounds, formula_repair_patch_ids = complete_body
                            complete_body = (source_region, bounds, ())
                    candidate_formula = (
                        _replace_formula_references(
                            str(cell.value),
                            {aggregate.reference.raw: expected_reference},
                        )
                        if aggregate.function in {"SUM", "AVERAGE"}
                        and label_source is not None
                        and complete_body is not None
                        else None
                    )
                    if (
                        candidate_formula is not None
                        and complete_body is not None
                        and label_source is not None
                        and _formula_patch_target_is_safe(
                            context,
                            worksheet,
                            cell,
                            candidate_formula,
                        )
                    ):
                        source_region, _bounds, normalization_patch_ids = complete_body
                        prerequisite_patch_ids = tuple(
                            sorted(
                                {
                                    *normalization_patch_ids,
                                    *formula_repair_patch_ids,
                                }
                            )
                        )
                        patch = _make_formula_patch(
                            worksheet=worksheet,
                            cell=cell,
                            formula=candidate_formula,
                            derivation=_range_derivation(
                                strategy="unique_complete_aggregate_range",
                                label_source=label_source,
                                source_sheet=aggregate.reference.sheet,
                                region=source_region,
                                synchronized_ranges=False,
                                prerequisite_patch_ids=normalization_patch_ids,
                                formula_repair_patch_ids=formula_repair_patch_ids,
                            ),
                            description=(
                                "Extend the aggregate to the unique complete inferred table body."
                            ),
                            prerequisite_patch_ids=prerequisite_patch_ids,
                        )
                        patches.append(patch)
                        _append_patch_with_prerequisites(result, context, patch)
                    evidence = Evidence(
                        summary="Aggregate range omits adjacent same-kind rows in one inferred table column",
                        observed=cell.value,
                        expected={"reviewed_reference": expected_reference},
                        peers=list(extension[:12]),
                        details={
                            "proof_level": "strong_structural",
                            "function": aggregate.function,
                            "source_sheet": aggregate.reference.sheet,
                            "observed_reference": aggregate.reference.raw,
                            "excluded_cells": list(extension),
                            "dominant_input_kind": kind,
                            "dominant_kind_ratio": round(ratio, 4),
                            "automatic_formula_change": bool(patches),
                            "candidate_formula": candidate_formula,
                            "formula_repair_prerequisites": list(formula_repair_patch_ids),
                        },
                    )
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=(
                                "A summary-labelled aggregate stops before or starts after adjacent "
                                "same-kind rows in the same inferred table column."
                            ),
                            severity=Severity.WARNING,
                            confidence=0.94 if len(extension) >= 2 else 0.9,
                            worksheet=worksheet,
                            location=cell.coordinate,
                            evidence=evidence,
                            expected=(
                                "Review the aggregate boundary against the complete table body; business "
                                "formulas are not changed automatically."
                            ),
                            suggested_action=(
                                f"Confirm whether {aggregate.reference.raw} should cover "
                                f"{expected_reference}."
                            ),
                            discriminator=(aggregate.reference.raw, extension),
                            patches=patches,
                        )
                    )
                    proven_by_column.setdefault(cell.column, {})[cell.coordinate] = cell
                conditional = _conditional_aggregate_ranges(ir)
                if conditional is None or not _row_has_unscoped_metric_label(
                    context,
                    worksheet,
                    cell.row,
                ):
                    continue
                function, ranges = conditional
                conditional_extension = _conditional_aggregate_extensions(
                    context,
                    worksheet,
                    cell,
                    ranges,
                )
                if conditional_extension is None:
                    continue
                range_extensions, excluded_rows = conditional_extension
                reviewed_references = {
                    role: _expanded_reference(reference, extension)
                    for role, reference, extension, _, _ in range_extensions
                }
                range_details = [
                    {
                        "role": role,
                        "observed_reference": reference.raw,
                        "reviewed_reference": reviewed_references[role],
                        "excluded_cells": list(extension),
                        "dominant_input_kind": kind,
                        "dominant_kind_ratio": round(ratio, 4),
                    }
                    for role, reference, extension, kind, ratio in range_extensions
                ]
                patches = []
                label_source = _aggregate_label_source(
                    context,
                    worksheet,
                    cell,
                    function,
                    cross_sheet=any(reference.is_cross_sheet for _, reference in ranges),
                )
                complete_bodies = []
                replacements: dict[str, str] = {}
                replacement_conflict = False
                for role, reference, extension, _, _ in range_extensions:
                    allowed_kinds = (
                        NUMERIC_AGGREGATE_KINDS
                        if role == "sum_range"
                        else CONDITIONAL_CRITERIA_KINDS
                    )
                    complete = _complete_reference_body(
                        context,
                        worksheet,
                        cell,
                        reference,
                        extension,
                        allowed_kinds=allowed_kinds,
                    )
                    complete_bodies.append(complete)
                    reviewed = reviewed_references[role]
                    existing = replacements.get(reference.raw)
                    if existing is not None and existing != reviewed:
                        replacement_conflict = True
                    replacements[reference.raw] = reviewed
                complete_regions = [
                    complete[0] for complete in complete_bodies if complete is not None
                ]
                complete_bounds = {
                    complete[1] for complete in complete_bodies if complete is not None
                }
                prerequisite_patch_ids = tuple(
                    sorted(
                        {
                            patch_id
                            for complete in complete_bodies
                            if complete is not None
                            for patch_id in complete[2]
                        }
                    )
                )
                candidate_formula = (
                    _replace_formula_references(str(cell.value), replacements)
                    if label_source is not None
                    and not replacement_conflict
                    and len(complete_regions) == len(complete_bodies)
                    and len(complete_bounds) == 1
                    else None
                )
                if (
                    candidate_formula is not None
                    and complete_regions
                    and label_source is not None
                    and _formula_patch_target_is_safe(
                        context,
                        worksheet,
                        cell,
                        candidate_formula,
                    )
                ):
                    patch = _make_formula_patch(
                        worksheet=worksheet,
                        cell=cell,
                        formula=candidate_formula,
                        derivation=_range_derivation(
                            strategy="unique_synchronized_conditional_aggregate_ranges",
                            label_source=label_source,
                            source_sheet=ranges[0][1].sheet,
                            region=complete_regions[0],
                            synchronized_ranges=True,
                            prerequisite_patch_ids=prerequisite_patch_ids,
                        ),
                        description=(
                            "Synchronize every conditional-aggregate range to the unique complete "
                            "inferred table body."
                        ),
                        prerequisite_patch_ids=prerequisite_patch_ids,
                    )
                    patches.append(patch)
                    _append_patch_with_prerequisites(result, context, patch)
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation=(
                            "A summary-labelled conditional aggregate uses aligned criteria and "
                            "sum ranges that all stop before or start after the same contiguous "
                            "source-table rows."
                        ),
                        severity=Severity.WARNING,
                        confidence=0.95 if len(excluded_rows) >= 2 else 0.91,
                        worksheet=worksheet,
                        location=cell.coordinate,
                        evidence=Evidence(
                            summary=(
                                "Conditional aggregate ranges synchronously omit contiguous "
                                "source-table rows"
                            ),
                            observed=cell.value,
                            expected={"reviewed_references": reviewed_references},
                            peers=[f"row {row}" for row in excluded_rows[:12]],
                            details={
                                "proof_level": "strong_structural",
                                "function": function,
                                "source_sheet": ranges[0][1].sheet,
                                "ranges": range_details,
                                "excluded_rows": list(excluded_rows),
                                "automatic_formula_change": bool(patches),
                                "candidate_formula": candidate_formula,
                            },
                        ),
                        expected=(
                            "Review every conditional-aggregate range against the complete "
                            "source table body; formulas are not changed automatically."
                        ),
                        suggested_action=(
                            "Confirm that the criteria and sum ranges should share the reviewed "
                            "source-table boundary."
                        ),
                        discriminator=(
                            function,
                            tuple((role, reference.raw) for role, reference in ranges),
                            excluded_rows,
                        ),
                        patches=patches,
                    )
                )
                proven_by_column.setdefault(cell.column, {})[cell.coordinate] = cell
            if worksheet.sheet_state != "visible":
                continue
            for column, proven_by_coordinate in sorted(proven_by_column.items()):
                proven_cells = sorted(
                    proven_by_coordinate.values(),
                    key=lambda item: (item.row, item.column),
                )
                if len(proven_cells) < 3:
                    continue
                rows = sorted(cell.row for cell in proven_cells)
                if rows[-1] - rows[0] > 5 or rows[-1] - rows[0] + 1 > len(proven_cells) + 2:
                    continue
                proven_coordinates = {cell.coordinate for cell in proven_cells}
                unproven_cells = []
                for candidate in formulas_by_column.get(column, ()):
                    if candidate.coordinate in proven_coordinates:
                        continue
                    if not rows[0] <= candidate.row <= rows[-1]:
                        continue
                    candidate_ir = _formula_ir(context, worksheet, candidate)
                    if not candidate_ir.aggregate_references:
                        continue
                    if not _row_has_summary_label(context, worksheet, candidate.row) and not (
                        any(
                            aggregate.reference.is_cross_sheet
                            for aggregate in candidate_ir.aggregate_references
                        )
                        and _row_has_unscoped_metric_label(context, worksheet, candidate.row)
                    ):
                        continue
                    unproven_cells.append(candidate)
                if not unproven_cells:
                    continue
                location = (
                    f"{get_column_letter(column)}{rows[0]}:{get_column_letter(column)}{rows[-1]}"
                )
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation=(
                            "Several aggregate formulas in one KPI block have independently proven "
                            "truncated ranges, while another formula in the same block lacks enough "
                            "evidence to classify it."
                        ),
                        severity=Severity.INFO,
                        confidence=0.82,
                        worksheet=worksheet,
                        location=location,
                        evidence=Evidence(
                            summary="KPI block contains multiple proven range omissions and one unproven formula",
                            observed={
                                "proven_cells": [cell.coordinate for cell in proven_cells],
                                "unproven_cells": [cell.coordinate for cell in unproven_cells],
                            },
                            expected={
                                "review_proven_boundaries": True,
                                "unproven_cells_are_not_classified": True,
                            },
                            peers=[cell.coordinate for cell in proven_cells],
                            details={
                                "proof_level": "advisory",
                                "automatic_formula_change": False,
                                "independent_extension_count": len(proven_cells),
                                "unproven_formulas": {
                                    cell.coordinate: cell.value for cell in unproven_cells
                                },
                            },
                        ),
                        expected=(
                            "Review each proven aggregate boundary separately; formulas without "
                            "independent boundary evidence are not classified as errors."
                        ),
                        suggested_action=(
                            "Review the listed KPI formulas against their source table bodies; do not "
                            "copy a neighboring range automatically."
                        ),
                        discriminator=("block", tuple(cell.coordinate for cell in proven_cells)),
                    )
                )
        return result


class UnlabeledAggregateRoleRule(WorkbookRule):
    rule_id = "WL052_UNLABELED_AGGREGATE_ROLE"
    title = "Aggregate lacks a clear summary label"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            if worksheet.sheet_state != "visible":
                continue
            aggregates_by_row: dict[int, list[tuple[Cell, Any]]] = {}
            for cell in _formula_cells(context, worksheet):
                cell_aggregates = _formula_ir(context, worksheet, cell).aggregate_references
                if len(cell_aggregates) == 1:
                    aggregates_by_row.setdefault(cell.row, []).append((cell, cell_aggregates[0]))
            for row, row_aggregates in aggregates_by_row.items():
                labels = _summary_label_cells(context, worksheet, row)
                if not labels:
                    continue
                first_label = min(labels, key=lambda cell: cell.column)
                right = [item for item in row_aggregates if item[0].column > first_label.column]
                if len(right) < 2:
                    continue
                dominant_function, dominant_count = Counter(
                    aggregate.function for _, aggregate in right
                ).most_common(1)[0]
                if dominant_count < 2 or dominant_count / len(right) < 0.75:
                    continue
                for cell, aggregate in row_aggregates:
                    if cell.column >= first_label.column or aggregate.function == dominant_function:
                        continue
                    if _nearby_explicit_aggregate_label(worksheet, cell, aggregate.function):
                        continue
                    reference = aggregate.reference
                    if reference.is_cross_sheet or not reference.is_single_column:
                        continue
                    target_worksheet = _worksheet_by_name(context, reference.sheet)
                    if target_worksheet is None:
                        continue
                    region = _region_for_reference(context, target_worksheet, reference)
                    if region is None or cell.row <= region.max_row:
                        continue
                    evidence = Evidence(
                        summary=(
                            "Aggregate uses a different function to the left of the row's summary label"
                        ),
                        observed={
                            "formula": cell.value,
                            "function": aggregate.function,
                            "reference": reference.raw,
                        },
                        expected={
                            "explicit_metric_label": True,
                            "labelled_summary_function": dominant_function,
                        },
                        peers=[peer.coordinate for peer, _ in right],
                        details={
                            "proof_level": "semantic_heuristic",
                            "summary_label": first_label.value,
                            "summary_label_cell": first_label.coordinate,
                            "dominant_labelled_function": dominant_function,
                            "labelled_aggregate_count": dominant_count,
                            "automatic_formula_change": False,
                        },
                    )
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=(
                                "An aggregate sits outside the labelled summary block and uses a different "
                                "function from the aggregates clearly governed by that label."
                            ),
                            severity=Severity.WARNING,
                            confidence=0.9,
                            worksheet=worksheet,
                            location=cell.coordinate,
                            evidence=evidence,
                            expected=(
                                "Distinct aggregate metrics have an explicit nearby label or a clear summary "
                                "block association."
                            ),
                            suggested_action=(
                                "Confirm the metric intent and add an explicit label or relocate it into the "
                                "labelled summary block; the formula is not changed automatically."
                            ),
                            discriminator=(
                                first_label.coordinate,
                                aggregate.function,
                                dominant_function,
                            ),
                        )
                    )
        return result


class CrossSheetReferenceValidityRule(WorkbookRule):
    rule_id = "WL043_CROSS_SHEET_REFERENCE_VALIDITY"
    title = "Cross-sheet reference target requires review"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for cell in _formula_cells(context, worksheet):
                ir = _formula_ir(context, worksheet, cell)
                aggregates_by_key = {
                    _reference_key(item.reference): item for item in ir.aggregate_references
                }
                for reference in ir.references:
                    if not reference.is_cross_sheet:
                        continue
                    target_worksheet = _worksheet_by_name(context, reference.sheet)
                    target = inspect_reference_target(
                        context.workbook,
                        reference,
                        content_index=(
                            worksheet_content_index(context, target_worksheet)
                            if target_worksheet is not None
                            else None
                        ),
                    )
                    reason: str | None = None
                    severity = Severity.WARNING
                    confidence = 0.9
                    details: dict[str, Any] = {
                        "proof_level": "strong_structural",
                        "source_sheet": reference.sheet,
                        "reference": reference.raw,
                        "automatic_formula_change": False,
                    }
                    peers: list[str] = []
                    expected: Any = {"review": "cross_sheet_target"}
                    aggregate = aggregates_by_key.get(_reference_key(reference))
                    role_columns = (
                        _heterogeneous_aggregate_column_roles(context, target_worksheet, reference)
                        if target_worksheet is not None
                        and aggregate is not None
                        and aggregate.function in {"AVERAGE", "SUM"}
                        else None
                    )
                    if not target.sheet_exists:
                        reason = "missing_worksheet"
                        severity = Severity.ERROR
                        confidence = 1.0
                    elif target.fully_blank and target.outside_content:
                        reason = "blank_outside_content"
                        severity = Severity.INFO
                        confidence = 0.7
                        details["proof_level"] = "advisory"
                        details["content_bounds"] = target.content_bounds
                    elif target.fully_blank:
                        if target_worksheet is not None and _blank_inside_dense_region(
                            context, target_worksheet, reference
                        ):
                            reason = "blank_inside_dense_table"
                            confidence = 0.94
                    elif role_columns is not None:
                        reason = "heterogeneous_aggregate_column_roles"
                        confidence = 0.97
                        details.update(
                            {
                                "aggregate_function": aggregate.function if aggregate else None,
                                "source_columns": role_columns,
                                "source_roles": sorted(
                                    {str(item["role"]) for item in role_columns}
                                ),
                            }
                        )
                        expected = {"compatible_source_column_roles": True}
                    elif _reference_key(
                        reference
                    ) not in aggregates_by_key and _row_has_summary_label(
                        context, worksheet, cell.row
                    ):
                        extension_data = _contiguous_extension(context, worksheet, cell, reference)
                        if extension_data is not None:
                            extension, kind, ratio = extension_data
                            reason = "truncated_cross_sheet_range"
                            confidence = 0.92 if len(extension) >= 2 else 0.88
                            peers = list(extension[:12])
                            expected = {
                                "reviewed_reference": _expanded_reference(reference, extension)
                            }
                            details.update(
                                {
                                    "excluded_cells": list(extension),
                                    "dominant_input_kind": kind,
                                    "dominant_kind_ratio": round(ratio, 4),
                                }
                            )
                    if reason is None:
                        continue
                    details["reason"] = reason
                    evidence = Evidence(
                        summary="Cross-sheet dependency target requires review",
                        observed=cell.value,
                        expected=expected,
                        peers=peers,
                        details=details,
                    )
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=(
                                "The referenced cells are currently blank outside the source sheet's "
                                "populated content and may be planned future input, so this advisory "
                                "does not prove the formula is erroneous."
                                if reason == "blank_outside_content"
                                else "A cross-sheet SUM or AVERAGE combines source columns whose explicit "
                                "headers imply incompatible amount, rate, or quantity roles."
                                if reason == "heterogeneous_aggregate_column_roles"
                                else "A cross-sheet dependency points to a missing, structurally blank, "
                                "out-of-content, or visibly truncated source target."
                            ),
                            severity=severity,
                            confidence=confidence,
                            worksheet=worksheet,
                            location=cell.coordinate,
                            evidence=evidence,
                            expected="Cross-sheet references resolve to the intended populated source region.",
                            suggested_action=(
                                "The referenced cells are currently blank outside the sheet's populated "
                                "content and may be reserved for future input; confirm the template intent "
                                "before editing the formula."
                                if reason == "blank_outside_content"
                                else "Confirm that every aggregated source column uses the same business unit; "
                                "WorkbookLens does not guess a replacement formula."
                                if reason == "heterogeneous_aggregate_column_roles"
                                else "Inspect the source worksheet and intended table boundary before editing the "
                                "formula; WorkbookLens does not guess replacement references."
                            ),
                            discriminator=(reference.raw, reason),
                        )
                    )
        return result


class FormulaInNotesColumnRule(WorkbookRule):
    rule_id = "WL044_FORMULA_IN_NOTES_COLUMN"
    title = "Formula appears in a notes or instructions column"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            if worksheet.sheet_state != "visible":
                continue
            for cell in _formula_cells(context, worksheet):
                profile = _column_note_profile(context, worksheet, cell)
                if profile is None:
                    continue
                header, peers, text_ratio = profile
                evidence = Evidence(
                    summary="A text-dominant notes column contains an isolated formula",
                    observed=cell.value,
                    expected={"column_role": "notes_or_instructions"},
                    peers=peers,
                    details={
                        "proof_level": "semantic_heuristic",
                        "header": header,
                        "text_peer_ratio": round(text_ratio, 4),
                        "automatic_formula_change": False,
                    },
                )
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation=(
                            "The inferred column role is free-text notes or instructions, while this "
                            "isolated cell contains an executable formula."
                        ),
                        severity=Severity.WARNING,
                        confidence=0.96 if text_ratio >= 0.9 else 0.92,
                        worksheet=worksheet,
                        location=cell.coordinate,
                        evidence=evidence,
                        expected="Notes and instruction columns normally contain literal text.",
                        suggested_action=(
                            "Confirm whether the formula was pasted into the wrong column or should be "
                            "stored as literal text; no automatic edit is proposed."
                        ),
                        discriminator=header,
                    )
                )
        return result


class DegenerateFormulaRule(WorkbookRule):
    rule_id = "WL045_DEGENERATE_FORMULA"
    title = "Formula collapses to an algebraic identity or constant"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        explanations = {
            "self_division": "The complete expression divides an expression by itself.",
            "self_subtraction": "The complete expression subtracts an expression from itself.",
            "multiply_by_zero": "The complete expression multiplies one side by the literal zero.",
        }
        for worksheet in context.workbook.worksheets:
            for cell in _formula_cells(context, worksheet):
                ir = _formula_ir(context, worksheet, cell)
                zero_denominator = _provable_zero_countif_denominator(
                    context,
                    worksheet,
                    cell,
                    ir,
                )
                if zero_denominator is not None:
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=(
                                "Every known numeric literal in the COUNTIF range fails the "
                                "comparison, so the denominator is statically zero."
                            ),
                            severity=Severity.ERROR,
                            confidence=1.0,
                            worksheet=worksheet,
                            location=cell.coordinate,
                            evidence=Evidence(
                                summary=(
                                    "COUNTIF denominator is proven zero from literal source values"
                                ),
                                observed=cell.value,
                                expected={"nonzero_denominator": True},
                                details=zero_denominator,
                            ),
                            expected=(
                                "The formula denominator is nonzero for the current literal inputs."
                            ),
                            suggested_action=(
                                "Review the COUNTIF condition and intended KPI denominator; "
                                "WorkbookLens does not infer a replacement formula."
                            ),
                            discriminator=(
                                zero_denominator["kind"],
                                zero_denominator["reference"],
                                zero_denominator["comparison"],
                            ),
                        )
                    )
                degeneracy = ir.degeneracy
                if degeneracy is not None:
                    details: dict[str, Any] = {
                        "proof_level": "proven_static_structure",
                        "kind": degeneracy.kind,
                        "operator": degeneracy.operator,
                        "automatic_formula_change": False,
                    }
                    if degeneracy.kind == "self_division":
                        details["zero_input_risk"] = (
                            "The expression is #DIV/0! when the common value is zero."
                        )
                    evidence = Evidence(
                        summary="Top-level formula structure is algebraically degenerate",
                        observed=cell.value,
                        expected={"review": "business_formula_logic"},
                        details=details,
                    )
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=explanations[degeneracy.kind],
                            severity=Severity.WARNING,
                            confidence=0.99,
                            worksheet=worksheet,
                            location=cell.coordinate,
                            evidence=evidence,
                            expected=(
                                "The formula expresses the intended non-degenerate business "
                                "calculation."
                            ),
                            suggested_action=(
                                "Review the operands and intended KPI logic; the structural fact "
                                "is proven, but WorkbookLens does not infer the replacement formula."
                            ),
                            discriminator=degeneracy.kind,
                        )
                    )
        return result


class FormulaFormatRoleMismatchRule(WorkbookRule):
    rule_id = "WL046_FORMULA_FORMAT_ROLE_MISMATCH"
    title = "Formula result role conflicts with its number format"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            if worksheet.sheet_state != "visible":
                continue
            for cell in _formula_cells(context, worksheet):
                label_data = _result_label(context, worksheet, cell)
                format_role = _format_role(cell.number_format)
                if label_data is None or format_role is None:
                    continue
                label, semantic_role = label_data
                mismatch = (semantic_role == "rate" and format_role == "currency") or (
                    semantic_role == "amount" and format_role == "percentage"
                )
                if not mismatch:
                    continue
                evidence = Evidence(
                    summary="A clear result label and the cell number format imply incompatible roles",
                    observed={
                        "formula": cell.value,
                        "number_format": cell.number_format,
                        "label": label,
                    },
                    expected={
                        "semantic_role": semantic_role,
                        "compatible_format": "percentage"
                        if semantic_role == "rate"
                        else "currency",
                    },
                    details={
                        "proof_level": "semantic_heuristic",
                        "observed_format_role": format_role,
                        "automatic_formula_change": False,
                        "automatic_format_change": False,
                    },
                )
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation=(
                            "The label strongly describes a rate or monetary amount, but the formula "
                            "cell uses the opposite percentage/currency display role."
                        ),
                        severity=Severity.WARNING,
                        confidence=0.96,
                        worksheet=worksheet,
                        location=cell.coordinate,
                        evidence=evidence,
                        expected="Formula results use a number format compatible with the labelled role.",
                        suggested_action=(
                            "Confirm the KPI meaning and then review the number format manually; the "
                            "formula and format are not changed automatically."
                        ),
                        discriminator=(label, semantic_role, format_role),
                    )
                )
        return result


FORMULA_SEMANTIC_RULES: tuple[type[WorkbookRule], ...] = (
    AggregateRangeCoverageRule,
    UnlabeledAggregateRoleRule,
    CrossSheetReferenceValidityRule,
    FormulaInNotesColumnRule,
    DegenerateFormulaRule,
    FormulaFormatRoleMismatchRule,
)


__all__ = [
    "FORMULA_SEMANTIC_RULES",
    "AggregateRangeCoverageRule",
    "CrossSheetReferenceValidityRule",
    "DegenerateFormulaRule",
    "FormulaFormatRoleMismatchRule",
    "FormulaInNotesColumnRule",
    "UnlabeledAggregateRoleRule",
]
