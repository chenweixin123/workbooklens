"""Conservative, report-only formula semantics rules."""

from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from datetime import date, datetime
from decimal import Decimal
from typing import Any, Literal, cast

from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import get_column_letter, quote_sheetname
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.formulas.ir import (
    FormulaIR,
    FormulaReference,
    inspect_reference_target,
    parse_formula_ir,
    worksheet_content_index,
)
from workbooklens.models import Confidence, Evidence, Finding, Region, Severity
from workbooklens.rules.base import RuleContext, RuleResult, WorkbookRule
from workbooklens.utils import stable_id

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
    "cost",
    "gross",
    "income",
    "inventory",
    "profit",
    "revenue",
    "salary",
    "sales",
    "turnover",
    "value",
}
AGGREGATE_METRIC_MARKERS = (
    "金额",
    "成本",
    "收入",
    "营收",
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
        safe_patch_available=False,
        patch_ids=[],
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


def _value_kind(cell: Cell | None) -> str | None:
    if cell is None or not _is_nonblank(cell.value):
        return None
    if cell.data_type == "f":
        return "formula"
    value = cell.value
    if isinstance(value, bool):
        return "logical"
    if isinstance(value, (int, float, Decimal)):
        return "number"
    if isinstance(value, (date, datetime)):
        return "date"
    if isinstance(value, str):
        return "text"
    return "other"


def _dominant_reference_kind(
    context: RuleContext,
    worksheet: Worksheet,
    reference: FormulaReference,
) -> tuple[str, float] | None:
    kinds = [
        kind
        for cell in worksheet_content_index(context, worksheet).iter_nonblank_rectangle(
            min_row=reference.min_row,
            max_row=reference.max_row,
            min_column=reference.min_column,
            max_column=reference.max_column,
        )
        if (kind := _value_kind(cell)) is not None
    ]
    if len(kinds) < 3:
        return None
    kind, count = Counter(kinds).most_common(1)[0]
    ratio = count / len(kinds)
    if kind not in {"formula", "number"} or ratio < 0.75:
        return None
    return kind, ratio


def _contiguous_extension(
    context: RuleContext,
    origin_worksheet: Worksheet,
    origin_cell: Cell,
    reference: FormulaReference,
) -> tuple[tuple[str, ...], str, float] | None:
    if not reference.is_single_column:
        return None
    target_worksheet = _worksheet_by_name(context, reference.sheet)
    if target_worksheet is None:
        return None
    region = _region_for_reference(context, target_worksheet, reference)
    if region is None:
        return None
    dominant = _dominant_reference_kind(context, target_worksheet, reference)
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
        return cell if _value_kind(cell) == kind else None

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
    column = get_column_letter(reference.min_column)
    local = f"{column}{min(rows)}:{column}{max(rows)}"
    return f"{quote_sheetname(reference.sheet)}!{local}" if reference.explicit_sheet else local


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
                            "automatic_formula_change": False,
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
                degeneracy = _formula_ir(context, worksheet, cell).degeneracy
                if degeneracy is None:
                    continue
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
                        expected="The formula expresses the intended non-degenerate business calculation.",
                        suggested_action=(
                            "Review the operands and intended KPI logic; the structural fact is proven, "
                            "but WorkbookLens does not infer the replacement formula."
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
