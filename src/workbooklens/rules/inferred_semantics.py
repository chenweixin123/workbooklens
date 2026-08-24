"""Conservative report-only inference for categorical and completeness anomalies."""

from __future__ import annotations

import math
from bisect import bisect_left, bisect_right
from collections import Counter
from datetime import date, datetime
from typing import Any

from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from workbooklens.models import Evidence, Severity
from workbooklens.rules.base import RuleContext, RuleResult, WorkbookRule
from workbooklens.rules.data_quality import (
    _cell_at,
    _is_error,
    _is_formula,
    _is_nonblank,
    _looks_like_identifier_header,
    _make_finding,
    _normalize_text,
    _table_profiles,
    _TableProfile,
    _tokens,
    _value_at,
)
from workbooklens.rules.profile_quality import (
    ColumnProfile,
    _context_profile,
    _resolved_tables,
    _ResolvedTable,
)

MIN_CATEGORY_ROWS = 15
MIN_COMPLETENESS_ROWS = 20
MAX_CATEGORY_CARDINALITY = 12
MAX_RARE_SHARE = 0.08
MIN_COMPLETE_COVERAGE = 0.95

FREE_TEXT_TOKENS = {
    "address",
    "comment",
    "comments",
    "description",
    "detail",
    "details",
    "instruction",
    "instructions",
    "link",
    "memo",
    "message",
    "note",
    "notes",
    "remark",
    "remarks",
    "url",
}
FREE_TEXT_MARKERS = (
    "备注",
    "说明",
    "描述",
    "详情",
    "地址",
    "评论",
    "留言",
    "网址",
    "链接",
)
PLACEHOLDER_WORDS = {
    "invalid",
    "notset",
    "tbd",
    "unassigned",
    "undefined",
    "unknown",
    "unset",
}
IGNORED_CATEGORY_MARKERS = {"-", "n/a", "n.a.", "na", "none", "null", "—", "无", "暂无"}


def _looks_like_free_text_header(value: Any) -> bool:
    normalized = _normalize_text(value)
    return bool(
        normalized
        and (
            _tokens(normalized) & FREE_TEXT_TOKENS
            or any(marker in normalized for marker in FREE_TEXT_MARKERS)
        )
    )


def _category_value(value: Any) -> str | None:
    normalized = _normalize_text(value)
    letters = [character for character in normalized if not character.isspace()]
    if (
        len(normalized) < 2
        or len(normalized) > 32
        or len(letters) < 2
        or not all(character.isalpha() for character in letters)
    ):
        return None
    return normalized


def _one_edit_apart(left: str, right: str) -> bool:
    """Return whether two normalized strings differ by one edit or adjacent transpose."""

    if left == right or abs(len(left) - len(right)) > 1:
        return False
    if len(left) == len(right):
        differences = [
            index for index, pair in enumerate(zip(left, right, strict=True)) if pair[0] != pair[1]
        ]
        if len(differences) == 1:
            return True
        return bool(
            len(differences) == 2
            and differences[1] == differences[0] + 1
            and left[differences[0]] == right[differences[1]]
            and left[differences[1]] == right[differences[0]]
        )
    shorter, longer = (left, right) if len(left) < len(right) else (right, left)
    short_index = long_index = mismatch_count = 0
    while short_index < len(shorter) and long_index < len(longer):
        if shorter[short_index] == longer[long_index]:
            short_index += 1
            long_index += 1
            continue
        mismatch_count += 1
        if mismatch_count > 1:
            return False
        long_index += 1
    return True


def _conservative_typo_candidate(value: str, peer: str) -> bool:
    """Return whether a rare value is a strong ASCII typo candidate for ``peer``."""

    return bool(
        value
        and peer
        and value.isascii()
        and peer.isascii()
        and value[0] == peer[0]
        and value[-1] == peer[-1]
        and _one_edit_apart(value, peer)
    )


def _compact_letters(value: str) -> str:
    return "".join(character for character in value if "a" <= character <= "z")


def _looks_like_placeholder(value: str, header: str) -> bool:
    compact = _compact_letters(value)
    if compact in PLACEHOLDER_WORDS:
        return True
    header_parts = {_compact_letters(token) for token in _tokens(header)}
    for prefix in ("invalid", "unknown", "undefined", "unset"):
        if compact.startswith(prefix) and compact[len(prefix) :] in header_parts:
            return True
    return False


def _value_family(value: Any) -> str | None:
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (date, datetime)):
        return "date"
    if isinstance(value, (int, float)):
        return "number"
    if isinstance(value, str) and value.strip():
        return "text"
    return None


def _profiles_overlap(profile: _TableProfile, table: _ResolvedTable) -> bool:
    return not (
        profile.region.max_row < table.min_row
        or table.max_row < profile.region.min_row
        or profile.region.max_column < table.min_column
        or table.max_column < profile.region.min_column
    )


def _explicit_column_contracts(
    context: RuleContext,
    worksheet: Worksheet,
    profile: _TableProfile,
) -> dict[int, ColumnProfile]:
    workbook_profile = _context_profile(context)
    contracts: dict[int, ColumnProfile] = {}
    for table in _resolved_tables(context, workbook_profile):
        if table.explicit and table.worksheet is worksheet and _profiles_overlap(profile, table):
            contracts.update(table.columns)
    return contracts


def _merged_row_index(
    worksheet: Worksheet,
    candidate_rows: tuple[int, ...],
) -> dict[int, tuple[tuple[int, int], ...]]:
    """Index merged-column intervals only for rows WL054 will inspect."""

    rows = tuple(sorted(set(candidate_rows)))
    if not rows:
        return {}
    intervals_by_row: dict[int, list[tuple[int, int]]] = {}
    for merged in worksheet.merged_cells.ranges:
        start = bisect_left(rows, merged.min_row)
        stop = bisect_right(rows, merged.max_row)
        for index in range(start, stop):
            intervals_by_row.setdefault(rows[index], []).append((merged.min_col, merged.max_col))
    return {row: tuple(sorted(intervals)) for row, intervals in intervals_by_row.items()}


def _indexed_merged_cell(
    merged_rows: dict[int, tuple[tuple[int, int], ...]], row: int, column: int
) -> bool:
    intervals = merged_rows.get(row)
    if not intervals:
        return False
    index = bisect_right(intervals, (column, 1 << 30)) - 1
    return index >= 0 and intervals[index][1] >= column


class InferredCategoryVariantRule(WorkbookRule):
    """Report rare typo-like or explicit placeholder values in inferred categories."""

    rule_id = "WL053_INFERRED_CATEGORY_VARIANT"
    title = "Rare variant in an inferred categorical column"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            for profile in _table_profiles(context, worksheet):
                column_contracts = _explicit_column_contracts(context, worksheet, profile)
                for column, header in profile.headers.items():
                    if _looks_like_identifier_header(header) or _looks_like_free_text_header(
                        header
                    ):
                        continue
                    configured = column_contracts.get(column)
                    if configured is not None and configured.allowed_values:
                        continue
                    cells_by_value: dict[str, list[Cell]] = {}
                    unsupported = 0
                    for row in profile.body_rows:
                        cell = _cell_at(worksheet, row, column)
                        if not isinstance(cell, Cell) or not _is_nonblank(cell.value):
                            continue
                        if _is_formula(cell) or _is_error(cell):
                            unsupported += 1
                            continue
                        normalized = _category_value(cell.value)
                        if normalized is None:
                            if _normalize_text(cell.value) in IGNORED_CATEGORY_MARKERS:
                                continue
                            unsupported += 1
                            continue
                        cells_by_value.setdefault(normalized, []).append(cell)
                    populated = sum(len(cells) for cells in cells_by_value.values())
                    if populated < MIN_CATEGORY_ROWS or unsupported > max(1, populated // 20):
                        continue
                    distinct = len(cells_by_value)
                    if (
                        distinct < 2
                        or distinct > MAX_CATEGORY_CARDINALITY
                        or distinct / populated > 0.35
                    ):
                        continue
                    common_threshold = max(3, math.ceil(populated * 0.10))
                    frequent = {
                        value: cells
                        for value, cells in cells_by_value.items()
                        if len(cells) >= common_threshold
                    }
                    if len(frequent) < 2:
                        continue
                    rare_limit = min(2, max(1, math.floor(populated * MAX_RARE_SHARE)))
                    for value, cells in sorted(cells_by_value.items()):
                        if len(cells) > rare_limit or len(cells) / populated > MAX_RARE_SHARE:
                            continue
                        matches = sorted(
                            peer for peer in frequent if _conservative_typo_candidate(value, peer)
                        )
                        placeholder = len(cells) == 1 and _looks_like_placeholder(value, header)
                        if not matches and not placeholder:
                            continue
                        if len(matches) == 1:
                            expected_value = matches[0]
                        elif matches:
                            expected_value = f"one of: {', '.join(matches)}"
                        else:
                            expected_value = "reviewed category value"
                        peer_coordinates = [
                            cell.coordinate for match in matches for cell in frequent[match][:3]
                        ][:6]
                        confidence = 0.94 if len(matches) == 1 else 0.90 if matches else 0.86
                        for cell in cells:
                            evidence = Evidence(
                                summary="A rare category value differs from a well-supported peer pattern",
                                observed=cell.value,
                                expected=expected_value,
                                peers=peer_coordinates,
                                details={
                                    "header": header,
                                    "normalized_value": value,
                                    "observed_occurrences": len(cells),
                                    "populated_rows": populated,
                                    "frequent_categories": {
                                        peer: len(peer_group)
                                        for peer, peer_group in sorted(frequent.items())
                                    },
                                    "reason": (
                                        "single_edit"
                                        if len(matches) == 1
                                        else "ambiguous_single_edit"
                                        if matches
                                        else "placeholder_marker"
                                    ),
                                },
                            )
                            result.findings.append(
                                _make_finding(
                                    context=context,
                                    rule_id=self.rule_id,
                                    title=self.title,
                                    explanation=(
                                        "A low-cardinality text column has several repeated categories, "
                                        "while this rare value is either one edit from exactly one frequent "
                                        "peer or uses an explicit placeholder marker."
                                    ),
                                    severity=Severity.WARNING,
                                    confidence=confidence,
                                    worksheet=worksheet,
                                    location=cell.coordinate,
                                    evidence=evidence,
                                    expected=(
                                        "Each category uses a consistent reviewed spelling without "
                                        "unresolved placeholder values."
                                    ),
                                    suggested_action=(
                                        "Confirm the intended category against an authoritative list; "
                                        "WorkbookLens does not replace semantic values automatically."
                                    ),
                                    discriminator=(column, value, cell.coordinate),
                                )
                            )
        return result


class IsolatedMissingFieldRule(WorkbookRule):
    """Report one interior blank in an otherwise highly complete literal column."""

    rule_id = "WL054_ISOLATED_MISSING_FIELD"
    title = "Isolated blank in a highly complete field"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for worksheet in context.workbook.worksheets:
            profiles = _table_profiles(context, worksheet)
            merged_rows = _merged_row_index(
                worksheet, tuple(row for profile in profiles for row in profile.body_rows)
            )
            for profile in profiles:
                column_contracts = _explicit_column_contracts(context, worksheet, profile)
                width = profile.region.max_column - profile.region.min_column + 1
                for column, header in profile.headers.items():
                    if (
                        _looks_like_identifier_header(header)
                        or _looks_like_free_text_header(header)
                        or not header
                    ):
                        continue
                    if column in column_contracts:
                        continue
                    eligible_rows = [
                        row
                        for row in profile.body_rows
                        if not _indexed_merged_cell(merged_rows, row, column)
                    ]
                    if len(eligible_rows) < MIN_COMPLETENESS_ROWS:
                        continue
                    cells = [_cell_at(worksheet, row, column) for row in eligible_rows]
                    if any(_is_formula(cell) or _is_error(cell) for cell in cells):
                        continue
                    missing_indexes = [
                        index
                        for index, row in enumerate(eligible_rows)
                        if not _is_nonblank(_value_at(worksheet, row, column))
                    ]
                    if len(missing_indexes) != 1:
                        continue
                    coverage = (len(eligible_rows) - 1) / len(eligible_rows)
                    if coverage < MIN_COMPLETE_COVERAGE:
                        continue
                    missing_index = missing_indexes[0]
                    if missing_index in {0, len(eligible_rows) - 1}:
                        continue
                    if not all(
                        _is_nonblank(_value_at(worksheet, eligible_rows[index], column))
                        for index in (missing_index - 1, missing_index + 1)
                    ):
                        continue
                    populated_values = [
                        _value_at(worksheet, row, column)
                        for row in eligible_rows
                        if _is_nonblank(_value_at(worksheet, row, column))
                    ]
                    families = Counter(
                        family
                        for value in populated_values
                        if (family := _value_family(value)) is not None
                    )
                    if not families or families.most_common(1)[0][1] / len(populated_values) < 0.95:
                        continue
                    row = eligible_rows[missing_index]
                    other_populated = sum(
                        _is_nonblank(_value_at(worksheet, row, peer_column))
                        for peer_column in range(
                            profile.region.min_column, profile.region.max_column + 1
                        )
                        if peer_column != column
                    )
                    minimum_record_support = max(2, math.ceil((width - 1) * 0.60))
                    if other_populated < minimum_record_support:
                        continue
                    coordinate = f"{get_column_letter(column)}{row}"
                    evidence = Evidence(
                        summary="One interior record is blank in an otherwise highly complete column",
                        observed=None,
                        expected="nonblank field value or an explicit missing-value marker",
                        peers=[
                            f"{get_column_letter(column)}{eligible_rows[missing_index - 1]}",
                            f"{get_column_letter(column)}{eligible_rows[missing_index + 1]}",
                        ],
                        details={
                            "header": header,
                            "peer_coverage": round(coverage, 4),
                            "eligible_rows": len(eligible_rows),
                            "other_populated_fields": other_populated,
                            "dominant_value_family": families.most_common(1)[0][0],
                        },
                    )
                    result.findings.append(
                        _make_finding(
                            context=context,
                            rule_id=self.rule_id,
                            title=self.title,
                            explanation=(
                                "After excluding merged cells, exactly one interior record is blank in a "
                                "literal column with at least 95% coverage, and that record contains strong "
                                "data in the other table fields."
                            ),
                            severity=Severity.WARNING,
                            confidence=0.88,
                            worksheet=worksheet,
                            location=coordinate,
                            evidence=evidence,
                            expected=(
                                "Highly complete record fields should not contain an unexplained isolated blank."
                            ),
                            suggested_action=(
                                "Check the source record and either restore the value or record an explicit "
                                "missing-value status; WorkbookLens does not invent the missing value."
                            ),
                            discriminator=(column, row),
                        )
                    )
        return result


INFERRED_SEMANTIC_RULES: tuple[type[WorkbookRule], ...] = (
    InferredCategoryVariantRule,
    IsolatedMissingFieldRule,
)

__all__ = ["INFERRED_SEMANTIC_RULES"]
