"""Conservative cross-sheet identifier relationship inference."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from typing import cast

from openpyxl.cell.cell import Cell
from openpyxl.utils.cell import get_column_letter

from workbooklens.models import Evidence, Severity
from workbooklens.rules.base import RuleContext, RuleResult, WorkbookRule
from workbooklens.rules.data_quality import (
    _cell_at,
    _identifier_key,
    _is_error,
    _is_formula,
    _looks_like_identifier_header,
    _make_finding,
    _table_profiles,
    _value_at,
)

IdentifierKey = tuple[str, str]
ColumnToken = tuple[str, int, int]

MIN_RELATION_ROWS = 8
MIN_PARENT_UNIQUENESS = 0.85
MIN_UNIQUENESS_GAP = 0.12
MIN_VALUE_OVERLAP = 0.80
MIN_DISTINCT_OVERLAP = 0.75
MIN_SHARED_DISTINCT = 6
MIN_PARENT_SCORE_ADVANTAGE = 0.08

_VALUE_OVERLAP_WEIGHT = 0.45
_DISTINCT_OVERLAP_WEIGHT = 0.45
_PARENT_UNIQUENESS_WEIGHT = 0.10

_RELATIONSHIP_CACHE_KEY = "relational_semantics.identifier_relationships.v1"


@dataclass(frozen=True, slots=True)
class _IdentifierColumn:
    sheet: str
    header_row: int
    column: int
    header: str
    normalized_header: str
    min_body_row: int
    max_body_row: int
    values: tuple[tuple[IdentifierKey, str], ...]
    distinct_keys: frozenset[IdentifierKey]
    uniqueness: float

    @property
    def token(self) -> ColumnToken:
        return self.sheet, self.header_row, self.column

    @property
    def populated(self) -> int:
        return len(self.values)

    @property
    def data_range(self) -> str:
        letter = get_column_letter(self.column)
        return f"{letter}{self.min_body_row}:{letter}{self.max_body_row}"


@dataclass(frozen=True, slots=True)
class InferredIdentifierRelationship:
    parent: _IdentifierColumn
    child: _IdentifierColumn
    value_overlap: float
    distinct_overlap: float
    shared_distinct: int
    unsupported_keys: frozenset[IdentifierKey]


@dataclass(frozen=True, slots=True)
class _RelationshipCandidate:
    parent: _IdentifierColumn
    value_overlap: float
    distinct_overlap: float
    shared_distinct: int

    @property
    def score(self) -> float:
        return (
            _VALUE_OVERLAP_WEIGHT * self.value_overlap
            + _DISTINCT_OVERLAP_WEIGHT * self.distinct_overlap
            + _PARENT_UNIQUENESS_WEIGHT * self.parent.uniqueness
        )


def _identifier_columns(context: RuleContext) -> tuple[_IdentifierColumn, ...]:
    columns: list[_IdentifierColumn] = []
    for worksheet in context.workbook.worksheets:
        for profile in _table_profiles(context, worksheet):
            for column, normalized_header in profile.headers.items():
                if not _looks_like_identifier_header(normalized_header):
                    continue
                values: list[tuple[IdentifierKey, str]] = []
                for row in profile.body_rows:
                    cell = _cell_at(worksheet, row, column)
                    if not isinstance(cell, Cell) or _is_formula(cell) or _is_error(cell):
                        continue
                    key = _identifier_key(cell.value)
                    if key is not None:
                        values.append((key, cell.coordinate))
                if len(values) < MIN_RELATION_ROWS or len(values) / len(profile.body_rows) < 0.75:
                    continue
                distinct_keys = frozenset(key for key, _coordinate in values)
                raw_header = _value_at(worksheet, profile.header_row, column)
                columns.append(
                    _IdentifierColumn(
                        sheet=worksheet.title,
                        header_row=profile.header_row,
                        column=column,
                        header=(
                            str(raw_header).strip() if raw_header is not None else normalized_header
                        ),
                        normalized_header=normalized_header,
                        min_body_row=min(profile.body_rows),
                        max_body_row=max(profile.body_rows),
                        values=tuple(values),
                        distinct_keys=distinct_keys,
                        uniqueness=len(distinct_keys) / len(values),
                    )
                )
    return tuple(columns)


def inferred_identifier_relationships(
    context: RuleContext,
) -> tuple[InferredIdentifierRelationship, ...]:
    cached = context.analysis_cache.get(_RELATIONSHIP_CACHE_KEY)
    if cached is not None:
        return cast(tuple[InferredIdentifierRelationship, ...], cached)

    groups: dict[str, list[_IdentifierColumn]] = defaultdict(list)
    for column in _identifier_columns(context):
        groups[column.normalized_header].append(column)

    relationships: list[InferredIdentifierRelationship] = []
    for group in groups.values():
        if len({column.sheet for column in group}) < 2:
            continue
        ordered_group = sorted(group, key=lambda column: column.token)
        for child in ordered_group:
            candidates: list[_RelationshipCandidate] = []
            for parent in ordered_group:
                if parent.token == child.token or parent.sheet == child.sheet:
                    continue
                uniqueness_gap = parent.uniqueness - child.uniqueness
                if parent.uniqueness < MIN_PARENT_UNIQUENESS or uniqueness_gap < MIN_UNIQUENESS_GAP:
                    continue
                matched = sum(key in parent.distinct_keys for key, _coordinate in child.values)
                shared_distinct = len(parent.distinct_keys & child.distinct_keys)
                value_overlap = matched / child.populated
                distinct_overlap = shared_distinct / len(child.distinct_keys)
                if (
                    value_overlap < MIN_VALUE_OVERLAP
                    or distinct_overlap < MIN_DISTINCT_OVERLAP
                    or shared_distinct < MIN_SHARED_DISTINCT
                ):
                    continue
                candidates.append(
                    _RelationshipCandidate(
                        parent=parent,
                        value_overlap=value_overlap,
                        distinct_overlap=distinct_overlap,
                        shared_distinct=shared_distinct,
                    )
                )
            if not candidates:
                continue
            candidates.sort(
                key=lambda candidate: (
                    -candidate.score,
                    -candidate.value_overlap,
                    -candidate.distinct_overlap,
                    -candidate.shared_distinct,
                    -candidate.parent.uniqueness,
                    candidate.parent.populated,
                    candidate.parent.token,
                )
            )
            best = candidates[0]
            if (
                len(candidates) > 1
                and best.score - candidates[1].score < MIN_PARENT_SCORE_ADVANTAGE
            ):
                # Sheet order and size make deterministic tie-breakers, but they are
                # not evidence that one competing table is the actual parent.
                continue
            parent = best.parent
            unsupported_keys = frozenset(
                key
                for key in child.distinct_keys - parent.distinct_keys
                if sum(key in column.distinct_keys for column in ordered_group) == 1
            )
            relationships.append(
                InferredIdentifierRelationship(
                    parent=parent,
                    child=child,
                    value_overlap=best.value_overlap,
                    distinct_overlap=best.distinct_overlap,
                    shared_distinct=best.shared_distinct,
                    unsupported_keys=unsupported_keys,
                )
            )

    result = tuple(sorted(relationships, key=lambda item: item.child.token))
    context.analysis_cache[_RELATIONSHIP_CACHE_KEY] = result
    return result


def inferred_foreign_key_columns(context: RuleContext) -> frozenset[ColumnToken]:
    return frozenset(
        relationship.child.token for relationship in inferred_identifier_relationships(context)
    )


class CrossSheetOrphanIdentifierRule(WorkbookRule):
    rule_id = "WL057_CROSS_SHEET_ORPHAN_IDENTIFIER"
    title = "Identifier absent from inferred parent table"

    def run(self, context: RuleContext) -> RuleResult:
        result = RuleResult()
        for relationship in inferred_identifier_relationships(context):
            worksheet = context.workbook[relationship.child.sheet]
            for key in sorted(relationship.unsupported_keys):
                coordinates = [
                    coordinate
                    for observed_key, coordinate in relationship.child.values
                    if observed_key == key
                ]
                result.findings.append(
                    _make_finding(
                        context=context,
                        rule_id=self.rule_id,
                        title=self.title,
                        explanation=(
                            "A value in a repeated identifier column is absent from a more unique "
                            "cross-sheet column with the same header and strong value overlap."
                        ),
                        severity=Severity.ERROR,
                        confidence=min(0.99, 0.90 + 0.10 * relationship.value_overlap),
                        worksheet=worksheet,
                        location=",".join(coordinates),
                        evidence=Evidence(
                            summary="Identifier value is absent from the inferred parent column",
                            observed={
                                "normalized_value": key[1],
                                "cells": coordinates,
                            },
                            expected={
                                "parent_sheet": relationship.parent.sheet,
                                "parent_range": relationship.parent.data_range,
                            },
                            details={
                                "header": relationship.child.header,
                                "parent_uniqueness": round(relationship.parent.uniqueness, 4),
                                "child_parent_match_rate": round(relationship.value_overlap, 4),
                                "distinct_match_rate": round(relationship.distinct_overlap, 4),
                                "shared_distinct_values": relationship.shared_distinct,
                                "evidence_level": "STRONG_STRUCTURAL",
                            },
                        ),
                        expected=(
                            "Foreign-key values exist in the inferred parent identifier column."
                        ),
                        suggested_action=(
                            "Review the child value and the inferred parent table; WorkbookLens "
                            "does not invent or replace identifiers."
                        ),
                        discriminator=(relationship.parent.token, key),
                    )
                )
        return result


RELATIONAL_SEMANTIC_RULES: tuple[type[WorkbookRule], ...] = (CrossSheetOrphanIdentifierRule,)

__all__ = [
    "RELATIONAL_SEMANTIC_RULES",
    "CrossSheetOrphanIdentifierRule",
    "InferredIdentifierRelationship",
    "inferred_foreign_key_columns",
    "inferred_identifier_relationships",
]
