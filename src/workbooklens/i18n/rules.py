"""Display-only localization for findings, patches, and semantic diff labels."""

from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
from typing import Any

from workbooklens.i18n._rule_catalog import (
    BUILTIN_RULE_TITLES,
    DYNAMIC_TRANSLATIONS,
    ZH_CANONICAL_TEXT,
)
from workbooklens.i18n.catalog import Language, normalize_language, translate
from workbooklens.models import Finding, PatchKind, PatchOperation, PatchRisk, Severity
from workbooklens.scanner import ScanResult


class MissingBuiltinTranslationError(RuntimeError):
    """A built-in human-readable message escaped its explicit translation catalog."""


# Finding evidence remains canonical in models and machine-readable exports.
# This explicit vocabulary is used only by the web presentation helper below.
# Arbitrary workbook text must never be guessed or translated.
_EVIDENCE_LABELS: dict[str, tuple[str, str]] = {
    "proof": ("Proof", "证明"),
    "source_proof": ("Source proof", "源证明"),
    "font_size": ("Font size", "字号"),
    "fixed_total_pages": ("Fixed total pages", "固定总页数"),
    "proof_level": ("Proof level", "证明级别"),
    "evidence_level": ("Evidence level", "证据级别"),
    "coverage": ("Coverage", "覆盖范围"),
    "source": ("Source", "来源"),
    "source_cell": ("Source cell", "源单元格"),
    "source_sheet": ("Source sheet", "源工作表"),
    "cell": ("Cell", "单元格"),
    "sheet": ("Sheet", "工作表"),
    "range": ("Range", "范围"),
    "formula": ("Formula", "公式"),
    "error": ("Error", "错误"),
    "reason": ("Reason", "原因"),
    "kind": ("Kind", "类型"),
    "role": ("Role", "角色"),
    "anomaly": ("Anomaly", "异常类型"),
    "observed": ("Observed", "实际值"),
    "expected": ("Expected", "预期值"),
    "excluded_cells": ("Excluded cells", "排除的单元格"),
    "merged_range": ("Merged range", "合并范围"),
    "fixed_rows": ("Fixed rows", "固定行数"),
    "fixed_columns": ("Fixed columns", "固定列数"),
    "total_pages": ("Total pages", "总页数"),
}

_EVIDENCE_VALUES: dict[str, tuple[str, str]] = {
    "propagated_formula_error": ("Propagated formula error", "传播的公式错误"),
    "numeric": ("Numeric", "数值"),
    "formula": ("Formula", "公式"),
    "Chart": ("Chart", "图表"),
    "Image": ("Image", "图片"),
    "hidden": ("Hidden", "隐藏"),
    "veryHidden": ("Very hidden", "深度隐藏"),
    "PROVEN_STATIC": ("Proven static", "静态证明"),
    "STRONG_STRUCTURAL": ("Strong structural", "强结构证据"),
    "strong_structural": ("Strong structural", "强结构证据"),
    "proven_static_structure": ("Proven static structure", "静态结构证明"),
    "advisory": ("Advisory", "提示性证据"),
    "semantic_heuristic": ("Semantic heuristic", "语义启发式"),
    "whole_percent_scale": ("Whole-percent scale", "整百分比倍率"),
    "outside_fraction_range": ("Outside fraction range", "超出小数范围"),
    "blank_outside_content": ("Blank outside content", "内容范围外空白"),
    "blank_inside_dense_table": ("Blank inside dense table", "密集表格内空白"),
    "single_edit": ("Single edit", "单点编辑"),
    "placeholder_marker": ("Placeholder marker", "占位标记"),
}
_EVIDENCE_VALUE_CONTEXTS = frozenset(
    {
        "anomaly",
        "candidate_kind",
        "dominant_input_kind",
        "dominant_kind",
        "evidence_level",
        "finding_kind",
        "issue_kind",
        "kind",
        "method",
        "proof",
        "proof_level",
        "reason",
        "role",
        "source_proof",
        "state",
        "type",
    }
)


def _localize_evidence_key(value: str, language: Language) -> str:
    """Translate a known structured-evidence key only."""

    labels = _EVIDENCE_LABELS.get(value)
    if labels is None:
        return value
    return labels[0] if language == "en" else labels[1]


def _localize_evidence_string(
    value: str,
    language: Language,
    *,
    context_key: str | None = None,
) -> str:
    """Translate a known enum-like evidence value, leaving workbook text intact."""

    # Short enum words (for example "formula" or "hidden") are translated only
    # in their canonical field context so ordinary workbook text is never
    # rewritten.
    if context_key not in _EVIDENCE_VALUE_CONTEXTS:
        return value
    labels = _EVIDENCE_VALUES.get(value)
    if labels is None:
        return value
    return labels[0] if language == "en" else labels[1]


def localize_evidence_value(value: Any, language: str | None = None) -> Any:
    """Return a recursively localized, non-mutating copy for web display.

    Evidence is part of the canonical finding contract, so this helper belongs
    at the presentation boundary. It translates only stable keys and enum-like
    values while preserving unknown text, scalar types, and container shape.
    """

    normalized = normalize_language(language)

    def visit(item: Any, context_key: str | None = None) -> Any:
        if isinstance(item, dict):
            localized: dict[Any, Any] = {}
            original_keys = set(item)
            for key, child in item.items():
                display_key = (
                    _localize_evidence_key(key, normalized)
                    if isinstance(key, str)
                    else deepcopy(key)
                )
                # A workbook/plugin may already use a translated-looking key.
                # Keep canonical spelling instead of silently overwriting it.
                if display_key != key and display_key in original_keys:
                    display_key = key
                if display_key in localized:
                    display_key = key
                localized[display_key] = visit(child, key if isinstance(key, str) else None)
            return localized
        if isinstance(item, list):
            return [visit(child, context_key) for child in item]
        if isinstance(item, tuple):
            return tuple(visit(child, context_key) for child in item)
        if isinstance(item, str):
            return _localize_evidence_string(item, normalized, context_key=context_key)
        return deepcopy(item)

    return visit(value)


def _dynamic_translation(text: str) -> str | None:
    for pattern, template in DYNAMIC_TRANSLATIONS:
        match = pattern.fullmatch(text)
        if match is None:
            continue
        values = match.groupdict()
        kind = values.get("kind")
        if kind:
            values["kind"] = {
                "numeric": "数值",
                "formula": "公式",
                "Chart": "图表",
                "Image": "图片",
            }.get(kind, kind)
        role = values.get("role")
        if role:
            values["role"] = {
                "email": "邮箱",
                "phone": "电话",
                "currency": "金额",
                "percentage": "百分比",
                "date": "日期",
            }.get(role, role)
        if values.get("state") == "hidden":
            values["state"] = "隐藏"
        elif values.get("state") == "veryHidden":
            values["state"] = "深度隐藏"
        edge = values.get("edge")
        if edge:
            values["edge"] = {
                "left": "左侧",
                "right": "右侧",
                "top": "上侧",
                "bottom": "下侧",
            }.get(edge, edge)
        return template.format_map(values)
    return None


def canonical_text_translation(text: str, language: str | None = None) -> str | None:
    """Translate one exact built-in sentence or supported dynamic template."""

    if normalize_language(language) == "en":
        return text
    return ZH_CANONICAL_TEXT.get(text) or _dynamic_translation(text)


def _localize_required(
    text: str,
    language: Language,
    *,
    rule_id: str,
    field: str,
    strict: bool,
) -> str:
    translated = canonical_text_translation(text, language)
    if translated is not None:
        return translated
    if strict:
        raise MissingBuiltinTranslationError(f"{rule_id} has no {field} translation: {text!r}")
    return translate("i18n.missing_builtin_text", language, rule_id=rule_id)


def localize_finding(
    finding: Finding,
    language: str | None = None,
    *,
    strict: bool = False,
) -> Finding:
    """Return a display copy while preserving IDs, fingerprints, and structured evidence."""

    normalized = normalize_language(language)
    titles = BUILTIN_RULE_TITLES.get(finding.rule_id)
    if titles is None:
        return finding.model_copy(
            update={"title": translate("i18n.plugin_provided", normalized, text=finding.title)}
        )
    expected_title = titles[0]
    if strict and finding.title != expected_title:
        raise MissingBuiltinTranslationError(
            f"{finding.rule_id} title changed from catalog: {finding.title!r}"
        )
    title = expected_title if normalized == "en" else titles[1]
    if normalized == "en":
        return finding.model_copy(update={"title": title})

    if finding.rule_id == "WL010_VOLATILE_OR_FRAGILE_FUNCTION":
        evidence_summary = finding.evidence.summary
    else:
        evidence_summary = _localize_required(
            finding.evidence.summary,
            normalized,
            rule_id=finding.rule_id,
            field="evidence.summary",
            strict=strict,
        )
    evidence = finding.evidence.model_copy(update={"summary": evidence_summary})
    return finding.model_copy(
        update={
            "title": title,
            "explanation": _localize_required(
                finding.explanation,
                normalized,
                rule_id=finding.rule_id,
                field="explanation",
                strict=strict,
            ),
            "evidence": evidence,
            "expected": _localize_required(
                finding.expected,
                normalized,
                rule_id=finding.rule_id,
                field="expected",
                strict=strict,
            ),
            "suggested_action": _localize_required(
                finding.suggested_action,
                normalized,
                rule_id=finding.rule_id,
                field="suggested_action",
                strict=strict,
            ),
        }
    )


def localize_patch(
    patch: PatchOperation,
    language: str | None = None,
    *,
    builtin: bool = True,
    strict: bool = False,
) -> PatchOperation:
    """Return a display copy; the localized copy must never authorize a repair."""

    normalized = normalize_language(language)
    if normalized == "en":
        return patch.model_copy()
    if builtin:
        description = _localize_required(
            patch.description,
            normalized,
            rule_id="built-in patch",
            field="description",
            strict=strict,
        )
    else:
        description = translate("i18n.plugin_provided", normalized, text=patch.description)
    return patch.model_copy(update={"description": description})


def localize_scan_result(
    scan: ScanResult,
    language: str | None = None,
    *,
    strict: bool = False,
) -> ScanResult:
    """Return a display-only scan copy; canonical JSON and repair data stay untouched."""

    normalized = normalize_language(language)
    builtin_ids = set(BUILTIN_RULE_TITLES)
    builtin_patch_ids = {
        patch_id
        for finding in scan.findings
        if finding.rule_id in builtin_ids
        for patch_id in finding.patch_ids
    }
    return replace(
        scan,
        findings=[
            localize_finding(finding, normalized, strict=strict) for finding in scan.findings
        ],
        patches=[
            localize_patch(
                patch,
                normalized,
                builtin=patch.id in builtin_patch_ids,
                strict=strict,
            )
            for patch in scan.patches
        ],
    )


def severity_label(severity: Severity | str, language: str | None = None) -> str:
    value = severity.value if isinstance(severity, Severity) else severity
    return translate(f"severity.{value}", language, default=value)


def risk_label(risk: PatchRisk | str, language: str | None = None) -> str:
    value = risk.value if isinstance(risk, PatchRisk) else risk
    return translate(f"risk.{value}", language, default=value)


def patch_kind_label(kind: PatchKind | str, language: str | None = None) -> str:
    value = kind.value if isinstance(kind, PatchKind) else kind
    return translate(f"patch_kind.{value}", language, default=value)


_CHANGE_TYPE_LABELS_ZH: dict[str, str] = {
    "value": "值",
    "formula": "公式",
    "style": "样式",
    "number_format": "数字格式",
    "sheet_renamed": "工作表重命名",
    "sheet_removed": "移除工作表",
    "sheet_added": "新增工作表",
    "sheet_reordered": "工作表重新排序",
    "sheet_visibility": "工作表可见性",
    "declared_dimension": "声明范围",
    "content_dimension": "内容范围",
    "view_top_left_cell": "视图左上角单元格",
    "view_zoom_scale": "视图缩放比例",
    "formatting_tail_cells_removed": "移除格式尾部单元格",
    "formatting_tail_cells_added": "新增格式尾部单元格",
    "defined_name": "定义名称",
    "calculation_mode": "计算模式",
    "merged_range_added": "新增合并区域",
    "merged_range_removed": "移除合并区域",
    "hidden_row_added": "新增隐藏行",
    "hidden_row_removed": "取消隐藏行",
    "hidden_column_added": "新增隐藏列",
    "hidden_column_removed": "取消隐藏列",
    "data_validation_added": "新增数据验证",
    "data_validation_removed": "移除数据验证",
    "row_height": "行高",
    "column_width": "列宽",
}


def change_type_label(change_type: str, language: str | None = None) -> str:
    """Translate a stable semantic-diff type for display only."""

    if normalize_language(language) == "zh-CN":
        return _CHANGE_TYPE_LABELS_ZH.get(change_type, change_type)
    return change_type.replace("_", " ").title()


__all__ = [
    "BUILTIN_RULE_TITLES",
    "MissingBuiltinTranslationError",
    "canonical_text_translation",
    "change_type_label",
    "localize_evidence_value",
    "localize_finding",
    "localize_patch",
    "localize_scan_result",
    "patch_kind_label",
    "risk_label",
    "severity_label",
]
