# ruff: noqa: RUF001
"""Stable English and Simplified Chinese product message catalog."""

from __future__ import annotations

from typing import Any, Literal, TypeAlias

Language: TypeAlias = Literal["en", "zh-CN"]
DEFAULT_LANGUAGE: Language = "en"
SUPPORTED_LANGUAGES: tuple[Language, ...] = ("en", "zh-CN")


def normalize_language(value: str | None, fallback: Language = DEFAULT_LANGUAGE) -> Language:
    """Normalize browser, CLI, and stored locale spellings."""

    if not value:
        return fallback
    normalized = value.strip().replace("_", "-").lower()
    if normalized == "zh" or normalized.startswith(("zh-cn", "zh-hans")):
        return "zh-CN"
    if normalized == "en" or normalized.startswith("en-"):
        return "en"
    return fallback


_MESSAGE_PAIRS: dict[str, tuple[str, str]] = {
    "language.en": ("English", "English"),
    "language.zh-CN": ("Simplified Chinese", "简体中文"),
    "i18n.plugin_provided": ("Plugin-provided: {text}", "第三方插件提供：{text}"),
    "i18n.missing_builtin_text": (
        "Localized explanation unavailable for built-in rule {rule_id}.",
        "内置规则 {rule_id} 的本地化说明暂不可用。",
    ),
    "severity.critical": ("Critical", "严重"),
    "severity.error": ("Error", "错误"),
    "severity.warning": ("Warning", "警告"),
    "severity.info": ("Info", "信息"),
    "risk.safe": ("Safe", "安全"),
    "risk.layout_review": ("Layout review", "需检查布局"),
    "sheet_state.visible": ("Visible", "可见"),
    "sheet_state.hidden": ("Hidden", "隐藏"),
    "sheet_state.veryHidden": ("Very hidden", "深度隐藏"),
    "common.yes": ("Yes", "是"),
    "common.no": ("No", "否"),
    "patch_kind.set_formula": ("Set formula", "设置公式"),
    "patch_kind.set_numeric": ("Set numeric value", "设置数值"),
    "patch_kind.copy_style": ("Copy style", "复制样式"),
    "patch_kind.extend_sum": ("Extend SUM", "扩展 SUM 范围"),
    "patch_kind.create_formula": ("Create formula", "创建公式"),
    "patch_kind.set_column_width": ("Set column width", "设置列宽"),
    "patch_kind.set_row_height": ("Set row height", "设置行高"),
    "patch_kind.set_wrap_text": ("Wrap text", "自动换行"),
    "patch_kind.set_shrink_to_fit": ("Shrink to fit", "缩小字体填充"),
    "patch_kind.set_text": ("Set text", "设置文本"),
    "patch_kind.set_sheet_view": ("Set sheet view", "设置工作表视图"),
    "patch_kind.copy_border": ("Copy border edge", "复制边框边线"),
    "patch_kind.clear_formatting_tail": ("Clear formatting tail", "清理格式尾部"),
    "patch_kind.remove_whitespace_tail_cells": (
        "Remove whitespace-only tail",
        "清理纯空白字符尾部",
    ),
    "web.brand_subtitle": ("Local Excel inspection and repair", "本地 Excel 检查与修复"),
    "web.local_status": ("Running locally", "正在本地运行"),
    "web.language_label": ("Language", "语言"),
    "web.nav_inspect": ("Inspect workbook", "检查工作簿"),
    "web.nav_convert": ("Convert XLS", "转换 XLS"),
    "web.home_title": ("Inspect and repair an Excel workbook", "检查并修复 Excel 工作簿"),
    "web.home_intro": (
        "Files are processed on this computer. Choose a workbook to begin.",
        "文件仅在这台电脑上处理。请选择工作簿开始操作。",
    ),
    "web.privacy_title": ("Private by design", "隐私优先"),
    "web.privacy_body": (
        "Files remain in a process-owned local temporary directory and are removed during normal shutdown. Scanning and repair do not execute formulas, macros, external links, or embedded objects.",
        "文件保存在当前进程专用的本地临时目录中，并在正常退出时删除。扫描和修复不会执行公式、宏、外部链接或嵌入对象。",
    ),
    "web.scan_title": ("Choose a workbook", "选择工作簿"),
    "web.scan_label": ("Supported: .xlsx; .xlsm is scan-only", "支持 .xlsx；.xlsm 仅可扫描"),
    "web.scan_hint": (
        "The workbook is inspected locally and is not uploaded to a cloud service.",
        "工作簿仅在本机检查，不会上传到云端服务。",
    ),
    "web.file_choose": ("Choose file", "选择文件"),
    "web.file_none": ("No file selected", "未选择文件"),
    "web.scan_button": ("Scan locally", "本地扫描"),
    "web.scan_limit": ("Maximum file size: {max_mb} MB.", "文件大小上限：{max_mb} MB。"),
    "web.scan_running": ("Scanning locally...", "正在本地扫描..."),
    "web.convert_title": ("Convert a legacy workbook", "转换旧版工作簿"),
    "web.convert_label": (
        "Choose a trusted binary .xls file",
        "选择可信的二进制 .xls 文件",
    ),
    "web.convert_hint": (
        "Creates a local macro-free .xlsx copy.",
        "创建不含宏的本地 .xlsx 副本。",
    ),
    "web.convert_trust_title": ("Trusted files only", "仅处理可信文件"),
    "web.convert_trust_body": (
        "Microsoft Excel or LibreOffice opens the workbook locally and may recalculate formulas or process workbook-defined behavior supported by that application.",
        "Microsoft Excel 或 LibreOffice 会在本机打开工作簿，并可能重新计算公式或处理该应用支持的工作簿内置行为。",
    ),
    "web.convert_button": ("Convert to .xlsx", "转换为 .xlsx"),
    "web.convert_running": ("Converting locally...", "正在本地转换..."),
    "web.convert_unavailable": ("Converter unavailable", "转换器不可用"),
    "web.convert_provider": (
        "Available locally: {providers}. Conversion fidelity depends on the application; review the downloaded copy.",
        "本机可用：{providers}。转换保真度取决于所用应用，请检查下载的副本。",
    ),
    "web.results_title": ("Inspection results", "检查结果"),
    "web.results_summary": (
        "{findings} findings and {patches} proposed repairs",
        "发现 {findings} 个问题，提出 {patches} 项修复",
    ),
    "web.results_actions": ("Downloads and actions", "下载与操作"),
    "web.results_report": ("Download HTML report", "下载 HTML 报告"),
    "web.results_plan": ("Download repair plan", "下载修复计划"),
    "web.results_review_title": ("Review proposed repairs", "检查建议修复"),
    "web.results_review_help": (
        "Only explicitly selected repairs are applied. Original files are never overwritten.",
        "只会应用明确选中的修复，原文件绝不会被覆盖。",
    ),
    "web.results_layout_consent": (
        "I reviewed and accept the layout changes",
        "我已检查并接受这些布局更改",
    ),
    "web.results_apply": ("Apply selected repairs", "应用所选修复"),
    "web.results_no_patches": ("No reviewable repairs were proposed.", "没有可供检查的修复建议。"),
    "web.results_findings": ("Findings", "发现的问题"),
    "web.results_evidence": ("Evidence", "依据"),
    "web.results_workbook_scope": ("Workbook", "工作簿"),
    "web.results_new_scan": ("Inspect another workbook", "检查另一个工作簿"),
    "web.applied_title": ("Repairs completed", "修复已完成"),
    "web.applied_body": (
        "A new repaired workbook was created; the original was not changed.",
        "已创建新的修复版工作簿，原文件未被更改。",
    ),
    "web.applied_resolved": ("Resolved findings", "已解决的问题"),
    "web.applied_new_findings": ("New findings", "新增问题"),
    "web.applied_download": ("Download repaired workbook", "下载修复后的工作簿"),
    "web.applied_diff": ("Download semantic diff", "下载语义差异报告"),
    "web.applied_apply_report": ("Download apply report", "下载修复执行报告"),
    "web.error_title": (
        "WorkbookLens could not complete this action",
        "WorkbookLens 无法完成此操作",
    ),
    "web.error_source_safe": ("The source workbook was not modified.", "源工作簿未被修改。"),
    "web.error_details": ("Details", "详细信息"),
    "web.error_diagnostic_id": ("Diagnostic ID", "诊断编号"),
    "web.error_retry": ("Return and try again", "返回并重试"),
    "web.language_apply": ("Apply language", "应用语言"),
    "web.footer_local": (
        "WorkbookLens runs locally on this computer.",
        "WorkbookLens 正在此电脑上本地运行。",
    ),
    "web.provider_separator": (", then ", "，然后 "),
    "web.value_empty": ("None", "无"),
    "web.scan_intro": (
        "Inspect formulas, values, styles, borders, dimensions, and saved views.",
        "检查公式、数值、样式、边框、尺寸和保存的视图。",
    ),
    "web.scan_running_detail": (
        "This may take a moment for a large workbook.",
        "大型工作簿可能需要一些时间。",
    ),
    "web.convert_intro": (
        "Create a modern .xlsx copy from a trusted legacy .xls workbook.",
        "将可信的旧版 .xls 工作簿转换为现代 .xlsx 副本。",
    ),
    "web.convert_running_detail": (
        "Keep WorkbookLens open while the local spreadsheet application completes conversion.",
        "本地电子表格应用完成转换前，请保持 WorkbookLens 打开。",
    ),
    "web.convert_install_provider": (
        "Install Microsoft Excel or LibreOffice, then restart WorkbookLens.",
        "请安装 Microsoft Excel 或 LibreOffice，然后重启 WorkbookLens。",
    ),
    "web.results_page_title": ("Workbook inspection results", "工作簿检查结果"),
    "web.results_intro": (
        "Review each finding and select only the repairs you understand.",
        "请逐项检查发现的问题，只选择您理解并接受的修复。",
    ),
    "web.metric_findings": ("Findings", "问题数"),
    "web.metric_patches": ("Proposed repairs", "建议修复数"),
    "web.metric_errors": ("Errors", "错误数"),
    "web.metric_warnings": ("Warnings", "警告数"),
    "web.confidence": ("Confidence", "置信度"),
    "web.layout_consent_body": (
        "Layout repairs can change widths, heights, wrapping, borders, or saved views. Review them before applying.",
        "布局修复可能更改列宽、行高、换行、边框或保存的视图，应用前请仔细检查。",
    ),
    "web.apply_running": ("Applying selected repairs...", "正在应用所选修复..."),
    "web.apply_running_detail": (
        "WorkbookLens is validating a new copy; the source file remains unchanged.",
        "WorkbookLens 正在验证新副本，源文件保持不变。",
    ),
    "web.no_findings": (
        "No findings were raised by the enabled deterministic rules.",
        "已启用的确定性规则未发现问题。",
    ),
    "web.applied_intro": (
        "The repaired copy passed WorkbookLens validation.",
        "修复副本已通过 WorkbookLens 验证。",
    ),
    "web.metric_applied": ("Applied repairs", "已应用修复"),
    "web.metric_resolved": ("Resolved findings", "已解决问题"),
    "web.metric_new_findings": ("New findings", "新增问题"),
    "web.metric_output_hash": ("Output SHA-256", "输出文件 SHA-256"),
    "web.applied_validation": (
        "Validation completed before the repaired copy was offered for download.",
        "修复副本仅在完成验证后才可下载。",
    ),
    "web.error_suggestion": ("Suggested action", "建议操作"),
    "web.error_code": ("Error code", "错误代码"),
    "cli.error": ("Error", "错误"),
    "cli.internal_error": ("Internal error", "内部错误"),
    "cli.diagnostic_id": ("Diagnostic ID", "诊断编号"),
    "cli.scanned": ("Scanned", "扫描完成"),
    "cli.scan_summary": (
        "{active} active, {suppressed} suppressed, {new} new; report {report}",
        "{active} 个有效问题，{suppressed} 个已抑制问题，{new} 个新问题；报告 {report}",
    ),
    "cli.expired_suppressions": ("Expired suppressions ignored:", "已忽略过期的抑制项："),
    "cli.planned": ("Planned", "计划已生成"),
    "cli.plan_summary": (
        "{patches} patches ({safe} safe, {layout} layout review) → {output}",
        "{patches} 项修复（{safe} 项安全修复，{layout} 项需检查布局）→ {output}",
    ),
    "cli.applied": ("Applied and validated", "已应用并通过验证"),
    "cli.apply_summary": ("{patches} patches → {output}", "{patches} 项修复 → {output}"),
    "cli.apply_report": ("Apply report", "修复执行报告"),
    "cli.compared": ("Compared", "比较完成"),
    "cli.diff_summary": (
        "{cells} cell and {structures} structural changes → {output}",
        "{cells} 项单元格更改和 {structures} 项结构更改 → {output}",
    ),
    "cli.assertions_title": ("Workbook assertions", "工作簿断言"),
    "cli.result": ("Result", "结果"),
    "cli.assertion": ("Assertion", "断言"),
    "cli.message": ("Message", "说明"),
    "cli.pass": ("PASS", "通过"),
    "cli.fail": ("FAIL", "失败"),
    "cli.suppressed_summary": (
        "{count} findings suppressed by documented waivers.",
        "已根据书面豁免抑制 {count} 个问题。",
    ),
    "cli.demo_complete": ("Demo complete", "演示完成"),
    "cli.before": ("Before", "修复前"),
    "cli.after": ("After", "修复后"),
    "cli.plan": ("Plan", "修复计划"),
    "cli.diff": ("Diff", "差异报告"),
    "cli.language_help": ("Output language: en or zh-CN.", "输出语言：en 或 zh-CN。"),
    "assertion.threshold": (
        "Observed {observed} {severity} findings; maximum is {maximum}",
        "发现 {observed} 个{severity}级问题；上限为 {maximum}",
    ),
    "assertion.prohibited": (
        "Matched {count} prohibited findings",
        "匹配到 {count} 个禁止出现的问题",
    ),
    "assertion.duplicates": ("Found {count} duplicate value groups", "发现 {count} 组重复值"),
    "assertion.allowed": (
        "Found {count} values outside the allowed domain",
        "发现 {count} 个不在允许范围内的值",
    ),
    "assertion.blank": ("Found {count} blank cells", "发现 {count} 个空白单元格"),
    "assertion.numeric": (
        "Found {count} nonnumeric or out-of-bounds cells",
        "发现 {count} 个非数值或越界单元格",
    ),
    "assertion.compared": ("Compared {left} with {right}", "已比较 {left} 与 {right}"),
    "assertion.unsafe": ("Assertion could not be evaluated safely.", "无法安全计算此断言。"),
    "report.scan_title": ("WorkbookLens scan", "WorkbookLens 扫描报告"),
    "report.workbook_summary": ("Workbook summary", "工作簿摘要"),
    "report.weighted_score": ("Weighted finding score", "问题加权分数"),
    "report.score_help": (
        "critical 35 · error 15 · warning 5 · info 1; additive, not a percentage",
        "严重 35 · 错误 15 · 警告 5 · 信息 1；分数累加，不是百分比",
    ),
    "report.critical": ("Critical", "严重"),
    "report.errors": ("Errors", "错误"),
    "report.warnings": ("Warnings", "警告"),
    "report.active_total": ("Active / total", "有效 / 总数"),
    "report.local_title": ("Local processing.", "本地处理。"),
    "report.local_body": (
        "WorkbookLens made no network request while inspecting this workbook. This report can contain workbook-derived evidence; whoever distributes or uploads it controls that data. Formulas, macros, links, and embedded objects were not executed.",
        "WorkbookLens 检查此工作簿时未发出网络请求。本报告可能包含来自工作簿的依据，分发或上传报告的人负责控制这些数据。公式、宏、链接和嵌入对象均未执行。",
    ),
    "report.policy_title": ("Finding policy", "问题处理策略"),
    "report.baseline_active": ("Baseline comparison is active", "已启用基线比较"),
    "report.new_only": ("only new findings are shown and gated", "仅显示新问题并用于判定"),
    "report.baseline_counts": (
        "Known baseline findings: {baseline_known} · new findings: {new}.",
        "基线已知问题：{baseline_known} · 新问题：{new}。",
    ),
    "report.suppressed_count": (
        "{count} findings were suppressed by documented configuration waivers.",
        "已根据配置中的书面豁免抑制 {count} 个问题。",
    ),
    "report.expired_suppressions": ("Expired suppressions ignored:", "已忽略过期的抑制项："),
    "report.sheet_overview": ("Sheet overview", "工作表概览"),
    "report.sheet": ("Sheet", "工作表"),
    "report.state": ("State", "状态"),
    "report.relevant_cells": ("Relevant cells", "相关单元格"),
    "report.content_range": ("Content range", "内容范围"),
    "report.declared_range": ("Declared range", "声明范围"),
    "report.saved_view": ("Saved view", "保存的视图"),
    "report.findings": ("Findings", "发现的问题"),
    "report.severity": ("Severity", "严重程度"),
    "report.all": ("All", "全部"),
    "report.rule": ("Rule", "规则"),
    "report.search": ("Search", "搜索"),
    "report.search_placeholder": ("title, cell, evidence", "标题、单元格、依据"),
    "report.sort": ("Sort", "排序"),
    "report.sort_risk": ("Risk order", "按风险"),
    "report.sort_rule": ("Rule", "按规则"),
    "report.sort_location": ("Sheet/location", "按工作表/位置"),
    "report.location": ("Location", "位置"),
    "report.confidence": ("Confidence", "置信度"),
    "report.explanation_evidence": ("Explanation and evidence", "说明与依据"),
    "report.evidence": ("Evidence", "依据"),
    "report.observed": ("Observed:", "观察值："),
    "report.expected": ("Expected:", "期望值："),
    "report.peers": ("Peers:", "参照单元格："),
    "report.suggested_action": ("Suggested action:", "建议操作："),
    "report.patch": ("Patch", "修复"),
    "report.safe_only": ("safe-only", "仅安全修复"),
    "report.no_findings": (
        "No findings were raised by the enabled deterministic rules.",
        "已启用的确定性规则未发现问题。",
    ),
    "report.no_active_findings": (
        "No active findings remain after baseline and suppression policy. The scan still recorded {baseline_known} baseline-known and {suppressed} suppressed findings.",
        "应用基线与抑制策略后没有有效问题。本次扫描仍记录了 {baseline_known} 个基线已知问题和 {suppressed} 个已抑制问题。",
    ),
    "report.suppressed_findings": ("Suppressed findings", "已抑制的问题"),
    "report.suppressed_help": (
        "Suppressed findings do not affect the active score, SARIF, or failure threshold.",
        "已抑制的问题不会影响有效分数、SARIF 或失败阈值。",
    ),
    "report.waiver": ("Waiver", "豁免项"),
    "report.reason": ("Reason", "原因"),
    "report.expires": ("expires", "到期"),
    "report.limitations": ("Limitations", "局限性"),
    "report.limitations_body": (
        "WorkbookLens does not calculate formulas and does not claim complete Excel compatibility. Cached formula values may be stale. Formula repairs remove stale cached values and request a full recalculation the next time the output is opened in a compatible spreadsheet application. Shared, array, data-table, dynamic-array, and ambiguous formulas are not automatically repaired.",
        "WorkbookLens 不计算公式，也不声称完全兼容 Excel。缓存的公式值可能已过期。公式修复会移除过期缓存值，并要求下次使用兼容的电子表格应用打开输出文件时执行完整重算。共享公式、数组公式、数据表公式、动态数组公式和含义不明确的公式不会自动修复。",
    ),
    "report.generated": (
        "Generated by WorkbookLens {version} · self-contained report",
        "由 WorkbookLens {version} 生成 · 独立完整报告",
    ),
    "diff.title": ("WorkbookLens semantic diff", "WorkbookLens 语义差异"),
    "diff.heading": ("Semantic workbook diff", "工作簿语义差异"),
    "diff.before": ("Before", "修复前"),
    "diff.after": ("After", "修复后"),
    "diff.summary": (
        "{cells} cell changes · {structures} structural changes. Formula values were not calculated.",
        "{cells} 项单元格更改 · {structures} 项结构更改。未计算公式值。",
    ),
    "diff.sheet": ("Sheet", "工作表"),
    "diff.type": ("Type", "类型"),
    "diff.importance": ("Importance", "重要程度"),
    "diff.search": ("Search", "搜索"),
    "diff.subject": ("Subject", "对象"),
    "diff.no_changes": ("No semantic differences detected.", "未检测到语义差异。"),
    "diff.privacy": (
        "Self-contained WorkbookLens report with no remote assets. It can contain workbook-derived evidence; whoever distributes or uploads it controls that data.",
        "这是不加载远程资源的独立完整 WorkbookLens 报告。报告可能包含来自工作簿的依据，分发或上传报告的人负责控制这些数据。",
    ),
    "change_type.value": ("Value", "值"),
    "change_type.formula": ("Formula", "公式"),
    "change_type.style": ("Style", "样式"),
    "change_type.number_format": ("Number format", "数字格式"),
    "change_type.structure": ("Structure", "结构"),
    "error.security.cross_origin.title": ("Request blocked", "请求已被阻止"),
    "error.security.cross_origin.message": (
        "The form was submitted from an untrusted origin.",
        "此表单来自不受信任的来源。",
    ),
    "error.security.cross_origin.suggestion": (
        "Open WorkbookLens from its installed shortcut and retry in the local application window.",
        "请从已安装的 WorkbookLens 快捷方式打开本地应用窗口后重试。",
    ),
    "error.security.csrf.title": ("Session verification failed", "会话验证失败"),
    "error.security.csrf.message": (
        "The local form security token is missing or expired.",
        "本地表单安全令牌缺失或已过期。",
    ),
    "error.security.csrf.suggestion": (
        "Reload the page and submit the form again.",
        "请重新加载页面后再次提交。",
    ),
    "error.upload.too_large.title": ("Workbook is too large", "工作簿过大"),
    "error.upload.too_large.message": (
        "The selected file exceeds the configured local upload limit.",
        "所选文件超过了本地上传大小限制。",
    ),
    "error.upload.too_large.suggestion": (
        "Choose a smaller workbook or increase the local limit before restarting WorkbookLens.",
        "请选择较小的工作簿，或提高本地限制并重启 WorkbookLens。",
    ),
    "error.upload.empty.title": ("Workbook is empty", "工作簿为空"),
    "error.upload.empty.message": (
        "The selected file contains no data.",
        "所选文件不包含任何数据。",
    ),
    "error.upload.empty.suggestion": (
        "Choose a valid workbook and try again.",
        "请选择有效的工作簿后重试。",
    ),
    "error.upload.invalid_type.title": ("Unsupported file type", "不支持的文件类型"),
    "error.upload.invalid_type.message": (
        "The selected file type is not supported for this action.",
        "当前操作不支持所选文件类型。",
    ),
    "error.upload.invalid_type.suggestion": (
        "Use .xlsx or .xlsm for inspection, or a trusted binary .xls file for conversion.",
        "检查请使用 .xlsx 或 .xlsm；转换请使用可信的二进制 .xls 文件。",
    ),
    "error.session.limit.title": ("Too many local sessions", "本地会话过多"),
    "error.session.limit.message": (
        "WorkbookLens reached its local session limit.",
        "WorkbookLens 已达到本地会话数量上限。",
    ),
    "error.session.limit.suggestion": (
        "Restart WorkbookLens and try again.",
        "请重启 WorkbookLens 后重试。",
    ),
    "error.session.not_found.title": ("Session expired", "会话已过期"),
    "error.session.not_found.message": (
        "The local workbook session no longer exists.",
        "本地工作簿会话已不存在。",
    ),
    "error.session.not_found.suggestion": (
        "Return to the home screen and inspect the workbook again.",
        "请返回首页并重新检查工作簿。",
    ),
    "error.conversion.invalid_input.title": (
        "XLS file could not be accepted",
        "无法接受此 XLS 文件",
    ),
    "error.conversion.invalid_input.message": (
        "The selected file is not a recognized binary Excel .xls workbook.",
        "所选文件不是可识别的二进制 Excel .xls 工作簿。",
    ),
    "error.conversion.invalid_input.suggestion": (
        "Open the source in a trusted spreadsheet application and save it as .xls, then retry.",
        "请在可信的电子表格应用中打开源文件并另存为 .xls，然后重试。",
    ),
    "error.conversion.unavailable.title": (
        "No local converter is available",
        "没有可用的本地转换器",
    ),
    "error.conversion.unavailable.message": (
        "WorkbookLens could not find Microsoft Excel or LibreOffice on this computer.",
        "WorkbookLens 未在此电脑上找到 Microsoft Excel 或 LibreOffice。",
    ),
    "error.conversion.unavailable.suggestion": (
        "Install Microsoft Excel or LibreOffice, restart WorkbookLens, and retry.",
        "请安装 Microsoft Excel 或 LibreOffice，重启 WorkbookLens 后重试。",
    ),
    "error.conversion.all_providers_failed.title": ("XLS conversion failed", "XLS 转换失败"),
    "error.conversion.all_providers_failed.message": (
        "Every available local converter stopped before producing a verified .xlsx file.",
        "所有可用的本地转换器都在生成通过验证的 .xlsx 文件之前停止。",
    ),
    "error.conversion.all_providers_failed.suggestion": (
        "Confirm the .xls file is trusted and opens normally in Microsoft Excel or LibreOffice, then retry. Use the diagnostic ID when reporting the problem.",
        "请确认该 .xls 文件可信且能在 Microsoft Excel 或 LibreOffice 中正常打开，然后重试。报告问题时请提供诊断编号。",
    ),
    "error.conversion.output_invalid.title": (
        "Converted workbook failed verification",
        "转换后的工作簿未通过验证",
    ),
    "error.conversion.output_invalid.message": (
        "The converter produced a file that WorkbookLens could not verify as a macro-free .xlsx workbook.",
        "转换器生成的文件未能通过 WorkbookLens 的无宏 .xlsx 验证。",
    ),
    "error.conversion.output_invalid.suggestion": (
        "Open the source in a trusted spreadsheet application and use Save As to create a new .xlsx copy.",
        "请在可信的电子表格应用中打开源文件，并使用“另存为”创建新的 .xlsx 副本。",
    ),
    "error.conversion.timeout.title": ("XLS conversion timed out", "XLS 转换超时"),
    "error.conversion.timeout.message": (
        "The local spreadsheet application did not finish conversion within the allowed time.",
        "本地电子表格应用未在规定时间内完成转换。",
    ),
    "error.conversion.timeout.suggestion": (
        "Close other spreadsheet windows, confirm the file opens normally, and retry.",
        "请关闭其他电子表格窗口，确认文件能正常打开后重试。",
    ),
    "error.scan.failed.title": ("Workbook inspection failed", "工作簿检查失败"),
    "error.scan.failed.message": (
        "WorkbookLens could not inspect this workbook within its safety limits.",
        "WorkbookLens 无法在安全限制内检查此工作簿。",
    ),
    "error.scan.failed.suggestion": (
        "Confirm the workbook is a valid .xlsx or .xlsm file and try again.",
        "请确认文件是有效的 .xlsx 或 .xlsm 工作簿后重试。",
    ),
    "error.repair.selection_required.title": ("No repairs selected", "尚未选择修复项"),
    "error.repair.selection_required.message": (
        "Select at least one reviewed repair.",
        "请至少选择一项已经检查的修复。",
    ),
    "error.repair.selection_required.suggestion": (
        "Select one or more proposed repairs and submit again.",
        "请选择一项或多项建议修复后再次提交。",
    ),
    "error.repair.stale_plan.title": ("Repair plan is out of date", "修复计划已过期"),
    "error.repair.stale_plan.message": (
        "The workbook or repair plan changed after inspection.",
        "检查完成后，工作簿或修复计划发生了变化。",
    ),
    "error.repair.stale_plan.suggestion": (
        "Inspect the current workbook again and create a new repair plan.",
        "请重新检查当前工作簿并创建新的修复计划。",
    ),
    "error.repair.validation_failed.title": ("Repair was not applied", "未应用修复"),
    "error.repair.validation_failed.message": (
        "A repair failed WorkbookLens validation, so the operation stopped without replacing the source.",
        "某项修复未通过 WorkbookLens 验证，因此操作已停止且未替换源文件。",
    ),
    "error.repair.validation_failed.suggestion": (
        "Inspect the workbook again or choose a smaller set of repairs. Use the diagnostic ID when reporting the problem.",
        "请重新检查工作簿，或减少所选修复项。报告问题时请提供诊断编号。",
    ),
    "error.download.not_ready.title": ("Download is not ready", "下载文件尚未准备好"),
    "error.download.not_ready.message": (
        "The requested local output has not been created or has expired.",
        "请求的本地输出尚未创建或已过期。",
    ),
    "error.download.not_ready.suggestion": (
        "Repeat the preceding action, then download the result again.",
        "请重新执行前一步操作，然后再次下载结果。",
    ),
    "error.request.invalid.title": ("Request could not be completed", "无法完成请求"),
    "error.request.invalid.message": (
        "The requested operation or input is not supported.",
        "不支持所请求的操作或输入。",
    ),
    "error.request.invalid.suggestion": (
        "Review the selected file and options, then try again.",
        "请检查所选文件和选项后重试。",
    ),
    "error.request.baseline_required.title": (
        "Baseline is required",
        "需要基线报告",
    ),
    "error.request.baseline_required.message": (
        "--new-only requires --baseline.",
        "使用 --new-only 时必须同时提供 --baseline。",
    ),
    "error.request.baseline_required.suggestion": (
        "Provide a previous findings.json file with --baseline, or remove --new-only.",
        "请使用 --baseline 提供先前的 findings.json，或移除 --new-only。",
    ),
    "error.request.baseline_scope_mismatch.title": (
        "Baseline source does not match",
        "基线来源不匹配",
    ),
    "error.request.baseline_scope_mismatch.message": (
        "The baseline source scope does not match the workbook source scope.",
        "基线来源范围与当前工作簿的来源范围不匹配。",
    ),
    "error.request.baseline_scope_mismatch.suggestion": (
        "Use a baseline created for this workbook scope, or correct --source-scope.",
        "请使用为当前工作簿范围创建的基线，或修正 --source-scope。",
    ),
    "error.internal.unexpected.title": ("Unexpected local error", "发生意外的本地错误"),
    "error.internal.unexpected.message": (
        "WorkbookLens encountered an unexpected local error. No source workbook was overwritten.",
        "WorkbookLens 遇到了意外的本地错误，源工作簿未被覆盖。",
    ),
    "error.internal.unexpected.suggestion": (
        "Restart WorkbookLens and retry. Use the diagnostic ID when reporting the problem.",
        "请重启 WorkbookLens 后重试。报告问题时请提供诊断编号。",
    ),
}


class _SafeFormatDict(dict[str, Any]):
    def __missing__(self, key: str) -> str:
        return "{" + key + "}"


def translate(
    key: str,
    language: str | None = None,
    *,
    default: str | None = None,
    **params: Any,
) -> str:
    """Translate a stable key, with English and caller-provided fallbacks."""

    normalized = normalize_language(language)
    pair = _MESSAGE_PAIRS.get(key)
    if pair is None:
        template = default if default is not None else key
    else:
        template = pair[0 if normalized == "en" else 1]
    return template.format_map(_SafeFormatDict(params))


def require_translation(key: str, language: str | None = None, **params: Any) -> str:
    """Translate a known product key and raise if the catalog is incomplete."""

    if key not in _MESSAGE_PAIRS:
        raise KeyError(f"Missing translation for {key}")
    return translate(key, language, **params)


def message_catalog(language: str | None = None, *, prefix: str | None = None) -> dict[str, str]:
    """Return a copy of the selected catalog, optionally filtered by prefix."""

    return {
        key: require_translation(key, language)
        for key in _MESSAGE_PAIRS
        if prefix is None or key.startswith(prefix)
    }


def assert_catalog_complete() -> None:
    """Validate pair shape and nonempty values for both supported languages."""

    invalid = [
        key
        for key, pair in _MESSAGE_PAIRS.items()
        if len(pair) != 2 or not pair[0].strip() or not pair[1].strip()
    ]
    if invalid:
        raise AssertionError(f"Incomplete translations: {sorted(invalid)}")


__all__ = [
    "DEFAULT_LANGUAGE",
    "SUPPORTED_LANGUAGES",
    "Language",
    "assert_catalog_complete",
    "message_catalog",
    "normalize_language",
    "require_translation",
    "translate",
]
