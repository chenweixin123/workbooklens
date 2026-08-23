# ruff: noqa: RUF001
"""Canonical built-in rule text and Simplified Chinese display translations."""

from __future__ import annotations

import re

BUILTIN_RULE_TITLES: dict[str, tuple[str, str]] = {
    "WL001_BROKEN_REFERENCE": ("Broken formula reference", "损坏的公式引用"),
    "WL002_FORMULA_PATTERN_OUTLIER": ("Formula pattern outlier", "公式模式异常"),
    "WL003_BLANK_IN_FORMULA_BAND": ("Blank interrupts formula band", "空白单元格中断公式序列"),
    "WL004_HARDCODED_VALUE_IN_FORMULA_BAND": (
        "Hardcoded value interrupts formula band",
        "硬编码值中断公式序列",
    ),
    "WL005_SUSPICIOUS_SUM_BOUNDARY": ("Suspicious SUM boundary", "可疑的 SUM 边界"),
    "WL006_NUMERIC_TEXT": ("Numeric text in numeric region", "数值区域中的文本数字"),
    "WL007_STYLE_OUTLIER": ("Style outlier in homogeneous region", "同质区域中的样式异常"),
    "WL008_HIDDEN_NONEMPTY_DATA": ("Hidden nonempty data", "隐藏的非空数据"),
    "WL009_EXTERNAL_LINK": ("External workbook link", "外部工作簿链接"),
    "WL010_VOLATILE_OR_FRAGILE_FUNCTION": (
        "Volatile or fragile formula construct",
        "易变或脆弱的公式结构",
    ),
    "WL011_ERROR_CELL": ("Stored Excel error value", "已存储的 Excel 错误值"),
    "WL012_DUPLICATE_CONFIGURED_KEY": ("Duplicate configured key", "重复的配置键值"),
    "WL013_BROKEN_DEFINED_NAME": ("Broken defined name", "损坏的定义名称"),
    "WL014_MERGED_CELL_IN_DATA_REGION": (
        "Merged cells intersect a data region",
        "合并单元格与数据区域相交",
    ),
    "WL015_INCONSISTENT_DATA_VALIDATION": (
        "Inconsistent data validation",
        "数据验证不一致",
    ),
    "WL016_TEXT_DISPLAY_RISK": (
        "Text may be clipped or cross a visible boundary",
        "文本可能被截断或越过可见边界",
    ),
    "WL017_BORDER_EDGE_INCONSISTENCY": (
        "Likely missing shared border edge",
        "可能缺失的共享边框",
    ),
    "WL018_USED_RANGE_INFLATION": (
        "Format-only tail inflates the worksheet used range",
        "仅格式尾部扩大工作表已用区域",
    ),
    "WL019_IDENTIFIER_SCIENTIFIC_NOTATION": (
        "Identifier may display in scientific notation",
        "标识符可能以科学记数法显示",
    ),
    "WL020_SAVED_VIEW_OFF_CONTENT": (
        "Saved worksheet view may hide meaningful content",
        "保存的工作表视图可能隐藏有效内容",
    ),
    "WL021_WHITESPACE_ONLY_TAIL": (
        "Whitespace-only cells extend beyond the visible layout",
        "纯空白字符单元格超出可见布局",
    ),
}


ZH_CANONICAL_TEXT: dict[str, str] = {
    "The formula contains an explicit #REF! token and cannot resolve as written.": "公式中含有明确的 #REF! 标记，按当前写法无法解析。",
    "Formula contains #REF!": "公式包含 #REF!",
    "Every formula reference resolves to an existing cell or range.": "每个公式引用都应指向现有单元格或区域。",
    "Review the deleted or moved source range; no automatic guess was made.": "请检查已删除或移动的源区域；WorkbookLens 未进行自动猜测。",
    "Replace the one-off formula with the exact translated peer consensus.": "用参照单元格一致推导出的精确公式替换此孤立公式。",
    "A formula has a different relative-reference signature from a strong band consensus.": "此公式的相对引用结构与该公式序列的强一致模式不同。",
    "Copied formulas in this band have the same structural signature.": "此序列中的复制公式应具有相同的结构签名。",
    "Multiple isolated anomalies were found, so no automatic patch is offered; compare each cell with the listed peers.": "发现多个孤立异常，因此不提供自动修复；请逐一与列出的参照单元格比较。",
    "Aggregate formulas are never replaced automatically; review the subtotal or total manually.": "聚合公式绝不会被自动替换；请手动检查小计或总计。",
    "The row context does not match stable detail-row semantics, so automatic replacement is withheld.": "该行上下文不符合稳定的明细行语义，因此不执行自动替换。",
    "Compare the cell with the listed peers and review any proposed formula.": "请将此单元格与列出的参照单元格比较，并检查建议公式。",
    "Create the missing cell with the exact translated formula agreed by peers.": "使用参照单元格一致推导出的精确公式创建缺失单元格。",
    "A single blank lies between formulas whose translations agree exactly at this cell.": "一个空白单元格位于公式之间，而两侧公式平移到此处后完全一致。",
    "Independent neighboring formulas translate to the same expression": "相邻公式独立平移后得到相同表达式",
    "The contiguous formula band has no unexplained blank.": "连续公式序列中不应存在无法解释的空白。",
    "The blank is inside a merged range and cannot safely receive a formula; review the merge manually.": "空白单元格位于合并区域内，无法安全写入公式；请手动检查合并区域。",
    "The blank is in a hidden sheet, row, or column, so automatic formula creation is withheld.": "空白单元格位于隐藏的工作表、行或列中，因此不自动创建公式。",
    "The blank is locked on a protected sheet, so automatic formula creation is withheld.": "空白单元格在受保护工作表中被锁定，因此不自动创建公式。",
    "The row context does not match stable detail-row semantics, so automatic formula creation is withheld.": "该行上下文不符合稳定的明细行语义，因此不自动创建公式。",
    "Review and select the proposed translated formula.": "请检查并选择建议的平移公式。",
    "Replace the isolated literal with the exact translated peer formula.": "用参照单元格精确平移得到的公式替换孤立常量。",
    "A literal value replaces one cell in an otherwise consistent copied-formula band.": "在原本一致的复制公式序列中，有一个单元格被常量值替代。",
    "Peer formulas translate to one exact replacement": "参照公式平移后得到唯一且精确的替换公式",
    "The formula band follows its consensus structure.": "公式序列应遵循其一致结构。",
    "The cell is inside a merged range, so automatic replacement is withheld.": "该单元格位于合并区域内，因此不执行自动替换。",
    "The cell is in a hidden sheet, row, or column, so automatic replacement is withheld.": "该单元格位于隐藏的工作表、行或列中，因此不执行自动替换。",
    "The cell is locked on a protected sheet, so automatic replacement is withheld.": "该单元格在受保护工作表中被锁定，因此不执行自动替换。",
    "Confirm the literal is not an intentional override before selecting the patch.": "选择修复前，请确认该常量不是有意的人工覆盖。",
    "A simple contiguous total includes its directly adjacent peer row.": "简单连续总计应包含紧邻的同类数据行。",
    "The adjacent row has subtotal or total semantics, so automatic SUM extension is withheld.": "相邻行具有小计或总计语义，因此不自动扩展 SUM 范围。",
    "The SUM target is merged, hidden, or locked on a protected sheet, so automatic extension is withheld.": "SUM 目标已合并、隐藏或在受保护工作表中锁定，因此不自动扩展。",
    "Review the adjacent peer manually. WorkbookLens reports the candidate formula but does not automatically extend SUM boundaries because inclusion semantics cannot be proven from adjacency alone.": "请手动检查相邻数据。WorkbookLens 会报告候选公式，但仅凭相邻关系无法证明应当纳入，因此不会自动扩展 SUM 边界。",
    "Convert an unambiguous numeric string to an OOXML numeric value.": "将含义明确的数字字符串转换为 OOXML 数值。",
    "A plain numeric string appears in a column dominated by numeric values.": "以数值为主的列中出现了普通数字字符串。",
    "Numeric measures use numeric cell storage, while identifiers remain text.": "数值度量应使用数值单元格存储，而标识符应保持文本。",
    "Identifier semantics or explicit text formatting prevent automatic numeric conversion; confirm storage intentionally.": "标识符语义或明确的文本格式阻止自动数值转换；请确认当前存储方式是否有意。",
    "The header does not explicitly identify a numeric measure, so automatic conversion is withheld; confirm the column semantics.": "表头未明确表示数值度量，因此不自动转换；请确认该列语义。",
    "Grouped numeric text is reported for review but is not converted automatically in this release.": "带分组符号的数字文本仅报告供检查，本版本不会自动转换。",
    "The row context does not match stable detail-row semantics, so automatic numeric conversion is withheld.": "该行上下文不符合稳定的明细行语义，因此不自动进行数值转换。",
    "Confirm the value is a measure rather than an identifier.": "请确认该值是度量值而不是标识符。",
    "Copy the existing consensus style ID from the nearest peer.": "从最近的参照单元格复制现有的一致样式 ID。",
    "A populated cell has a different visual style from its column peers.": "一个非空单元格的视觉样式与同列参照单元格不同。",
    "A homogeneous measure column uses its consensus style.": "同质度量列应使用一致样式。",
    "Multiple isolated style anomalies were found, so no automatic patch is offered.": "发现多个孤立样式异常，因此不提供自动修复。",
    "The row context does not match stable detail-row semantics, so no automatic style patch is offered.": "该行上下文不符合稳定的明细行语义，因此不提供自动样式修复。",
    "Check whether the visual distinction is intentional.": "请检查这种视觉差异是否有意。",
    "A hidden worksheet contains data or formulas that may affect interpretation.": "隐藏工作表中含有可能影响理解的数据或公式。",
    "Hidden content is reviewed and documented.": "隐藏内容应经过检查并有相应说明。",
    "Inspect the hidden sheet manually; WorkbookLens never unhides it automatically.": "请手动检查隐藏工作表；WorkbookLens 绝不会自动取消隐藏。",
    "A hidden row contains values or formulas.": "隐藏行中含有值或公式。",
    "Hidden rows with consequential content are intentionally documented.": "含有重要内容的隐藏行应有明确说明。",
    "Review the row manually; no automatic unhide is offered.": "请手动检查该行；不提供自动取消隐藏。",
    "A hidden column contains values or formulas.": "隐藏列中含有值或公式。",
    "Hidden columns with consequential content are intentionally documented.": "含有重要内容的隐藏列应有明确说明。",
    "Review the column manually; no automatic unhide is offered.": "请手动检查该列；不提供自动取消隐藏。",
    "The formula depends on another workbook; WorkbookLens does not fetch it.": "此公式依赖另一个工作簿；WorkbookLens 不会获取该文件。",
    "Formula contains external workbook reference": "公式包含外部工作簿引用",
    "External dependencies are explicit, available, and reviewed.": "外部依赖应当明确、可用且经过检查。",
    "Verify the linked workbook and consider replacing fragile dependencies.": "请验证链接的工作簿，并考虑替换脆弱依赖。",
    "A defined name refers to another workbook.": "一个定义名称引用了另一个工作簿。",
    "Defined name contains external reference": "定义名称包含外部引用",
    "Defined-name dependencies remain local or are explicitly reviewed.": "定义名称的依赖应保持本地，或经过明确检查。",
    "Review the external target; WorkbookLens never opens it.": "请检查外部目标；WorkbookLens 绝不会打开它。",
    "The formula uses constructs that can recalculate frequently or resist static tracing.": "公式使用了可能频繁重算或难以静态追踪的结构。",
    "Performance-sensitive and auditable models avoid unnecessary fragile constructs.": "对性能敏感且要求可审计的模型应避免不必要的脆弱结构。",
    "Review whether a bounded direct reference can express the same intent.": "请检查是否可以用有界的直接引用表达相同意图。",
    "The cell stores a recognized Excel error value.": "单元格存储了可识别的 Excel 错误值。",
    "Stored error cell": "存储错误值的单元格",
    "Calculated or imported values do not contain Excel error tokens.": "计算值或导入值不应包含 Excel 错误标记。",
    "Trace the producing formula or upstream data; no value is fabricated.": "请追踪生成该值的公式或上游数据；WorkbookLens 不会编造替代值。",
    "A value repeats in a column explicitly configured as a unique key.": "在明确配置为唯一键的列中出现了重复值。",
    "Resolve the duplicate records or revise the explicit key configuration.": "请处理重复记录，或修改明确的键配置。",
    "A workbook defined name cannot resolve to an existing valid range.": "工作簿中的定义名称无法解析到现有有效区域。",
    "Defined names resolve to valid local sheets and ranges.": "定义名称应解析到有效的本地工作表和区域。",
    "Repair or remove the name in Excel after confirming downstream usage.": "确认下游用途后，请在 Excel 中修复或删除该名称。",
    "target contains #REF!": "目标包含 #REF!",
    "range target could not be parsed": "无法解析区域目标",
    "range target has no resolvable destination": "区域目标没有可解析的位置",
    "A merge intersects the body of a dense table-like region.": "合并区域与密集表格状区域的主体相交。",
    "Merged range overlaps inferred data body": "合并区域与推断的数据主体重叠",
    "Table-like data bodies use one logical value per cell.": "表格状数据主体中每个单元格应对应一个逻辑值。",
    "Review downstream sort/filter behavior; no automatic unmerge is offered.": "请检查后续排序和筛选行为；不提供自动取消合并。",
    "One populated input cell lacks or differs from the validation used by its peers.": "一个非空输入单元格缺少参照单元格使用的数据验证，或与其不同。",
    "Cells in a homogeneous input column share validation constraints.": "同质输入列中的单元格应共享数据验证约束。",
    "Review and restore the intended validation rule manually.": "请手动检查并恢复预期的数据验证规则。",
    "Widen the repeatedly overflowing text column to the measured local maximum.": "将反复溢出的文本列加宽到本地测量的最大需求宽度。",
    "Wrap the blocked text within its existing cell after review.": "检查后，在现有单元格内对受阻文本启用自动换行。",
    "Unwrapped text exceeds its cell and natural overflow is blocked": "未换行文本超出单元格，且自然溢出受阻",
    "The text is wider than the available cell width and either an adjacent value or a visible border prevents a clean natural overflow.": "文本宽于可用单元格宽度，且相邻值或可见边框阻止其正常自然溢出。",
    "Wrapped or multiline text exceeds an explicit row height": "换行或多行文本超出明确设置的行高",
    "Static text measurement indicates that the saved explicit row height is too small for all wrapped lines.": "静态文本测量表明，保存的明确行高不足以容纳全部换行内容。",
    "Visible text remains inside its intended cell boundary without clipping.": "可见文本应完整显示在预期单元格边界内，不被截断。",
    "The row would exceed Excel's maximum height; widen the layout or shorten the content manually.": "所需行高会超过 Excel 上限；请手动加宽布局或缩短内容。",
    "Review the proposed local wrap/row-height change in Excel; font rendering can vary by device.": "请在 Excel 中检查建议的局部换行和行高更改；不同设备的字体渲染可能不同。",
    "Both sides of one or more shared edges are absent inside a dense rectangular table, or a table perimeter edge is absent, while parallel edges show a stable style. A border present on either side remains visually continuous and is not reported.": "密集矩形表格内部一个或多个共享边的两侧均缺失，或表格外周边缺失，而平行边呈现稳定样式。只要任一侧存在边框，视觉上仍连续，就不会报告。",
    "Shared and perimeter table edges remain visually continuous.": "表格的共享边和外周边应保持视觉连续。",
    "Review the proposed edge-only border copy; no fill, font, or number format is changed. Findings below 0.95 confidence are report-only.": "请检查仅复制边线的建议；填充、字体和数字格式不会改变。置信度低于 0.95 的问题仅报告、不修复。",
    "Clear only the exact reviewed format-only cells and empty row records.": "仅清理经过检查的精确纯格式单元格和空行记录。",
    "A large, separated set of blank styled cells or empty row records extends far beyond the populated content. Broad column-dimension styling by itself is ignored.": "大量彼此分离的空白样式单元格或空行记录远远超出实际内容。仅有整列范围样式不会触发此规则。",
    "Worksheet dimensions reflect meaningful content and intentional structures.": "工作表范围应反映有效内容和有意保留的结构。",
    "Apply the exact-cell cleanup only after reviewing names, print settings, comments, links, breaks, and drawing anchors.": "检查名称、打印设置、批注、链接、分页符和绘图锚点后，方可应用精确单元格清理。",
    "Widen the identifier column while preserving the stored numeric value and type.": "在保留已存储数值及类型的同时加宽标识符列。",
    "A long integer under an identifier-like header uses General formatting and is wider than its column. Excel may display it in scientific notation.": "标识符类表头下的长整数使用常规格式，且宽于所在列。Excel 可能以科学记数法显示。",
    "Long numeric identifier uses General format in a narrow column": "长数字标识符在窄列中使用常规格式",
    "General format forces a long numeric identifier into scientific notation": "常规格式会强制将长数字标识符显示为科学记数法",
    "Phone, ID, account, and similar identifiers display fully without changing their stored value or type.": "电话、ID、账号及类似标识符应完整显示，且不改变存储值或类型。",
    "Review the font-aware width-only proposal and source value.": "请检查考虑字体后的仅调整宽度建议及源值。",
    "General format still uses scientific notation for 12-15 digit integers even in wide columns; choose an explicit integer format or intentional text storage after semantic review.": "即使列足够宽，常规格式仍可能将 12 至 15 位整数显示为科学记数法；请在语义检查后选择明确的整数格式或有意的文本存储。",
    "The value may already have lost precision and is not auto-patched.": "该值可能已经丢失精度，因此不自动修复。",
    "Reset the saved viewport when safe and reduce zoom using conservative two-dimensional layout estimates.": "在安全时重置保存的视口，并根据保守的二维布局估计降低缩放比例。",
    "The saved viewport either starts beyond the intended visible layout or uses a zoom level that makes the intended visible layout difficult to use.": "保存的视口起点超出预期可见布局，或缩放比例使预期布局难以使用。",
    "Compact content is too wide to fit above the automatic zoom safety floor": "紧凑内容过宽，无法在自动缩放安全下限以上完整显示",
    "Saved zoom is too large for the visible width of a vertically scrollable sheet": "保存的缩放比例相对于可纵向滚动工作表的可见宽度过大",
    "Saved viewport is offset or too zoomed-in for the compact visible layout": "保存的视口存在偏移，或对紧凑可见布局放大过度",
    "Opening the sheet starts at the intended content origin and uses a readable zoom for the visible layout.": "打开工作表时应从预期内容起点开始，并使用适合可见布局的可读缩放比例。",
    "The estimated width fit zoom is below the automatic safety floor; review the layout, column widths, or saved zoom manually.": "估计的宽度适配缩放低于自动安全下限；请手动检查布局、列宽或保存的缩放比例。",
    "Review the frozen-pane scroll origin manually; WorkbookLens will not change the number of frozen rows or columns.": "请手动检查冻结窗格的滚动起点；WorkbookLens 不会更改冻结行数或列数。",
    "Review the saved viewport. The proposed zoom is a conservative estimate for a typical desktop window, not a cross-device guarantee.": "请检查保存的视口。建议缩放比例是针对典型桌面窗口的保守估计，并非跨设备保证。",
    "Remove reviewed default-style whitespace nodes and clear only the values of non-default-style cells, preserving their styles and row dimensions.": "移除已检查的默认样式空白字符节点；对非默认样式单元格仅清除值，并保留其样式和行尺寸。",
    "A connected tail of literal whitespace strings extends the worksheet cell bounds beyond its visible content. These cells can render as stray values in some preview tools even though Excel shows them as blank.": "相连的纯空白字符字符串尾部使工作表单元格范围超出可见内容。即使 Excel 将其显示为空白，某些预览工具仍可能把这些单元格渲染为多余值。",
    "Literal whitespace values outside the intended layout are cleared; only intentional non-default style nodes may continue to define stored bounds.": "预期布局之外的纯空白字符值应被清除；只有有意保留的非默认样式节点可以继续定义存储范围。",
    "Apply the exact value cleanup after reviewing the preserved cell styles and blank-row layout.": "检查保留的单元格样式和空白行布局后，应用精确值清理。",
    "Review the referenced structure or formula blocker; no automatic cleanup is offered.": "请检查相关结构或公式阻碍因素；不提供自动清理。",
}


DYNAMIC_TRANSLATIONS: tuple[tuple[re.Pattern[str], str], ...] = (
    (
        re.compile(r"(?P<a>\d+) of (?P<b>\d+) formulas share one signature"),
        "{a}/{b} 个公式共享同一结构签名",
    ),
    (
        re.compile(
            r"A simple total stops one row before a directly adjacent (?:non-hidden )?"
            r"(?P<kind>numeric|formula) peer\."
        ),
        "简单总计在紧邻的{kind}数据行前一行提前结束。",
    ),
    (
        re.compile(r"SUM ends at row (?P<row>\d+), while (?P<cell>[A-Z]+\d+) is adjacent"),
        "SUM 在第 {row} 行结束，而 {cell} 与其相邻",
    ),
    (
        re.compile(r"(?P<count>\d+) peer cells are stored as numbers"),
        "{count} 个参照单元格以数值形式存储",
    ),
    (
        re.compile(r"One visual style appears in (?P<count>\d+) of (?P<total>\d+) peer cells"),
        "一种视觉样式出现在 {total} 个参照单元格中的 {count} 个",
    ),
    (
        re.compile(r"(?P<state>hidden|veryHidden) sheet contains (?P<count>\d+) nonempty cells"),
        "{state}工作表包含 {count} 个非空单元格",
    ),
    (
        re.compile(r"Hidden row (?P<row>\d+) contains (?P<count>\d+) nonempty cells"),
        "隐藏行 {row} 包含 {count} 个非空单元格",
    ),
    (
        re.compile(r"Hidden column range (?P<location>.+) contains (?P<count>\d+) nonempty cells"),
        "隐藏列区域 {location} 包含 {count} 个非空单元格",
    ),
    (
        re.compile(r"Configured key value appears (?P<count>\d+) times"),
        "配置的键值出现了 {count} 次",
    ),
    (re.compile(r"Values in (?P<range>.+) are unique\."), "{range} 中的值应唯一。"),
    (re.compile(r"target sheet (?P<sheet>.+) does not exist"), "目标工作表 {sheet} 不存在"),
    (re.compile(r"target range (?P<range>.+) is invalid"), "目标区域 {range} 无效"),
    (
        re.compile(r"(?P<count>\d+) peer cells share one validation signature"),
        "{count} 个参照单元格共享同一数据验证签名",
    ),
    (
        re.compile(
            r"Increase row (?P<row>\d+) height to (?P<height>[0-9.]+) points after review\."
        ),
        "检查后将第 {row} 行行高增加到 {height} 磅。",
    ),
    (
        re.compile(r"Copy the parallel-consensus (?P<edge>[a-z]+) border edge after review\."),
        "检查后复制平行边一致的{edge}边框边线。",
    ),
    (
        re.compile(r"Missing parallel-consensus edge\(s\): (?P<edges>.+)"),
        "缺少具有平行边一致性的边线：{edges}",
    ),
    (
        re.compile(
            r"(?P<cells>\d+) exact blank styled cells and (?P<rows>\d+) empty row "
            r"records form a separated tail"
        ),
        "{cells} 个精确空白样式单元格和 {rows} 条空行记录形成分离尾部",
    ),
    (
        re.compile(r"(?P<count>\d+) literal-whitespace cells form an outer tail"),
        "{count} 个纯空白字符单元格形成外部尾部",
    ),
)


__all__ = ["BUILTIN_RULE_TITLES", "DYNAMIC_TRANSLATIONS", "ZH_CANONICAL_TEXT"]
