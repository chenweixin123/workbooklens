# Changelog

All notable changes are documented here. WorkbookLens follows Semantic Versioning.

## [2.2.0] - 2026-08-22

### Added

- Add `WL016_TEXT_DISPLAY_RISK` for deterministic detection of vertically clipped wrapped or
  multiline text and horizontally blocked overflow. Repeated overflow can propose a bounded column
  width; local overflow uses an atomic wrap-and-row-height group.
- Add `WL017_BORDER_EDGE_INCONSISTENCY` with visual shared-edge semantics. A border stored on either
  adjoining cell is treated as present. Dense rectangular tables also detect completely borderless
  internal holes and missing perimeter edges; reviewed repairs require at least 95% peer consensus.
- Add `WL018_USED_RANGE_INFLATION` for separated format-only tails. Proposed cleanup enumerates the
  exact blank styled cells and empty row records instead of clearing a rectangular range.
- Add `WL019_IDENTIFIER_SCIENTIFIC_NOTATION` for long integers under identifier-like headers.
  Ten- and eleven-digit values can receive a font-aware width-only proposal that preserves their
  stored numeric value and type. General-formatted values of 12-15 digits remain findings-only
  because Excel can force scientific notation regardless of column width; longer values remain
  findings-only because their original precision cannot be recovered safely.
- Add `WL020_SAVED_VIEW_OFF_CONTENT` for visible sheets saved below or right of their first content,
  or at a zoom unlikely to show a compact sheet's full estimated width and height in a typical
  desktop window. Reviewed repairs can reset the top-left cell and reduce excessive zoom while
  accounting for earlier width/height proposals. Zoom-only repairs preserve an unshifted frozen pane;
  shifted frozen panes and split panes remain findings-only.
- Add `WL021_WHITESPACE_ONLY_TAIL` for connected literal-space cells beyond the visible layout
  envelope. Reviewed cleanup removes default-style nodes, clears only the value of styled nodes, and
  preserves their style IDs, fonts, alignment, protection, and custom row dimensions.
- Record declared/content dimensions, explicit row heights, column widths, saved top-left cells, and
  zoom in snapshots and semantic diffs. Large format-tail removals are summarized instead of emitted
  as hundreds of blank-cell style changes.

### Changed

- Classify column width, row height, alignment, saved-view, edge-border, and exact-tail
  operations as `layout_review`. They are always `safe=false`, are excluded from `--safe-only`, and
  require explicit patch selection plus `--accept-layout-risk` at confidence 0.95 or greater.
- Close related wrap/height operations into atomic groups so a wrapped cell cannot be applied without
  the row-height change needed to display it.
- Extend direct OOXML patching to reviewed row, column, alignment, text, view, border-edge, and
  format-tail changes without saving the whole package through a spreadsheet application.
- Preserve identifier value semantics in `WL019`; the low-level text-replacement primitive is not
  proposed by this rule and requests formula recalculation if used by another reviewed workflow.
- Size identifier columns with the cell font, include every newly wrapped cell in its atomic row
  height, and treat hidden columns as zero-width when checking natural text overflow.
- Detect peer-consensus border gaps in styled blank table cells, preserve merged-range extents during
  exact UsedRange-tail cleanup, include visible empty border/fill templates in view fitting, and report
  compact sheets that cannot fit above the automatic 50% zoom safety floor instead of silently
  treating them as acceptable.
- Treat single-row merged text as bounded by its merge, retain saved zoom above 100% when the content
  still fits, include full merged extents and visible row heights in two-dimensional viewport
  estimates, include prior row-height and column-width proposals in the final fit, preserve frozen-pane
  XML during zoom-only repair, and coalesce same-column width proposals to their largest sufficient
  value.
- Require a materialized internal peer before proposing a shared border edge, reject conflicting
  duplicate patch identities, and preserve intentional custom-height rows during both formatting-
  tail and whitespace-tail cleanup.
- Keep text findings review-only when wrapping would exceed Excel's maximum row height, avoid
  snapshot-side creation of default row dimensions, and reject an applied patch when the same rule
  still reports the targeted sheet and cell under a changed finding identity.
- Show safe and layout-review counts in CLI plans and require a separate layout-risk confirmation in
  the loopback web UI.

### Security

- Harden the loopback web UI with exact Host validation, Origin/Referer checks, an HttpOnly
  SameSite CSRF cookie plus form token, restrictive browser headers, and declared and streaming
  request-body limits enforced before multipart/form parsing.
- Bind layout patches to row, column, view, or exact-tail fingerprints and revalidate them against a
  fresh canonical scan before writing.
- Make format-tail cleanup fail closed on formulas, defined names, tables, validation and conditional
  formatting ranges, hyperlinks, comments, page breaks, drawing anchors, hidden or outlined rows,
  unsupported row metadata, and other intersecting worksheet structures.
- Apply the same reference checks to whitespace-tail removal and reject rich strings, formulas,
  dynamic `INDIRECT`/`OFFSET` references, metadata-bearing cells, and any stale dimension or layout
  fingerprint.
- Continue to write a new output file, enforce the changed-part allowlist, reopen with two readers,
  rescan, and delete partial output after validation failure.

### Compatibility

- The scan-report schema remains version 2, while the patch-plan schema advances from version 1 to 2.
  Existing rule IDs retain their stable identities. Consumers that exhaustively enumerate rules,
  patch kinds, risk values, or snapshot fields must accept the new version-2 members.
- Repair plans remain version-bound and source-bound; regenerate a 2.1 plan with WorkbookLens 2.2
  before applying it.
- `.xlsm` remains read-only, and `.xls`, `.xlsb`, `.ods`, and Google Sheets remain unsupported.

## [2.1.0] - 2026-08-20

### Changed

- Preserve repair safety for protected input cells and explicitly text-formatted cells by refusing
  style-copy patches that would change their protection or text-storage semantics.
- Restrict automatic numeric-text conversion to an explicit measure-header allowlist. Identifier
  columns such as IDs, SKUs, account numbers, postal codes, and Chinese identifier fields, unknown
  columns, grouped numeric strings, and explicitly text-formatted values are findings-only.
  A separate numeric-text anomaly in an identifier column no longer prevents an otherwise safe
  explicit-measure conversion on the same ordinary detail row.
- Keep every merged-range cell, summary or subtotal row, protected worksheet, non-visible worksheet,
  and hidden row or column out of automatic repair. Grouped hidden column spans are now recognized
  across their complete range and represented accurately in snapshots.
- Require formula and style repairs to establish stable detail-row context across all text columns
  and peer visual styles. Secondary override labels, labels outside the inferred data rectangle,
  intentionally highlighted rows, unique notes, and free-form-only row labels are review-only.
- Require style-copy repairs to preserve number format, protection, quote-prefix, and pivot-button
  semantics; protection-only differences are not reported as visual style anomalies.
- Skip valid Chartsheet relationships during worksheet analysis and patching while preserving
  Chartsheet, chart, drawing, and relationship parts byte-for-byte.
- Exclude boundary totals and subtotals from formula-outlier replacement, report multiple isolated
  formula/style anomalies without bulk auto-repair, and make all suspicious SUM-boundary findings
  review-only because adjacency cannot prove inclusion semantics.

### Security

- Recheck non-visible sheets, hidden rows, grouped hidden columns, protected sheets, and semantic
  style fields in the low-level OOXML patch preconditions.
- Reject hidden column spans outside Excel's A:XFD limit before expanding snapshot metadata.

### Compatibility

- JSON schemas and rule IDs remain unchanged from 2.0.0.
- `PatchKind.EXTEND_SUM` remains in the serialized enum for compatibility, but 2.1 does not generate
  it as a canonical automatic repair.
- Serialized patch plans continue to be revalidated against a fresh canonical scan before repair
  authority is granted.

## [2.0.0] - 2026-08-19

### Added

- Source-scoped and aggregate-manifest baselines with stable rule/location identities,
  evidence-content fingerprints, and --new-only gating.
- Version-2 YAML finding suppressions with reasons, optional expiry, and auditable test output.
- Findings report schema version 2 with total, active, suppressed, known, and new counts.
- Composite Action scan/test modes with config, baseline, new-only, .xlsx, and .xlsm support.
- Action manifest schema version 2 and explicit aggregate exit-code output.
- Wheel/sdist allowlists and executable artifact-content auditing.
- Linux, Windows, and macOS fresh-wheel smoke tests, dependency review, Dependabot, issue forms,
  pull-request checklist, and tag-triggered release-candidate verification.

### Changed

- Package, documentation, SARIF, and Action links use
  https://github.com/chenweixin123/workbooklens.
- Version metadata is 2.0.0 across the package, lock file, documentation, Action examples, and
  release checks.
- Public scan JSON consumers must handle schema version 2. Stable finding IDs remain the baseline
  comparison key; source_scope prevents a findings report from being reused for another workbook.
- Semantic diff compares value types and canonical styles, and formula analysis ignores constructs
  inside string literals.

### Security

- Distribution inspection rejects virtual environments, caches, bytecode, key/certificate
  material, environment files, unexpected archive roots, and oversized members.
- Action config, baseline, scan root, and output paths must remain inside GITHUB_WORKSPACE.
- Test mode rejects baseline-only semantics rather than silently ignoring them.
- Repair execution rejects any plan whose patch fields differ from a fresh canonical scan and treats
  macro content or extension/content-type mismatches as read-only.

## [0.1.0] - 2026-08-17

- Initial deterministic scan, plan, apply, diff, YAML test, local web UI, demo, reports, security
  limits, direct-OOXML repair, cross-platform CI, and CodeQL release.
