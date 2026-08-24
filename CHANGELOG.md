# Changelog

All notable changes are documented here. WorkbookLens follows Semantic Versioning.

## [Unreleased]

## [2.4.0] - 2026-08-24

### Added

- Add an optional user-editable Workbook Profile with bounded sheet/column contracts for required
  fields, enumerations and near-match evidence, email and phone structure, fixed-width identifiers,
  trailing whitespace, and currency/percentage/date role conflicts. Business values remain
  report-only; an explicitly requested ASCII-space cleanup is always `layout_review`.
- Add a conservative Formula IR and report-only rules for truncated aggregate ranges, invalid or
  structurally blank cross-sheet targets, isolated formulas in notes columns, algebraically
  degenerate top-level formulas, and formula-label/number-format role conflicts.
- Add report-only layout geometry for chart/image overlap with non-source data or KPI regions,
  fixed-format numeric display width risk, unexplained explicit row-height outliers, and
  role-aware title/header/body/total style fragmentation.
- Add report-only inferred data-quality rules for duplicate or missing identifiers, mixed numeric
  storage, robust numeric outliers, percentage-scale anomalies, date-storage anomalies, and
  sign-domain violations: negative values under nonnegative headers, and zero or negative values
  under positive headers.
- Add report-only structure checks for partial AutoFilter coverage, chart category/value range
  mismatches, deeply displaced freeze panes, and print areas that truncate inferred tables or chart
  frames already included by the print area.
- Add formula-text, bounded static circular-dependency, and provable formula-error findings. The
  provable error rule is intentionally narrow and recognizes only deterministic expressions such as
  direct NA(), simple constant division by zero, and clearly nonnumeric literal VALUE(...) calls.
  Saved OOXML formula-error caches are reported separately as advisory evidence and may be stale.

### Changed

- Add bilingual Select all/Clear all controls and a live selected-count summary to the local repair
  review page. Bulk selection is scoped only to proposed repairs and never checks the separate
  layout-risk consent on the user's behalf.
- Keep references to missing worksheets as `ERROR`, while classifying an empty target outside
  recorded content as low-confidence advisory `INFO`. Add scan-scoped sparse-index,
  formula-cell, and row-label caches; these formula checks remain findings-only.
- Apply one 1,000,000-cell Profile-range limit to direct scanner configuration and YAML Profiles,
  and resolve sparse Profile body rows without walking the full declared rectangle.
- Make Profile configuration fail closed for boolean-to-integer coercion, out-of-range header rows,
  duplicate worksheet/table definitions, ambiguous or missing headers, out-of-range columns, and
  multiple selectors resolving to the same physical column. A bounded range now defaults its
  header row to the range's first row, and Profiles require configuration version 2.
- Resolve chart source worksheet names plus global and worksheet-local defined names
  case-insensitively, matching Excel semantics. Suppress overlap findings only for exact chart
  source cells, preserving covered non-source columns for WL047 detection.
- Interpret zero-offset TwoCellAnchor end markers at the leading edge of the marker cell, avoiding
  false print-area truncation findings while retaining real row or column overflow findings.
- Expand the built-in deterministic registry from 35 to 50 rules and require exact Chinese/English
  title, explanation, evidence-summary, expectation, and suggested-action coverage for every
  built-in module.
- Extend version-2 YAML validation with a strict, bounded `profile.sheets[].columns[]` schema while
  preserving direct scanner calls that omit a Profile.
- Bound Hatchling below 1.32 because 1.32 emits Core Metadata 2.5, which current Twine 6.2 rejects;
  the compatible build emits Core Metadata 2.4 and passes strict distribution checks.
- Build copied-formula consensus across inferred data-body columns so several anomalies no longer
  split the evidence into short bands. Multiple formula, blank, or hardcoded anomalies are reported
  but are never patched automatically; a formula patch still requires one anomaly, at least 0.95
  confidence, exact translation agreement, and stable visible detail-row semantics.
- Exclude fully blank structural separator rows from missing-formula findings, preventing false
  creation of formulas immediately before subtotal or total rows.
- Cache circular-reference analysis once per scan and bound range expansion to avoid quadratic work
  on broad references while retaining direct and bounded-range cycle detection.
- Prefer an atomic bounded-width, wrap, and row-height proposal for one isolated extreme long-text
  row. Multiple dependent rows remain findings-only, and automatic row-height proposals no longer
  grow to visually disruptive sizes.
- Remove the stale fixed-version wording from the OOXML input safety message.
- Harden role-aware style inspection when an OOXML border side is absent instead of represented by
  an empty `Side`, preventing a valid workbook from raising `WL-INT-001`.
- Keep WL007 from treating a data region's first-row top edge or last-row bottom edge as a column
  style anomaly when every other visible style component matches; WL017 remains responsible for
  border continuity findings.
- Generalize the Windows installer upgrade smoke test to accept a pinned earlier numeric release,
  while retaining the legacy portable-profile exception only for version 2.2.1.

## [2.3.0] - 2026-08-23

### Added

- Add a native Windows desktop shell built with pywebview and Microsoft Edge WebView2. The
  windowed `WorkbookLens.exe` starts the loopback-only FastAPI service silently, renders it inside
  the named application window, supports downloads, and stops the service when the window closes.
- Add a persistent Chinese/English selector. Pages, built-in findings, repair suggestions, scan
  reports, semantic-diff reports, and safe error messages follow the selected language while stable
  IDs, fingerprints, formulas, paths in structured data, JSON, and SARIF contracts remain canonical.
- Add separate `WorkbookLens.exe` and `WorkbookLensCLI.exe` Windows entry points, with the branded
  WorkbookLens icon applied to the desktop executable, installer, Start-menu entry, and shortcut.

### Changed

- Windows shortcuts and `Start-WorkbookLens.cmd` now open the no-console desktop application rather
  than a server console and external browser. The console CLI remains available in
  `WorkbookLensCLI.exe` and through normal Python package installation.
- The desktop launcher verifies that pywebview selected EdgeChromium before starting the local
  service and checks again after the window closes. A missing WebView2 runtime therefore produces
  the localized startup dialog instead of silently falling back to the legacy MSHTML renderer.
- The workbook picker is rendered by WorkbookLens itself, so its button, selected-file state, and
  validation guidance follow the chosen Chinese or English language instead of mixing browser- or
  operating-system-provided labels into the interface.
- Excel-backed `.xls` conversion no longer fails merely because Office or endpoint-security policy
  denies `Process.Path`. The path is an optional identity signal; the converter continues to bind
  ownership to the held process handle, PID, creation time, and Windows session before cleanup.
- After Excel automation requests a normal shutdown, conversion waits up to 15 seconds for that
  verified process to exit before using its held process handle for termination. This avoids an
  intermittent Windows access-denied race while refusing to stop pre-existing or identity-mismatched
  Excel processes.
- Conversion and desktop failures now return localized, stable error codes and diagnostic IDs.
  `WL-CNV-001` through `WL-CNV-005` distinguish invalid input, unavailable providers, provider
  failure, invalid output, and timeout without exposing PowerShell CLIXML or raw exception details.

### Security

- The native desktop shell remains local-only: its HTTP service binds to `127.0.0.1`, workbook
  content is not uploaded to a cloud service, and persistent WebView data stays under the current
  user's local application-data directory.
- User-facing conversion errors redact workbook names, private paths, commands, stack traces, and
  CLIXML payloads. Rotating local diagnostics use the stable diagnostic identifier to correlate a
  failure without returning the raw provider payload to the interface.

### Compatibility

- The native Windows interface requires Microsoft Edge WebView2 Runtime. Windows 10 and 11 normally
  provide it; a localized startup dialog identifies the missing runtime without showing a traceback.
- Existing CLI commands, machine-readable schemas, rule and patch IDs, and Action inputs/outputs are
  unchanged. Third-party plugin text remains in its original language when no trusted translation is
  available.

## [2.2.1] - 2026-08-22

### Added

- Add an opt-in local web conversion button for OLE-based `.xls` workbooks. It prefers an installed
  Microsoft Excel automation server, falls back to an isolated LibreOffice profile, never uses a
  cloud converter, and releases only outputs that pass the normal macro-free OOXML safety gate.
  This separate trust boundary is intended only for trusted legacy files because the installed
  provider opens the workbook and may recalculate formulas or process workbook-defined behavior.
- Add an official Windows x64 portable ZIP built from the validated wheel with CPython 3.12 and
  PyInstaller. It runs without a separately installed Python, writes no registry keys, requires no
  administrator access, and includes project, Python, and bundled-dependency license notices.
- Add an official per-user Windows x64 setup executable. It provides a normal Install wizard,
  creates a named Start-menu entry and an optional desktop shortcut, registers with Windows
  Installed apps, and includes both Start-menu and Settings-based uninstall paths.
- Add a double-click launcher that starts the existing local UI, waits for its health endpoint, and
  then opens the default browser. The full console CLI remains available in the same executable.
- Add frozen-application smoke tests for CLI workflows, HTML templates, Chinese and spaced paths,
  real CSRF-protected multipart workbook uploads, loopback-only serving, graceful shutdown, and
  release archive safety.

### Changed

- The `serve` command can explicitly open a browser and can optionally fall back to a system-chosen
  port. Both behaviors are opt-in, so existing CLI automation retains its previous defaults.
- The double-click launcher opens a dedicated console so `Ctrl+C` stops the server without a second
  batch-job confirmation. Redirected frozen-process output uses UTF-8 for non-ASCII paths.
- The local OpenAPI document reports the installed WorkbookLens version.
- Release artifact checks reject extra files in the distribution directory as well as duplicate,
  case-colliding, encrypted, linked, oversized, or suspicious archive members. The only optional
  build marker is uv's exact generated `.gitignore`, and the new frozen entry modules are required.
- Portable license collection follows declared `License-File` metadata and recognizes both
  `LICENSE*` and `LICENCE*` names, failing the build when a declared license file is missing.
- Excel-backed `.xls` conversion materializes locale-specific built-in number formats as explicit
  OOXML formats, so downstream readers preserve date semantics such as Chinese year-month cells.
- Setup upgrades remove the previously managed portable payload before copying the new version,
  preventing renamed or removed runtime files from surviving an in-place upgrade.

### Security

- Accept `localhost` and `127.0.0.1` as equivalent form origins only when scheme and port match, and
  accept sandboxed-browser `Origin: null` submissions only after the existing CSRF cookie and form
  token both validate. External origins, changed ports or schemes, and invalid tokens remain blocked.
- Scope the non-execution guarantee to normal OOXML scan, test, diff, and repair workflows. Optional
  `.xls` conversion invokes an installed spreadsheet application and is explicitly documented as a
  trusted-file-only boundary.
- The local server reserves its `127.0.0.1` socket before Uvicorn starts, removing the port-probe
  race. Browser launch waits for `/health` and bypasses system HTTP proxies.

### Compatibility

- The installer and portable package support 64-bit Windows 10 and 11. They are currently unsigned
  and may trigger Microsoft SmartScreen; users should verify the published SHA-256 checksum.
- Portable builds include the built-in rules. Third-party entry-point plugins continue to require a
  normal Python wheel installation.

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
