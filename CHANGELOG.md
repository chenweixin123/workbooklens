# Changelog

All notable changes are documented here. WorkbookLens follows Semantic Versioning.

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
