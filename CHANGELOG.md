# Changelog

All notable changes are documented here. WorkbookLens follows Semantic Versioning.

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
