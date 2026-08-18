# Changelog

All notable changes are documented here. This project follows Semantic Versioning.

## [0.1.0] - Unreleased

### Added

- Seven CLI workflows: scan, plan, apply, semantic diff, YAML tests, local web UI, and demo.
- Fifteen deterministic workbook rules with evidence, stable IDs, confidence, and conservative
  region/formula analysis.
- Five declarative safe patch types implemented through direct OOXML ZIP/XML editing.
- Source/cell fingerprint checks, changed-part manifests, recalculation metadata, reopen/rescan
  validation, and `.xlsm` read-only enforcement.
- Self-contained HTML, JSON, SARIF 2.1.0, semantic diff, and apply reports.
- ZIP/XML/relationship safety limits and adversarial security tests.
- Generated demo workbooks including a chart-preservation fixture.
- Composite GitHub Action, Linux/Windows/macOS CI, and CodeQL workflow.
- Architecture guide, direct-OOXML ADR, security policy, contribution guide, and release checklist.

### Security

- Reject path traversal, backslash/absolute names, duplicates, symlinks, encryption, suspicious
  compression, oversized packages/parts, DTD/entity declarations, malformed XML, missing internal
  targets, and relationship root escapes before workbook parsing.
