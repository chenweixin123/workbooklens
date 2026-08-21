# CLI exit codes

Exit codes are part of the version-2 CLI and GitHub Action automation contract.

| Code | Constant | Meaning |
|---:|---|---|
| 0 | OK | command completed successfully |
| 1 | FINDINGS_OR_ASSERTIONS | active fail-on threshold or YAML assertion failed |
| 2 | USAGE | unsupported format/operation, invalid selection, or unsafe output path |
| 3 | UNSAFE_INPUT | malformed or resource-unsafe workbook package |
| 4 | STALE_PLAN | source SHA-256 or a target cell precondition differs |
| 5 | VALIDATION_FAILED | OOXML apply or post-apply validation failed |
| 10 | INTERNAL_ERROR | unexpected implementation failure |

Typer may also use code 2 for command-line syntax errors before WorkbookLens runs. A scan does not
fail merely because findings exist unless --fail-on is supplied. With --baseline and --new-only,
only active new findings participate in the threshold. The test command exits 1 when a configured
gate fails.

`apply --safe-only` selects only operations marked safe with confidence at least 0.95 and always
excludes `layout_review` operations. A layout operation requires explicit `--patch-id` selection and
`--accept-layout-risk`; the flag never enables another unsafe operation. Selecting one member of an
atomic repair group expands the selection to the complete group. Missing layout consent, a patch
below the confidence threshold, an incomplete group presented to the low-level patcher, an unknown
patch ID, or incompatible selection flags exits with code 2 before an output workbook is written.

A source hash or cell/layout fingerprint mismatch exits with code 4. An OOXML structure that makes a
reviewed change unsafe to perform, or a post-write semantic/layout mismatch, exits with code 5 and
the partial output is removed.

The composite Action records the aggregate code in its exit-code output after writing the manifest
and report-path outputs. This lets artifact upload run before the final enforcement step.
