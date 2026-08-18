# CLI exit codes

Exit codes are part of the v0.1 automation contract.

| Code | Constant | Meaning |
|---:|---|---|
| 0 | `OK` | command completed successfully |
| 1 | `FINDINGS_OR_ASSERTIONS` | `--fail-on` threshold or YAML assertion failed |
| 2 | `USAGE` | unsupported format/operation, invalid plan selection, or unsafe output path |
| 3 | `UNSAFE_INPUT` | malformed or resource-unsafe workbook package |
| 4 | `STALE_PLAN` | source SHA-256 or a target cell precondition differs |
| 5 | `VALIDATION_FAILED` | OOXML apply or post-apply validation failed |
| 10 | `INTERNAL_ERROR` | unexpected implementation failure |

Typer may also use code 2 for command-line syntax errors before WorkbookLens code runs. A scan does
not fail merely because findings exist unless `--fail-on` is supplied. `test` always exits 1 when a
configured gate fails.
