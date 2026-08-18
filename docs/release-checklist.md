# Release checklist

## Correctness

- [ ] `uv lock --check` and `uv sync --locked` succeed from a clean checkout.
- [ ] Format, lint, strict typing, and the complete test suite pass.
- [ ] Coverage remains above the enforced threshold.
- [ ] The demo produces before/after workbooks, reports, a plan, an apply manifest, and a diff.
- [ ] Every required CLI command and `--help` works.
- [ ] JSON models remain backward-compatible or the schema version changes.

## Workbook preservation

- [ ] Search the repair package for any `openpyxl.save()` call.
- [ ] Source hashes remain unchanged in integration tests.
- [ ] Only expected worksheet/calculation parts change.
- [ ] Chart, drawing, image, relationship, theme, and unknown fixture parts remain byte-identical.
- [ ] Shared, array, data-table, dynamic-array, stale-plan, and malformed-package cases fail closed.
- [ ] Formula changes remove caches and request recalculation without claiming it occurred.

## Security

- [ ] ZIP traversal, duplicate entry, symlink, encryption, size/count/ratio, DTD/entity, malformed
  XML, and relationship escape tests pass.
- [ ] Web binding is `127.0.0.1`; upload and session limits are enforced.
- [ ] Temporary output is removed after validation failure and web temporary storage is cleaned.
- [ ] Dependencies and GitHub Actions are reviewed for advisories and supported major versions.
- [ ] CodeQL passes.

## Documentation and distribution

- [ ] README commands match a clean checkout.
- [ ] Rule catalog and limitations match code rather than aspirations.
- [ ] `CHANGELOG.md` has a dated version entry.
- [ ] Build wheel/sdist and inspect their contents.
- [ ] Test the wheel in a fresh virtual environment on Windows, macOS, and Linux.
- [ ] Tag `v0.1.0`, publish signed checksums, and verify `uvx`/`pipx` installation.
