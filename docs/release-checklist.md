# Release checklist

## Version and contracts

- [ ] pyproject.toml, src/workbooklens/__init__.py, uv.lock, README, and Action examples use the
  same version.
- [ ] The tag is exactly v plus the project version.
- [ ] CHANGELOG.md has a dated entry and identifies schema or compatibility changes.
- [ ] JSON schemas, source-scoped baseline contracts, CLI exit codes, Action inputs/outputs, and
  plugin contracts are reviewed.

## Correctness and preservation

- [ ] uv lock --check, format, lint, strict typing, and the complete test suite pass.
- [ ] Demo, scan, baseline/new-only scan, test, plan/apply, and diff commands pass.
- [ ] Source hashes remain unchanged and only expected OOXML parts change.
- [ ] Charts, drawings, images, relationships, themes, and unknown fixtures remain byte-identical.
- [ ] Shared, array, data-table, dynamic-array, stale-plan, and malformed inputs fail closed.
- [ ] Formula edits remove caches and request recalculation without claiming it occurred.
- [ ] `--safe-only` excludes every `layout_review` patch; explicit layout selection requires
  `--accept-layout-risk`, confidence at least 0.95, and complete atomic groups.
- [ ] Row heights, column widths, wrapping, width-only identifier display, saved views, edge-only borders, and
  exact format-tail cleanup are represented accurately in scan, apply, rescan, and semantic diff.
- [ ] Format-tail cleanup refuses formulas, names, tables, guarded ranges, links, comments, page
  breaks, drawings, hidden/outlined rows, and unsupported row metadata.

## Security and supply chain

- [ ] Adversarial ZIP/XML/relationship and web-limit tests pass.
- [ ] CodeQL, dependency review, and locked-runtime vulnerability audit pass.
- [ ] GitHub Actions and dependencies have reviewed updates.
- [ ] uv build --out-dir dist and twine check --strict dist/* pass.
- [ ] scripts/check_release_artifacts.py passes for the release version.
- [ ] Unpacked wheel/sdist contain no environment, cache, bytecode, environment file, credential,
  private key, or unexpected top-level path.

## Installation

- [ ] The exact wheel is installed into fresh Linux, Windows, and macOS environments.
- [ ] workbooklens --version, --help, and demo run from the wheel rather than the source tree.
- [ ] README Bash and PowerShell commands match a clean checkout.
- [ ] Representative layout repairs are opened in the target spreadsheet application; text is not
  clipped, identifiers are exact, the initial viewport is useful, and printer/page layout is reviewed.

## Release candidate and publication

- [ ] Push v2.2.0; the release-candidate workflow validates the tag, builds distributions,
  generates SHA256SUMS, and uploads workflow artifacts.
- [ ] Inspect the downloaded artifacts before external publication.
- [ ] Record whether the release is GitHub-only or also publishes to PyPI.
- [ ] Attach the validated wheel, sdist, and checksums to the GitHub Release. Attach a provenance
  attestation only when the release workflow produces one.
- [ ] If publishing to PyPI, configure the GitHub pypi environment and trusted publisher for
  chenweixin123/workbooklens, then publish only from the validated tag using OIDC; do not add a
  long-lived PyPI token.
- [ ] If publishing to PyPI, verify the project page, uvx installation, and pipx installation for
  version 2.2.0.
