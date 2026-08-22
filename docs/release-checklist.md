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
- [ ] Windows CI and release jobs compile an `Output=no` probe and accept only the ISCC compiler
  banner version exactly 6.7.0; record the `CompilerEngineVersion` emitted by the workflow.
- [ ] uv build --out-dir dist and twine check --strict dist/* pass.
- [ ] scripts/check_release_artifacts.py passes for the release version.
- [ ] Unpacked wheel/sdist contain no environment, cache, bytecode, environment file, credential,
  private key, or unexpected top-level path.
- [ ] The Windows portable ZIP passes traversal, duplicate/case-collision, symlink, encryption,
  compression, expanded-size, sensitive-file, PE x64, and version checks.
- [ ] The Windows setup executable passes filename, size, PE bootstrap, version-resource, and
  exact SHA-256 sidecar checks.

## Installation

- [ ] The exact wheel is installed into fresh Linux, Windows, and macOS environments.
- [ ] workbooklens --version, --help, and demo run from the wheel rather than the source tree.
- [ ] The exact portable ZIP runs from a path containing spaces and Chinese characters with
  `PYTHONHOME` and `PYTHONPATH` cleared and no Python directory on `PATH`.
- [ ] The frozen executable runs scan, diff, plan/apply, and report-template workflows; the real
  server answers `/health` and `/`, listens only on `127.0.0.1`, exits cleanly, and releases its port.
- [ ] `Start-WorkbookLens.cmd` opens the browser only after readiness and falls back when port 8765
  is occupied. The console documents `Ctrl+C` as the clean stop mechanism.
- [ ] The exact setup executable installs for the current user, preserves every portable payload
  hash, creates the WorkbookLens Start-menu and selected desktop shortcuts, registers the correct
  Installed apps version, rejects forged `HKCU` uninstall paths without changing their files,
  writes and validates the installer ownership marker, rejects reparse points in the install tree,
  runs without Python on `PATH`, and uninstalls without residual test state.
- [ ] When the immediately previous release has Windows installer and portable artifacts, run
  `scripts/smoke_installer_windows.py` with both `--previous-installer` and
  `--previous-portable-zip` and record the exact previous artifact hashes. Version 2.2.1 is the
  first Windows-installer release, so the published 2.2.0 wheel/sdist-only assets cannot supply this
  baseline; do not substitute an unpinned or same-version download.
- [ ] README Bash and PowerShell commands match a clean checkout.
- [ ] Representative layout repairs are opened in the target spreadsheet application; text is not
  clipped, identifiers are exact, the initial viewport is useful, and printer/page layout is reviewed.

## Release candidate and publication

- [ ] Push v2.2.1; the release-candidate workflow validates the tag, builds distributions,
  generates SHA256SUMS, and uploads workflow artifacts.
- [ ] Inspect the downloaded artifacts before external publication.
- [ ] Record whether the release is GitHub-only or also publishes to PyPI.
- [ ] Attach the validated wheel, sdist, Windows x64 setup executable, portable ZIP, and checksums
  to the GitHub Release. Attach a provenance attestation only when the release workflow produces one.
- [ ] If publishing to PyPI, configure the GitHub pypi environment and trusted publisher for
  chenweixin123/workbooklens, then publish only from the validated tag using OIDC; do not add a
  long-lived PyPI token.
- [ ] If publishing to PyPI, verify the project page, uvx installation, and pipx installation for
  version 2.2.1.
