# Contributing

WorkbookLens values source preservation, explainable evidence, deterministic behavior, and tests
above feature count.

## Setup

~~~bash
git clone https://github.com/chenweixin123/workbooklens.git
cd workbooklens
uv sync --locked
uv run workbooklens demo --out .artifacts/demo
~~~

PowerShell uses the same commands after Set-Location workbooklens.

Do not add Office, LibreOffice, network calls, or an LLM to the default execution path. Public
models and schemas must be typed and versioned.

## Rules and repairs

For a rule, allocate a stable WLxxx_UPPER_SNAKE_CASE ID and document the invariant, observed
evidence, peers, confidence basis, ambiguity, and suggested action. Generate fixtures in Python.
Test positive, negative, boundary, hidden/merged, mixed-reference, and unsupported cases.

Only propose a patch when independent deterministic evidence establishes the output. Safe patches
require confidence at least 0.95. Direct OOXML code—not openpyxl.save()—must implement writes.
Tests must prove the source hash is unchanged, the exact expected parts changed, unrelated package
parts remain byte-identical, the output reopens, and stale or advanced-formula cases fail closed.
WorkbookLens never deletes user data or overwrites the source.

## Public contracts

- Do not recycle finding or patch IDs.
- Changing JSON shape requires a schema-version decision and migration note.
- CLI exit codes, Action inputs/outputs, and plugin entry points are compatibility surfaces.
- Baseline behavior is based on stable finding IDs; do not silently suppress a rule by title or
  location alone.

## Quality gates

~~~bash
uv lock --check
uv run ruff format --check .
uv run ruff check .
uv run mypy src
uv run python -m pytest -q
uv run workbooklens demo --out .artifacts/demo
uv build --out-dir dist
uvx --from twine twine check --strict dist/*
python scripts/check_release_artifacts.py dist --version 2.2.1
~~~

On 64-bit Windows, reproduce and exercise the portable artifact from the built wheel with:

~~~powershell
uv sync --locked --python 3.12
New-Item -ItemType Directory -Force -Path .artifacts | Out-Null
uv build --out-dir .artifacts/portable-python-dist
uv export --locked --python 3.12 --no-dev --group portable --no-emit-project --no-hashes --no-header --no-annotate --output-file .artifacts/portable-constraints.txt
$wheel = @(Get-ChildItem -LiteralPath .artifacts/portable-python-dist -Filter '*.whl')
if ($wheel.Count -ne 1) { throw "Expected exactly one wheel, found $($wheel.Count)." }
$python = (Resolve-Path .venv/Scripts/python.exe).Path
& $python scripts/build_portable_windows.py --wheel $wheel[0].FullName --python $python --expected-version 2.2.1 --constraints .artifacts/portable-constraints.txt --output-dir .artifacts/portable-dist
$archive = @(Get-ChildItem -LiteralPath .artifacts/portable-dist -Filter '*.zip')
if ($archive.Count -ne 1) { throw "Expected exactly one ZIP, found $($archive.Count)." }
& $python scripts/check_portable_artifact.py $archive[0].FullName --expected-version 2.2.1
& $python scripts/smoke_portable.py $archive[0].FullName --expected-version 2.2.1
$iscc = (Resolve-Path "$env:LOCALAPPDATA/Programs/Inno Setup 6/ISCC.exe").Path
& $python scripts/build_installer_windows.py --portable-zip $archive[0].FullName --expected-version 2.2.1 --iscc $iscc --output-dir .artifacts/installer-dist
$installer = @(Get-ChildItem -LiteralPath .artifacts/installer-dist -Filter '*.exe')
if ($installer.Count -ne 1) { throw "Expected exactly one installer, found $($installer.Count)." }
& $python scripts/check_installer_artifact.py $installer[0].FullName --expected-version 2.2.1
& $python scripts/smoke_installer_windows.py $installer[0].FullName --portable-zip $archive[0].FullName --expected-version 2.2.1
~~~

The installer smoke test is intentionally destructive only to its own temporary installation. It
refuses to run when it detects an existing WorkbookLens installation or shortcut, verifies the
installed payload and shell integration, and requires the uninstaller to remove all test-created
state.

Review for unsafe ZIP handling, unbounded memory/ranges, inaccurate documentation, source rewrites,
missing cleanup, schema drift, and release-artifact leakage. Security issues follow
[SECURITY.md](SECURITY.md), never a public proof-of-concept issue.
