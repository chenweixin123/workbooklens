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
uv run pytest -q
uv run workbooklens demo --out .artifacts/demo
uv build --out-dir dist
uvx --from twine twine check dist/*
python scripts/check_release_artifacts.py dist --version 2.0.0
~~~

Review for unsafe ZIP handling, unbounded memory/ranges, inaccurate documentation, source rewrites,
missing cleanup, schema drift, and release-artifact leakage. Security issues follow
[SECURITY.md](SECURITY.md), never a public proof-of-concept issue.
