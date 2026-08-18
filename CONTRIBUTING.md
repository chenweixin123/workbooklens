# Contributing

WorkbookLens values source preservation, explainable evidence, and tests above feature count.

## Setup

Use Python 3.11 or newer:

```bash
uv sync --locked
uv run workbooklens demo --out .artifacts/demo
```

Do not add Node or an Office/LibreOffice dependency to the default path. Keep public models typed
and documented. Core behavior must remain deterministic without an LLM or network call.

## Adding a rule

1. Implement `WorkbookRule.run(context) -> RuleResult` in a focused module.
2. Allocate a stable `WLxxx_UPPER_SNAKE_CASE` ID; never recycle a released ID.
3. State the invariant, observed evidence, peers, confidence basis, and actionable explanation.
4. Suppress ambiguous/decorative regions rather than broadening a heuristic for coverage.
5. Generate the workbook fixture in Python; avoid unexplained binary fixtures.
6. Add the rule to the built-in registry or publish the documented `workbooklens.rules` entry point.
7. Test positive, negative, boundary, hidden/merged, absolute/mixed-reference, and unsupported cases.

Only propose a patch when independent deterministic evidence proves the output. A safe patch must
have confidence at least 0.95. Report-only rules are often the right design.

## Adding a patch type

Patch operations are declarative Pydantic models with source/cell preconditions. The direct OOXML
engine—not `openpyxl.save()`—must implement writes. Add tests that prove:

- the input hash never changes;
- only the exact expected package parts change;
- unrelated chart/drawing/image/theme/custom entries remain byte-identical;
- the output reopens and has the requested semantic value;
- stale, shared/array/dynamic, malformed, and unexpected-part cases fail closed;
- formula caches and calculation metadata are handled honestly.

Never delete user data in v0.1.

## Quality gates

```bash
uv run ruff format --check .
uv run ruff check .
uv run mypy src
uv run pytest -q
uv run workbooklens demo --out .artifacts/demo
uv run workbooklens scan .artifacts/demo/before.xlsx --out .artifacts/scan
uv run workbooklens diff .artifacts/demo/before.xlsx .artifacts/demo/after.xlsx --out .artifacts/diff.html
```

Review the final diff for unsafe ZIP handling, unbounded ranges/memory, inaccurate documentation,
dead code, accidental source rewrites, and missing cleanup. Security issues follow `SECURITY.md`.
