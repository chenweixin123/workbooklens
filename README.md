# WorkbookLens

**Lint, test, diff, and safely repair Excel workbooks—locally, explainably, and without
Microsoft Excel.**

WorkbookLens helps analysts, accountants, operations teams, consultants, and developers answer
four practical questions about an inherited workbook: where it is fragile, why a finding was
raised, exactly what a proposed repair would change, and whether the repaired copy preserved the
rest of the OOXML package.

The v0.1 engine is deterministic. It does not require an AI key, a cloud service, Microsoft Office,
or LibreOffice. It never evaluates formulas, runs VBA, opens embedded objects, or fetches external
links.

## 60-second demo

From a source checkout:

```bash
uv sync --locked
uv run workbooklens demo --out .artifacts/demo
```

The command generates a real, intentionally flawed workbook and then performs the complete
workflow:

```text
.artifacts/demo/
├── before.xlsx                 original generated workbook
├── repair-plan.json            five reviewable safe patch types
├── after.xlsx                  directly patched OOXML copy
├── apply-report.json           source/output hashes and package-change manifest
├── before-report/report.html   interactive scan report
├── after-report/report.html    post-repair scan report
└── diff.html                   filterable semantic before/after diff
```

Open `before-report/report.html`, `repair-plan.json`, and `diff.html` in that order. This is also
the supported way to create screenshots or a GIF for project documentation: record those generated
local assets. The repository intentionally contains no fabricated screenshot.

## Installation

WorkbookLens requires Python 3.11 or newer.

After a package release, run without installing permanently:

```bash
uvx --from workbooklens workbooklens --help
```

Or install as an isolated CLI:

```bash
pipx install workbooklens
workbooklens --help
```

From a source checkout:

```bash
git clone https://github.com/workbooklens/workbooklens.git
cd workbooklens
uv sync --locked
uv run workbooklens --help
```

## CLI workflows

```bash
# Read-only inspection: HTML, JSON, snapshot, and GitHub-compatible SARIF
workbooklens scan INPUT.xlsx --out report/

# Preview deterministic repairs; INPUT.xlsx is not changed
workbooklens plan INPUT.xlsx --out repair-plan.json

# Apply reviewed IDs, or every eligible safe patch, to a new path
workbooklens apply INPUT.xlsx repair-plan.json \
  --patch-id patch-0123456789abcdef --out INPUT.fixed.xlsx
workbooklens apply INPUT.xlsx repair-plan.json --safe-only --out INPUT.fixed.xlsx

# Semantic comparison: writes diff.html and diff.json
workbooklens diff BEFORE.xlsx AFTER.xlsx --out diff.html

# YAML assertions; exits 1 on a failed gate
workbooklens test INPUT.xlsx --config workbooklens.yml

# Local web review UI; always binds to loopback in v0.1
workbooklens serve --port 8765
```

`scan` accepts `.xlsx` and `.xlsm`. `.xlsm` remains read-only: scans suppress patch proposals, and
`plan`/`apply` refuse it so that macro and unknown package content cannot be rewritten accidentally.
`apply` refuses to overwrite either the source or an existing destination.

## What a scan produces

- `report.html`: accessible, self-contained HTML with severity/sheet/rule filters, evidence, peer
  formulas, patch previews, privacy statement, and limitations.
- `findings.json`: stable finding IDs, locations, confidence, evidence, expected invariant, and
  suggested action.
- `snapshot.json`: deterministic values/formulas/styles/structure snapshot.
- `results.sarif`: SARIF 2.1.0 suitable for GitHub code-scanning upload.

No report loads a CDN, font, tracker, or remote script.

## Rule catalog

| ID | Detects | Automatic behavior |
|---|---|---|
| `WL001_BROKEN_REFERENCE` | `#REF!` in formulas | Report; no guess |
| `WL002_FORMULA_PATTERN_OUTLIER` | one structural R1C1-like anomaly in a strong formula consensus | Replace only when independent translations agree and confidence is at least 0.95 |
| `WL003_BLANK_IN_FORMULA_BAND` | one blank interrupting copied formulas | Create an exact translated formula when peers agree |
| `WL004_HARDCODED_VALUE_IN_FORMULA_BAND` | isolated literal inside a formula band | Replace only at very high consensus |
| `WL005_SUSPICIOUS_SUM_BOUNDARY` | simple `SUM` excluding one adjacent numeric peer | Extend conservatively; hidden rows and ambiguous cases suppress it |
| `WL006_NUMERIC_TEXT` | numeric storage mismatch in a numeric region | Convert plain numeric text; preserve leading-zero IDs, long IDs, phone/postal-like values |
| `WL007_STYLE_OUTLIER` | one style/number-format mismatch in a homogeneous region | Copy an existing consensus style ID |
| `WL008_HIDDEN_NONEMPTY_DATA` | hidden rows, columns, or sheets with content | Report only; never unhide |
| `WL009_EXTERNAL_LINK` | formulas or names that refer to another workbook | Report only; never fetch |
| `WL010_VOLATILE_OR_FRAGILE_FUNCTION` | volatile functions and whole-column references | Report only |
| `WL011_ERROR_CELL` | stored Excel error values | Report only |
| `WL012_DUPLICATE_CONFIGURED_KEY` | duplicates in a YAML-declared key column | Report only |
| `WL013_BROKEN_DEFINED_NAME` | names targeting missing sheets, invalid ranges, or `#REF!` | Report only |
| `WL014_MERGED_CELL_IN_DATA_REGION` | merges crossing a dense data body | Report only |
| `WL015_INCONSISTENT_DATA_VALIDATION` | one missing/different validation among homogeneous peers | Report only |

The rules favor false-negative ambiguity over a confident-looking false repair. Findings distinguish
deterministic evidence and confidence from unsupported formula cases.

## Repair plan and safety threshold

A plan is strict versioned JSON. Its essential shape is:

```json
{
  "schema_version": 1,
  "tool_version": "0.1.0",
  "source_name": "INPUT.xlsx",
  "source_sha256": "<full SHA-256>",
  "findings": [{"id": "finding-…", "evidence": {"summary": "…"}}],
  "patches": [
    {
      "id": "patch-<stable digest>",
      "kind": "create_formula",
      "sheet": "Sales",
      "cell": "D8",
      "before": null,
      "after": "=B8*C8",
      "source_cell": "D7",
      "confidence": 0.99,
      "safe": true,
      "description": "...",
      "precondition": {
        "cell_fingerprint": "<full SHA-256>",
        "expected_value": null,
        "expected_formula": null,
        "expected_style_id": 2
      }
    }
  ]
}
```

The embedded patch-linked findings preserve the evidence behind each preview. `apply` verifies the
full source hash and every target fingerprint. It accepts only patches marked
`safe=true` with confidence at least `0.95`, even when an ID is explicitly selected. A stale plan,
unknown patch, unsupported formula, malformed package, unexpected changed part, or failed reopen
causes a closed failure and removes the partial output.

## Direct OOXML preservation

`openpyxl` reads workbooks and generates test/demo inputs, but it never saves a repaired workbook.
The repair engine:

1. validates the untrusted ZIP and relationships;
2. rewrites only selected worksheet XML plus `xl/workbook.xml` calculation metadata when formulas
   change;
3. copies every other package part with byte-identical uncompressed content;
4. records SHA-256 before/after hashes for each changed part;
5. rejects shared, array, data-table, spilled, and dynamic-array formula ranges;
6. reopens the output through both the secure package reader and `openpyxl`;
7. rescans it and rejects newly introduced error/critical findings.

Cached values are removed from changed formula cells after a repair. WorkbookLens sets `calcMode=auto`,
`fullCalcOnLoad=1`, and `forceFullCalc=1`; it does **not** claim that formulas were recalculated.
See [ADR 0001](docs/adr/0001-direct-ooxml-patching.md).

## Workbook tests

Copy `workbooklens.example.yml` and configure only invariants you understand. v0.1 supports:

- finding-count gates and `no_findings`;
- `unique`, `allowed_values`, and `nonblank` ranges;
- numeric lower/upper bounds;
- equality between a direct cell, numeric literal, or simple `SUM(Sheet!A1:A10)`.

Ranges are capped at 100,000 cells. An unsupported expression fails clearly rather than being
partially evaluated.

## Python rule plugins

Plugins are explicit so an ordinary scan never executes unreviewed third-party code:

```python
from pathlib import Path

from workbooklens.rules import RuleRegistry, default_registry
from workbooklens.scanner import scan_workbook

registry: RuleRegistry = default_registry()
registry.register(MyCompanyRule())  # MyCompanyRule implements WorkbookRule.run(context)
result = scan_workbook(Path("book.xlsx"), registry=registry)
```

Packages may also publish a `workbooklens.rules` entry point. Call
`registry.load_entry_points()` only after deciding to trust installed plugins. No scanner-core edit
is required.

## GitHub Action

The composite action detects changed `.xlsx` files, emits reports for each, uploads one immutable
artifact, and can optionally upload SARIF:

```yaml
permissions:
  contents: read
  security-events: write
steps:
  - uses: actions/checkout@v7
    with:
      fetch-depth: 0
  - uses: workbooklens/workbooklens@v0.1.0
    with:
      base-sha: ${{ github.event.pull_request.base.sha || github.event.before }}
      fail-on: error
      upload-sarif: "true"
```

Scanning runs locally on the runner, and workbook files themselves are not uploaded as artifacts.
Generated reports can contain workbook-derived values, formulas, and evidence and are uploaded by
the example workflow; review repository access and artifact-retention settings before enabling it.

## Security boundaries

Before any workbook parser runs, WorkbookLens enforces configurable limits (defaults shown):

- compressed file: 100 MiB;
- ZIP entries: 10,000;
- total uncompressed data: 1,000 MiB;
- one entry: 100 MiB;
- one XML part: 50 MiB;
- per-entry compression ratio: 100:1.

It rejects path traversal, absolute/backslash member paths, duplicate/encrypted/symlink entries,
DTD/entity declarations, malformed XML, missing internal relationship targets, and relationship
targets that escape the package. The web server stores uploads in a process-owned temporary
directory and cleans it when the server stops.

## Honest limitations

- `.xlsx` scan/diff/test/repair and `.xlsm` scan/diff/test are supported; legacy `.xls`, `.xlsb`,
  `.ods`, and Google Sheets are not.
- WorkbookLens does not implement an Excel calculation engine or execute VBA.
- Structured references can be scanned but are not translated for repair.
- Shared, array, data-table, and dynamic-array formulas are not automatically patched.
- Region inference is conservative and may miss sparse or unusually formatted tables.
- `SUM(range)` is the only formula expression evaluated by YAML assertions; cached formula values
  may be stale.
- OOXML is large and extensible. The package-preservation design minimizes changes but does not
  claim complete Excel compatibility.

## Development

```bash
uv sync --locked
uv run ruff format --check .
uv run ruff check .
uv run mypy src
uv run pytest -q
uv run workbooklens demo --out .artifacts/demo
```

The test suite generates workbooks programmatically, covers malicious package inputs, and enforces
75% branch-aware coverage (the v0.1 suite currently exceeds that threshold). See
[`CONTRIBUTING.md`](CONTRIBUTING.md), [`SECURITY.md`](SECURITY.md), and the
[`architecture guide`](docs/architecture.md).

## CLI exit codes

| Code | Meaning |
|---:|---|
| 0 | success |
| 1 | configured finding threshold or workbook assertion failed |
| 2 | unsupported operation or invalid usage |
| 3 | unsafe/malformed workbook input |
| 4 | stale source or cell precondition |
| 5 | patched output validation failed |
| 10 | unexpected internal failure |

## Roadmap

1. Broader preservation fixtures for pivots, slicers, rich data types, and vendor extensions.
2. More conservative table/structured-reference analysis without widening automatic repair.
3. Signed plugin metadata and a rule SDK compatibility test kit.
4. Streaming report generation for very large but valid workbooks.
5. Optional explanation adapters that receive redacted findings only and remain disabled by default.

Licensed under Apache-2.0.
