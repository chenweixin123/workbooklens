# WorkbookLens

**Deterministic linting, regression testing, semantic diffing, and conservative repair for
Excel workbooks.**

WorkbookLens 2.1 works locally without Microsoft Excel, LibreOffice, an AI key, or a cloud service.
It does not calculate formulas, execute VBA, open embedded objects, or fetch external links.
.xlsx files support scan, test, diff, and safe-copy repair; .xlsm files remain read-only.

> **Release status:** GitHub Releases are authoritative for source archives and attached artifacts.
> Version 2.1.0 may not be published to [PyPI](https://pypi.org/project/workbooklens/); use the
> source-install instructions below unless PyPI explicitly lists that version. Do not assume pipx
> or uvx can install a GitHub-only release.

## What is new in 2.1

- Chartsheet workbooks scan and repair their ordinary worksheets without losing chart parts.
- Numeric text is auto-converted only under an explicit measure header such as Amount, Price,
  Units, Balance, 金额, or 数量. Unknown columns, identifiers, and explicitly text-formatted cells
  remain review-only.
- Merged ranges, summary/subtotal rows, non-visible worksheets, hidden rows or columns (including
  grouped columns), and protected worksheets remain review-only for automatic repair.
- Formula and style patches also require stable detail-row evidence. Unknown summary labels,
  secondary override/adjustment labels, intentionally highlighted rows, and free-form-only row
  labels are findings-only when ordinary row semantics cannot be established conservatively.
- Style repair preserves number format, protection, quote-prefix, and pivot-button semantics.
  Multiple isolated anomalies are reported without bulk auto-repair.
- Suspicious SUM boundaries now report the candidate formula without an automatic patch because
  adjacency alone cannot prove that a tax, adjustment, subtotal, or statistic belongs in the SUM.

## What was new in 2.0

- Baseline-aware scans accept a source-scoped findings.json or an aggregate baseline manifest.
- The --new-only gate uses stable rule/location identities; evidence changes are tracked separately.
- findings.json schema version 2 records total, active, suppressed, known, and new counts.
- The GitHub Action supports mode scan or test, YAML config, baselines, and .xlsx/.xlsm.
- Release builds use explicit wheel/sdist allowlists and reject environments, caches, credentials,
  private keys, bytecode, unexpected roots, or oversized members.
- CI tests Python 3.11–3.14 on Linux, Windows, and macOS, then installs the built wheel in fresh
  environments.

## Install from this source checkout

WorkbookLens requires Python 3.11 or newer and [uv](https://docs.astral.sh/uv/).

### Bash

~~~bash
git clone https://github.com/chenweixin123/workbooklens.git
cd workbooklens
uv sync --locked
uv run workbooklens --version
uv run workbooklens demo --out .artifacts/demo
~~~

### PowerShell

~~~powershell
git clone https://github.com/chenweixin123/workbooklens.git
Set-Location workbooklens
uv sync --locked
uv run workbooklens --version
uv run workbooklens demo --out .artifacts/demo
~~~

If PyPI lists version 2.1.0, isolated installation is:

~~~bash
uvx --from workbooklens==2.1.0 workbooklens --help
pipx install workbooklens==2.1.0
~~~

## Core workflows

~~~bash
# Self-contained HTML, JSON, snapshot, and SARIF
workbooklens scan INPUT.xlsx --out report/

# Optional YAML rule configuration
workbooklens scan INPUT.xlsx --config workbooklens.yml --out report/

# Gate only findings absent from an earlier report
workbooklens scan INPUT.xlsx --baseline previous/findings.json --new-only --fail-on warning --out report/

# Pin a portable logical path when invoking from varying working directories
workbooklens scan INPUT.xlsx --source-scope models/INPUT.xlsx --baseline previous/findings.json --new-only --out report/

# Preview and apply source-bound deterministic repairs to a new file
workbooklens plan INPUT.xlsx --config workbooklens.yml --out repair-plan.json
workbooklens apply INPUT.xlsx repair-plan.json --patch-id patch-0123456789abcdef --out INPUT.fixed.xlsx
workbooklens apply INPUT.xlsx repair-plan.json --safe-only --out INPUT.fixed.xlsx

# Semantic before/after comparison
workbooklens diff BEFORE.xlsx AFTER.xlsx --out diff.html

# Bounded YAML assertions and finding-count gates
workbooklens test INPUT.xlsx --config workbooklens.yml --out test-results.json

# Loopback-only local review UI
workbooklens serve --port 8765
~~~

PowerShell uses the same arguments:

~~~powershell
workbooklens scan INPUT.xlsx --baseline previous\findings.json --new-only --fail-on warning --out report
~~~

--new-only requires --baseline. A findings.json baseline is bound to its source_scope (normally the
input path relative to the current directory), so a report for one workbook cannot silently hide a
finding in another. Without --new-only, reports retain all current findings while identifying which
were already known. Reusing a --new-only report preserves both its new and previously known IDs.

For a repository containing several workbooks, use an aggregate manifest keyed by repository-
relative POSIX paths:

~~~json
{
  "workbooks": {
    "workbooks/model-a.xlsx": ["finding-0123456789abcdef"],
    "workbooks/model-b.xlsx": ["finding-fedcba9876543210"]
  }
}
~~~

A top-level `finding_ids` array is deliberately unscoped and authoritative. Use that form only when
the same global IDs are intended to apply to every scanned workbook. Suppressions are documented
waivers in version-2 YAML and are never imported into a baseline when they appear in a report.

## Scan outputs

Each scan directory contains:

- report.html: self-contained, filterable review report;
- findings.json: schema-v2 findings, source_scope, stable IDs, content fingerprints, baseline
  metadata, and summary counts;
- snapshot.json: deterministic workbook structure/value/formula/style snapshot;
- results.sarif: SARIF 2.1.0 for GitHub code scanning.

Reports can contain workbook values, formulas, sheet names, and evidence. Treat them with the same
confidentiality as the input workbook.

## GitHub Action

After the v2.1.0 tag exists:

~~~yaml
permissions:
  contents: read
  security-events: write

steps:
  - uses: actions/checkout@v7
    with:
      fetch-depth: 0

  - uses: chenweixin123/workbooklens@v2.1.0
    with:
      mode: scan
      path: workbooks
      config: workbooklens.yml
      baseline: baselines/workbooks.json
      new-only: "true"
      fail-on: warning
      upload-sarif: "true"
~~~

Assertion mode requires config and deliberately rejects baseline/new-only:

~~~yaml
- uses: chenweixin123/workbooklens@v2.1.0
  with:
    mode: test
    path: workbooks
    config: workbooklens.yml
~~~

The Action detects changed tracked .xlsx and .xlsm files and supplies each repository-relative path
as source_scope. For multi-file scans, use the aggregate `workbooks` baseline form shown above. Set
fetch-depth to 0 when using base-sha.
Generated reports are uploaded for 14 days by default; workbook files themselves are not uploaded
by the Action.

## Repair safety

Repair plans bind every operation to the complete source SHA-256 and a target-cell fingerprint.
Before writing, the engine rescans the source and requires the serialized plan to match the canonical
plan field-for-field; only canonical operations marked safe with confidence at least 0.95 execute.
Automatic operations are withheld for merged targets, summary rows, non-visible sheets, hidden rows
or columns, and protected worksheets. The low-level patcher independently rechecks protection and
visibility so an edited or stale plan cannot bypass the scanner's safety boundary. Numeric conversion
requires an explicit measure-column signal; an unknown header is not treated as permission to change
stored text into a number. Style-copy operations must preserve number-format, cell-protection,
quote-prefix, and pivot-button semantics.

For formula and style operations, the target row must also agree with stable peer-row text and visual
patterns. A secondary adjustment label, a unique note, a whole-row highlight, or only free-form labels
without a dominant template causes automatic repair to be withheld.

The engine writes a new OOXML package directly, verifies the exact changed-part allowlist, reopens
the result, rescans it, and removes partial output after validation failure. Formula edits remove
stale caches and request Excel recalculation; WorkbookLens never claims to have calculated the result.

## Security boundaries

Before parsing, WorkbookLens validates ZIP paths, duplicates, encryption, symlinks, compression
ratios, entry and package sizes, required parts, XML size, DTD/entity declarations, and internal
relationships. Default limits include a 100 MiB compressed file, 10,000 entries, 1 GiB total
uncompressed data, 100 MiB per entry, and 50 MiB per XML part.

See [SECURITY.md](SECURITY.md) for private reporting. Never attach a confidential production
workbook to a public issue.

## Honest limitations

- No Excel calculation engine, VBA execution, external-link fetching, or embedded-object opening.
- No automatic repair for shared, array, data-table, spilled, or dynamic-array formulas.
- Suspicious SUM boundaries are findings-only; the expected formula is evidence for human review,
  not proof that the adjacent row belongs in the aggregate.
- Structured references can be scanned but are not translated for repair.
- Region inference is conservative and can miss sparse or unusually formatted tables.
- Rule and region detection are deterministic heuristics. Passing a scan does not prove that every
  real-world workbook error has been found, and a finding does not by itself prove the proposed
  business meaning.
- .xls, .xlsb, .ods, and Google Sheets are unsupported.
- OOXML is extensible; preservation tests reduce risk but do not prove compatibility with every
  vendor extension.

## Development

~~~bash
uv sync --locked
uv run ruff format --check .
uv run ruff check .
uv run mypy src
uv run python -m pytest -q
uv build --out-dir dist
uvx --from twine twine check --strict dist/*
python scripts/check_release_artifacts.py dist --version 2.1.0
~~~

See [CONTRIBUTING.md](CONTRIBUTING.md), [the architecture guide](docs/architecture.md),
[CLI exit codes](docs/cli-exit-codes.md), and [the release checklist](docs/release-checklist.md).

Licensed under Apache-2.0.
