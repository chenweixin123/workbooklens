# WorkbookLens

**Deterministic linting, regression testing, semantic diffing, and conservative repair for
Excel workbooks.**

WorkbookLens 2.2 works locally without Microsoft Excel, LibreOffice, an AI key, or a cloud service.
It does not calculate formulas, execute VBA, open embedded objects, or fetch external links.
.xlsx files support scan, test, diff, and safe-copy repair; .xlsm files remain read-only.

> **Release status:** GitHub Releases are authoritative for source archives and attached artifacts.
> Version 2.2.1 may not be published to [PyPI](https://pypi.org/project/workbooklens/); use the
> downloaded wheel or source checkout instructions below unless PyPI explicitly lists that version.
> Do not assume pipx or uvx can install a GitHub-only release by package name.

## What is new in 2.2

- Layout-aware snapshots and semantic diffs now record declared/content dimensions, explicit row
  heights, column widths, and saved worksheet views.
- `WL016_TEXT_DISPLAY_RISK` detects likely clipped wrapped text and blocked horizontal overflow. It
  can propose bounded row-height, wrap, or repeated-column-width changes.
- `WL017_BORDER_EDGE_INCONSISTENCY` treats either side of a shared cell edge as visually present,
  reducing false positives. It also detects fully borderless holes and missing outer edges inside a
  dense rectangular table, and proposes only edges supported by at least 95% peer consensus.
- `WL018_USED_RANGE_INFLATION` identifies separated, format-only tails and can remove only the
  enumerated blank styled cells and empty row records after reference and structure checks.
- `WL019_IDENTIFIER_SCIENTIFIC_NOTATION` detects long numeric identifiers under identifier-like
  headers. Ten- and eleven-digit values can receive a font-aware width-only proposal that preserves
  the stored numeric value and type. General-formatted values of 12-15 digits remain findings-only
  because Excel can keep scientific notation even in a wide column; longer values may already have
  lost precision.
- `WL020_SAVED_VIEW_OFF_CONTENT` detects sheets saved with their first content scrolled away or at a
  zoom unlikely to show a compact sheet's full content width and height in a typical desktop window.
  A reviewed repair can restore the first visible content cell and lower the saved zoom only when the
  saved zoom actually exceeds the two-dimensional estimated fit. Merged content contributes its full
  extent, visible border/fill templates count toward the layout, and earlier width/height proposals are
  included in the estimate. A pure zoom repair may preserve an unshifted frozen pane byte-for-byte;
  shifted frozen panes and split panes remain findings-only.
- `WL021_WHITESPACE_ONLY_TAIL` detects connected literal-space cells outside the visible layout
  envelope. Its reviewed cleanup deletes only default-style nodes; styled cells are cleared while
  their font, alignment, protection, style ID, and custom blank-row heights are preserved. Referenced
  or structurally significant targets are refused.
- Single-row merged titles are treated as bounded display regions, same-column width requests are
  coalesced to the largest sufficient proposal, and border repair is limited to genuine internal
  shared edges with materialized peers.
- Layout-changing proposals are labeled `layout_review`, excluded from `--safe-only`, and require
  explicit patch selection plus `--accept-layout-risk`. Related wrap/height changes form atomic
  groups and cannot be applied partially.

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

## Windows portable ZIP

The official `WorkbookLens-2.2.1-windows-x64-portable.zip` is the simplest local Windows option.
It bundles a 64-bit CPython 3.12 runtime, so users do not need to install Python, `uv`, Microsoft
Excel, LibreOffice, an AI key, or a cloud client. It is a portable folder rather than an installer:
it does not request administrator access, modify the registry, create file associations, or add an
automatic updater.

Download the ZIP and `SHA256SUMS` from the official
[GitHub Release](https://github.com/chenweixin123/workbooklens/releases/tag/v2.2.1), compare the
published SHA-256 value, and extract the complete folder. In PowerShell:

~~~powershell
(Get-FileHash .\WorkbookLens-2.2.1-windows-x64-portable.zip -Algorithm SHA256).Hash
Expand-Archive .\WorkbookLens-2.2.1-windows-x64-portable.zip -DestinationPath .\WorkbookLens
Set-Location .\WorkbookLens\WorkbookLens-2.2.1-windows-x64
.\Start-WorkbookLens.cmd
~~~

The launcher opens a dedicated WorkbookLens console, starts only on `127.0.0.1`, waits for the
health check, and opens the default browser. It prefers port 8765 and selects an available local
port if that port is already in use. Press `Ctrl+C` in the WorkbookLens console to stop the server
cleanly. Workbook files and generated reports stay on the computer; the portable runtime does not
upload them.

The executable is currently unsigned, so Microsoft SmartScreen may display a warning. Verify the
checksum and download only from the official release. The portable build guarantees the built-in
rules; install the wheel in a normal Python environment when third-party entry-point plugins are
required.

## Install the downloaded wheel

The `.whl` attached to the GitHub Release is a local Python application package, not a hosted
upload service. Workbook processing stays on the user's computer. Installing the wheel may contact
the configured Python package index once to obtain dependencies; scanning, planning, repair, diff,
and the loopback web UI do not upload workbooks.

Download `workbooklens-2.2.1-py3-none-any.whl`, `workbooklens-2.2.1.tar.gz`, and
`SHA256SUMS` from the official
[GitHub Release](https://github.com/chenweixin123/workbooklens/releases/tag/v2.2.1), verify the
checksum, and install it with Python 3.11+ and `uv`.

### PowerShell

~~~powershell
Get-Content .\SHA256SUMS
(Get-FileHash .\workbooklens-2.2.1-py3-none-any.whl -Algorithm SHA256).Hash
uv tool install .\workbooklens-2.2.1-py3-none-any.whl
workbooklens --version
workbooklens serve --port 8765
~~~

### Bash

~~~bash
sha256sum -c SHA256SUMS
uv tool install ./workbooklens-2.2.1-py3-none-any.whl
workbooklens --version
workbooklens serve --port 8765
~~~

Then open `http://127.0.0.1:8765/`. The browser page and temporary upload directory are local to
that process. Sensitive workbooks should use the CLI or this loopback UI rather than a hosted CI
runner.

The wheel and portable ZIP are two distributions of the same local application. The wheel is the
extensible option for Python users; the portable ZIP is the no-install option for Windows users.

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

If PyPI lists version 2.2.1, isolated installation is:

~~~bash
uvx --from workbooklens==2.2.1 workbooklens --help
pipx install workbooklens==2.2.1
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

# Apply an explicitly reviewed layout patch or atomic layout group
workbooklens apply INPUT.xlsx repair-plan.json --patch-id patch-fedcba9876543210 --accept-layout-risk --out INPUT.layout-fixed.xlsx

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

After the v2.2.1 tag exists:

~~~yaml
permissions:
  contents: read
  security-events: write

steps:
  - uses: actions/checkout@v7
    with:
      fetch-depth: 0

  - uses: chenweixin123/workbooklens@v2.2.1
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
- uses: chenweixin123/workbooklens@v2.2.1
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

Repair plans bind every operation to the complete source SHA-256 and target preconditions. Before
writing, the engine rescans the source and requires the serialized plan to match the canonical plan
field-for-field. Ordinary operations execute only when marked safe with confidence at least 0.95.
Layout-changing operations are always `safe=false` with risk `layout_review`; `--safe-only` cannot
select them. They require one or more explicit `--patch-id` selections, confidence at least 0.95,
and `--accept-layout-risk`. This opt-in never authorizes any other unsafe operation. Selecting one
member of an atomic group selects the complete group.

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
Layout repairs additionally verify row, column, view, or exact-tail fingerprints. Formatting-tail
cleanup fails closed when an authorized cell or row intersects formulas, names, table or validation
ranges, comments, links, page breaks, drawing anchors, or other guarded worksheet structures.

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
- Text width and row-height estimates are deterministic approximations, not Excel's rendering
  engine. Results can vary with fonts, printer metrics, DPI, locale, and application version; review
  layout changes in the target spreadsheet application.
- A saved zoom reset is a conservative opening view, not a guarantee that every column fits every
  screen. Frozen or split panes are preserved and block automatic saved-view repair.
- Border repair requires strong local peer consensus and changes one reviewed edge only; irregular
  tables can remain findings-only or undetected.
- Format-tail cleanup ignores broad column styling by itself and refuses ambiguous references or
  structures. It is not a general "reset UsedRange" command.
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
python scripts/check_release_artifacts.py dist --version 2.2.1
~~~

To reproduce the Windows x64 portable release candidate from the built wheel, run the following in
PowerShell on 64-bit Windows. The final smoke test executes the frozen program, including the CLI,
loopback web UI, occupied-port fallback, clean shutdown, and installation-directory hash check.

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
~~~

See [CONTRIBUTING.md](CONTRIBUTING.md), [the architecture guide](docs/architecture.md),
[CLI exit codes](docs/cli-exit-codes.md), and [the release checklist](docs/release-checklist.md).

Licensed under Apache-2.0.
