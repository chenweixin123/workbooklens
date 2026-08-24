# WorkbookLens

**Deterministic linting, regression testing, semantic diffing, and conservative repair for
Excel workbooks.**

WorkbookLens 2.4 core scan, test, diff, and repair workflows work locally without Microsoft Excel,
LibreOffice, an AI key, or a cloud service. Those normal OOXML workflows do not calculate formulas,
execute VBA, open embedded objects, or fetch external links. .xlsx files support scan, test, diff,
and safe-copy repair; .xlsm files remain read-only. The optional local `.xls` conversion button is a
separate trust boundary and should be used only with trusted files: Microsoft Excel (preferred) or
LibreOffice opens the workbook locally and may recalculate formulas or process workbook-defined
behavior supported by that application. Conversion never uses a cloud service and does not claim
that every legacy workbook feature can be preserved.

> **Release status:** GitHub Releases are authoritative for source archives and attached artifacts.
> Version 2.4.0 may not be published to [PyPI](https://pypi.org/project/workbooklens/); use the
> downloaded wheel or source checkout instructions below unless PyPI explicitly lists that version.
> Do not assume pipx or uvx can install a GitHub-only release by package name.

## What is new in 2.4

- Formula analysis now profiles the full inferred data body instead of relying only on short
  contiguous bands. It reports multiple copied-formula outliers, hardcoded interruptions,
  formula-looking text, bounded static circular dependencies, and a small set of provable formula
  errors without evaluating workbook formulas. Saved OOXML formula-error caches are advisory,
  lower-confidence evidence and may be stale. Multiple anomalies remain findings-only; automatic
  formula replacement is limited to a single high-confidence anomaly with stable detail-row
  evidence.
- New report-only data-quality rules infer likely identifier columns, mixed numeric storage, robust
  numeric outliers, percentage-scale mistakes, date-storage anomalies, and sign-domain violations
  (negative values under nonnegative headers, or zero/negative values under positive headers). They
  never invent replacement business values.
- New report-only structure rules check partial AutoFilter coverage, chart source-shape mismatches,
  freeze panes saved deep inside or beyond content, and print areas that truncate inferred tables or
  included chart frames.
- A conservative Formula IR now checks aggregate coverage, cross-sheet targets, notes-column
  formulas, algebraic degeneracy, and label/number-format role conflicts without evaluating or
  rewriting business formulas.
- Optional Workbook Profiles add user-owned semantics for required fields, enumerations, contacts,
  fixed-width identifiers, and number-format roles. Without a Profile, rules continue to require
  strong structural evidence and do not invent allowed-value sets.
- Layout geometry now reports drawings that cover non-source data or KPI cells, fixed-format numeric
  width risks, unexplained row-height outliers, and role-aware title/header/body/total style
  fragmentation. These checks are report-only.
- Isolated extreme long text can receive one atomic bounded-width, wrap, and row-height proposal.
  Repeated extreme rows remain findings-only so a shared width dependency cannot be selected
  partially, and generated row-height proposals are capped at a readable size.

## What is new in 2.3

- The Windows installer and portable package now launch `WorkbookLens.exe` as a native, branded
  desktop window. The local FastAPI service remains bound to `127.0.0.1`, but no console or external
  browser is shown; closing the window stops the background service and releases its port.
- A persistent language selector supports Chinese and English. Built-in pages, findings, repair
  suggestions, scan reports, semantic-diff reports, and safe error messages follow the selected
  language while canonical rule IDs, patch IDs, fingerprints, formulas, and JSON/SARIF data remain
  unchanged. Third-party plugin text is retained when no trusted translation is available.
- Local `.xls` conversion tolerates Windows or Office configurations that deny access to
  `Process.Path`. WorkbookLens verifies the Excel process with stronger held-handle identity checks
  and reports stable `WL-CNV-001` through `WL-CNV-005` codes plus a diagnostic ID instead of exposing
  PowerShell CLIXML, private paths, or tracebacks in the interface.
- Windows artifacts contain a windowed `WorkbookLens.exe`, a separate console
  `WorkbookLensCLI.exe`, and the WorkbookLens application icon. The CLI and Python package workflows
  remain available for automation and third-party plugins.

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

## Windows installer (recommended)

For the normal Windows experience, download
`WorkbookLens-2.4.0-windows-x64-setup.exe` and `SHA256SUMS` from the official
[GitHub Release](https://github.com/chenweixin123/workbooklens/releases/tag/v2.4.0). Verify the
published checksum before running the installer:

~~~powershell
(Get-FileHash .\WorkbookLens-2.4.0-windows-x64-setup.exe -Algorithm SHA256).Hash
~~~

The installer uses the current Windows account and does not require administrator access. Its
wizard provides the normal **Install** button, adds **WorkbookLens** to the Start menu, offers a
desktop shortcut by default, and registers a standard entry in Windows **Installed apps**. Open the
Start-menu or desktop shortcut to launch the local `WorkbookLens` window directly. The application
uses pywebview with the Microsoft Edge WebView2 Runtime to render its loopback-only interface; it
does not show a console or open an external browser. Closing the desktop window stops its background
server cleanly. The language selector offers Chinese and English and remembers the choice locally.

Uninstall from Windows **Settings > Apps > Installed apps**, or choose
**Uninstall WorkbookLens** in the WorkbookLens Start-menu folder. The installer and application are
currently unsigned, so Microsoft SmartScreen may display an unfamiliar-app warning. Download only
from the official release and verify the SHA-256 value; do not disable antivirus protection. The
installed application exposes the same local scan, repair, report, and optional `.xls` conversion
interface as the portable build.

## Windows portable ZIP

The official `WorkbookLens-2.4.0-windows-x64-portable.zip` is the no-install Windows option.
It bundles a 64-bit CPython 3.12 runtime, so users do not need to install Python, `uv`, Microsoft
Excel, LibreOffice, an AI key, or a cloud client. It is a portable folder rather than an installer:
it does not request administrator access, modify the registry, create file associations, or add an
automatic updater.

Download the ZIP and `SHA256SUMS` from the official
[GitHub Release](https://github.com/chenweixin123/workbooklens/releases/tag/v2.4.0), compare the
published SHA-256 value, and extract the complete folder. In PowerShell:

~~~powershell
(Get-FileHash .\WorkbookLens-2.4.0-windows-x64-portable.zip -Algorithm SHA256).Hash
Expand-Archive .\WorkbookLens-2.4.0-windows-x64-portable.zip -DestinationPath .\WorkbookLens
Set-Location .\WorkbookLens\WorkbookLens-2.4.0-windows-x64
.\Start-WorkbookLens.cmd
~~~

The launcher starts the no-console `WorkbookLens.exe` desktop application. Its background service
listens only on `127.0.0.1`, prefers port 8765, and selects an available local port if that port is
already in use. The interface opens inside the WorkbookLens window through Edge WebView2 rather than
an external browser, and closing the window stops the service cleanly. Workbook files, generated
reports, language settings, and diagnostic logs stay on the computer; the portable runtime does not
upload them. `WorkbookLensCLI.exe` remains available in the same folder for command-line workflows.

The local home page also offers **Convert to .xlsx** for legacy binary `.xls` workbooks. Conversion
is enabled only when WorkbookLens detects Microsoft Excel or LibreOffice on that computer. Excel is
preferred for fidelity; LibreOffice is a fallback. Macros and events are disabled for the automated
Excel open, external links are not updated, and the generated file must pass the same bounded OOXML
safety inspection used by normal WorkbookLens inputs before it is downloaded. Conversion can still
change unsupported legacy features, cached formula results, fonts, print layout, or vendor-specific
objects, so review the `.xlsx` copy before relying on it.

The executable is currently unsigned, so Microsoft SmartScreen may display a warning. Verify the
checksum and download only from the official release. The portable build guarantees the built-in
rules; install the wheel in a normal Python environment when third-party entry-point plugins are
required.

## Install the downloaded wheel

The `.whl` attached to the GitHub Release is a local Python application package, not a hosted
upload service. Workbook processing stays on the user's computer. Installing the wheel may contact
the configured Python package index once to obtain dependencies; scanning, planning, repair, diff,
and the loopback web UI do not upload workbooks.

Download `workbooklens-2.4.0-py3-none-any.whl`, `workbooklens-2.4.0.tar.gz`, and
`SHA256SUMS` from the official
[GitHub Release](https://github.com/chenweixin123/workbooklens/releases/tag/v2.4.0), verify the
checksum, and install it with Python 3.11+ and `uv`.

### PowerShell

~~~powershell
Get-Content .\SHA256SUMS
(Get-FileHash .\workbooklens-2.4.0-py3-none-any.whl -Algorithm SHA256).Hash
uv tool install .\workbooklens-2.4.0-py3-none-any.whl
workbooklens --version
workbooklens serve --port 8765
~~~

### Bash

~~~bash
grep -E '  workbooklens-2\.4\.0-py3-none-any\.whl$' SHA256SUMS | sha256sum -c -
grep -E '  workbooklens-2\.4\.0\.tar\.gz$' SHA256SUMS | sha256sum -c -
uv tool install ./workbooklens-2.4.0-py3-none-any.whl
workbooklens --version
workbooklens serve --port 8765
~~~

Then open `http://127.0.0.1:8765/`. The browser page and temporary upload directory are local to
that process. Sensitive workbooks should use the CLI or this loopback UI rather than a hosted CI
runner.

The setup executable, portable ZIP, and wheel are distributions of the same local application. The
installer is the normal Windows option, the portable ZIP makes no installation changes, and the
wheel is the extensible option for Python users.

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

If PyPI lists version 2.4.0, isolated installation is:

~~~bash
uvx --from workbooklens==2.4.0 workbooklens --help
pipx install workbooklens==2.4.0
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

### Optional Workbook Profile

Some spreadsheet mistakes cannot be proved from layout alone. An optional version-2 YAML
`profile` lets a user declare field roles and constraints without sending the workbook anywhere:

~~~yaml
version: 2
profile:
  infer_semantics: true
  report_trailing_whitespace: true
  review_trailing_whitespace_patches: false
  sheets:
    - sheet: Invoices
      range: A1:G10000
      header_row: 1
      columns:
        - header: Invoice ID
          role: identifier
          required: true
          identifier_width: 8
          preserve_leading_zeros: true
        - header: Status
          role: category
          required: true
          allowed_values: [Open, Paid, Void]
        - header: Customer Email
          role: email
        - header: Amount
          role: currency
        - header: Tax Rate
          role: percentage
~~~

Supported roles are `identifier`, `category`, `email`, `phone`, `currency`, `percentage`,
`date`, `number`, and `text`. Select a field by `header` or bounded Excel `column`. Profile rules
report missing required values, invalid enumerations or contact formats, possible lost leading
zeros, trailing whitespace, and number-format role conflicts. They never invent business values.
Trailing ASCII-space cleanup is available only when explicitly requested and remains a reviewed
layout-risk patch, so `--safe-only` never selects it.

Profile contracts fail closed instead of silently disabling checks: each worksheet may appear once,
each column must use exactly one header or column selector, and resolved headers/columns must be
unique and inside the configured table. If range is present and header_row is omitted, its first
row is used. Missing or hidden worksheets, ambiguous headers, invalid types, and version-1 Profile
configuration produce a clear configuration error before findings are evaluated.

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

After the v2.4.0 tag exists:

~~~yaml
permissions:
  contents: read
  security-events: write

steps:
  - uses: actions/checkout@v7
    with:
      fetch-depth: 0

  - uses: chenweixin123/workbooklens@v2.4.0
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
- uses: chenweixin123/workbooklens@v2.4.0
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

Legacy `.xls` conversion is a separate, explicit trust boundary because binary BIFF is not parsed by
the normal OOXML safety layer. WorkbookLens first requires the OLE compound-file signature, then
opens the temporary copy with an installed local spreadsheet application under restricted automation
settings. Use this path only for trusted legacy files: Microsoft Excel or LibreOffice may recalculate
formulas or process workbook-defined behavior supported by that application while opening or saving
the workbook. Only the resulting macro-free `.xlsx` package enters the normal safety layer.
Temporary conversion files are deleted after the download response completes and again on normal
server shutdown.

See [SECURITY.md](SECURITY.md) for private reporting. Never attach a confidential production
workbook to a public issue.

## Honest limitations

- Normal OOXML scan, test, diff, and repair have no Excel calculation engine, VBA execution,
  external-link fetching, or embedded-object opening. This guarantee does not describe the optional
  local `.xls` conversion path.
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
- `.xls` cannot be scanned directly. The local web UI can convert an OLE-based binary `.xls` to
  `.xlsx` when Microsoft Excel or LibreOffice is installed. Use only trusted input because that
  application may recalculate formulas or process workbook-defined behavior; exact fidelity is not
  guaranteed. `.xlsb`, `.ods`, and Google Sheets remain unsupported.
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
python scripts/check_release_artifacts.py dist --version 2.4.0
~~~

To reproduce the Windows x64 portable release candidate from the built wheel, run the following in
PowerShell on 64-bit Windows. The final smoke test exercises the console CLI and native desktop
window, including the loopback-only service, occupied-port fallback, clean shutdown, dual-executable
subsystems, and installation-directory hash check.

~~~powershell
uv sync --locked --python 3.12
New-Item -ItemType Directory -Force -Path .artifacts | Out-Null
uv build --out-dir .artifacts/portable-python-dist
uv export --locked --python 3.12 --no-dev --group portable --no-emit-project --no-hashes --no-header --no-annotate --output-file .artifacts/portable-constraints.txt
$wheel = @(Get-ChildItem -LiteralPath .artifacts/portable-python-dist -Filter '*.whl')
if ($wheel.Count -ne 1) { throw "Expected exactly one wheel, found $($wheel.Count)." }
$python = (Resolve-Path .venv/Scripts/python.exe).Path
& $python scripts/build_portable_windows.py --wheel $wheel[0].FullName --python $python --expected-version 2.4.0 --constraints .artifacts/portable-constraints.txt --output-dir .artifacts/portable-dist
$archive = @(Get-ChildItem -LiteralPath .artifacts/portable-dist -Filter '*.zip')
if ($archive.Count -ne 1) { throw "Expected exactly one ZIP, found $($archive.Count)." }
& $python scripts/check_portable_artifact.py $archive[0].FullName --expected-version 2.4.0
& $python scripts/smoke_portable.py $archive[0].FullName --expected-version 2.4.0
$iscc = (Resolve-Path "$env:LOCALAPPDATA/Programs/Inno Setup 6/ISCC.exe").Path
& $python scripts/build_installer_windows.py --portable-zip $archive[0].FullName --expected-version 2.4.0 --iscc $iscc --output-dir .artifacts/installer-dist
$installer = @(Get-ChildItem -LiteralPath .artifacts/installer-dist -Filter '*.exe')
if ($installer.Count -ne 1) { throw "Expected exactly one installer, found $($installer.Count)." }
& $python scripts/check_installer_artifact.py $installer[0].FullName --expected-version 2.4.0
& $python scripts/smoke_installer_windows.py $installer[0].FullName --portable-zip $archive[0].FullName --expected-version 2.4.0
~~~

The installer smoke test requires a user account without an existing WorkbookLens installation. It
rejects custom directories and a forged per-user uninstall entry, installs only to the fixed
`%LOCALAPPDATA%\Programs\WorkbookLens` directory, verifies the ownership marker, exact portable
payload, Start-menu and desktop shortcuts, the Installed apps record, and the installed executable,
then uninstalls and confirms that all test-created state was removed. Also pass the pinned preceding
release's setup executable and portable ZIP through `--previous-installer` and
`--previous-portable-zip`; for 2.4.0 the upgrade baseline is 2.3.0. The smoke test validates that the
baseline is an earlier numeric release and retains the legacy ZIP profile only for 2.2.1.

See [CONTRIBUTING.md](CONTRIBUTING.md), [the architecture guide](docs/architecture.md),
[CLI exit codes](docs/cli-exit-codes.md), and [the release checklist](docs/release-checklist.md).

Licensed under Apache-2.0.
