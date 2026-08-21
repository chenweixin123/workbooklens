# Architecture

## Trust boundary

Every workbook path enters through `ooxml.safety.inspect_package`. The safety layer validates the
filesystem type and extension, ZIP member/resource limits, XML parser configuration, and internal
relationship resolution before `openpyxl` or a rule sees the file. External relationships are
recorded but never dereferenced.

```mermaid
flowchart LR
    A["Untrusted .xlsx/.xlsm"] --> B["ZIP and XML safety gate"]
    B --> C["Read-only semantic and layout snapshot"]
    C --> D["Region, formula, and layout analysis"]
    D --> E["21-rule registry"]
    E --> F["HTML, JSON, SARIF"]
    E --> G["Source-bound patch plan"]
    G --> H["Safe-only or explicitly reviewed layout operations"]
    H --> I["Direct OOXML patch engine"]
    I --> J["Part manifest and reopen"]
    J --> K["Rescan and semantic diff"]
```

## Modules

- `models`: strict Pydantic public schemas and enums.
- `ooxml`: package limits, hardened XML parsing, relationship checks, and direct patching.
- `snapshot`: sparse extraction of populated/styled cells and workbook structure.
- `layout`: deterministic text measurement, visual shared-border semantics, content bounds, and
  format-tail discovery without invoking a spreadsheet renderer.
- `formulas`: non-executing A1 token analysis, R1C1-like signatures, and conservative translation.
- `regions`: connected dense regions and formula bands with bounded gaps.
- `rules`: plugin interface, registry, and the version-2 built-in catalog.
- `repair`: deterministic plans, cell/layout preconditions, atomic groups, direct XML operations,
  manifests, and validation.
- `diff`: semantic and layout comparison independent of ZIP serialization.
- `testing`: bounded YAML assertions, not a general formula evaluator.
- `reports`: self-contained HTML/JSON/SARIF serialization.
- `web`: loopback-only FastAPI workflow backed by process-owned temporary storage.
- `demo`: generated defects and end-to-end learning artifact.

## Data flow and invariants

The scanner holds one in-memory `openpyxl` workbook only while rules run, then closes it. Rules may
propose declarative operations but do not write files. The snapshot distinguishes populated-content
bounds from the worksheet's declared dimension and records explicit row heights, column widths, and
saved views. Finding, patch, and atomic-group IDs are hashes of stable semantic inputs rather than
list positions.

The repair engine checks these invariants:

1. input SHA-256 equals the plan source hash;
2. each target semantic fingerprint equals its precondition;
3. every selected operation has confidence at least 0.95 and is either safe-only eligible or an
   explicitly accepted `layout_review` operation;
4. `--safe-only` excludes every layout-changing operation, and layout consent never authorizes a
   different unsafe risk class;
5. each atomic group is selected and applied in full;
6. target/source cell fingerprints and applicable row, column, view, or exact-tail fingerprints
   still match;
7. `.xlsm` is never written;
8. target/source cells do not use unsupported formula modes;
9. changed package parts equal the exact expected part set;
10. unchanged entry contents compare byte-for-byte;
11. output reopens through two readers and matches every requested semantic or layout value;
12. source SHA-256 remains unchanged;
13. a rescan introduces no error- or critical-level finding.

## Resource model

The safety layer performs central-directory checks before decompression and reads XML with a
per-part bound. Snapshot, region, and layout logic iterate the sparse cell store rather than trusting
an OOXML dimension such as `A1:XFD1048576`. Format-tail cleanup authorizes bounded exact cell and row
lists, then rejects intersections with formulas, names, tables, validations, comments, hyperlinks,
page breaks, drawings, or unsupported row metadata. YAML assertions independently cap expanded
ranges at 100,000 cells. Web uploads stream in 1 MiB chunks and stop at the configured compressed-file
limit.

## Extension model

`WorkbookRule.run(RuleContext) -> RuleResult` is the public rule interface. A caller can register a
trusted instance directly. Installed packages can expose the `workbooklens.rules` entry-point group,
which is loaded only by an explicit call. This keeps deterministic built-ins available without
implicitly executing arbitrary installed plugin code.
