# ADR 0001: Repair through direct OOXML patching

- Status: accepted
- Date: 2026-08-17

## Context

`openpyxl` is a strong high-level reader and fixture generator, but loading and saving a workbook
reconstructs package parts. A repair to one cell could therefore rewrite or discard charts, pivot
metadata, drawings, custom XML, rich extensions, cached values, or relationships that WorkbookLens
does not understand. This conflicts with the product promise that a reviewed repair changes only
what it declares.

## Decision

Repaired workbooks are written by a dedicated ZIP/XML engine. It copies every input entry and
replaces only worksheet parts containing selected operations. A reviewed alignment, text-storage,
or border-edge operation may also add an exact style record in `xl/styles.xml`. If any formula
changes, the engine updates workbook calculation metadata, removes the changed formula's stale
cached value, and requests full recalculation on next open.

Before replacement, the engine validates the source and cell fingerprints. Layout operations also
carry row, column, saved-view, or exact-tail fingerprints. Column width, row height, wrapping,
shrink-to-fit, text storage, saved view, edge-border, and format-tail operations are classified as
`layout_review`: they are never eligible for `--safe-only` and require explicit patch selection plus
`--accept-layout-risk`. Related wrap/height operations use atomic groups. The identifier-display rule
does not rewrite numeric values as text; it proposes a font-aware width-only repair for ten- and
eleven-digit values so COUNT, SUM, MATCH, validation, and pivot semantics remain unchanged. Longer
General-formatted identifiers remain review-only when width alone cannot guarantee literal display.

Exact format-tail cleanup is not a rectangular clear or a generic UsedRange reset. It removes only
enumerated blank styled cells and empty row records and fails closed when they intersect worksheet
formulas, defined names, tables, validation or formatting ranges, hyperlinks, comments, page breaks,
drawing anchors, or unsupported row metadata.

After replacement, the engine compares every package entry, rejects an unexpected changed, added,
or removed part, reopens the output, checks requested cell and layout semantics, rescans findings,
and deletes the output on failure.

## Consequences

Benefits:

- unknown and unrelated package content is preserved;
- changes are reviewable at both semantic and package-part levels;
- stale plans and ambiguous formula types fail closed;
- layout risk is separated from confidence and requires an explicit second opt-in;
- atomic groups prevent partial repairs that would introduce a new clipping or display defect;
- no Office installation is required.

Costs:

- the patch surface stays deliberately small;
- XML namespace, cell-order, dimension, style-table, shared-formula, view, and recalculation rules
  require focused tests;
- unsupported formula modes are reported but cannot be auto-repaired;
- deterministic text measurement and saved zoom are conservative estimates and require review in
  the target spreadsheet application;
- changed worksheet XML is reserialized, so byte identity is guaranteed for untouched parts, not
  for a modified worksheet or styles part itself.

## Rejected alternatives

- Save through `openpyxl`: too broad a package rewrite for a preservation tool.
- Automate Excel/LibreOffice: not reliably cross-platform or local-installation independent, and it
  would recalculate or normalize unrelated content.
- Patch cached values: WorkbookLens is not an Excel calculation engine and cannot prove them.
