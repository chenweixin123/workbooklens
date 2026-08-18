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
replaces only worksheet parts containing selected cells. If any formula changes, it also updates
workbook calculation metadata, removes the changed formula's stale cached value, and requests full
recalculation on next open.

Before replacement, the engine validates the source and cell fingerprints. After replacement, it
compares every package entry, rejects an unexpected changed/added/removed part, reopens the output,
checks requested cell semantics, rescans findings, and deletes the output on failure.

## Consequences

Benefits:

- unknown and unrelated package content is preserved;
- changes are reviewable at both semantic and package-part levels;
- stale plans and ambiguous formula types fail closed;
- no Office installation is required.

Costs:

- the patch surface stays deliberately small;
- XML namespace, cell-order, dimension, shared-formula, and recalculation rules require focused
  tests;
- unsupported formula modes are reported but cannot be auto-repaired;
- changed worksheet XML is reserialized, so byte identity is guaranteed for untouched parts, not
  for the modified worksheet part itself.

## Rejected alternatives

- Save through `openpyxl`: too broad a package rewrite for a preservation tool.
- Automate Excel/LibreOffice: not reliably cross-platform or local-installation independent, and it
  would recalculate or normalize unrelated content.
- Patch cached values: WorkbookLens is not an Excel calculation engine and cannot prove them.
