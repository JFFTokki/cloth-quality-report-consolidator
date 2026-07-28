---
name: quality-report-consolidator
description: Coordinate and harden an adapter-backed workflow that consolidates Excel-listed quality inspection reports and linked PDFs into a traceable product overview and item-level audit workbook. Use for batch QC or laboratory report projects that need bounded TLS-verified PDF acquisition, native/OCR extraction adapters, metadata and test-item normalization, parent/child grouping, explicit exceptions, strict completeness validation, and non-overwriting Excel delivery across heterogeneous layouts.
---

# Quality Report Consolidator

Build a traceable report pipeline, not a one-off workbook edit. Treat unknown layouts and missing reports as explicit states. Never claim completeness from a successful export alone.

## Establish the contract

Before processing, state:

1. scope: workbook, sheets, reporting period, product/SKU selection, and whether the output is preview or formal delivery;
2. required source fields and report relationships;
3. output directory, filename policy, and historical overwrite rule;
4. platform capabilities for native PDF parsing, rendering, and OCR;
5. business taxonomy overrides and unresolved decisions.

Inspect the target repository before choosing commands. If it already has a production pipeline, use it through the adapter contract in [references/adapter-contract.md](references/adapter-contract.md). Do not replace a proven parser with an ad-hoc script merely to make the workflow look generic.

## Read references progressively

- Read [references/adapter-contract.md](references/adapter-contract.md) when integrating an existing repository or defining stage inputs and outputs.
- Read [references/workflow.md](references/workflow.md) before acquisition, PDF parsing, metadata recognition, or exception handling.
- Read [references/schema.md](references/schema.md) before creating records or exporting a workbook.
- Read [references/classification.md](references/classification.md) before classifying textile test parents; use a supplied company taxonomy first.

## Process the source workbook

Inspect every non-empty worksheet and identify business sheets by header meaning. Record ignored sheets and why they were excluded. Locate product code, color, sample type, source judgment, report number, last-modified metadata, and report links semantically.

Extract URLs from plain text, cell hyperlink targets, and `HYPERLINK()` formulas when the workbook library exposes them. Split every URL in a cell. Preserve one source relationship per workbook, sheet, row, cell, URL, product, color, and sample type. Generate a stable `sourceRelationshipId`; do not reduce relationships to the first `URL + product` match.

## Acquire reports safely

Allow only absolute `http` and `https` URLs. Verify TLS with system trust or an explicitly supplied CA file; never disable certificate validation. Apply timeouts, a maximum byte size, streaming download, temporary-file writes, atomic rename, HTTP status checks, `%PDF-` signature checks, and an actual PDF-open check before marking success.

Record original URL, final URL, byte size, SHA-256, local path, cache status, server validators when available, and exact failure reason. Use the bundled `scripts/download_reports.mjs` only as a generic acquisition fallback; the parsing stage must still verify that each file opens successfully.

## Extract and normalize

Convert complete native and OCR text from traditional Chinese to simplified Chinese before recognizing metadata, headers, items, results, or judgments. Preserve original text for audit.

Process each page in this order:

1. native grid tables;
2. no-grid text positioning;
3. OCR when native text is absent, insufficient, or visibly garbled.

Extract report-level metadata once per PDF and inherit it to related records. Keep report number, issue date, institution, CMA, and CNAS evidence separate. Validate dates with a real calendar. Do not infer institutions or accreditation from report prefixes or numbers.

Normalize each reliable result to a parent test item plus child context. Preserve material, part, direction, condition, component, analyte, method, limit, result, judgment, unit, CAS number, raw row, page, URL, source relationship, parser version, and rule version. Keep fragments and uncertain mappings out of official overview counts.

## Resume without hiding failures

Persist one atomic result per PDF. Use `succeeded`, `partial`, `retryable_failed`, and `permanent_failed`. Skip only successful or explicitly permanent outcomes. Keep partial and retryable outcomes eligible for an explicit retry.

Invalidate a stage when its input hash or relevant parser, OCR, header, table, or mapping version changes. Rebuild aggregate JSON/JSONL from per-PDF results so a crash between result and state writes cannot duplicate business rows.

## Validate before export

Run the bundled validator against structured records and the download manifest:

```bash
python quality-report-consolidator/scripts/validate_records.py report_data.json --downloads download_manifest.json
```

Strict mode is the default. A required failed download must block formal delivery even when an exception row exists. Use `--allow-partial` only for a clearly labeled preview; `部分完成` is never formally deliverable.

For formal workbook creation, use the fail-closed wrapper rather than calling the exporter directly:

```bash
QC_PYTHON=python3 node quality-report-consolidator/scripts/validate_and_export.mjs report_data.json download_manifest.json output.xlsx
```

The wrapper writes validation and export-diagnostic sidecars and runs the exporter only after `完整通过`. It refuses to overwrite any existing output. Formula errors or an unavailable/blank built-in preview block formal completion; the generated workbook is retained with an `external-review-required` filename for external visual inspection. Direct exporter invocation is an adapter/debug surface, not evidence of formal validation.

Require a non-empty source relationship index, one stable ID per relationship, coverage of every relationship by a structured or exception record, at least one structured result, successful per-PDF process states, and report metadata for every successfully acquired PDF. Also require zero unexplained gaps for structured item names, overview eligibility, metadata evidence, CAS/detection-limit transfer, and formula errors. Reconcile unique PDF counts separately from source relationship counts.

## Export and inspect

Default to two business sheets:

- `货号检测总览`: one row per product code, using only reliable parent items;
- `质检全项目明细`: every structured result, review row, exception, and trace field.

Create a new filename and preserve prior runs. Render and visually inspect both sheets. Reject blank or `1x1` preview placeholders as evidence. Keep the run manifest, versions, counts, validation output, workbook inspection, and known limitations beside the output.

## Handle project-specific rules

Keep quarter filters, source column aliases, laboratory layout rules, manual mapping workbooks, business classifications, and delivery naming in a project adapter or extension. Apply newer explicit human rules before older seeds. Change the reusable rule source and rerun; never repair only the final Excel.
