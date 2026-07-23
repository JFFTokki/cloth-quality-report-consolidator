---
name: quality-report-consolidator
description: Batch-read quality inspection report lists from Excel, download or open every linked PDF, extract all test items and results across heterogeneous laboratory formats, classify basic versus functional textile tests, normalize parent and child test names, and generate a searchable product-code overview plus complete auditable detail workbook. Use for third-party inspection reports, quality report consolidation, product/SKU test coverage, PDF table extraction, and large batches requiring no-silent-omission QA.
---

# Quality Report Consolidator

Use the leader-version workflow as the controlling structure: process every report into two business sheets, split basic and functional tests, and never silently discard an unfamiliar layout. Use normalization, traceability, and exception rules only to strengthen that workflow.

## Global preflight rule

Before every quality-report processing task, reread this section and follow it as the final processing logic. If a requested shortcut conflicts with this logic, surface the conflict before processing.

最终处理逻辑：

1. 读取 Excel 全部工作表和表头，自动定位报告链接列。
2. 拆分单元格中的多个 PDF 链接，保留来源行、货号、样品类型和判定。
3. 下载并验证 PDF 完整性。
4. 每份 PDF 逐页读取全文和表格；扫描件使用 OCR。
5. 识别不同机构的表头、合并单元格、跨页表格和无重复表头页面。
6. 提取货号、版单号、检测项目、方法、限值、实测结果、判定等全部字段。
7. 将汇总页与同一报告的明细页匹配回填。
8. 清除供应商、地址、日期、订单等非检测内容。
9. 区分基础检测与功能性检测。
10. 生成 `货号检测总览` 和 `质检全项目明细`。
11. 自动检查下载失败、有结果无项目、疑似未解析结果及未知格式。
12. 未知格式进入 `待复核`，不会静默遗漏。

## Required workflow

1. Read the `Global preflight rule` above, then run `.venv/bin/python pipeline/qc_all/check_environment.py --smoke-ocr` from the project root. Stop before processing if any Python module, Node exporter dependency, Poppler command, Swift/macOS SDK component, PDF rendering smoke test, or Vision OCR smoke test fails. Never downgrade an environment failure into per-report `待复核` records.
2. Read [references/workflow.md](references/workflow.md).
3. Read [references/schema.md](references/schema.md) before creating records or the workbook.
4. Read [references/classification.md](references/classification.md) before classifying tests.
5. Use the bundled spreadsheet and PDF skills for `.xlsx` authoring and PDF visual verification.
6. Inspect every input sheet and locate fields by header meaning. Never assume a fixed report-link column.
7. Extract source metadata and every PDF URL. Split multiple URLs in one cell, and retain every source row and report even when a newer report exists.
8. In this project, download PDFs with `.venv/bin/python pipeline/qc_all/download_pdfs.py`, which preserves the existing `source_records.json` and `download_manifest.json` contract; use `scripts/download_reports.mjs` only when the project pipeline files are not in scope. Cache successes and resume safely.
9. Process every PDF page in this order: native grid table, no-grid text positioning, then OCR for scanned or insufficient text. Visually verify representative and uncertain pages.
10. Extract report-level metadata once per PDF, including report issue date, issuing institution, and whether CMA/CNAS marks appear; inherit it to every record from that report. Metadata that is unavailable or uncertain must carry a visible status, reason, and evidence column in the workbook, not a silent blank. Normalize every detected result into the schema. Preserve original text, page, table, filename, URL, source row, report number, parser version, and rule version.
11. Convert statistical names to simplified Chinese while retaining original wording. Normalize to a unified parent item; retain material, part, direction, condition, component, and specific substance as child items.
12. Join overview/conclusion rows to detail-page values only within the same report. Mark joined values as derived from the detail page.
13. Classify unified parent items as basic or functional. Keep uncertain parent assignments as `待确认归属` or `待复核`; do not put them into official overview item lists.
14. Build exactly two business sheets by default:
    - `货号检测总览`: one row per product code, separate basic and functional parent-item counts/lists;
    - `质检全项目明细`: one row per normalized or audit record, including raw text and exception records.
15. Run `scripts/validate_records.py` before export. Stop final delivery on validation failures or unresolved missing reports.
16. Render both sheets, inspect key ranges, scan formula errors, and export a new `.xlsx` filename without overwriting historical output.

## Non-negotiable safeguards

- Treat different laboratories and page layouts as separate layout families.
- Support merged cells, multi-line cells, repeated headers, and continuation pages without headers.
- Do not infer a value that the PDF does not provide. Leave the value cell blank only when the paired status/reason column explains why, or use `待复核` when uncertain.
- Put download failures, unreadable scans, encrypted or broken PDFs, OCR failures, unmatched headers, missing result columns, and unknown layouts into explicit exception records.
- Keep raw text and source coordinates so every normalized value can be traced back.
- Do not claim completeness when any report failed download, parsing, OCR, or validation.
- Do not count customer information, suppliers, addresses, dates, standards descriptions, remarks, or order fields as test items.
- Treat CMA/CNAS fields only as evidence that the corresponding mark appears in the report. Do not claim certificate validity or that a specific item is within the accredited scope without separate verification.
- For report issue date, CMA, and CNAS, do not rely on text keywords alone. Use the evidence chain defined in `references/workflow.md`, and output explicit reasons for `未发现`, `待复核`, or blank values.
- For a stylized CMA logo, accept degraded OCR such as `MA` or `MMA` only when a clear 12-digit qualification number is spatially aligned below it with valid coordinates and horizontal overlap. Normalize spacing and punctuation in the number; do not accept the degraded token or an unaligned number by itself.
- Identify the issuing institution from legal-name content in OCR blocks, headers, footers, and repeated full-report text. Do not require an institution identity registry, and do not infer an institution from a report-number prefix or accreditation number alone.
- Carry CAS numbers and detection/report limits from source tables through structured records and export. A defined workbook column must not be populated with a fixed blank placeholder.
- Apply newer explicit human rules before older mapping seeds. Never repair only the final Excel; update the reusable rule source.
- Keep all child semantics in detail. Use only a confidently unified parent item for overview counting and listing.

## Scale behavior

Process reports in batches and cache downloaded PDFs plus extracted page JSON. Resume without reprocessing successful files. Record parser and rule versions in audit fields. For a newly observed layout, preserve its source page as `待复核`; add a layout rule only after comparing normalized rows with the rendered PDF page.

## Example invocation

`使用 $quality-report-consolidator 读取这个Excel中的全部质检报告，以两张表输出货号基础/功能检测总览和全部检测项目明细，并列出所有待复核与异常。`
