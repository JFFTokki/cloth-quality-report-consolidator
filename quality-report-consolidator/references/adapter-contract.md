# Project adapter contract

Use this contract when a repository already contains a quality-report pipeline.

## Required stages

| Stage | Required input | Required output |
| --- | --- | --- |
| Source selection | workbook and scope | source records plus one row per source relationship |
| Acquisition | unique report URLs | download manifest and validated local PDFs |
| Text/OCR | download manifest | page text, OCR evidence, PDF hash, extractor versions |
| Report metadata | page evidence | report number, issue date, institution, CMA/CNAS status and reasons |
| Item extraction | pages and layout evidence | raw item/result rows plus per-PDF status |
| Normalization | raw rows and rules | parent/child records, mapping evidence, overview eligibility |
| Validation/export | structured records and downloads | validation status, two-sheet workbook, inspection evidence |

## Required source relationship

Generate a stable ID from:

```text
workbook + worksheet + row + cell + URL + product code + color + sample type
```

Carry the ID into every structured and exception record. If one parsed PDF belongs to several source rows, expand the relationship instead of selecting the first match.

## Required status and version fields

Record PDF SHA-256 and independent versions for acquisition, text extraction, OCR configuration, report-header parsing, table parsing, mapping rules, and export rules. Reuse a cached stage only when its input identity and relevant versions match.

Use `succeeded`, `partial`, `retryable_failed`, and `permanent_failed` for resumable parsing. Use `完整通过`, `部分完成`, and `校验失败` for delivery validation. Only `完整通过` permits formal delivery.

Provide an explicit parent classification or `classification_default`. New generic adapters should prefer `待分类`; project adapters that intentionally use `基础检测` as a fallback must state it explicitly in the envelope.

## Adapter discovery

Read repository documentation and identify the actual production entry point before editing. Mark historical and diagnostic scripts separately. A generic Skill may coordinate the stages, but the project adapter owns source headers, business scope, OCR backend, laboratory layouts, manual mappings, taxonomy, and delivery naming.

The bundled downloader requires a Node.js runtime with `fetch` and a `pdfinfo` command, configurable through `QC_PDFINFO`. The bundled leader exporter requires `@oai/artifact-tool`; when it is unavailable, use the target repository's spreadsheet adapter while preserving the same schema and validation gates.
