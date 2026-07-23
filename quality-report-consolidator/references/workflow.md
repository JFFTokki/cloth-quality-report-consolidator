# End-to-end workflow

## First rule: traditional before recognition

Convert the complete native and OCR report text from traditional Chinese to simplified Chinese before metadata, header, item, result, and judgment recognition. Preserve the original traditional page text, OCR block text, or source row in audit fields. This rule precedes every other extraction and normalization rule.

## 0. Environment gate

Run `.venv/bin/python pipeline/qc_all/check_environment.py --smoke-ocr` from the project root before selecting or downloading reports. The gate must verify the project Python modules, Node workbook exporter, Poppler `pdftoppm`/`pdfinfo`, Swift, a dynamically discovered macOS SDK, PDF-to-PNG rendering, and one actual Vision OCR invocation. Stop the batch on any failure and report the exact missing or failing component. Do not continue and convert a shared environment failure into hundreds of report-level `待复核` values.

## 1. Input workbook

- Inspect all sheets and headers.
- Locate report order, product/SKU code, color, sample type, source judgment, last-modified time, and report link by header meaning.
- Treat column J only as a discovered source column, never as a fixed rule.
- Split every URL from cells containing semicolons, newlines, spaces, or multiple hyperlinks.
- Create one manifest row per source-row and report-URL relationship. Preserve the original sheet, row number, cell address, product code, color, sample type, and judgment.
- Keep all valid reports for the same product and color. Last-modified time is trace metadata or a sorting aid, not a reason to discard an older report.

## 2. Report acquisition

- Download with bounded concurrency and stable filenames.
- Verify HTTP success and `%PDF-` signature.
- Record status, byte size, error, local path, source URL, and source relationship.
- Reuse a successful cached file only when its URL and file signature still match.
- Stop final delivery if any required report is missing; list every failure explicitly.

## 3. Page extraction

For every PDF page, save:

- full native text and native character count;
- extracted tables, table count, and cell coordinates when available;
- page number and page count;
- whether the page appears scanned or text-based;
- OCR text, OCR character count, engine status, and error when OCR is needed.

For every PDF, also extract report-level metadata once and inherit it to every record from that PDF:

- first-page report header: run a dedicated parser before the general full-report metadata rules. Treat the top band of the first page as one full-width semantic region; do not assume that the institution is on the left or that CMA/CNAS marks are on the right. Read native text first. If the institution or either mark is still unresolved, render the first page and OCR the complete top band, retaining OCR blocks, coordinates, confidence, parser version, status, and errors. Locate fields by their content and nearby evidence rather than fixed horizontal position. OCR completion without strong mark evidence does not prove absence and remains `待复核`; write `未发现` only after the complete mark evidence chain supports absence;
- report issue date: inspect the first page first, then collect candidates from the full report if the first page is inconclusive. Use this label priority: `签发日期` or `报告签发日期`; `出具日期` or `Date of Issue`; `报告日期`; `检测日期` or `检验日期`. For each candidate, retain page number, source label, original nearby text, normalized date, whether it is a date range, and why it was selected or rejected. Use a lower-priority label only when no reliable higher-priority value exists. If the selected value is a date range, use the end date and retain the original range in diagnostics. If multiple dates appear under the same priority label, use the later date as the updated date, mark the field as `已识别`, and retain all candidates plus the selection reason in diagnostics. Format confirmed dates as `yyyy-mm-dd`; other genuinely ambiguous candidates are `待复核` and must not be guessed. Populate `报告签发日期识别状态` as `已识别`, `未发现`, or `待复核`, and populate `报告签发日期异常原因` for every blank or uncertain date;
- issuing institution: prefer the complete legal name in the dedicated first-page report-header result, title, cover, beside the institution seal, repeated footer, or repeated full-report text. Parse OCR blocks independently before reconstructed visual lines so CMA/CNAS wording on the same row does not discard a valid legal-name block. Support spaced Chinese header text, simplified/traditional quality terms, and legal names ending in `中心`, `检验中心`, `检验检测中心`, `检测实验室`, `研究院`, `集团有限公司`, or `有限公司`. Prefer a repeated or longer, more specific legal name when the same report also shows an alias. Do not use the client, supplier, customer, applicant, manufacturer, or production-unit name. Do not require an institution identity registry and do not infer a legal name from a report prefix, CNAS number, CMA number, domain, or logo alone. Keep `机构名称待确认` when the report does not support a stable identification;
- CMA and CNAS marks: inspect each mark independently with this evidence chain: native PDF text; OCR text from the first page and top area of every page; rendered top-region image evidence for common visible marks; then a manual-review state for suspected marks with insufficient confidence. For a stylized CMA logo, normalize punctuation and accept degraded OCR tokens `MA` or `MMA` only when a clear 12-digit qualification number is below it, both blocks have valid coordinates, and their horizontal ranges overlap; a degraded token, phone number, ordinary identifier, or unaligned number alone is insufficient. A negative authorization or applicability sentence is not positive mark evidence, while a strong top-region visual combination is not negated by unrelated body disclaimers because the field records mark appearance only. Store each mark independently as `有`, `未发现`, or `待复核`. `有` means the mark visibly appears in the report; it does not verify current certificate validity or whether a specific test is within the accredited scope. Write `未发现` only after native text, OCR text, and top-region image checks provide no mark evidence. Write `待复核` when a top-region image suggests a CMA/CNAS seal but text/OCR evidence is weak, cropped, blurred, or conflicting. Populate `CMA识别证据/异常原因` and `CNAS识别证据/异常原因` with the concrete evidence or failure reason.

Use this extraction order:

1. Parse native grid tables.
2. If the table has no usable grid, locate headers and fields by text position.
3. If native text is absent, insufficient, or visibly garbled, run OCR.
4. Render representative pages and every uncertain page for visual comparison.

Do not reject a page merely because a fixed keyword set is absent. Use multiple signals to find candidate result tables, but only emit a structured item when the item-to-result or item-to-judgment relationship is reliable.

## 4. Layout recognition

Recognize headers semantically, not by column number. Common aliases:

- item: 测试项目, 检测项目, 检验项目, 项目名称;
- method: 测试方法, 检测方法, 判定依据;
- requirement: 技术要求, 限值, 标准值, 要求;
- result: 测试结果, 检测结果, 实测值;
- judgment: 单项判定, 结论, 评价;
- auxiliary: 序号, 单位, CAS号, 报告限, 检出限, 备注.

Build explicit column mappings for auxiliary fields rather than emitting fixed blank values:

- recognize CAS headers such as `CAS号`, `CAS No.`, `CAS No`, and `CAS编号`;
- read CAS from a dedicated column when present, or extract it from the same result row/item text using the standard hyphenated form such as `101-14-4`;
- keep all CAS values associated with the same structured result row, validate the numeric structure/check digit when possible, and send OCR-ambiguous or invalid candidates to `待复核` without deleting the original text;
- recognize `报告限`, `检出限`, `检测限`, `定量限`, `LOD`, and `LOQ` as report/detection-limit fields and preserve their units and row association.

Prioritize `测试项目(计量单位)` as the item column, not the unit column.

Handle:

- vertically merged item cells by inheriting the nearest valid parent item;
- continuation rows by appending method, requirement, result, or judgment fragments to the active item;
- continuation pages without headers by reusing only a verified layout signature;
- repeated headers across pages without treating them as result rows;
- summary pages separately from detailed result pages;
- chemical or composition sub-tables using category/title plus analyte/component rows;
- the full context of nested names, such as a composition parent plus each component child.

## 5. Record normalization

- Preserve exact source wording in `rawItem` and `rawRow`.
- Preserve report-level metadata evidence in `diagnostics`, including source page, source label or mark evidence, original date text, candidate list, selected/rejected reasons, top-region image/OCR status, and uncertainty reason.
- Convert statistical names to simplified Chinese, while retaining the original wording for audit.
- Normalize whitespace, Unicode compatibility forms, full-width punctuation, unit spelling, private-use bullets, and bracket variants before matching.
- Strip a bracketed unit from the parent name; retain material, part, direction, condition, component, and specific substance as `childItem`.
- Map semantically equivalent institution-specific names to `parentItem` without changing meaning.
- Apply newest explicit human rules first, then confirmed mapping seeds, then conservative general rules.
- When the parent is uncertain, set `mappingStatus=待确认归属` or `待复核` and keep the row out of official overview item lists. Complete new parent items may use `自动新增`; fragments or isolated symbols must never become new parents.
- Never overwrite source values with guesses.
- When a summary row contains only item plus conclusion, match detail rows from the same report and fill method, requirement, result, and unit. Add note `汇总行由同一报告检测明细页回填` and mark the derivation.
- When the same product, report, sample type/source relationship, color, normalized parent, and child contain both a conclusion-only row and a row with a concrete result plus method or requirement, retain only the detailed row. References such as `见附表1` and compliance text such as `符合GB/T 14272-2021` are not concrete results. Preserve qualitative exceptions such as `不考核`, `不适用`, and `免测`; a conclusion row carrying a method may be replaced only by a detailed row with the same method. Never apply this rule across different reports, samples, colors, children, methods, or multiple rows that each contain concrete results.

## 6. Product overview

- Split product codes using newlines and verified separators.
- Keep identifiers as text and never allow scientific notation.
- Exclude empty, `0`, and obvious placeholder product codes from the overview, but retain them in detail as exceptions.
- Deduplicate confidently normalized parent items per product code.
- Do not create separate overview items for child material, part, direction, condition, component, or substance.
- Exclude metadata such as supplier, company, address, date, order, style number, sample description, and report remarks.
- Separate basic and functional parent items according to `classification.md`.
- Preserve different results from multiple reports; never let a newer report overwrite a distinct older result.

## 7. Exception statuses

Use the most specific status supported by evidence:

- `下载失败`, `PDF打开失败`;
- `OCR待处理`, `OCR失败`, `PDF文本异常`;
- `无可识别表格结构`, `表头字段未匹配`, `缺少结果或判定列`;
- `检测项被有效性规则排除`, `项名不完整待复核`, `新增检测项待确认`;
- `缺失判定或结果`, `机构名称待确认`, `同项结果差异`;
- `无有效货号`, `无PDF链接`, `未知格式待复核`.

Retain page text, diagnostic counts, header candidates, engine error, and source coordinates. A report with some reliable items plus page-level failures is `部分解析`, not fully successful and not generically unrecognized.

## 8. Quality gates

Required zero counts:

- structured result rows with blank raw or normalized item;
- official overview items whose `mappingStatus` is not confirmed;
- numbered rows containing a standard and outcome but left unstructured;
- summary items without a same-report detail match when a detail page exists;
- report downloads marked failed but omitted from exception records;
- PDFs absent from both structured records and exception records;
- formulas with Excel errors;
- private-use or control characters and symbol-only unified parent items.
- a structured source row containing a reliable CAS value while its `casNumber` is blank;
- a structured source row containing a reliable report/detection limit while its `detectionLimit` is blank;
- inconsistent report issue date, issuing institution, CMA mark, or CNAS mark values among records from the same PDF.

Also reconcile:

- all Excel sheets and source rows inspected;
- all URLs found versus manifest relationships;
- unique report URLs versus success and failure totals;
- PDF page count versus extracted-page count;
- structured item, pending-review, excluded-fragment, and exception counts;
- duplicate report URLs counted once in quality totals while retaining all source relationships.
- every PDF has one report-level metadata result for issue date, issuing institution, CMA mark, and CNAS mark; unavailable or uncertain values must carry an explicit reason instead of being silently blank.
- report issue date, CMA, and CNAS diagnosis columns are not blank when their corresponding value is blank, `未发现`, or `待复核`.

## 9. Output

Put `货号检测总览` first and `质检全项目明细` second. Apply the exact leader visual contract in `schema.md`: no freeze panes, table filters, wrapped long lists, leader conditional colors, and product codes preserved as text. Include normalized item, child context, trace fields, audit records, and exceptions in the detail sheet. Render both sheets before export. Reject 1x1 or blank renderer placeholders as visual evidence and record that an external headless visual review is required; never silently save them as valid previews. Create a new output filename; never overwrite a historical result.
