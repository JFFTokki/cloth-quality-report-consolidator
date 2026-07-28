import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { inflateSync } from "node:zlib";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const [, , inputPath, outputPath] = process.argv;
if (!inputPath || !outputPath) {
  throw new Error("Usage: node export_leader_workbook.mjs report_data.json output.xlsx");
}

async function assertPathAbsent(filePath) {
  try {
    await fs.access(filePath);
  } catch {
    return;
  }
  throw new Error(`Refusing to overwrite existing output: ${filePath}`);
}

await assertPathAbsent(outputPath);
await assertPathAbsent(outputPath.replace(/\.xlsx$/i, ".export.json"));
await assertPathAbsent(outputPath.replace(/\.xlsx$/i, "_overview_preview.png"));
await assertPathAbsent(outputPath.replace(/\.xlsx$/i, "_detail_preview.png"));

const generatedAt = new Date().toISOString();
const rawData = JSON.parse(await fs.readFile(inputPath, "utf8"));

function adaptCanonical(payload) {
  if (payload.schemaVersion !== "quality-report-consolidator/v1") return payload;
  const reportMetadataByUrl = Object.fromEntries(
    Object.entries(payload.reportMetadataByUrl || {}).map(([url, row]) => [url, {
      report_no: row.reportNumber || row.reportInternalId || "",
      report_issue_date: row.reportIssueDate || "",
      report_issue_date_status: row.reportIssueDateStatus || "",
      report_issue_date_reason: row.reportIssueDateReason || "",
      institution: row.issuingInstitution || "",
      cma_mark: row.cmaMark || "",
      cnas_mark: row.cnasMark || "",
      cma_evidence: row.cmaRecognitionNote || "",
      cnas_evidence: row.cnasRecognitionNote || "",
      report_product_codes: row.reportProductCodes || [],
      plate_numbers: row.plateNumbers || [],
      material_numbers: row.materialNumbers || [],
      path: row.localPath || "",
    }]),
  );
  const sourceIndex = (payload.sourceRelationships || []).map((row) => ({
    source_relationship_id: row.sourceRelationshipId || "",
    source_sheet: row.sourceSheet || "",
    source_row: row.sourceRow || "",
    source_cell: row.sourceCell || "",
    sku: row.sourceProductCode || "",
    color: row.color || "",
    order_no: row.sourceReportNumber || "",
    sample_type: row.sampleType || "",
    overall_result: row.sourceJudgment || "",
    url: row.url || "",
    processing_status: row.processStatus || "",
  }));
  const adaptedRows = (payload.records || []).map((row) => ({
    record_id: row.recordId || "",
    record_type: row.recordType || "检测结果明细",
    source_relationship_id: row.sourceRelationshipId || "",
    source_sheet: row.sourceSheet || "",
    source_row: row.sourceRow || "",
    source_cell: row.sourceCell || "",
    sku: row.sourceProductCode || "",
    subcategory: row.subcategory || "",
    color: row.color || "",
    source_order_no: row.sourceReportNumber || "",
    sample_type: row.sampleType || "",
    report_no: row.reportNumber || row.reportInternalId || "",
    report_issue_date: row.reportIssueDate || "",
    report_issue_date_status: row.reportIssueDateStatus || "",
    report_issue_date_reason: row.reportIssueDateReason || "",
    report_product_code: row.reportProductCode || "",
    plate_number: row.plateNumber || "",
    material_number: row.materialNumber || "",
    institution: row.issuingInstitution || "",
    cma_mark: row.cmaMark || "",
    cnas_mark: row.cnasMark || "",
    cma_recognition_note: row.cmaRecognitionNote || "",
    cnas_recognition_note: row.cnasRecognitionNote || "",
    raw_item: row.rawItem || "",
    simple_item: row.simplifiedItem || "",
    standard_item: row.parentItem || "",
    subitem: row.childItem || "",
    method: row.method || "",
    requirement: row.requirement || "",
    result: row.result || "",
    verdict: row.judgment || "",
    unit: row.unit || "",
    cas_number: row.casNumber || "",
    detection_limit: row.detectionLimit || "",
    result_detail: row.note || "",
    raw_row: row.rawRow || "",
    page: row.pageNumber || "",
    table: row.tableNumber || "",
    item_number: row.itemNumber || "",
    table_title: row.tableTitle || "",
    status: row.mappingStatus || "",
    merge_evidence: row.mappingEvidence || "",
    include_horizontal: row.includeInOverview === true || row.includeInOverview === "是",
    classification: row.classification || "",
    local_path: row.localPath || "",
    parser_version: row.parserVersion || "",
    rule_version: row.ruleVersion || "",
    processing_status: row.processStatus || "",
    metadata_diagnostics: typeof row.diagnostics === "string" ? row.diagnostics : JSON.stringify(row.diagnostics || {}),
    url: row.url || "",
  }));
  const detailRows = adaptedRows.filter((_, index) => !["异常", "待复核"].includes(payload.records[index]?.recordType));
  const errors = adaptedRows
    .map((row, index) => ({ row, original: payload.records[index] }))
    .filter(({ original }) => ["异常", "待复核"].includes(original?.recordType))
    .map(({ row, original }) => ({
      ...row,
      type: original.processStatus || original.recordType,
      detail: original.note || original.diagnostics || "",
      raw_item: original.rawItem || "",
      simple_item: original.simplifiedItem || "",
      suggested_item: original.parentItem || "",
    }));
  return {
    ...payload,
    source_workbook: payload.sourceWorkbook || "",
    source_index: sourceIndex,
    detail_rows: detailRows,
    errors,
    sample_skus: [...new Set(adaptedRows.map((row) => row.sku).filter(Boolean))],
    sample_urls: [...new Set(adaptedRows.map((row) => row.url).filter(Boolean))],
    report_metadata_by_url: reportMetadataByUrl,
    pipeline_versions: payload.pipelineVersions || {},
    classification_default: payload.classificationDefault || "待分类",
  };
}

const data = adaptCanonical(rawData);
const parserVersion = Object.entries(data.pipeline_versions || {}).map(([key, value]) => `${key}=${value}`).join(";") || "project-adapter-unspecified";
const ruleVersion = data.pipeline_versions?.mapping_rules || data.rule_version || data.ruleVersion || "rule-unspecified";

const functionalTerms = [
  "远红外", "红外升温", "吸湿发热", "发热", "升温", "保温", "蓄热", "暖感",
  "凉感", "接触瞬间凉感", "抗菌", "抑菌", "抗病毒", "防螨", "防霉", "消臭",
  "防紫外", "紫外线", "防泼水", "拒水", "防水", "耐水压", "静水压", "拒油", "防油",
  "防污", "易去污", "吸湿速干", "速干", "透湿", "透气", "防风", "防钻绒",
  "钻绒", "防静电", "抗静电", "负离子", "阻燃", "防火", "自清洁", "驱蚊", "防蚊",
];

const placeholderValues = new Set(["", "-", "---", "——", "/", "无", "null", "undefined"]);

function compactText(value) {
  return String(value ?? "").trim();
}

function cleanPlateNumbers(value) {
  const cleaned = [];
  for (const line of compactText(value).split(/\n+/)) {
    const compact = line.replace(/\s+/g, "");
    const plateNumber = compact.match(/^[A-Za-z0-9]+(?:[-/_][A-Za-z0-9]+)*/)?.[0] ?? "";
    if (/[A-Za-z0-9]/.test(plateNumber) && !cleaned.includes(plateNumber)) cleaned.push(plateNumber);
  }
  return cleaned.join("\n");
}

function uniqueJoined(values, separator = "\n") {
  const out = [];
  for (const value of values) {
    const text = compactText(value);
    if (!text || out.includes(text)) continue;
    out.push(text);
  }
  return out.join(separator);
}

function normalizedJudgment(values) {
  let hasQualified = false;
  let hasUnqualified = false;
  for (const value of values) {
    for (const part of compactText(value).split(/\n|,|，|;|；/)) {
      const text = part.trim();
      if (!text || placeholderValues.has(text)) continue;
      if (/不合格|不符合|未通过/.test(text)) {
        hasUnqualified = true;
      } else if (/合格|符合|通过/.test(text)) {
        hasQualified = true;
      }
    }
  }
  if (hasUnqualified) return "不合格";
  if (hasQualified) return "合格";
  return "";
}

function classify(parentItem, mappingStatus, explicitClassification = "") {
  if (!parentItem || ["待确认归属", "待复核", "解析残片"].includes(mappingStatus)) return "";
  if (["基础检测", "功能性检测", "其他", "待分类"].includes(explicitClassification)) return explicitClassification;
  if (functionalTerms.some((term) => parentItem.includes(term))) return "功能性检测";
  return data.classification_default || "待分类";
}

function reviewRequired(mappingStatus) {
  return ["待确认归属", "待复核", "解析残片", "自动归入父项（子项类型未完全识别）"].includes(mappingStatus);
}

function localFileName(url) {
  if (!url) return "";
  try {
    const parsed = new URL(url);
    return decodeURIComponent(path.basename(parsed.pathname));
  } catch {
    return path.basename(String(url));
  }
}

function reportInternalId(rawReportNo, url) {
  const raw = compactText(rawReportNo);
  if (raw && !/^(INSPECTION|REPORT|CATEGORY)$/i.test(raw)) return raw;
  return "";
}

function normalizeItemList(values) {
  return uniqueJoined([...values].sort((a, b) => a.localeCompare(b, "zh-Hans-CN")), "、");
}

const sourceRowsByUrlSku = new Map();
const sourceRowsByUrl = new Map();
const sourceRowsByRelationshipId = new Map();
for (const row of data.source_index || []) {
  const url = compactText(row.url);
  const sku = compactText(row.sku);
  if (row.source_relationship_id) sourceRowsByRelationshipId.set(compactText(row.source_relationship_id), row);
  if (!url) continue;
  const rowList = sourceRowsByUrl.get(url) || [];
  rowList.push(row);
  sourceRowsByUrl.set(url, rowList);
  if (sku) {
    const key = `${url}\u0000${sku}`;
    const skuList = sourceRowsByUrlSku.get(key) || [];
    skuList.push(row);
    sourceRowsByUrlSku.set(key, skuList);
  }
}

function relationshipId(row) {
  if (row.source_relationship_id) return compactText(row.source_relationship_id);
  const values = [data.source_workbook || "", row.source_sheet, row.source_row, row.source_cell, row.url, row.sku, row.color, row.sample_type];
  if (!values.some((value) => compactText(value))) return "";
  return `src-${crypto.createHash("sha256").update(values.map(compactText).join("\u001f")).digest("hex").slice(0, 20)}`;
}

function sourceFor(row) {
  const requestedRelationshipId = relationshipId(row);
  const relationshipSource = sourceRowsByRelationshipId.get(requestedRelationshipId);
  if (row.source_relationship_id && !relationshipSource) {
    throw new Error(`Unknown source_relationship_id: ${row.source_relationship_id}`);
  }
  const candidates = sourceRowsByUrlSku.get(`${compactText(row.url)}\u0000${compactText(row.sku)}`) || sourceRowsByUrl.get(compactText(row.url)) || [];
  const exactCandidates = candidates.filter((candidate) => {
    const comparisons = [
      [row.source_sheet, candidate.source_sheet],
      [row.source_row, candidate.source_row],
      [row.source_cell, candidate.source_cell],
      [row.sample_type, candidate.sample_type],
    ].filter(([expected]) => compactText(expected));
    const candidateColors = compactText(candidate.color || candidate.selected_colors).split(/[，,\n]/).filter(Boolean);
    const colorMatches = !compactText(row.color) || candidateColors.length === 0 || candidateColors.includes(compactText(row.color));
    return colorMatches && comparisons.every(([expected, actual]) => compactText(expected) === compactText(actual));
  });
  const resolved = relationshipSource || (exactCandidates.length === 1 ? exactCandidates[0] : candidates.length === 1 ? candidates[0] : null);
  if (!resolved && candidates.length > 1) {
    throw new Error(`Ambiguous source relationship for URL=${compactText(row.url)} SKU=${compactText(row.sku)}`);
  }
  const fallback = resolved || {};
  return {
    ...fallback,
    source_relationship_id: row.source_relationship_id || fallback.source_relationship_id || "",
    source_sheet: row.source_sheet || fallback.source_sheet || "",
    source_row: row.source_row || fallback.source_row || "",
    source_cell: row.source_cell || fallback.source_cell || "",
    url: row.url || fallback.url || "",
    sku: row.sku || fallback.sku || "",
    color: row.color || fallback.color || "",
    sample_type: row.sample_type || fallback.sample_type || "",
    order_no: row.source_order_no || fallback.order_no || "",
  };
}

const reportIndexByUrl = new Map();
for (const [index, url] of (data.sample_urls || []).entries()) {
  reportIndexByUrl.set(url, index + 1);
}

const reportMetadataByUrl = new Map(Object.entries(data.report_metadata_by_url || {}));

function reportMetadata(url) {
  return reportMetadataByUrl.get(compactText(url)) || {};
}

function joinedMetadata(meta, key) {
  const value = meta?.[key];
  return Array.isArray(value) ? uniqueJoined(value) : compactText(value);
}

function excelDate(value) {
  const match = /^(20\d{2})-(\d{2})-(\d{2})$/.exec(compactText(value));
  if (!match) return compactText(value);
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const parsed = new Date(Date.UTC(year, month - 1, day));
  if (parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day) return "";
  return parsed;
}

function dateFields(row, meta) {
  const raw = compactText(row.report_issue_date || meta.report_issue_date);
  const parsed = excelDate(raw);
  const valid = parsed instanceof Date;
  if (raw && !valid) {
    const priorReason = compactText(row.report_issue_date_reason || meta.report_issue_date_reason);
    const reason = `待复核｜非法日历日期：${raw}${priorReason ? `；原识别信息：${priorReason}` : ""}`;
    return { value: "", status: "待复核", reason };
  }
  const status = row.report_issue_date_status || meta.report_issue_date_status || (valid ? "已识别" : raw ? "待复核" : "未发现");
  const reason = row.report_issue_date_reason || meta.report_issue_date_reason || "";
  return { value: valid ? parsed : "", status, reason };
}

function excelColumnName(index) {
  let value = index;
  let name = "";
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
}

const detailHeaders = [
  "记录ID", "记录类型", "来源表行", "报告序号", "报告单号",
  "检测报告文件名", "报告内编号", "来源货号/款号", "小类", "报告签发日期", "报告内货号/款号", "颜色",
  "报告内版单号", "面料/物料编号", "样品类型", "来源判定", "PDF页码", "表序号", "项目序号",
  "项目类别/表标题", "检测项目", "统一检测父项", "标准检测子项", "检测方法/判定依据", "技术要求/限值",
  "实测/测试结果", "单项判定/结论", "单位", "是否有国际认证", "CMA标识", "CNAS标识", "CAS号",
  "报告出具机构名称", "报告限/检出限", "备注", "项目原始行/页面完整原文", "检测报告来源URL",
  "简体检测项目", "归并状态", "归并证据",
  "报告签发日期识别状态", "报告签发日期异常原因", "CMA识别证据/异常原因", "CNAS识别证据/异常原因",
  "是否进入总览", "检测分类", "本地PDF路径", "解析器版本", "规则版本", "诊断信息",
  "处理状态", "来源关系ID", "来源工作表", "来源单元格",
];

const grayAddedDetailHeaders = new Set([
  "报告签发日期识别状态", "报告签发日期异常原因", "CMA识别证据/异常原因", "CNAS识别证据/异常原因",
  "处理状态", "来源关系ID", "来源工作表", "来源单元格", "简体检测项目", "归并状态", "归并证据",
  "是否进入总览", "检测分类", "本地PDF路径", "解析器版本", "规则版本", "诊断信息",
]);

const rowsBySku = new Map();
const detailRows = [];

function ensureSku(sku) {
  const key = compactText(sku);
  if (!rowsBySku.has(key)) {
    rowsBySku.set(key, {
      sku: key,
      subcategories: new Set(),
      parents: new Set(),
      basic: new Set(),
      functional: new Set(),
      urls: new Set(),
      sourceReportNos: new Set(),
      reportNos: new Set(),
      fileNames: new Set(),
      plateNumbers: new Set(),
      sampleTypes: new Set(),
      judgments: [],
      pending: 0,
      errorUrls: new Set(),
    });
  }
  return rowsBySku.get(key);
}

function pushDetail(row, overrides = {}) {
  const source = sourceFor(row);
  const meta = reportMetadata(row.url);
  const internalId = reportInternalId(row.report_no || meta.report_no, row.url);
  const reportProductCode = compactText(row.report_product_code) || joinedMetadata(meta, "report_product_codes");
  const plateNumbers = cleanPlateNumbers(compactText(row.plate_number) || joinedMetadata(meta, "plate_numbers"));
  const materialNumber = compactText(row.material_number) || joinedMetadata(meta, "material_numbers");
  const parentItem = compactText(row.standard_item);
  const mappingStatus = compactText(row.status);
  const includeInOverview = Boolean(row.include_horizontal) && !["待确认归属", "待复核", "解析残片"].includes(mappingStatus);
  const classification = classify(parentItem, mappingStatus, compactText(row.classification));
  const cmaMark = row.cma_mark || meta.cma_mark || "";
  const cnasMark = row.cnas_mark || meta.cnas_mark || "";
  const reportDate = dateFields(row, meta);

  const metadataDiagnostics = row.metadata_diagnostics || JSON.stringify({
    reportIssueDateLabel: meta.report_issue_date_label || "",
    reportIssueDateOriginal: meta.report_issue_date_original || "",
    reportIssueDateReason: meta.report_issue_date_reason || "",
    cmaEvidence: meta.cma_evidence || "",
    cnasEvidence: meta.cnas_evidence || "",
  });
  const detailRecord = {
    "记录ID": row.record_id || detailRows.length + 1,
    "记录类型": overrides.recordType || row.record_type || "检测结果明细",
    "处理状态": row.processing_status || source.processing_status || "",
    "来源关系ID": relationshipId(source),
    "来源工作表": source.source_sheet || "",
    "来源表行": source.source_row || "",
    "来源单元格": source.source_cell || "",
    "报告序号": reportIndexByUrl.get(row.url) || "",
    "报告单号": internalId,
    "检测报告文件名": localFileName(row.url),
    "报告内编号": internalId,
    "来源货号/款号": row.sku || "",
    "小类": row.subcategory || source.subcategory || "",
    "报告签发日期": reportDate.value,
    "报告签发日期识别状态": reportDate.status,
    "报告签发日期异常原因": reportDate.reason,
    "报告内货号/款号": reportProductCode,
    "颜色": row.color || "",
    "报告内版单号": plateNumbers,
    "面料/物料编号": materialNumber,
    "样品类型": source.sample_type || "",
    "来源判定": normalizedJudgment([source.overall_result]),
    "PDF页码": row.page || "",
    "表序号": row.table || "",
    "项目序号": row.item_number || "",
    "项目类别/表标题": row.table_title || "",
    "检测项目": row.raw_item || row.simple_item || parentItem,
    "统一检测父项": parentItem,
    "标准检测子项": row.subitem || "",
    "检测方法/判定依据": row.method || "",
    "技术要求/限值": row.requirement || "",
    "实测/测试结果": row.result || "",
    "单项判定/结论": row.verdict || "",
    "单位": row.unit || "",
    "是否有国际认证": cmaMark === "有" || cnasMark === "有" ? "是" : "否",
    "CMA标识": cmaMark,
    "CMA识别证据/异常原因": row.cma_recognition_note || meta.cma_evidence || "",
    "CNAS标识": cnasMark,
    "CNAS识别证据/异常原因": row.cnas_recognition_note || meta.cnas_evidence || "",
    "CAS号": row.cas_number || "",
    "报告出具机构名称": row.institution || meta.institution || "",
    "报告限/检出限": row.detection_limit || "",
    "备注": row.result_detail || overrides.note || "",
    "项目原始行/页面完整原文": row.raw_row || row.raw_item || "",
    "检测报告来源URL": row.url || "",
    "简体检测项目": row.simple_item || "",
    "归并状态": mappingStatus || overrides.mappingStatus || "",
    "归并证据": row.merge_evidence || overrides.mappingEvidence || "",
    "是否进入总览": includeInOverview ? "是" : "否",
    "检测分类": classification,
    "本地PDF路径": row.local_path || meta.path || "",
    "解析器版本": row.parser_version || parserVersion,
    "规则版本": row.rule_version || ruleVersion,
    "诊断信息": overrides.diagnostics || JSON.stringify({
      metadata: JSON.parse(metadataDiagnostics || "{}"),
      mergeConfidence: row.merge_confidence ?? "",
      fullStandardItem: row.full_standard_item || "",
    }),
  };
  detailRows.push(detailHeaders.map((header) => detailRecord[header] ?? ""));
}

for (const row of data.detail_rows || []) {
  const skuAgg = ensureSku(row.sku);
  const source = sourceFor(row);
  if (row.subcategory || source.subcategory) skuAgg.subcategories.add(row.subcategory || source.subcategory);
  const parentItem = compactText(row.standard_item);
  const mappingStatus = compactText(row.status);
  const includeInOverview = Boolean(row.include_horizontal) && !["待确认归属", "待复核", "解析残片"].includes(mappingStatus);
  const classification = classify(parentItem, mappingStatus, compactText(row.classification));

  if (includeInOverview && parentItem) {
    skuAgg.parents.add(parentItem);
    if (classification === "功能性检测") skuAgg.functional.add(parentItem);
    if (classification === "基础检测") skuAgg.basic.add(parentItem);
  }
  if (row.url) {
    skuAgg.urls.add(row.url);
    skuAgg.fileNames.add(localFileName(row.url));
    const meta = reportMetadata(row.url);
    const plateText = cleanPlateNumbers(compactText(row.plate_number) || joinedMetadata(meta, "plate_numbers"));
    for (const plate of plateText.split("\n")) {
      if (plate) skuAgg.plateNumbers.add(plate);
    }
  }
  if (row.report_no || reportMetadata(row.url).report_no) skuAgg.reportNos.add(row.report_no || reportMetadata(row.url).report_no);
  if (row.source_order_no || source.order_no) skuAgg.sourceReportNos.add(row.source_order_no || source.order_no);
  if (source.sample_type) skuAgg.sampleTypes.add(source.sample_type);
  if (source.overall_result) skuAgg.judgments.push(source.overall_result);
  if (reviewRequired(mappingStatus)) skuAgg.pending += 1;

  pushDetail(row);
}

for (const error of data.errors || []) {
  const skuAgg = ensureSku(error.sku);
  const source = sourceFor(error);
  if (error.subcategory || source.subcategory) skuAgg.subcategories.add(error.subcategory || source.subcategory);
  if (error.url) {
    skuAgg.urls.add(error.url);
    skuAgg.fileNames.add(localFileName(error.url));
    skuAgg.errorUrls.add(error.url);
  }
  if (error.report_no || reportMetadata(error.url).report_no) skuAgg.reportNos.add(error.report_no || reportMetadata(error.url).report_no);
  if (error.source_order_no || source.order_no) skuAgg.sourceReportNos.add(error.source_order_no || source.order_no);
  if (source.sample_type) skuAgg.sampleTypes.add(source.sample_type);
  if (source.overall_result) skuAgg.judgments.push(source.overall_result);

  pushDetail({
    record_id: error.record_id || "",
    record_type: error.record_type || "异常",
    processing_status: error.processing_status || "",
    sku: error.sku || "",
    subcategory: error.subcategory || source.subcategory || "",
    color: error.color || "",
    source_order_no: error.source_order_no || "",
    source_relationship_id: error.source_relationship_id || "",
    source_sheet: error.source_sheet || "",
    source_row: error.source_row || "",
    source_cell: error.source_cell || "",
    sample_type: error.sample_type || "",
    report_no: error.report_no || "",
    raw_item: error.raw_item || "",
    simple_item: error.simple_item || "",
    standard_item: error.suggested_item || "",
    subitem: "",
    method: "",
    requirement: "",
    result: "",
    verdict: "",
    unit: "",
    result_detail: error.detail || "",
    page: error.page || "",
    status: "待复核",
    include_horizontal: false,
    local_path: error.local_path || "",
    parser_version: error.parser_version || "",
    rule_version: error.rule_version || "",
    url: error.url || "",
  }, {
    recordType: error.record_type || "异常",
    note: error.detail || "",
    mappingStatus: "待复核",
    mappingEvidence: error.detail || "",
    diagnostics: error.diagnostic_metrics || "",
  });
}

const overviewHeaders = [
  "货号",
  "小类",
  "检测项目总数",
  "基础检测项目数",
  "基础检测项目清单",
  "功能性检测项目数",
  "功能性检测项目清单",
  "涉及报告数",
  "报告单号",
  "检测报告文件名",
  "对应版单号",
  "样品类型",
  "总体判定",
  "待复核项目数",
  "异常报告数",
];

const addedOverviewHeaders = new Set(["待复核项目数", "异常报告数"]);

const overviewRows = [...rowsBySku.values()]
  .filter((row) => row.sku)
  .sort((a, b) => a.sku.localeCompare(b.sku, "zh-Hans-CN"))
  .map((row) => [
    row.sku,
    normalizeItemList(row.subcategories),
    row.parents.size,
    row.basic.size,
    normalizeItemList(row.basic),
    row.functional.size,
    normalizeItemList(row.functional),
    row.urls.size,
    uniqueJoined([...row.reportNos]),
    uniqueJoined([...row.fileNames]),
    uniqueJoined([...row.plateNumbers]),
    uniqueJoined([...row.sampleTypes]),
    normalizedJudgment(row.judgments),
    row.pending,
    row.errorUrls.size,
  ]);

const simplificationVerified = data.traditional_to_simplified_applied === true || data.traditionalToSimplifiedApplied === true;
const simplificationAuditText = simplificationVerified
  ? "适配器声明已执行繁体转简体识别并保留原文。"
  : "适配器未提供繁体转简体处理完成证据。";
const auditRecord = {
  "记录ID": detailRows.length + 1,
  "记录类型": "审计摘要",
  "项目序号": "导出摘要",
  "备注": `generatedAt=${generatedAt}`,
  "项目原始行/页面完整原文": `领导版两表导出；${simplificationAuditText}报告级元数据诊断列统一置于AM列之后。`,
  "是否有国际认证": "否",
  "是否进入总览": "否",
  "解析器版本": parserVersion,
  "规则版本": ruleVersion,
  "诊断信息": JSON.stringify({
    sampleSkus: data.sample_skus?.length ?? 0,
    sampleRecords: data.sample_records ?? 0,
    uniqueReports: data.sample_urls?.length ?? 0,
    structuredItems: data.detail_rows?.length ?? 0,
    exceptionCount: data.errors?.length ?? 0,
    stats: data.stats || {},
  }),
};
detailRows.push(detailHeaders.map((header) => auditRecord[header] ?? ""));

const workbook = Workbook.create();
const overviewSheet = workbook.worksheets.add("货号检测总览");
const detailSheet = workbook.worksheets.add("质检全项目明细");

function writeSheet(sheet, headers, rows, options) {
  sheet.showGridLines = false;
  const preRows = options.preRows || [];
  const matrix = [...preRows, headers, ...rows];
  const headerRowIndex = preRows.length;
  const lastColumn = options.lastColumn || excelColumnName(headers.length);
  for (const col of options.textColumns || []) {
    if (col < 0) continue;
    sheet.getRangeByIndexes(0, col, matrix.length, 1).setNumberFormat("@");
  }
  for (const col of options.numericIdentifierColumns || []) {
    if (col < 0) continue;
    sheet.getRangeByIndexes(0, col, matrix.length, 1).setNumberFormat("0");
  }
  sheet.getRangeByIndexes(0, 0, matrix.length, headers.length).values = matrix;

  if (options.title) {
    sheet.mergeCells(`A1:${lastColumn}1`);
    sheet.getRange("A1").format.fill.color = "#17365D";
    sheet.getRange("A1").format.font.color = "#FFFFFF";
    sheet.getRange("A1").format.font.bold = true;
    sheet.getRange("A1").format.font.size = 18;
    sheet.getRange("A1").format.horizontalAlignment = "Left";
    sheet.getRange("A1").format.verticalAlignment = "Center";
    sheet.getRange("A1").format.rowHeight = 34;
  }

  const fullRange = sheet.getRangeByIndexes(0, 0, matrix.length, headers.length);
  fullRange.format.font.name = "Carlito";
  fullRange.format.font.size = 11;
  fullRange.format.wrapText = true;
  if (options.title) sheet.getRange("A1").format.font.size = 18;

  const headerRange = sheet.getRangeByIndexes(headerRowIndex, 0, 1, headers.length);
  headerRange.format.fill.color = "#1F4E78";
  headerRange.format.font.color = "#FFFFFF";
  headerRange.format.font.bold = true;
  headerRange.format.horizontalAlignment = "Center";
  headerRange.format.verticalAlignment = "Center";
  headerRange.format.rowHeight = options.headerRowHeight;

  if (rows.length) {
    const tableAddress = sheet.getRangeByIndexes(headerRowIndex, 0, rows.length + 1, headers.length).address;
    const table = sheet.tables.add(tableAddress, true, options.tableName);
    table.style = "TableStyleMedium2";
    table.showFilterButton = true;
    const bodyRange = sheet.getRangeByIndexes(headerRowIndex + 1, 0, rows.length, headers.length);
    bodyRange.format.rowHeight = options.dataRowHeight;
    bodyRange.format.verticalAlignment = "Top";
    bodyRange.format.wrapText = true;
    for (const col of options.centerColumns || []) {
      if (col < 0) continue;
      sheet.getRangeByIndexes(headerRowIndex + 1, col, rows.length, 1).format.horizontalAlignment = "Center";
    }
  }

  for (const [col, width] of Object.entries(options.widths || {})) {
    sheet.getRangeByIndexes(0, Number(col), matrix.length, 1).format.columnWidth = width;
  }
  for (const col of options.dateColumns || []) {
    if (col < 0) continue;
    sheet.getRangeByIndexes(headerRowIndex + 1, col, rows.length, 1).setNumberFormat("yyyy-mm-dd");
  }
  for (const header of options.addedHeaders || []) {
    const col = headers.indexOf(header);
    if (col < 0) continue;
    const columnRange = sheet.getRangeByIndexes(headerRowIndex, col, rows.length + 1, 1);
    columnRange.format.fill.color = "#F2F2F2";
    columnRange.format.font.color = "#666666";
    const addedHeader = sheet.getRangeByIndexes(headerRowIndex, col, 1, 1);
    addedHeader.format.fill.color = "#D9D9D9";
    addedHeader.format.font.color = "#404040";
    addedHeader.format.font.bold = true;
  }

  for (const rule of options.conditionalFormats || []) {
    if (rule.column < 0) continue;
    const range = sheet.getRangeByIndexes(headerRowIndex + 1, rule.column, rows.length, 1);
    range.conditionalFormats.add("containsText", {
      text: rule.text,
      format: {
        fill: rule.fill,
        font: { color: rule.fontColor, bold: rule.bold },
      },
    });
  }
}

writeSheet(overviewSheet, overviewHeaders, overviewRows, {
  title: true,
  tableName: "LeaderOverview",
  textColumns: ["货号", "报告单号", "对应版单号"].map((header) => overviewHeaders.indexOf(header)),
  numericIdentifierColumns: [],
  headerRowHeight: 38,
  dataRowHeight: 58,
  centerColumns: ["货号", "小类", "检测项目总数", "功能性检测项目数", "涉及报告数", "样品类型", "总体判定", "待复核项目数", "异常报告数"].map((header) => overviewHeaders.indexOf(header)),
  addedHeaders: addedOverviewHeaders,
  preRows: [
    ["货号检测总览｜每个货号对应的检测项目", ...Array(overviewHeaders.length - 1).fill("")],
    [`共 ${overviewRows.length} 个货号；项目已去重。功能性检测按远红外/升温、保温、凉感、抗菌、防紫外、防泼水、拒水拒油、防钻绒、透湿透气、防静电、阻燃等功能关键词归类；其余归基础检测。浅灰列为相对领导原始版本补足字段。`, ...Array(overviewHeaders.length - 1).fill("")],
    Array(overviewHeaders.length).fill(""),
  ],
  widths: Object.fromEntries(overviewHeaders.map((header, index) => [index, ({
    "货号": 18, "小类": 16, "检测项目总数": 12, "基础检测项目数": 14, "基础检测项目清单": 68,
    "功能性检测项目数": 15, "功能性检测项目清单": 52, "涉及报告数": 12, "报告单号": 24,
    "检测报告文件名": 46, "对应版单号": 28, "样品类型": 14, "总体判定": 14,
    "待复核项目数": 14, "异常报告数": 14,
  })[header] || 14])),
  conditionalFormats: [
    { column: overviewHeaders.indexOf("总体判定"), text: "不合格", fill: "#FFC7CE", fontColor: "#9C0006", bold: true },
    { column: overviewHeaders.indexOf("总体判定"), text: "合格", fill: "#C6EFCE", fontColor: "#006100", bold: false },
  ],
});

const overviewLastColumn = excelColumnName(overviewHeaders.length);
overviewSheet.mergeCells(`A2:${overviewLastColumn}2`);
overviewSheet.getRange("A2").format.fill.color = "#FFF4CE";
overviewSheet.getRange("A2").format.font.color = "#7A4E00";
overviewSheet.getRange("A2").format.font.italic = true;
overviewSheet.getRange("A2").format.rowHeight = 28;

writeSheet(detailSheet, detailHeaders, detailRows, {
  title: true,
  tableName: "LeaderDetail",
  textColumns: ["报告单号", "报告内编号", "来源货号/款号", "报告内货号/款号", "报告内版单号", "面料/物料编号", "CAS号"].map((header) => detailHeaders.indexOf(header)),
  numericIdentifierColumns: [],
  dateColumns: [detailHeaders.indexOf("报告签发日期")],
  headerRowHeight: 44,
  dataRowHeight: 54,
  centerColumns: ["记录ID", "记录类型", "处理状态", "来源表行", "报告序号", "报告单号", "来源货号/款号", "报告签发日期", "报告签发日期识别状态", "报告内货号/款号", "颜色", "样品类型", "来源判定", "PDF页码", "表序号", "项目序号", "单项判定/结论", "单位", "是否有国际认证", "CMA标识", "CNAS标识", "CAS号", "归并状态", "是否进入总览", "检测分类"].map((header) => detailHeaders.indexOf(header)),
  addedHeaders: grayAddedDetailHeaders,
  preRows: [
    [`质检报告全项目明细库｜${data.sample_urls?.length ?? 0}份报告`, ...Array(detailHeaders.length - 1).fill("")],
    ["报告数量", data.sample_urls?.length ?? 0, "", "结构化项目行", data.detail_rows?.length ?? 0, "", "异常记录", data.errors?.length ?? 0, "", "总记录数", detailRows.length, ...Array(detailHeaders.length - 11).fill("")],
    [`使用说明：${simplificationAuditText}统一检测父项、标准检测子项紧跟检测项目；国际认证列位于CMA/CNAS标识之前；四个报告级识别诊断灰列统一置于AM列之后。`, ...Array(detailHeaders.length - 1).fill("")],
    Array(detailHeaders.length).fill(""),
  ],
  widths: Object.fromEntries(detailHeaders.map((header, index) => [index, ({
    "记录ID": 8, "记录类型": 18, "处理状态": 16, "来源工作表": 18, "来源表行": 9, "来源单元格": 12,
    "报告序号": 9, "报告单号": 18, "检测报告文件名": 38, "报告内编号": 22, "来源货号/款号": 22,
    "报告签发日期": 14, "报告签发日期识别状态": 16, "报告签发日期异常原因": 42,
    "报告内货号/款号": 28, "颜色": 18, "报告内版单号": 28, "面料/物料编号": 24,
    "样品类型": 11, "来源判定": 11, "PDF页码": 9, "表序号": 8, "项目序号": 9, "项目类别/表标题": 28,
    "检测项目": 30, "统一检测父项": 28, "标准检测子项": 28, "检测方法/判定依据": 38, "技术要求/限值": 34,
    "实测/测试结果": 38, "单项判定/结论": 16, "单位": 12, "是否有国际认证": 14, "CMA标识": 11, "CMA识别证据/异常原因": 42,
    "CNAS标识": 11, "CNAS识别证据/异常原因": 42, "CAS号": 20,
    "报告出具机构名称": 34, "报告限/检出限": 22, "备注": 40, "项目原始行/页面完整原文": 48,
    "检测报告来源URL": 60, "简体检测项目": 28, "归并状态": 16, "归并证据": 32, "是否进入总览": 14,
    "检测分类": 14, "本地PDF路径": 48, "解析器版本": 32, "规则版本": 24, "诊断信息": 48,
  })[header] || 18])),
  conditionalFormats: [
    { column: detailHeaders.indexOf("记录类型"), text: "检测结果明细", fill: "#E2F0D9", fontColor: "#375623", bold: true },
    { column: detailHeaders.indexOf("记录类型"), text: "项目判定汇总", fill: "#DDEBF7", fontColor: "#1F4E78", bold: true },
    { column: detailHeaders.indexOf("记录类型"), text: "补充检测结果", fill: "#E4DFEC", fontColor: "#60497A", bold: true },
    { column: detailHeaders.indexOf("记录类型"), text: "页面完整原文", fill: "#E7E6E6", fontColor: "#666666", bold: false },
    { column: detailHeaders.indexOf("单项判定/结论"), text: "不符合", fill: "#FFC7CE", fontColor: "#9C0006", bold: true },
    { column: detailHeaders.indexOf("单项判定/结论"), text: "符合", fill: "#C6EFCE", fontColor: "#006100", bold: false },
    { column: detailHeaders.indexOf("检测项目"), text: "远红外", fill: "#FCE4D6", fontColor: "#C65911", bold: true },
  ],
});

detailSheet.getRange("A2:K2").format.verticalAlignment = "Center";
for (const labelColumn of [0, 3, 6, 9]) {
  const label = detailSheet.getRangeByIndexes(1, labelColumn, 1, 1);
  label.format.fill.color = "#D9EAF7";
  label.format.font.color = "#17365D";
  label.format.font.bold = true;
  label.format.horizontalAlignment = "Center";
  const value = detailSheet.getRangeByIndexes(1, labelColumn + 1, 1, 1);
  value.format.fill.color = "#F3F8FC";
  value.format.font.color = "#17365D";
  value.format.font.bold = true;
  value.format.font.size = 13;
  value.format.horizontalAlignment = "Center";
}
const detailLastColumn = excelColumnName(detailHeaders.length);
detailSheet.mergeCells(`A3:${detailLastColumn}3`);
detailSheet.getRange("A3").format.fill.color = "#FFF4CE";
detailSheet.getRange("A3").format.font.color = "#7A4E00";
detailSheet.getRange("A3").format.font.italic = true;
detailSheet.getRange("A3").format.verticalAlignment = "Center";
detailSheet.getRange("A3").format.rowHeight = 32;

const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 20 },
  maxChars: 4000,
});
console.log(errorScan.ndjson);
const formulaErrorCount = compactText(errorScan.ndjson).split(/\n+/).filter(Boolean).length;

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);

function paethPredictor(left, above, upperLeft) {
  const estimate = left + above - upperLeft;
  const leftDistance = Math.abs(estimate - left);
  const aboveDistance = Math.abs(estimate - above);
  const upperLeftDistance = Math.abs(estimate - upperLeft);
  if (leftDistance <= aboveDistance && leftDistance <= upperLeftDistance) return left;
  if (aboveDistance <= upperLeftDistance) return above;
  return upperLeft;
}

function inspectPng(bytes) {
  if (bytes.length < 24 || bytes[0] !== 0x89 || bytes[1] !== 0x50 || bytes[2] !== 0x4e || bytes[3] !== 0x47) {
    return { width: 0, height: 0, contentDetected: false, nonWhiteSamples: 0 };
  }
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  const bitDepth = bytes[24];
  const colorType = bytes[25];
  const interlace = bytes[28];
  const channels = ({ 0: 1, 2: 3, 4: 2, 6: 4 })[colorType] || 0;
  if (width <= 1 || height <= 1) {
    return { width, height, contentDetected: false, nonWhiteSamples: 0 };
  }
  if (!width || !height || bitDepth !== 8 || !channels || interlace !== 0) {
    return { width, height, contentDetected: false, nonWhiteSamples: 0 };
  }

  const idatChunks = [];
  let offset = 8;
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset);
    const type = String.fromCharCode(...bytes.slice(offset + 4, offset + 8));
    if (offset + 12 + length > bytes.length) break;
    if (type === "IDAT") idatChunks.push(Buffer.from(bytes.slice(offset + 8, offset + 8 + length)));
    offset += 12 + length;
    if (type === "IEND") break;
  }
  if (!idatChunks.length) return { width, height, contentDetected: false, nonWhiteSamples: 0 };

  let inflated;
  try {
    inflated = inflateSync(Buffer.concat(idatChunks));
  } catch {
    return { width, height, contentDetected: false, nonWhiteSamples: 0 };
  }
  const rowBytes = width * channels;
  if (inflated.length < height * (rowBytes + 1)) {
    return { width, height, contentDetected: false, nonWhiteSamples: 0 };
  }

  let prior = new Uint8Array(rowBytes);
  let sourceOffset = 0;
  let sampledPixels = 0;
  let nonWhiteSamples = 0;
  const sampleEvery = Math.max(1, Math.floor((width * height) / 200000));
  const requiredSamples = Math.max(1, Math.min(10, Math.floor((width * height) / sampleEvery / 1000)));
  for (let y = 0; y < height; y += 1) {
    const filter = inflated[sourceOffset];
    sourceOffset += 1;
    const row = new Uint8Array(rowBytes);
    for (let index = 0; index < rowBytes; index += 1) {
      const raw = inflated[sourceOffset + index];
      const left = index >= channels ? row[index - channels] : 0;
      const above = prior[index] || 0;
      const upperLeft = index >= channels ? prior[index - channels] : 0;
      let value = raw;
      if (filter === 1) value += left;
      else if (filter === 2) value += above;
      else if (filter === 3) value += Math.floor((left + above) / 2);
      else if (filter === 4) value += paethPredictor(left, above, upperLeft);
      else if (filter !== 0) return { width, height, contentDetected: false, nonWhiteSamples: 0 };
      row[index] = value & 0xff;
    }
    sourceOffset += rowBytes;
    for (let x = 0; x < width; x += 1) {
      const pixelNumber = y * width + x;
      if (pixelNumber % sampleEvery !== 0) continue;
      const pixelOffset = x * channels;
      const gray = row[pixelOffset];
      const red = colorType === 0 || colorType === 4 ? gray : row[pixelOffset];
      const green = colorType === 0 || colorType === 4 ? gray : row[pixelOffset + 1];
      const blue = colorType === 0 || colorType === 4 ? gray : row[pixelOffset + 2];
      const alpha = colorType === 4 ? row[pixelOffset + 1] : colorType === 6 ? row[pixelOffset + 3] : 255;
      sampledPixels += 1;
      if (alpha > 16 && (red < 245 || green < 245 || blue < 245)) nonWhiteSamples += 1;
      if (nonWhiteSamples >= requiredSamples) {
        return { width, height, contentDetected: true, nonWhiteSamples, sampledPixels };
      }
    }
    prior = row;
  }
  return { width, height, contentDetected: false, nonWhiteSamples, sampledPixels };
}

const previewDiagnostics = {};
const overviewPreview = await workbook.render({ sheetName: "货号检测总览", range: `A1:${overviewLastColumn}24`, scale: 1, format: "png" });
const overviewPreviewBytes = new Uint8Array(await overviewPreview.arrayBuffer());
previewDiagnostics.overview = inspectPng(overviewPreviewBytes);
if (previewDiagnostics.overview.contentDetected) {
  await fs.writeFile(outputPath.replace(/\.xlsx$/i, "_overview_preview.png"), overviewPreviewBytes);
}
const detailPreview = await workbook.render({ sheetName: "质检全项目明细", range: `A1:${detailLastColumn}24`, scale: 1, format: "png" });
const detailPreviewBytes = new Uint8Array(await detailPreview.arrayBuffer());
previewDiagnostics.detail = inspectPng(detailPreviewBytes);
if (previewDiagnostics.detail.contentDetected) {
  await fs.writeFile(outputPath.replace(/\.xlsx$/i, "_detail_preview.png"), detailPreviewBytes);
}
const visualPreviewStatus = Object.values(previewDiagnostics).some(({ contentDetected }) => !contentDetected)
  ? "external_review_required"
  : "passed";
const exportDiagnostics = {
  output: outputPath,
  formulaErrorCount,
  visualPreviewStatus,
  previewDiagnostics,
};
await fs.writeFile(
  outputPath.replace(/\.xlsx$/i, ".export.json"),
  `${JSON.stringify(exportDiagnostics, null, 2)}\n`,
  { flag: "wx" },
);
if (visualPreviewStatus !== "passed") console.warn(JSON.stringify(exportDiagnostics));

console.log(JSON.stringify({
  output: outputPath,
  overviewRows: overviewRows.length,
  detailRows: detailRows.length,
  sampleSkus: data.sample_skus?.length ?? 0,
  sampleRecords: data.sample_records ?? 0,
  samplePdfs: data.sample_urls?.length ?? 0,
  structuredRows: data.detail_rows?.length ?? 0,
  exceptions: data.errors?.length ?? 0,
  formulaErrorCount,
  visualPreviewStatus,
  previewDiagnostics,
}, null, 2));
