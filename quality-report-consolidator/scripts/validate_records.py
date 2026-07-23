import argparse
import json
import re
import unicodedata
from collections import Counter
from pathlib import Path


def text(value):
    return str(value or "").strip()


def is_true(value):
    return value is True or text(value).lower() in {"true", "1", "yes", "是"}


def confirmed_mapping(value):
    status = text(value)
    return status == "已确认" or status.startswith("自动归并") or status.startswith("自动归入") or status == "自动新增"


parser = argparse.ArgumentParser()
parser.add_argument("records", type=Path)
parser.add_argument("--downloads", type=Path)
args = parser.parse_args()

payload = json.loads(args.records.read_text("utf-8"))
report_metadata = {}
if isinstance(payload, dict):
    report_metadata = payload.get("report_metadata_by_url") or {}
    source_index_keys = {
        (text(row.get("url")), text(row.get("sku")))
        for row in payload.get("source_index") or []
        if text(row.get("url"))
    }
    records = []
    for row in payload.get("detail_rows") or []:
        records.append({
            **row,
            "recordType": "检测结果明细",
            "sourceProductCode": row.get("sku", ""),
            "color": row.get("color", ""),
            "rawItem": row.get("raw_item", ""),
            "parentItem": row.get("standard_item", ""),
            "mappingStatus": row.get("status", ""),
            "includeInOverview": row.get("include_horizontal", False),
            "casNumber": row.get("cas_number", ""),
            "detectionLimit": row.get("detection_limit", ""),
        })
    for row in payload.get("errors") or []:
        records.append({**row, "recordType": "异常", "rawItem": row.get("raw_item", "")})
else:
    records = payload
structured = {"检测结果明细", "项目判定汇总", "补充检测结果"}
exception_types = {"异常", "待复核"}

blank_items = []
invalid_overview_rows = []
invalid_parent_names = []
suspects = []
missing_cas = []
missing_detection_limits = []
missing_structured_trace = []
invalid_mark_values = []
pending_institution_rows = []
covered_urls = set()
exception_urls = set()

standard = re.compile(r"(GB/?T|GB\s|FZ/?T|ISO|AATCC|ASTM|QB/?T|SN/?T|EN\s|DIN)", re.I)
outcome = re.compile(r"(符合|不符合|合格|不合格|未检出|不适用|\bND\b|[<>≤≥]?\d)", re.I)

for row in records:
    record_type = text(row.get("recordType"))
    raw_item = text(row.get("rawItem") or row.get("item"))
    normalized_item = text(row.get("parentItem") or row.get("item"))
    url = text(row.get("url"))
    if url:
        covered_urls.add(url)
        if record_type in exception_types:
            exception_urls.add(url)

    if record_type in structured and (not raw_item or not normalized_item):
        blank_items.append(row)
    if record_type in structured:
        sku = text(row.get("sourceProductCode") or row.get("sku"))
        color = text(row.get("color"))
        if not sku or not color or not url:
            missing_structured_trace.append(row)
        elif isinstance(payload, dict) and source_index_keys and (url, sku) not in source_index_keys:
            missing_structured_trace.append(row)
        if text(row.get("institution") or row.get("issuingInstitution")) == "机构名称待确认":
            pending_institution_rows.append(row)

    if is_true(row.get("includeInOverview")) and not confirmed_mapping(row.get("mappingStatus")):
        invalid_overview_rows.append(row)

    if normalized_item:
        has_semantic_char = any(unicodedata.category(ch)[0] in {"L", "N"} for ch in normalized_item)
        has_forbidden_char = any(unicodedata.category(ch) in {"Cc", "Co"} for ch in normalized_item)
        if not has_semantic_char or has_forbidden_char:
            invalid_parent_names.append(row)

    if record_type == "报告信息/表格原文":
        raw = text(row.get("rawRow"))
        if re.match(r"^\d+\s*\|", raw) and standard.search(raw) and outcome.search(raw):
            suspects.append(row)

    raw_context = " ".join(filter(None, (text(row.get("rawItem")), text(row.get("rawRow")))))
    if re.search(r"(?<!\d)\d{2,7}-\d{2}-\d(?!\d)", re.sub(r"\s+", "", raw_context)) and not text(row.get("casNumber")):
        missing_cas.append(row)
    if re.search(r"(?:报告限|检出限|检测限|定量限|LOD|LOQ)\s*[：:=]?\s*[<>≤≥＜＞]?\s*\d", raw_context, re.I) and not text(row.get("detectionLimit")):
        missing_detection_limits.append(row)

metadata_errors = []
for url, metadata in report_metadata.items():
    if metadata.get("cma_mark") not in {"有", "未发现", "待复核"}:
        invalid_mark_values.append((url, "CMA", metadata.get("cma_mark")))
    if metadata.get("cnas_mark") not in {"有", "未发现", "待复核"}:
        invalid_mark_values.append((url, "CNAS", metadata.get("cnas_mark")))
    if metadata.get("report_issue_date_status") not in {"已识别", "未发现", "待复核"}:
        metadata_errors.append((url, f"报告签发日期识别状态无效: {metadata.get('report_issue_date_status')}"))
    if not metadata.get("report_issue_date") and not metadata.get("report_issue_date_reason"):
        metadata_errors.append((url, "报告签发日期缺失且无原因"))
    if metadata.get("report_issue_date_status") in {"未发现", "待复核"} and not metadata.get("report_issue_date_reason"):
        metadata_errors.append((url, "报告签发日期未识别/待复核但无原因"))
    if metadata.get("cma_mark") in {"未发现", "待复核"} and not metadata.get("cma_evidence"):
        metadata_errors.append((url, "CMA未识别/待复核但无证据或原因"))
    if metadata.get("cnas_mark") in {"未发现", "待复核"} and not metadata.get("cnas_evidence"):
        metadata_errors.append((url, "CNAS未识别/待复核但无证据或原因"))
    if not metadata.get("institution"):
        metadata_errors.append((url, "报告出具机构字段为空"))
    if text(metadata.get("institution")) == "机构名称待确认":
        metadata_errors.append((url, "报告出具机构仍为待确认占位值"))

failed_downloads = []
missing_failure_exceptions = []
unaccounted_downloads = []
if args.downloads:
    downloads = json.loads(args.downloads.read_text("utf-8"))
    failed_downloads = [row for row in downloads if row.get("status") not in {"ok", "cached", "downloaded"}]
    failed_urls = {text(row.get("url")) for row in failed_downloads if text(row.get("url"))}
    all_download_urls = {text(row.get("url")) for row in downloads if text(row.get("url"))}
    missing_failure_exceptions = sorted(failed_urls - exception_urls)
    unaccounted_downloads = sorted(all_download_urls - covered_urls)

result = {
    "records": len(records),
    "recordTypes": Counter(row.get("recordType", "") for row in records),
    "structuredRows": sum(row.get("recordType") in structured for row in records),
    "blankStructuredItems": len(blank_items),
    "invalidOverviewRows": len(invalid_overview_rows),
    "invalidParentNames": len(invalid_parent_names),
    "unparsedSuspectRows": len(suspects),
    "reliableCasMissing": len(missing_cas),
    "reliableDetectionLimitMissing": len(missing_detection_limits),
    "structuredTraceMissing": len(missing_structured_trace),
    "invalidMarkValues": len(invalid_mark_values),
    "pendingInstitutionRows": len(pending_institution_rows),
    "reportMetadataErrors": len(metadata_errors),
    "failedDownloads": len(failed_downloads),
    "failedDownloadsWithoutException": len(missing_failure_exceptions),
    "downloadsWithoutRecordOrException": len(unaccounted_downloads),
}
print(json.dumps(result, ensure_ascii=False, indent=2, default=dict))

if any((blank_items, invalid_overview_rows, invalid_parent_names, suspects, missing_cas, missing_detection_limits, missing_structured_trace, invalid_mark_values, pending_institution_rows, metadata_errors, missing_failure_exceptions, unaccounted_downloads)):
    raise SystemExit(2)
