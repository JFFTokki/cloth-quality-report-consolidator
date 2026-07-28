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


def records_from_payload(payload):
    if not isinstance(payload, dict):
        return payload
    if isinstance(payload.get("records"), list):
        return payload["records"]
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
    return records


def download_error_category(row):
    error = text(row.get("error"))
    if not error:
        return text(row.get("status")) or "unknown"
    return error.split(":", 1)[0][:80]


def validate_payload(payload, downloads=None, allow_partial=False):
    report_metadata = (payload.get("report_metadata_by_url") or payload.get("reportMetadataByUrl") or {}) if isinstance(payload, dict) else {}
    source_index = (payload.get("source_index") or payload.get("sourceRelationships") or []) if isinstance(payload, dict) else []
    payload_shape_invalid = bool(
        isinstance(payload, dict)
        and not any(key in payload for key in ("detail_rows", "errors", "records"))
    )
    canonical_schema_invalid = bool(
        isinstance(payload, dict)
        and "records" in payload
        and payload.get("schemaVersion") != "quality-report-consolidator/v1"
    )
    source_ids = {
        text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
        for row in source_index
        if text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
    }
    source_rows_without_id = [
        row for row in source_index
        if not text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
    ]
    source_identity_by_id = {
        text(row.get("source_relationship_id") or row.get("sourceRelationshipId")): (
            text(row.get("url")), text(row.get("sku") or row.get("sourceProductCode")), text(row.get("color")),
            text(row.get("source_sheet") or row.get("sourceSheet")),
            text(row.get("source_row") or row.get("sourceRow")),
            text(row.get("source_cell") or row.get("sourceCell")),
            text(row.get("sample_type") or row.get("sampleType")),
        )
        for row in source_index
        if text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
    }
    source_status_by_id = {
        text(row.get("source_relationship_id") or row.get("sourceRelationshipId")):
            text(row.get("processing_status") or row.get("processStatus"))
        for row in source_index
        if text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
    }
    source_url_sku_keys = {
        (text(row.get("url")), text(row.get("sku") or row.get("sourceProductCode")))
        for row in source_index if text(row.get("url"))
    }
    records = records_from_payload(payload)
    structured = {"检测结果明细", "项目判定汇总", "补充检测结果"}
    exception_types = {"异常", "待复核"}

    blank_items = []
    invalid_overview_rows = []
    invalid_parent_names = []
    suspects = []
    missing_cas = []
    missing_detection_limits = []
    missing_structured_trace = []
    missing_exception_trace = []
    invalid_mark_values = []
    pending_institution_rows = []
    covered_urls = set()
    exception_urls = set()
    covered_relationship_ids = set()
    incomplete_record_process_rows = []

    standard = re.compile(r"(GB/?T|GB\s|FZ/?T|ISO|AATCC|ASTM|QB/?T|SN/?T|EN\s|DIN)", re.I)
    outcome = re.compile(r"(符合|不符合|合格|不合格|未检出|不适用|\bND\b|[<>≤≥]?\d)", re.I)

    for row in records:
        record_type = text(row.get("recordType"))
        raw_item = text(row.get("rawItem") or row.get("item"))
        normalized_item = text(row.get("parentItem") or row.get("item"))
        url = text(row.get("url"))
        relationship_id = text(row.get("source_relationship_id") or row.get("sourceRelationshipId"))
        if relationship_id and record_type in structured | exception_types:
            covered_relationship_ids.add(relationship_id)
        if url:
            covered_urls.add(url)
            if record_type in exception_types:
                exception_urls.add(url)

        if record_type in structured and (not raw_item or not normalized_item):
            blank_items.append(row)
        if record_type in structured or record_type in exception_types:
            sku = text(row.get("sourceProductCode") or row.get("sku"))
            color = text(row.get("color"))
            row_identity = (
                url, sku, color, text(row.get("source_sheet") or row.get("sourceSheet")),
                text(row.get("source_row") or row.get("sourceRow")),
                text(row.get("source_cell") or row.get("sourceCell")),
                text(row.get("sample_type") or row.get("sampleType")),
            )
            trace_missing = record_type in structured and (not sku or not color or not url)
            if source_ids:
                trace_missing = (
                    trace_missing
                    or not relationship_id
                    or relationship_id not in source_ids
                    or source_identity_by_id.get(relationship_id) != row_identity
                )
            elif source_url_sku_keys:
                trace_missing = trace_missing or (bool(url) and (url, sku) not in source_url_sku_keys)
            if trace_missing:
                target = missing_structured_trace if record_type in structured else missing_exception_trace
                target.append(row)
            record_process_status = text(row.get("processing_status") or row.get("processStatus"))
            effective_process_status = record_process_status or source_status_by_id.get(relationship_id, "")
            if record_process_status and record_process_status not in {
                "已解析", "已解析（OCR）", "succeeded", "success", "complete", "completed"
            }:
                incomplete_record_process_rows.append(row)
            elif record_type in structured and not effective_process_status:
                incomplete_record_process_rows.append(row)
            if record_type in structured and text(row.get("institution") or row.get("issuingInstitution")) == "机构名称待确认":
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
        cma_mark = metadata.get("cma_mark") or metadata.get("cmaMark")
        cnas_mark = metadata.get("cnas_mark") or metadata.get("cnasMark")
        date_value = metadata.get("report_issue_date") or metadata.get("reportIssueDate")
        date_status = metadata.get("report_issue_date_status") or metadata.get("reportIssueDateStatus")
        date_reason = metadata.get("report_issue_date_reason") or metadata.get("reportIssueDateReason")
        cma_evidence = metadata.get("cma_evidence") or metadata.get("cmaRecognitionNote")
        cnas_evidence = metadata.get("cnas_evidence") or metadata.get("cnasRecognitionNote")
        institution = metadata.get("institution") or metadata.get("issuingInstitution")
        if cma_mark not in {"有", "未发现", "待复核"}:
            invalid_mark_values.append((url, "CMA", cma_mark))
        if cnas_mark not in {"有", "未发现", "待复核"}:
            invalid_mark_values.append((url, "CNAS", cnas_mark))
        if date_status not in {"已识别", "未发现", "待复核"}:
            metadata_errors.append((url, f"报告签发日期识别状态无效: {date_status}"))
        if not date_value and not date_reason:
            metadata_errors.append((url, "报告签发日期缺失且无原因"))
        if date_status in {"未发现", "待复核"} and not date_reason:
            metadata_errors.append((url, "报告签发日期未识别/待复核但无原因"))
        if cma_mark in {"未发现", "待复核"} and not cma_evidence:
            metadata_errors.append((url, "CMA未识别/待复核但无证据或原因"))
        if cnas_mark in {"未发现", "待复核"} and not cnas_evidence:
            metadata_errors.append((url, "CNAS未识别/待复核但无证据或原因"))
        if not institution:
            metadata_errors.append((url, "报告出具机构字段为空"))
        if text(institution) == "机构名称待确认":
            metadata_errors.append((url, "报告出具机构仍为待确认占位值"))

    downloads_supplied = downloads is not None
    downloads = downloads or []
    failed_downloads = [row for row in downloads if row.get("status") not in {"ok", "cached", "downloaded"}]
    successful_urls = {
        text(row.get("url")) for row in downloads
        if row.get("status") in {"ok", "cached", "downloaded"} and text(row.get("url"))
    }
    incomplete_record_process_rows = [
        row for row in incomplete_record_process_rows
        if text(row.get("url")) in successful_urls
    ]
    failed_urls = {text(row.get("url")) for row in failed_downloads if text(row.get("url"))}
    all_download_urls = {text(row.get("url")) for row in downloads if text(row.get("url"))}
    missing_failure_exceptions = sorted(failed_urls - exception_urls)
    unaccounted_downloads = sorted(all_download_urls - covered_urls)
    records_without_download_manifest = sorted(covered_urls - all_download_urls)
    download_manifest_missing = not downloads_supplied or not downloads
    empty_records = not records
    structured_rows_count = sum(row.get("recordType") in structured for row in records)
    source_relationship_index_missing = not source_index
    uncovered_source_relationships = sorted(source_ids - covered_relationship_ids)
    missing_metadata_urls = sorted(successful_urls - {text(url) for url in report_metadata})
    successful_process_statuses = {"已解析", "已解析（OCR）", "succeeded", "success", "complete", "completed"}
    incomplete_source_relationships = [
        row for row in source_index
        if text(row.get("url")) in successful_urls
        and text(row.get("processing_status") or row.get("processStatus")) not in successful_process_statuses
    ]

    blocking_groups = (
        blank_items, invalid_overview_rows, invalid_parent_names, suspects, missing_cas,
        missing_detection_limits, missing_structured_trace, invalid_mark_values,
        missing_exception_trace, pending_institution_rows, metadata_errors, missing_failure_exceptions, unaccounted_downloads,
        records_without_download_manifest, source_rows_without_id, uncovered_source_relationships,
        missing_metadata_urls, incomplete_source_relationships, incomplete_record_process_rows,
        ["download manifest missing"] if download_manifest_missing else [],
        ["records empty"] if empty_records else [],
        ["structured records empty"] if not structured_rows_count else [],
        ["source relationship index missing"] if source_relationship_index_missing else [],
        ["payload shape invalid"] if payload_shape_invalid else [],
        ["canonical schema invalid"] if canonical_schema_invalid else [],
    )
    has_data_failure = any(blocking_groups)
    if has_data_failure or (failed_downloads and not allow_partial):
        validation_status = "校验失败"
    elif failed_downloads:
        validation_status = "部分完成"
    else:
        validation_status = "完整通过"

    return {
        "validationStatus": validation_status,
        "formalDeliveryAllowed": validation_status == "完整通过",
        "allowPartial": allow_partial,
        "records": len(records),
        "recordTypes": Counter(row.get("recordType", "") for row in records),
        "structuredRows": structured_rows_count,
        "blankStructuredItems": len(blank_items),
        "invalidOverviewRows": len(invalid_overview_rows),
        "invalidParentNames": len(invalid_parent_names),
        "unparsedSuspectRows": len(suspects),
        "reliableCasMissing": len(missing_cas),
        "reliableDetectionLimitMissing": len(missing_detection_limits),
        "structuredTraceMissing": len(missing_structured_trace),
        "exceptionTraceMissing": len(missing_exception_trace),
        "invalidMarkValues": len(invalid_mark_values),
        "pendingInstitutionRows": len(pending_institution_rows),
        "reportMetadataErrors": len(metadata_errors),
        "failedDownloads": len(failed_downloads),
        "failedDownloadReasons": Counter(download_error_category(row) for row in failed_downloads),
        "failedDownloadsWithoutException": len(missing_failure_exceptions),
        "downloadsWithoutRecordOrException": len(unaccounted_downloads),
        "recordsWithoutDownloadManifest": len(records_without_download_manifest),
        "sourceRelationshipIndexMissing": source_relationship_index_missing,
        "sourceRelationshipsWithoutId": len(source_rows_without_id),
        "uncoveredSourceRelationships": len(uncovered_source_relationships),
        "successfulReportsWithoutMetadata": len(missing_metadata_urls),
        "successfulReportsWithIncompleteProcessStatus": len(incomplete_source_relationships),
        "recordsWithIncompleteProcessStatus": len(incomplete_record_process_rows),
        "downloadManifestMissing": download_manifest_missing,
        "emptyRecords": empty_records,
        "payloadShapeInvalid": payload_shape_invalid,
        "canonicalSchemaInvalid": canonical_schema_invalid,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("records", type=Path)
    parser.add_argument("--downloads", type=Path)
    parser.add_argument("--allow-partial", action="store_true")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    payload = json.loads(args.records.read_text("utf-8"))
    downloads = json.loads(args.downloads.read_text("utf-8")) if args.downloads else []
    result = validate_payload(payload, downloads=downloads, allow_partial=args.allow_partial)
    rendered = json.dumps(result, ensure_ascii=False, indent=2, default=dict)
    print(rendered)
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        temp = args.output.with_suffix(args.output.suffix + ".tmp")
        temp.write_text(rendered + "\n", encoding="utf-8")
        temp.replace(args.output)
    if result["validationStatus"] == "校验失败":
        raise SystemExit(2)


if __name__ == "__main__":
    main()
