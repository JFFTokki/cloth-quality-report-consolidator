import importlib.util
from pathlib import Path

from source_relationship import source_relationship_id


ROOT = Path(__file__).resolve().parents[2]
VALIDATOR_PATH = ROOT / "quality-report-consolidator" / "scripts" / "validate_records.py"
spec = importlib.util.spec_from_file_location("validate_records", VALIDATOR_PATH)
validator = importlib.util.module_from_spec(spec)
spec.loader.exec_module(validator)


def relation(row_number, url="https://example.test/shared.pdf", color="黑色"):
    return source_relationship_id(
        workbook="source.xlsx",
        sheet="Sheet1",
        row=row_number,
        cell=f"J{row_number}",
        url=url,
        sku="200000000001",
        color=color,
        sample_type="成衣",
    )


def metadata():
    return {
        "report_issue_date": "",
        "report_issue_date_status": "未发现",
        "report_issue_date_reason": "测试数据未提供签发日期",
        "institution": "测试检验机构",
        "cma_mark": "未发现",
        "cnas_mark": "未发现",
        "cma_evidence": "测试数据未发现CMA标识",
        "cnas_evidence": "测试数据未发现CNAS标识",
    }


def main():
    first_id = relation(2)
    second_id = relation(3)
    assert first_id != second_id
    source_index = [
        {
            "source_relationship_id": relationship_id,
            "source_sheet": "Sheet1",
            "source_row": row_number,
            "source_cell": f"J{row_number}",
            "url": "https://example.test/shared.pdf",
            "sku": "200000000001",
            "color": "黑色",
            "sample_type": "成衣",
            "processing_status": "已解析",
        }
        for relationship_id, row_number in ((first_id, 2), (second_id, 3))
    ]
    detail_rows = [
        {
            "source_relationship_id": relationship_id,
            "source_sheet": "Sheet1",
            "source_row": row_number,
            "source_cell": f"J{row_number}",
            "sku": "200000000001",
            "color": "黑色",
            "sample_type": "成衣",
            "url": "https://example.test/shared.pdf",
            "raw_item": "pH值",
            "standard_item": "pH值",
            "status": "自动归并",
            "include_horizontal": True,
        }
        for relationship_id, row_number in ((first_id, 2), (second_id, 3))
    ]
    payload = {
        "source_index": source_index,
        "detail_rows": detail_rows,
        "errors": [],
        "report_metadata_by_url": {"https://example.test/shared.pdf": metadata()},
    }
    shared_download = [{"url": "https://example.test/shared.pdf", "status": "downloaded"}]
    complete = validator.validate_payload(payload, downloads=shared_download)
    assert complete["validationStatus"] == "完整通过"
    assert complete["structuredTraceMissing"] == 0

    payload["detail_rows"][1]["source_relationship_id"] = "src-missing"
    broken = validator.validate_payload(payload, downloads=shared_download)
    assert broken["validationStatus"] == "校验失败"
    assert broken["structuredTraceMissing"] == 1
    payload["detail_rows"][1]["source_relationship_id"] = second_id
    payload["detail_rows"][1]["source_row"] = 99
    mismatched_identity = validator.validate_payload(payload, downloads=shared_download)
    assert mismatched_identity["structuredTraceMissing"] == 1
    payload["detail_rows"][1]["source_row"] = 3

    uncovered_payload = {**payload, "detail_rows": detail_rows[:1]}
    uncovered = validator.validate_payload(uncovered_payload, downloads=shared_download)
    assert uncovered["validationStatus"] == "校验失败"
    assert uncovered["uncoveredSourceRelationships"] == 1

    audit_cover_payload = {
        **payload,
        "schemaVersion": "quality-report-consolidator/v1",
        "records": [
            {
                **detail_rows[0],
                "recordType": "检测结果明细",
                "sourceProductCode": "200000000001",
                "rawItem": "pH值",
                "parentItem": "pH值",
                "mappingStatus": "自动归并",
                "processStatus": "succeeded",
            },
            {
                "recordType": "审计摘要",
                "sourceRelationshipId": second_id,
                "url": "https://example.test/shared.pdf",
            },
        ],
    }
    audit_cover = validator.validate_payload(audit_cover_payload, downloads=shared_download)
    assert audit_cover["validationStatus"] == "校验失败"
    assert audit_cover["uncoveredSourceRelationships"] == 1

    exception_payload = {
        "source_index": source_index,
        "detail_rows": detail_rows,
        "errors": [{
            "source_relationship_id": first_id,
            "source_sheet": "Sheet1",
            "source_row": 2,
            "source_cell": "J2",
            "sku": "200000000001",
            "color": "黑色",
            "sample_type": "成衣",
            "url": "https://example.test/shared.pdf",
            "type": "待复核",
        }],
        "report_metadata_by_url": {"https://example.test/shared.pdf": metadata()},
    }
    exception_complete = validator.validate_payload(exception_payload, downloads=shared_download)
    assert exception_complete["exceptionTraceMissing"] == 0
    exception_payload["errors"][0]["source_relationship_id"] = ""
    exception_broken = validator.validate_payload(exception_payload, downloads=shared_download)
    assert exception_broken["validationStatus"] == "校验失败"
    assert exception_broken["exceptionTraceMissing"] == 1

    failed_url = "https://example.test/missing.pdf"
    failed_id = relation(4, url=failed_url)
    failed_source = {
        "source_relationship_id": failed_id,
        "source_sheet": "Sheet1",
        "source_row": 4,
        "source_cell": "J4",
        "url": failed_url,
        "sku": "200000000001",
        "color": "黑色",
        "sample_type": "成衣",
        "processing_status": "retryable_failed",
    }
    partial_payload = {
        "source_index": source_index[:1] + [failed_source],
        "detail_rows": detail_rows[:1],
        "errors": [{
            **failed_source,
            "type": "下载失败",
        }],
        "report_metadata_by_url": {"https://example.test/shared.pdf": metadata()},
    }
    downloads = shared_download + [{"url": failed_url, "status": "failed", "error": "SSLError: certificate verify failed"}]
    strict = validator.validate_payload(partial_payload, downloads=downloads)
    assert strict["validationStatus"] == "校验失败"
    assert not strict["formalDeliveryAllowed"]
    partial = validator.validate_payload(partial_payload, downloads=downloads, allow_partial=True)
    assert partial["validationStatus"] == "部分完成"
    assert not partial["formalDeliveryAllowed"]
    assert partial["failedDownloadReasons"]["SSLError"] == 1
    assert partial["exceptionTraceMissing"] == 0

    assert validator.validate_payload([], downloads=[])["validationStatus"] == "校验失败"
    assert validator.validate_payload({"unexpected": []}, downloads=[])["payloadShapeInvalid"]
    canonical = {
        "schemaVersion": "quality-report-consolidator/v1",
        "sourceRelationships": [{
            "sourceRelationshipId": first_id,
            "sourceSheet": "Sheet1",
            "sourceRow": 2,
            "sourceCell": "J2",
            "url": "https://example.test/shared.pdf",
            "sourceProductCode": "200000000001",
            "color": "黑色",
            "sampleType": "",
            "processStatus": "succeeded",
        }],
        "records": [{
            "recordType": "检测结果明细",
            "sourceRelationshipId": first_id,
            "sourceSheet": "Sheet1",
            "sourceRow": 2,
            "sourceCell": "J2",
            "url": "https://example.test/shared.pdf",
            "sourceProductCode": "200000000001",
            "color": "黑色",
            "rawItem": "pH值",
            "parentItem": "pH值",
            "mappingStatus": "自动归并",
            "includeInOverview": True,
        }],
        "reportMetadataByUrl": {
            "https://example.test/shared.pdf": {
                "reportIssueDate": "",
                "reportIssueDateStatus": "未发现",
                "reportIssueDateReason": "测试数据未提供签发日期",
                "issuingInstitution": "测试检验机构",
                "cmaMark": "未发现",
                "cnasMark": "未发现",
                "cmaRecognitionNote": "测试数据未发现CMA标识",
                "cnasRecognitionNote": "测试数据未发现CNAS标识",
            },
        },
    }
    assert validator.validate_payload(canonical, downloads=shared_download)["validationStatus"] == "完整通过"
    canonical["records"][0]["processStatus"] = "partial"
    canonical_partial = validator.validate_payload(canonical, downloads=shared_download)
    assert canonical_partial["validationStatus"] == "校验失败"
    assert canonical_partial["recordsWithIncompleteProcessStatus"] == 1
    print("source relationship and validation contract ok")


if __name__ == "__main__":
    main()
