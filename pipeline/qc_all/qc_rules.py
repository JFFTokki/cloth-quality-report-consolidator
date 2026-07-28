import re

from opencc import OpenCC


_TRADITIONAL_TO_SIMPLIFIED = OpenCC("t2s")
_CONCRETE_RESULT = re.compile(r"\d|未检出|未测出|N\.?D\.?|<|>|≤|≥", re.I)
_REFERENCE_ONLY_RESULT = re.compile(
    r"^(?:符合|合格|通过|见|详见|参见|按|依据|参照)|(?:附表|标准|方法|条款)",
    re.I,
)
_QUALITATIVE_EXCEPTION = re.compile(r"不考核|不适用|免测|无需(?:检测|测试)|未测试|无法测试")


def to_simplified(value):
    """Convert report text before recognition while leaving non-text values unchanged."""
    if not isinstance(value, str) or not value:
        return value
    return _TRADITIONAL_TO_SIMPLIFIED.convert(value)


def simplify_ocr_blocks(blocks):
    for block in blocks or []:
        original = str(block.get("original_text") or block.get("text") or "")
        simplified = to_simplified(original)
        if simplified != original:
            block.setdefault("original_text", original)
        block["text"] = simplified
    return blocks


def simplify_document_text(document):
    """Normalize cached/native/OCR text before any report-level recognition."""
    for key in ("header_native_text", "header_ocr_text"):
        original = str(document.get(f"original_{key}") or document.get(key) or "")
        simplified = to_simplified(original)
        if simplified != original:
            document.setdefault(f"original_{key}", original)
        if original:
            document[key] = simplified
    simplify_ocr_blocks(document.get("header_ocr_blocks") or [])
    for page in document.get("pages") or []:
        original = str(page.get("original_text") or "")
        if not original and page.get("ocr_blocks"):
            original = "\n".join(
                str(block.get("original_text") or block.get("text") or "")
                for block in page.get("ocr_blocks") or []
                if block.get("original_text") or block.get("text")
            )
        if not original:
            original = str(page.get("text") or "")
        simplified = to_simplified(original)
        if simplified != original:
            page.setdefault("original_text", original)
        page["text"] = simplified
        simplify_ocr_blocks(page.get("ocr_blocks") or [])
    return document


def _has_concrete_result(row):
    result = str(row.get("result") or "").strip()
    if not result or _REFERENCE_ONLY_RESULT.search(result):
        return False
    return bool(_CONCRETE_RESULT.search(result))


def prefer_specific_detail_rows(rows):
    """Drop conclusion-only duplicates when the same sample/item has a detailed value row."""
    groups = {}
    for index, row in enumerate(rows):
        key = (
            str(row.get("sku") or ""),
            str(row.get("color") or ""),
            str(row.get("source_order_no") or ""),
            str(row.get("source_sheet") or ""),
            str(row.get("source_row") or ""),
            str(row.get("source_cell") or ""),
            str(row.get("sample_type") or ""),
            str(row.get("url") or row.get("report_no") or ""),
            str(row.get("standard_item") or ""),
            str(row.get("subitem") or ""),
        )
        groups.setdefault(key, []).append((index, row))

    remove = set()
    for entries in groups.values():
        detailed = [
            (index, row) for index, row in entries
            if _has_concrete_result(row) and (row.get("method") or row.get("requirement"))
        ]
        if not detailed:
            continue
        for index, row in entries:
            conclusion_only = not _has_concrete_result(row) and bool(row.get("verdict"))
            qualitative_text = " ".join(
                str(row.get(field) or "") for field in ("result", "verdict", "requirement", "result_detail")
            )
            method = re.sub(r"\s+", "", str(row.get("method") or ""))
            method_matches_detail = not method or any(
                re.sub(r"\s+", "", str(detail_row.get("method") or "")) == method
                for _, detail_row in detailed
            )
            if conclusion_only and not _QUALITATIVE_EXCEPTION.search(qualitative_text) and method_matches_detail:
                remove.add(index)
    return [row for index, row in enumerate(rows) if index not in remove]


def international_certification(cma_mark, cnas_mark):
    return "是" if "有" in {str(cma_mark or "").strip(), str(cnas_mark or "").strip()} else "否"
