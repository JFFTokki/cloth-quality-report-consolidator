import hashlib


def source_relationship_id(*, workbook, sheet, row, cell, url, sku, color, sample_type):
    parts = (workbook, sheet, row, cell, url, sku, color, sample_type)
    canonical = "\x1f".join(str(value or "").strip() for value in parts)
    return "src-" + hashlib.sha256(canonical.encode("utf-8")).hexdigest()[:20]


def relationship_from_record(record, *, workbook, url, color):
    return source_relationship_id(
        workbook=workbook,
        sheet=record.get("source_sheet", ""),
        row=record.get("source_row", ""),
        cell=record.get("source_cell", ""),
        url=url,
        sku=record.get("sku", ""),
        color=color,
        sample_type=record.get("sample_type", ""),
    )
