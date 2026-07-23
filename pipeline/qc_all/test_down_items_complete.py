"""Regression contract for complete down/feather item names.

This test intentionally uses an actual 2026Q4 report instead of a synthetic table.
It protects the merged-cell relationship between the parent item
``羽绒成分测定`` and each of its eight child items.

Run from the repository root::

    python3 pipeline/qc_all/test_down_items_complete.py
"""

from __future__ import annotations

import importlib.util
import json
import os
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
WORK = ROOT / "tmp/qc_random_100_20260715_v2"
PDF_PATH = WORK / "pdfs_by_url/514391714e45d582fe5ab5d7510ff73e2187ad2f.pdf"
REPORT_NO = "HY26020113"
SOURCE_ORDER_NO = "SM26020500045"
SKU = "208426120205"
COLOR = "白色"

EXPECTED_DOWN_COMPONENTS = {
    "羽绒成分测定-含绒量": "-",
    "羽绒成分测定-含绒量极限偏差": "-",
    "羽绒成分测定-绒子含量": "83.9",
    "羽绒成分测定-长毛片含量": "-",
    "羽绒成分测定-异色毛绒": "0.1",
    "羽绒成分测定-绒丝+羽丝": "6.5",
    "羽绒成分测定-杂质": "0.2",
    "羽绒成分测定-陆禽毛": "0",
}

EXPECTED_INDEPENDENT_ITEMS = {
    "蓬松度": "16.6",
    "浊度": "1000",
    "水分率": "10.5",
    "残脂率": "0.6",
    "耗氧量": "0.8",
    "气味": "-",
    "羽毛羽绒品种-鹅毛绒含量": "-",
}


def _load_target_module(temp_work: Path):
    """Load the parser with a minimal data context for the real HY report."""
    docs = json.loads((WORK / "pdf_text.json").read_text(encoding="utf-8"))
    document = next(doc for doc in docs if doc.get("report_no") == REPORT_NO)
    url = document["url"]
    record = {
        "order_no": SOURCE_ORDER_NO,
        "sku": SKU,
        "selected_colors": [COLOR],
        "urls": [url],
    }
    (temp_work / "source_records.json").write_text(
        json.dumps({"selected_records": [record]}, ensure_ascii=False), encoding="utf-8"
    )
    (temp_work / "download_manifest.json").write_text("[]", encoding="utf-8")
    (temp_work / "pdf_text.json").write_text(
        json.dumps([document], ensure_ascii=False), encoding="utf-8"
    )

    os.environ["QC_WORK_DIR"] = str(temp_work)
    module_path = ROOT / "pipeline/qc_all/extract_table_items_checkpoint.py"
    spec = importlib.util.spec_from_file_location("down_items_parser_under_test", module_path)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module, url


def _rows_by_item(rows):
    by_item = {}
    for row in rows:
        by_item.setdefault(row["item"], []).append(row)
    return by_item


def main():
    assert PDF_PATH.exists(), f"missing regression PDF: {PDF_PATH}"
    sys.path.insert(0, str(ROOT / "tmp/vendor"))
    with tempfile.TemporaryDirectory(prefix="qc_down_contract_") as temp_dir:
        parser, url = _load_target_module(Path(temp_dir))
        rows, stats = parser.process_pdf(
            {"url": url, "path": str(PDF_PATH), "status": "downloaded"}
        )

    by_item = _rows_by_item(rows)
    missing = sorted((EXPECTED_DOWN_COMPONENTS | EXPECTED_INDEPENDENT_ITEMS) - by_item.keys())
    assert not missing, (
        f"{REPORT_NO} missing complete down items: {missing}; "
        f"extracted={sorted(by_item)}; stats={dict(stats)}"
    )

    for item, expected_result in (EXPECTED_DOWN_COMPONENTS | EXPECTED_INDEPENDENT_ITEMS).items():
        matching = [row for row in by_item[item] if row["page"] == 2]
        assert matching, f"{item} was not traced to {REPORT_NO} page 2"
        assert any(row["result"] == expected_result for row in matching), (
            item,
            expected_result,
            [row["result"] for row in matching],
        )
        for row in matching:
            assert row["report_no"] == REPORT_NO
            assert row["source_order_no"] == SOURCE_ORDER_NO
            assert row["sku"] == SKU
            assert row["color"] == COLOR
            assert row["url"] == url

    # A merged-cell child must never escape as an ambiguous standalone item.
    forbidden_standalone = {
        "含绒量",
        "含绒量极限偏差",
        "绒子含量",
        "长毛片含量",
        "异色毛绒",
        "绒丝+羽丝",
        "杂质",
        "陆禽毛",
        "鹅毛绒含量",
    }
    leaked = sorted(forbidden_standalone & by_item.keys())
    assert not leaked, f"merged-cell children lost their parent name: {leaked}"

    # The same report has a second fluffiness method on page 3; it must remain
    # traceable and must not overwrite the page-2 GB/T result.
    page3_fluffiness = [row for row in by_item["蓬松度"] if row["page"] == 3]
    assert any(row["result"] == "660" for row in page3_fluffiness), page3_fluffiness

    print(
        f"down item completeness contract ok: report={REPORT_NO}, "
        f"page2_required={len(EXPECTED_DOWN_COMPONENTS) + len(EXPECTED_INDEPENDENT_ITEMS)}, "
        f"rows={len(rows)}"
    )


if __name__ == "__main__":
    main()
