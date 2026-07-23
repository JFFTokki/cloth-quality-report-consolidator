import os
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
os.environ.setdefault("QC_WORK_DIR", str(ROOT / "tmp" / "qc_random_100_20260716_context_v7"))

import extract_table_items_checkpoint as parser


def main():
    os.environ["QC_TEST_ONLY_SKUS"] = "not-a-sku"
    try:
        parser.parse_requested_skus("QC_TEST_ONLY_SKUS")
        raise AssertionError("invalid target scope must fail closed")
    except ValueError:
        pass
    os.environ["QC_TEST_ONLY_SKUS"] = "200125132201"
    try:
        parser.parse_requested_skus("QC_TEST_ONLY_SKUS", {"200326108102"})
        raise AssertionError("unknown target SKU must fail closed")
    except ValueError:
        pass
    finally:
        os.environ.pop("QC_TEST_ONLY_SKUS", None)

    schema = parser.header_indices([
        "", "", "标准（称）值", "实测值", "单项判定",
    ])
    assert schema["inferred"] is True
    assert schema["item_cols"] == [0]

    item, detail, parent = parser.item_parts_from_row(
        ["5", "除腰", "带外，背部不应有", "符合", ""],
        schema["item_cols"],
        "绳带要求",
    )
    assert (item, detail, parent) == ("", "", "绳带要求")

    item, detail, parent = parser.item_parts_from_row(
        ["水洗尺寸变化率", "长度", "-5.0~+3.0", "0.0", "符合"],
        schema["item_cols"],
        "",
    )
    assert (item, detail, parent) == ("水洗尺寸变化率", "", "水洗尺寸变化率")

    wide_schema = parser.header_indices([
        "", "", "", "标准（称）值", "实测值", "", "单项判定",
    ])
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["接缝性能（缝子纰裂程", "度）（cm）", "", "", "", "", "符合"],
        wide_schema,
        "",
    )
    assert item == "接缝性能（缝子纰裂程度）"
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["外观", "", "符", "合森马提供梭织服", "该样品经规定程序", "", ""],
        wide_schema,
        "水洗后外观质量",
        force_child=True,
    )
    assert item == "水洗后外观质量-外观"
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["外观质量", "符 装", "合森马提供针织服", "试样经规定程序"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 2, "result": 3},
        "水洗后外观质量",
        force_child=True,
    )
    assert item == "水洗后外观质量-外观质量"
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["除腰", "带外，背部不应有", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "result": 2, "verdict": 3},
        "绳带要求",
        force_child=True,
    )
    assert (item, detail, parent) == ("", "", "绳带要求")
    assert parser.inferred_item_fields(
        ["除腰", "带外，背部不应有", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "result": 2, "verdict": 3},
    ) == ("", "除腰带外，背部不应有", "")
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["肩带-应是固定的", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "result": 1, "verdict": 2},
        "绳带要求",
        force_child=True,
    )
    assert (item, detail, parent) == ("", "", "绳带要求")
    assert parser.inferred_item_fields(
        ["肩带-应是固定的", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "result": 1, "verdict": 2},
    ) == ("", "肩带-应是固定的", "")
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["附件抗拉强力符合森马提供针织服装", "要求", "结果"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 1, "result": 2},
        "",
    )
    assert item == "附件抗拉强力"
    assert parser.inferred_item_fields(
        ["附件抗拉强力符合森马提供针织服装", "要求", "结果"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 1, "result": 2},
    ) == ("附件抗拉强力", "符合森马提供针织服装", "")
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["外观质量符装", "要求", "结果"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 1, "result": 2},
        "水洗后外观质量",
        force_child=True,
    )
    assert item == "水洗后外观质量-外观质量"
    assert parser.inferred_item_fields(
        ["直向-~+3.0", "", "0.3", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 1, "result": 2, "verdict": 3},
    ) == ("直向", "~+3.0", "")
    assert parser.inferred_item_fields(
        ["方法A", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "result": 1, "verdict": 2},
    ) == ("", "", "方法A")
    assert parser.split_mixed_item_text("附件抗拉强力符合森马提供针织服装") == (
        "附件抗拉强力", "符合森马提供针织服装", "",
    )
    assert parser.split_mixed_item_text("附件抗拉强力-长度超过75mm") == (
        "附件抗拉强力", "长度超过75mm", "",
    )
    assert parser.split_mixed_item_text("检测时绳带长度超过75mm") == (
        "", "检测时绳带长度超过75mm", "",
    )
    assert parser.split_mixed_item_text("除腰带外，背部不应有") == (
        "", "除腰带外，背部不应有", "",
    )
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["水 洗尺寸变化率（", "%）", "-", "深蓝80861款", "", "-5.0～+3.0", "+0.3", "", "符合"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 5, "result": 6, "verdict": 8},
        "附件抗拉强力（N）",
    )
    assert (item, parent) == ("水洗尺寸变化率（%）", "水洗尺寸变化率（%）")
    assert parser.result_subitems("+ 0.3") == []
    assert parser.complete_split_verdict(["项目", "符", "合"], 2, "合") == "符合"
    assert parser.complete_split_verdict(["项目", "不", "符", "合"], 3, "合") == "不符合"
    assert parser.sanitize_requirement_result("629-2017、GB/T 8630", "-2013") == ("", "")
    assert parser.sanitize_requirement_result("彩色印花面料", "/白色腰头罗纹 GB/T") == ("彩色印花面料", "")
    assert parser.requirement_with_left_fragment(["外观质量", "符", "合森马提供针织服"], 2, "合森马提供针织服") == "符合森马提供针织服"
    assert parser.requirement_with_left_fragment(["", "装", "采购内控标准"], 2, "采购内控标准") == "装采购内控标准"
    assert parser.join_field_text("符", "符合森马提供针织服") == "符合森马提供针织服"
    assert parser.join_field_text("符合森马提供针织服", "装采购内控标准") == "符合森马提供针织服装采购内控标准"
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["符", "合Q/BALABALA", "内容齐全"],
        {"item": 0, "item_cols": [0], "inferred": True, "requirement": 1, "result": 2},
        "使用说明",
        force_child=True,
    )
    assert (item, detail, parent) == ("", "", "使用说明")

    assert not parser.valid_item("方法A")
    assert not parser.valid_item("要求")
    assert not parser.valid_item("除腰带外，背部不应有")
    assert not parser.valid_item("肩带-应是固定的")
    assert not parser.valid_item("外观质量符装")

    split_header = parser.header_indices([
        "", "序号 检测", "项目名称 单位", "技术要求", "", "检测结果", "单项判定", "备注",
    ])
    assert split_header["wrapped_item_cols"] is True
    assert split_header["item_cols"] == [1, 2]
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["", "可分解 2", "致癌芳香胺 mg/kg 染料", "禁用（≤20）", "", "未检出", "符合", "定量限5mg/kg"],
        split_header,
        "",
    )
    assert item == "可分解致癌芳香胺染料"
    item, detail, parent = parser.item_parts_from_inferred_row(
        ["", "1 蓬松度", "In^3/30 （IDFB） g", "≥600", "", "660", "符合", "还原时间=50hours"],
        split_header,
        "",
    )
    assert item == "蓬松度（IDFB）"

    split_requirement = parser.header_indices([
        "标准", "（称）值", "实测值", "", "单项判定",
    ])
    assert split_requirement["requirement"] == 1
    assert split_requirement["item_cols"] == [0]
    split_verdict = parser.header_indices([
        "", "", "标准", "（称）值", "实测值", "单", "", "项判定",
    ])
    assert split_verdict["verdict"] == 7
    assert parser.valid_item("钻绒值-方法A(大箱体)")
    print("table item line-break contract ok")


if __name__ == "__main__":
    main()
