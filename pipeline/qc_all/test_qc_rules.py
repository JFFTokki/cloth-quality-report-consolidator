import json

from qc_rules import international_certification, prefer_specific_detail_rows, to_simplified


assert to_simplified("檢測項目 技術要求 檢測結果 單項判定 報告編號 簽發日期") == (
    "检测项目 技术要求 检测结果 单项判定 报告编号 签发日期"
)

summary = {
    "sku": "208426107107", "color": "咖色调00455", "source_order_no": "A",
    "sample_type": "成品",
    "url": "JST-BW202603189SE", "standard_item": "蓬松度", "subitem": "",
    "method": "", "requirement": "", "result": "", "verdict": "合格", "page": 2,
}
detail = {
    **summary,
    "method": "GB/T 14272—2021附录C", "requirement": "≥15.5", "result": "16.3", "page": 4,
}
other_report = {**summary, "url": "OTHER", "page": 2}
other_sample = {**summary, "sample_type": "填充物", "page": 2}
inherited_method_summary = {**summary, "method": "GB/T 14272—2021附录C"}
reference_only = {**summary, "result": "见附表1", "requirement": "见附表1"}
standard_only = {**summary, "result": "符合GB/T 14272-2021"}
not_assessed = {
    **summary,
    "method": "GB/T 30157-2013",
    "verdict": "样品无涂层，重金属项目不考核",
}
different_method = {**summary, "method": "OTHER METHOD"}
kept = prefer_specific_detail_rows([
    summary, detail, other_report, other_sample, inherited_method_summary, reference_only, standard_only,
    not_assessed, different_method,
])
assert kept == [detail, other_report, other_sample, not_assessed, different_method], kept

assert international_certification("有", "未发现") == "是"
assert international_certification("待复核", "有") == "是"
assert international_certification("待复核", "未发现") == "否"

print(json.dumps({"status": "passed", "kept": len(kept)}, ensure_ascii=False))
