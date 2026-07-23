from report_header import (
    CODE_VALUE_PATTERN,
    HEADER_TOP_MIN_Y,
    extract_labeled_values,
    extract_same_line_labeled_values,
    find_cma_combination,
    guess_institution,
    guess_institution_display,
    normalize_institution_display_value,
    has_mark_text,
    has_nonnegative_mark_text,
    select_header_blocks,
    extract_report_issue_date,
)


def main():
    blocks = [
        {"text": "福建省纤维检验中心", "mid_y": 0.91, "x": 0.04},
        {"text": "CMA 220011349576", "mid_y": 0.87, "x": 0.36},
        {"text": "CNAS L1968", "mid_y": 0.86, "x": 0.74},
        {"text": "报告页脚", "mid_y": 0.08, "x": 0.45},
    ]
    selected = select_header_blocks(blocks)
    texts = {block["text"] for block in selected}
    assert "福建省纤维检验中心" in texts
    assert "CMA 220011349576" in texts
    assert "CNAS L1968" in texts
    assert "报告页脚" not in texts
    assert {block["x"] for block in selected} == {0.04, 0.36, 0.74}
    number, degraded_mark = find_cma_combination([
        {"text": "MA", "mid_x": 0.60, "mid_y": 0.79, "confidence": 1},
        {"text": "220011349576", "mid_x": 0.604, "mid_y": 0.69, "confidence": 1},
    ])
    assert (number, degraded_mark) == ("220011349576", "MA")
    assert find_cma_combination([
        {"text": "MA", "mid_x": 0.20, "mid_y": 0.79, "confidence": 1},
        {"text": "220011349576", "mid_x": 0.80, "mid_y": 0.69, "confidence": 1},
    ]) == ("", "")

    # Stylized CMA logos are often recognized as low-confidence MA/MMA while
    # the aligned 12-digit qualification number remains clear.
    assert find_cma_combination([
        {"text": "MA", "mid_x": 0.169, "mid_y": 0.662, "confidence": 0.30, "width": 0.087},
        {"text": "240020349903", "mid_x": 0.169, "mid_y": 0.568, "confidence": 1, "width": 0.097},
    ]) == ("240020349903", "MA")
    assert find_cma_combination([
        {"text": "MMA", "mid_x": 0.724, "mid_y": 0.822, "confidence": 0.30, "width": 0.055},
        {"text": "2300 1134 9614", "mid_x": 0.725, "mid_y": 0.740, "confidence": 1, "width": 0.061},
    ]) == ("230011349614", "MMA")
    assert find_cma_combination([
        {"text": "MRA", "mid_x": 0.60, "mid_y": 0.79, "confidence": 1},
        {"text": "220011349576", "mid_x": 0.60, "mid_y": 0.69, "confidence": 1},
    ]) == ("", "")
    assert find_cma_combination([
        {"text": "MA", "mid_x": 0.60, "mid_y": 0.79, "confidence": 0.3, "width": 0.08},
        {"text": "13800138000", "mid_x": 0.60, "mid_y": 0.69, "confidence": 1, "width": 0.08},
    ]) == ("", "")
    assert find_cma_combination([
        {"text": "MA", "confidence": 0.3},
        {"text": "220011349576", "confidence": 1},
    ]) == ("", "")
    assert find_cma_combination([
        {"text": "MA", "mid_x": 0.50, "mid_y": 0.79, "confidence": 0.3, "width": 0.02},
        {"text": "220011349576", "mid_x": 0.58, "mid_y": 0.69, "confidence": 1, "width": 0.02},
    ]) == ("", "")
    assert has_mark_text("通过CMA资质认定", "CMA")
    assert not has_nonnegative_mark_text("未获得CMA 220011349576", "CMA")
    assert not has_nonnegative_mark_text("CMA不适用本检测项目", "CMA")

    assert guess_institution(
        "广州检验检测认证集团有限公司\t中国认可\n广检集团\tMA\t国际互认"
    ) == "广州检验检测认证集团有限公司"
    assert guess_institution(
        "供应商名称 青岛贵华针织有限公司\n"
        "Intertek Testing Services Ltd., Shanghai\n"
        "上海天祥质量技术服务有限公司\n"
        "上海天祥质量技术服务有限公司"
    ) == "上海天祥质量技术服务有限公司"
    assert guess_institution(
        "中国认可\t安徽古麒绒材股份有限公司检测实验室\nCNAS L12268"
    ) == "安徽古麒绒材股份有限公司检测实验室"
    assert guess_institution("批准签发 深圳市英柏检测技术有限公司") == "深圳市英柏检测技术有限公司"
    assert guess_institution("南京海关工业产品：\t检测中心") == "南京海关工业产品检测中心"
    assert guess_institution("匦\t南京海关工业产品检测中心") == "南京海关工业产品检测中心"
    assert guess_institution("CTI华测检测\n上海华测品标检测技术有限公司") == "上海华测品标检测技术有限公司"
    assert guess_institution("MAug中联品检（福建）检测服务有限公司") == "中联品检（福建）检测服务有限公司"
    assert guess_institution("委托方\t浙江某某检测技术有限公司") == "机构名称待确认"
    assert guess_institution("生产单位\n上海某某质量检测有限公司") == "机构名称待确认"
    assert guess_institution("生产单位/地址\n__QC_SOURCE_BREAK__\n天纺标检测认证股份有限公司") == "天纺标检测认证股份有限公司"
    assert guess_institution("上海高质量服饰有限公司") == "机构名称待确认"
    assert guess_institution("某某设计中心") == "机构名称待确认"
    assert guess_institution("附设机构 国家体育场馆及健身器材质量检验检测中心（浙江）") == "机构名称待确认"
    assert guess_institution("Semir森馬安徽华英新塘羽绒有限公司检测中心") == "安徽华英新塘羽绒有限公司检测中心"
    assert guess_institution("CNAS L1234\t台灣檢驗科技股份有限公司") == "台湾检验科技股份有限公司"
    assert guess_institution("建新（广东）纺织有限公司\n测试中心") == "建新（广东）纺织有限公司测试中心"
    assert guess_institution_display("防伪查询网址 https://kf.cttc.net.cn") == "中纺标检验认证股份有限公司"
    assert guess_institution_display("检测 ECT远东正大") == "远东正大检验集团有限公司"
    assert guess_institution_display("委托方 Semir 森馬") == "未识别"
    assert guess_institution_display("实验室的说明如下") == "未识别"
    assert guess_institution_display("320003-SC26-SM0076 230011349614") == "未识别"
    assert guess_institution_display("只有报告编号 A123456") == "未识别"
    assert guess_institution("生产单位\n\n上海某某质量检测有限公司") == "机构名称待确认"
    assert normalize_institution_display_value("机构名称待确认") == "未识别"
    assert normalize_institution_display_value("") == "未识别"
    assert guess_institution(
        "国家某某产品质量检验检测中心\n"
        "某省纺织产品质量监督检验研究院 批准人：张三\n"
        "某省纺织产品质量监督检验研究院 地址：某市"
    ) == "某省纺织产品质量监督检验研究院"
    issue_date, status, label, original, reason, candidates = extract_report_issue_date({
        "pages": [{
            "page": 1,
            "text": "检测日期 2026.06.10；复检 检测日期 2026.06.12 检测类别 委托检测",
        }]
    })
    assert issue_date == "2026-06-12"
    assert status == "已识别"
    assert label == "检测日期"
    assert "2026-06-10; 2026-06-12" in reason
    assert {"2026-06-10", "2026-06-12"}.issubset({candidate["date"] for candidate in candidates})
    assert extract_labeled_values(
        "产品款号或货号： 款号：202426107130\n版单号：426FZ- BS-925\n年/季度：2026/Q4",
        [r"版单号"],
        CODE_VALUE_PATTERN,
    ) == ["426FZ-BS-925"]
    assert extract_labeled_values(
        "产品款号或货号： 款号：202426107130\n版单号：426FZ-BS-925\n年/季度：2026/Q4",
        [r"版单号"],
        CODE_VALUE_PATTERN,
    ) == ["426FZ-BS-925"]
    assert extract_labeled_values(
        "版单号：426FZ-BS-925 年/季度：2026/Q4 颜色：黄绿色调00334",
        [r"版单号"],
        CODE_VALUE_PATTERN,
    ) == ["426FZ-BS-925"]
    assert extract_same_line_labeled_values(
        "版单号：426FZ- BS-925 特殊补充字段",
        [r"版单号"],
        stop_labels=(r"年/季度", r"颜色"),
    ) == ["426FZ-BS-925"]
    assert extract_same_line_labeled_values(
        "版单号：125FB-UZ-801 面料编号：202134731AD 批次号：",
        [r"版单号"],
    ) == ["125FB-UZ-801"]
    assert extract_same_line_labeled_values(
        "版单号：125FB-UZ-801面料编号：—",
        [r"版单号"],
        stop_labels=(r"面料编号", r"物料编号", r"料号"),
    ) == ["125FB-UZ-801"]
    assert extract_same_line_labeled_values("版单号：A1 中文说明", [r"版单号"]) == ["A1"]
    assert extract_same_line_labeled_values("版单号：ABC 中文说明", [r"版单号"]) == ["ABC"]
    assert extract_same_line_labeled_values("版单号：ABC- 中文说明", [r"版单号"]) == ["ABC"]
    assert extract_same_line_labeled_values(
        "版单号：426FZ-BS-925 年/季度：2026/Q4 颜色：黄绿色调00334",
        [r"版单号"],
        stop_labels=(r"年/季度", r"颜色"),
    ) == ["426FZ-BS-925"]
    print(f"report header contract ok: min_y={HEADER_TOP_MIN_Y}, full_width_blocks={len(selected)}")


if __name__ == "__main__":
    main()
