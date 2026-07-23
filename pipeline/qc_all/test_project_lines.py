from pathlib import Path

source_path = Path("pipeline/qc_all/build_data.py")
source = source_path.read_text(encoding="utf-8")
source = source.replace(
    'source = json.loads((ROOT / "source_records.json").read_text(encoding="utf-8"))',
    'source = {}',
).replace(
    'documents = json.loads((ROOT / "pdf_text.json").read_text(encoding="utf-8"))',
    'documents = []',
)
cut = source.index("parsed_documents =")
namespace = {"__file__": str(source_path)}
exec(compile(source[:cut], str(source_path), "exec"), namespace)
project_lines_only = namespace["project_lines_only"]

lines = [
    (1, "检测项目 耐贮存色牢度"),
    (1, "判定依据 森马集团-毛针织服装V2.0 采购内控标准：V2.0"),
    (1, "本报告检验检测结论是根据检验检测依据/判定依据仅对所检项目得出的"),
    (2, "序号 检验项目 判定依据 单项判定"),
    (2, "1 异味 GB 31701—2015 合格"),
    (2, "2 附件锐利性 GB 31701—2015 合格"),
    (2, "备注 ——"),
]

filtered = [line for _, line in project_lines_only(lines)]
assert "检测项目 耐贮存色牢度" in filtered, filtered
assert "1 异味 GB 31701—2015 合格" in filtered, filtered
assert "2 附件锐利性 GB 31701—2015 合格" in filtered, filtered
assert not any("判定依据 森马" in line for line in filtered), filtered
assert not any("本报告检验检测结论" in line for line in filtered), filtered
print("project lines contract ok")
