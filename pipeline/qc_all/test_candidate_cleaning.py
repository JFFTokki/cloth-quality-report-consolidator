from pathlib import Path

path = Path("pipeline/qc_all/build_data.py")
source = path.read_text(encoding="utf-8")
source = source.replace(
    'source = json.loads((ROOT / "source_records.json").read_text(encoding="utf-8"))',
    'source = {}',
).replace(
    'documents = json.loads((ROOT / "pdf_text.json").read_text(encoding="utf-8"))',
    'documents = []',
)
cut = source.index("parsed_documents =")
namespace = {"__file__": str(path)}
exec(compile(source[:cut], str(path), "exec"), namespace)
fn = namespace["candidate_item_from_line"]

cases = {
    "(级) 2008, A(1), 符合": "",
    "（Cd）,mg/kg 2006(测定低限 ≤0.1 未检出 符合": "",
    "--- cm ≤0.4 符合": "",
    "干摩擦 ≥ 4 4-5 符合": "",
    "镉 = 0.25 mg/kg 符合": "镉",
    "沾色 级 ≥3 4 合格": "",
    "2,4,5-三甲基苯胺 137-17-7 N.D. ≤20 符合": "三甲基苯胺",
    "序号 检验项目 检测方法 标准值 实测值 单项判定": "",
    "检验类别 委托送样 到样日期 2026-06-29 样品状态 符合检验要求": "",
    "除非委托方要求，本报告检测结果及符合性判定不考虑测量结果的不确定度。": "",
    "标准（称）值 实测值 单项判定": "",
    "年份/季节 26Q4 产品等级 合格品": "",
    "样品等级 合格品": "",
    "缝子纰裂 客户要求 合格": "",
    "防泼水性能（拒水性）（洗前） Q/BALABALA 600—2024 合格": "",
    "残脂率 Semir(G)-2/0872020-2026V1.0 符合": "",
}

failures = []
for line, expected in cases.items():
    got = fn(line)
    if got != expected:
        failures.append((line, expected, got))

assert not failures, failures
print("candidate cleaning contract ok")
