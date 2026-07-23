import ast
from pathlib import Path

source = Path('pipeline/qc_all/build_workbook.py').read_text(encoding='utf-8')
module = ast.parse(source)
assigns = {}
for node in ast.walk(module):
    if isinstance(node, ast.Assign):
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id.endswith('_headers'):
                if isinstance(node.value, ast.List):
                    assigns[target.id] = ast.literal_eval(node.value)
expected = {
    'detail_headers': ['状态'],
    'mapping_headers': ['首次出现报告号', '首次出现PDF链接'],
    'source_headers': ['异常原因'],
    'error_headers': ['源报告单号', 'PDF报告号', '原始检测项', '简体检测项', '建议统一检测项'],
}
missing = {name: [col for col in cols if col not in assigns.get(name, [])] for name, cols in expected.items()}
missing = {k: v for k, v in missing.items() if v}
assert not missing, missing
print('workbook contract ok')
