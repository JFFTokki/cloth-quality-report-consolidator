import json
import os
import subprocess
import sys
import tempfile
import unicodedata
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
FIXTURE_WORK = ROOT / "tmp" / "qc_random_100_20260716_context_v7"
BUILDER = ROOT / "pipeline" / "qc_all" / "build_100_sku_parent_child_report.py"


def has_private_or_control(value):
    return any(unicodedata.category(char) in {"Cc", "Cf", "Co", "Cs"} for char in str(value or ""))


with tempfile.TemporaryDirectory(prefix="qc-normalization-") as temp_dir:
    data_path = Path(temp_dir) / "report_data.json"
    env = os.environ.copy()
    env.update({
        "QC_WORK_DIR": str(FIXTURE_WORK),
        "QC_DATA_ONLY": "1",
        "QC_DATA_PATH": str(data_path),
    })
    completed = subprocess.run(
        [sys.executable, str(BUILDER)],
        cwd=ROOT,
        env=env,
        text=True,
        capture_output=True,
        check=False,
    )
    if completed.returncode:
        raise AssertionError(completed.stderr or completed.stdout)
    data = json.loads(data_path.read_text(encoding="utf-8"))

rows = data["detail_rows"]

for row in rows:
    for field in ("standard_item", "subitem", "full_standard_item"):
        assert not has_private_or_control(row.get(field)), (field, row)

material_rows = [
    row for row in rows
    if row.get("raw_item") in {"材质鉴别[帮面]", "材质鉴定（帮面）"}
]
assert len(material_rows) == 2, material_rows
assert {row["standard_item"] for row in material_rows} == {"材质鉴定"}, material_rows
assert {row["subitem"] for row in material_rows} == {"帮面"}, material_rows

fat_rows = [row for row in rows if str(row.get("raw_item", "")).startswith("残脂率")]
assert fat_rows, fat_rows
assert {row["standard_item"] for row in fat_rows} == {"残脂率"}, fat_rows

private_aromatic_rows = [row for row in rows if "\uf0a2" in str(row.get("raw_item", ""))]
assert len(private_aromatic_rows) == 6, private_aromatic_rows
assert {row["standard_item"] for row in private_aromatic_rows} == {"可分解致癌芳香胺染料"}, private_aromatic_rows
assert all(row.get("subitem") for row in private_aromatic_rows), private_aromatic_rows

print(json.dumps({
    "status": "passed",
    "material_rows": len(material_rows),
    "fat_rows": len(fat_rows),
    "private_aromatic_rows": len(private_aromatic_rows),
}, ensure_ascii=False))
