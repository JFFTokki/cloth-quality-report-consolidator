import argparse
import importlib
import json
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from runtime_env import inspect_ocr_environment, smoke_test_ocr


ROOT = Path(__file__).resolve().parents[2]
OCR_HELPER = Path(__file__).resolve().parent / "macos_vision_ocr.swift"


def check_python_modules():
    results = {}
    for module_name in ("openpyxl", "pdfplumber", "PIL"):
        try:
            module = importlib.import_module(module_name)
            results[module_name] = {
                "ready": True,
                "version": str(getattr(module, "__version__", "unknown")),
            }
        except Exception as exc:
            results[module_name] = {
                "ready": False,
                "error": f"{type(exc).__name__}: {exc}",
            }
    return results


def check_node_exporter():
    exporter = ROOT / "quality-report-consolidator" / "scripts" / "export_leader_workbook.mjs"
    artifact_tool = ROOT / "quality-report-consolidator" / "node_modules" / "@oai" / "artifact-tool" / "index.mjs"
    errors = []
    if not exporter.is_file():
        errors.append(f"缺少导出脚本：{exporter}")
    if not artifact_tool.is_file():
        errors.append(f"缺少 Node 依赖 @oai/artifact-tool：{artifact_tool}")
    node = shutil.which("node")
    if not node:
        errors.append("缺少系统命令 node")
    if errors:
        return {"ready": False, "errors": errors}
    with tempfile.TemporaryDirectory(prefix="qc_export_smoke_") as temp_dir:
        output = Path(temp_dir) / "smoke.xlsx"
        script = (
            f"const m=await import({json.dumps(artifact_tool.as_uri())});"
            "const w=m.Workbook.create();const s=w.worksheets.add('环境自检');"
            "s.getRange('A1').values=[['ok']];"
            "const f=await m.SpreadsheetFile.exportXlsx(w);"
            f"await f.save({json.dumps(str(output))});"
        )
        result = subprocess.run(
            [node, "--input-type=module", "-e", script],
            capture_output=True,
            text=True,
            timeout=60,
        )
        if result.returncode != 0:
            errors.append(f"Node 工作簿导出自检失败：{result.stderr.strip()}")
        elif not output.is_file() or output.stat().st_size < 1000:
            errors.append("Node 工作簿导出自检未生成有效 XLSX")
    return {"ready": not errors, "errors": errors}


def main():
    parser = argparse.ArgumentParser(description="检查质检报告项目运行环境")
    parser.add_argument("--smoke-ocr", action="store_true", help="兼容参数；环境检查始终执行真实PDF渲染和Vision OCR")
    parser.parse_args()

    python_modules = check_python_modules()
    ocr = inspect_ocr_environment(OCR_HELPER)
    node_exporter = check_node_exporter()
    result = {
        "python": {"executable": sys.executable, "version": sys.version.split()[0], "modules": python_modules},
        "ocr": ocr,
        "node_exporter": node_exporter,
    }
    result["ocr_smoke"] = smoke_test_ocr(ocr)
    print(json.dumps(result, ensure_ascii=False, indent=2))

    modules_ready = all(item["ready"] for item in python_modules.values())
    smoke_ready = result["ocr_smoke"]["ready"]
    if not (modules_ready and ocr["ready"] and node_exporter["ready"] and smoke_ready):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
