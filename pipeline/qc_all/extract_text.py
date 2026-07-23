import json
import os
import re
import subprocess
import sys
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "vendor"))
import pdfplumber
from PIL import Image
from qc_rules import simplify_document_text, simplify_ocr_blocks, to_simplified
from report_header import (
    CODE_VALUE_PATTERN,
    HEADER_TOP_MIN_Y,
    extract_labeled_values,
    extract_report_issue_date,
    extract_same_line_labeled_values,
    find_cma_combination,
    guess_institution,
    guess_institution_display,
    normalize_institution_display_value,
    has_mark_text,
    has_nonnegative_mark_text,
    select_header_blocks,
)
from runtime_env import format_ocr_environment_error, inspect_ocr_environment, smoke_test_ocr


ROOT = Path(os.environ.get("QC_WORK_DIR", Path(__file__).resolve().parent)).resolve()
manifest = json.loads((ROOT / "download_manifest.json").read_text(encoding="utf-8"))
previous = {}
pdf_text_path = ROOT / "pdf_text.json"
OCR_HELPER = Path(__file__).resolve().parent / "macos_vision_ocr.swift"
OCR_ENABLED = os.environ.get("QC_ENABLE_OCR", "1") != "0"
HEADER_OCR_ENABLED = os.environ.get("QC_ENABLE_HEADER_OCR", "1") != "0"
HEADER_PARSER_VERSION = "report-header-v6"
EXTRACT_WORKERS = max(1, int(os.environ.get("QC_EXTRACT_WORKERS", "1")))
OCR_RUNTIME = inspect_ocr_environment(OCR_HELPER)
if OCR_ENABLED and not OCR_RUNTIME["ready"]:
    raise SystemExit(
        "OCR环境预检失败："
        + format_ocr_environment_error(OCR_RUNTIME)
        + "。请运行 scripts/setup_environment.sh 完成配置后重试。"
    )
if OCR_ENABLED:
    OCR_SMOKE = smoke_test_ocr(OCR_RUNTIME)
    if not OCR_SMOKE["ready"]:
        raise SystemExit(
            "OCR运行自检失败："
            + OCR_SMOKE["error"]
            + "。批处理已在读取PDF前停止，未生成降级结果。"
        )
if pdf_text_path.exists():
    for row in json.loads(pdf_text_path.read_text(encoding="utf-8")):
        if row.get("url") and row.get("pages"):
            previous[row["url"]] = row
for cache_path in filter(None, os.environ.get("QC_TEXT_CACHE_PATHS", "").split(os.pathsep)):
    path = Path(cache_path)
    if not path.exists() or path.resolve() == pdf_text_path.resolve():
        continue
    for row in json.loads(path.read_text(encoding="utf-8")):
        if row.get("url") and row.get("pages"):
            previous.setdefault(row["url"], row)


def compact(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def guess_report_no(text: str, fallback: str) -> str:
    patterns = [
        r"报告[编編]号\s*[：:]?\s*([A-Z0-9-]{6,})",
        r"(TTTS[-–—]?WT\d{8,})",
        r"\b(W0\d{10,})\b",
        r"\b(W\d{8,})\b",
        r"\b(TNL\d{8,}[A-Z]?)\b",
        r"\b((?=[A-Z0-9_-]*\d)[A-Z]{2,8}[-_]?[A-Z0-9-]{6,})\b",
    ]
    for candidate_text in (text or "", compact(text or "")):
        for pattern in patterns:
            match = re.search(pattern, candidate_text, re.I)
            if match:
                value = compact(match.group(1)).replace("—", "-").replace("–", "-")
                if value.upper() not in {"INSPECTION", "REPORT", "CATEGORY", "BALABALA", "SAMPLENAME"}:
                    return value
    fallback = compact(fallback)
    return fallback if re.search(r"\d", fallback) and not re.fullmatch(r"[0-9a-f]{32,64}", fallback, re.I) and fallback.upper() not in {"BALABALA"} else ""


def top_region_text(document):
    parts = []
    if document.get("header_native_text"):
        parts.append("首页顶部原生文本：" + str(document["header_native_text"]))
    if document.get("header_ocr_text"):
        parts.append("首页顶部报告头：" + str(document["header_ocr_text"]))
    for page in document.get("pages") or []:
        page_no = int(page.get("page") or len(parts) + 1)
        blocks = page.get("ocr_blocks") or []
        if not blocks:
            continue
        top_blocks = []
        for block in blocks:
            try:
                mid_y = float(block.get("mid_y") or 0)
            except (TypeError, ValueError):
                mid_y = 1
            if mid_y >= HEADER_TOP_MIN_Y:
                text = str(block.get("text") or "").strip()
                if text:
                    top_blocks.append((float(block.get("x") or 0), text))
        if top_blocks:
            parts.append(f"第{page_no}页顶部：" + " ".join(text for _, text in sorted(top_blocks)))
    return "\n".join(parts)


def mark_status(document, mark: str):
    full_text = "\n".join(str(page.get("text") or "") for page in document.get("pages") or [])
    top_text = top_region_text(document)
    text_source = document.get("text_source", "native")
    negative = re.compile(rf"(?:未|不|暂未).{{0,12}}{mark}|{mark}.{{0,12}}(?:未授权|不适用|未获得)", re.I)
    cma_number, cma_ocr_text = find_cma_combination(document.get("header_ocr_blocks") or []) if mark == "CMA" else ("", "")
    top_positive = bool(cma_number) or has_nonnegative_mark_text(top_text, mark)
    full_positive = has_nonnegative_mark_text(full_text, mark)
    if top_positive or full_positive:
        scope = "首页顶部报告头OCR" if document.get("header_ocr_text") and has_mark_text(document.get("header_ocr_text", ""), mark) else ("顶部区域OCR" if top_positive else ("OCR全文" if text_source == "vision_ocr" else "PDF原生文本"))
        if cma_number:
            return "有", f"有｜首页顶部报告头OCR识别到CMA组合证据：图形字母识别为{cma_ocr_text}，其下方同水平位置识别到资质编号{cma_number}"
        evidence = top_text if top_positive else re.sub(r"\s+", " ", (full_text or "")[:500])
        return "有", f"有｜{scope}识别到{mark}标识证据：{evidence[:180]}"
    if document.get("header_parse_attempted") and document.get("header_parse_error"):
        return "待复核", f"待复核｜{mark}未在文本中识别，首页顶部报告头OCR失败：{document.get('header_parse_error')}"
    if document.get("header_parse_status") == "OCR未启用":
        return "待复核", f"待复核｜{mark}未在原生文本中识别，且首页顶部报告头OCR未启用"
    if negative.search(full_text or "") and document.get("header_parse_status") == "已完成":
        return "未发现", f"未发现｜仅识别到{mark}未授权/未获得/不适用说明，首页报告头也无强标识证据"
    if document.get("header_parse_status") == "已完成":
        return "待复核", f"待复核｜首页顶部报告头OCR已完成，但未获得{mark}强文字或图章组合证据，不能据此确认未出现"
    if any(page.get("ocr_blocks") for page in document.get("pages") or []):
        return "未发现", f"未发现｜PDF原生文本、OCR全文及每页顶部区域均未识别到{mark}标识"
    if document.get("ocr_attempted") and document.get("ocr_error"):
        return "待复核", f"待复核｜{mark}未在文本中识别，且OCR失败：{document.get('ocr_error')}"
    return "待复核", f"待复核｜PDF原生文本/缓存文本未识别到{mark}标识，且未完成有效顶部图像检查"


def extract_report_metadata(document):
    pages = document.get("pages") or []
    full_text = "\n".join(str(page.get("text") or "") for page in pages)
    first_page_text = str(pages[0].get("text") or "") if pages else ""
    labeled_text = full_text
    for pattern, replacement in (
        (r"货\s*号", "货号"), (r"款\s*号", "款号"), (r"版\s*单\s*号", "版单号"),
        (r"面\s*料\s*编\s*号", "面料编号"), (r"物\s*料\s*编\s*号", "物料编号"), (r"料\s*号", "料号"),
    ):
        labeled_text = re.sub(pattern, replacement, labeled_text)
    issue_date, date_status, date_label, date_original, date_reason, date_candidates = extract_report_issue_date(document)
    header_text = "\n__QC_SOURCE_BREAK__\n".join(filter(None, [str(document.get("header_native_text") or ""), str(document.get("header_ocr_text") or "")]))
    institution_text = "\n__QC_SOURCE_BREAK__\n".join(filter(None, [header_text, full_text]))
    strict_institution = guess_institution(institution_text)
    institution = guess_institution_display(institution_text)
    header_institution = guess_institution(header_text) if header_text else "机构名称待确认"
    institution_source = "首页顶部报告头" if header_institution == strict_institution else "报告全文或重复页脚文本"
    cma_mark, cma_evidence = mark_status(document, "CMA")
    cnas_mark, cnas_evidence = mark_status(document, "CNAS")
    report_product_codes = extract_labeled_values(
        labeled_text,
        [r"(?:报告内)?货号", r"款号"],
        CODE_VALUE_PATTERN,
    )
    plate_numbers = extract_same_line_labeled_values(
        labeled_text,
        [r"版单号"],
        stop_labels=(
            r"面料编号", r"物料编号", r"料号", r"年/季度", r"年季度", r"码段/号型",
            r"使用用途", r"批次号", r"颜色", r"样品描述", r"判定标准",
        ),
    )
    material_numbers = extract_labeled_values(labeled_text, [r"面料编号", r"物料编号", r"料号"], CODE_VALUE_PATTERN)
    return {
        "report_no": guess_report_no(full_text, document.get("report_no", "")),
        "report_issue_date": issue_date,
        "report_issue_date_status": date_status,
        "report_issue_date_label": date_label,
        "report_issue_date_original": date_original,
        "report_issue_date_reason": date_reason,
        "report_issue_date_candidates": date_candidates,
        "institution": normalize_institution_display_value(institution),
        "institution_evidence": (
            f"{institution_source}识别：{institution}"
            if strict_institution != "机构名称待确认" else
            (
                f"宽松展示｜报告可识别简称、品牌或稳定版式标识：{institution}；按业务规则直接展示，不要求人工确认"
                if institution != "未识别" else
                "未识别｜首页顶部报告头及全文没有可用机构文字；按业务规则不再转人工确认"
            )
        ),
        "cma_mark": cma_mark,
        "cnas_mark": cnas_mark,
        "cma_evidence": cma_evidence,
        "cnas_evidence": cnas_evidence,
        "report_product_codes": report_product_codes,
        "plate_numbers": plate_numbers,
        "material_numbers": material_numbers,
    }


def ocr_blocks_to_text(blocks, original=False):
    """按Vision坐标将识别块重组成视觉行，同时保留原始块供后续列定位。"""
    lines = []
    current = []
    current_y = None
    for block in blocks:
        source_text = block.get("original_text") if original else block.get("text")
        if source_text is None:
            source_text = block.get("text") or ""
        text = str(source_text if original else to_simplified(source_text)).strip()
        if not text:
            continue
        mid_y = float(block.get("mid_y") or 0)
        tolerance = max(0.008, float(block.get("height") or 0) * 0.45)
        if current and current_y is not None and abs(mid_y - current_y) > tolerance:
            lines.append("\t".join(part[1] for part in sorted(current, key=lambda part: part[0])))
            current = []
        current.append((float(block.get("x") or 0), text))
        current_y = mid_y if current_y is None else (current_y * (len(current) - 1) + mid_y) / len(current)
    if current:
        lines.append("\t".join(part[1] for part in sorted(current, key=lambda part: part[0])))
    return "\n".join(lines)


def run_vision_ocr(pdf_path: str, first_page_only: bool = False, header_only: bool = False):
    if not OCR_ENABLED:
        return [], "OCR disabled"
    runtime = OCR_RUNTIME
    pdftoppm = runtime["pdftoppm"]
    swift = runtime["swift"]
    sdk = runtime["sdk"]
    ocr_root = ROOT / "ocr_runtime"
    ocr_root.mkdir(parents=True, exist_ok=True)
    module_cache = ocr_root / "swift_module_cache"
    module_cache.mkdir(parents=True, exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix="qc_ocr_", dir=ocr_root) as temp_dir:
            prefix = str(Path(temp_dir) / "page")
            render_command = [pdftoppm]
            if first_page_only:
                render_command.extend(["-f", "1", "-l", "1"])
            render_command.extend(["-r", "150", "-png", pdf_path, prefix])
            render_timeout = 90 if first_page_only else 180
            render = subprocess.run(
                render_command,
                capture_output=True, text=True, timeout=render_timeout,
            )
            if render.returncode != 0:
                return [], f"PDF render failed: {render.stderr.strip()[:300]}"
            images = sorted(Path(temp_dir).glob("page-*.png"))
            if not images:
                return [], "PDF render produced no pages"
            if header_only:
                cropped_images = []
                for image_path in images:
                    with Image.open(image_path) as source_image:
                        crop_height = max(1, int(source_image.height * (1 - HEADER_TOP_MIN_Y)))
                        cropped_path = image_path.with_name(image_path.stem + "-header.png")
                        source_image.crop((0, 0, source_image.width, crop_height)).save(cropped_path)
                        cropped_images.append(cropped_path)
                images = cropped_images
            command = [
                swift, "-sdk", sdk, "-module-cache-path", str(module_cache),
                str(OCR_HELPER), *[str(path) for path in images],
            ]
            ocr_timeout = 120 if header_only else 600
            result = subprocess.run(command, capture_output=True, text=True, timeout=ocr_timeout)
            if result.returncode != 0:
                return [], f"Vision OCR failed: {result.stderr.strip()[:500]}"
            pages = []
            for line in result.stdout.splitlines():
                if not line.strip():
                    continue
                payload = json.loads(line)
                if payload.get("status") != "ok":
                    return [], f"Vision OCR page failed: {payload.get('error') or payload.get('status')}"
                blocks = simplify_ocr_blocks(payload.get("blocks") or [])
                page_text = ocr_blocks_to_text(blocks)
                original_page_text = ocr_blocks_to_text(blocks, original=True)
                pages.append({
                    "page": int(payload.get("page_index") or len(pages) + 1),
                    "text": page_text,
                    "original_text": original_page_text,
                    "ocr_blocks": blocks,
                })
            if not pages:
                return [], "Vision OCR produced no page output"
            return pages, ""
    except Exception as exc:
        return [], f"{type(exc).__name__}: {exc}"


def ensure_report_header(document):
    """独立解析首页顶部整宽报告头；字段定位不依赖左右位置。"""
    if document.get("header_parse_attempted") and document.get("header_parser_version") == HEADER_PARSER_VERSION and document.get("header_parse_status") not in {"失败", "OCR未启用"}:
        return
    document["header_parse_attempted"] = True
    document["header_parser_version"] = HEADER_PARSER_VERSION
    pages = document.get("pages") or []
    native_header = str(document.get("header_native_text") or "")
    if not native_header and pages:
        native_header = "\n".join(str(pages[0].get("text") or "").splitlines()[:35])
    document["header_native_text"] = native_header

    native_all = "\n".join(str(page.get("text") or "") for page in pages)
    unresolved = (
        guess_institution(native_header) == "机构名称待确认"
        or not has_mark_text(native_all, "CMA")
        or not has_mark_text(native_all, "CNAS")
    )
    if not unresolved:
        document["header_parse_status"] = "原生文本已解决"
        document["header_parse_error"] = ""
        return
    if document.get("header_ocr_blocks"):
        top_blocks = document.get("header_ocr_blocks") or []
        document["header_ocr_text"] = ocr_blocks_to_text(top_blocks)
        document["header_parse_status"] = "已完成"
        document["header_parse_error"] = ""
        return
    if not HEADER_OCR_ENABLED or not OCR_ENABLED:
        document["header_parse_status"] = "OCR未启用"
        document["header_parse_error"] = ""
        return
    if pages and pages[0].get("ocr_blocks"):
        top_blocks = select_header_blocks(pages[0].get("ocr_blocks") or [])
        document["header_ocr_blocks"] = top_blocks
        document["header_ocr_text"] = ocr_blocks_to_text(top_blocks)
        document["header_parse_status"] = "已完成"
        document["header_parse_error"] = ""
        return
    if not document.get("path"):
        document["header_parse_status"] = "失败"
        document["header_parse_error"] = "缺少本地PDF路径"
        return

    ocr_pages, error = run_vision_ocr(str(document["path"]), first_page_only=True, header_only=True)
    if error or not ocr_pages:
        document["header_parse_status"] = "失败"
        document["header_parse_error"] = error or "首页OCR无输出"
        return
    top_blocks = ocr_pages[0].get("ocr_blocks") or []
    document["header_ocr_blocks"] = top_blocks
    document["header_ocr_text"] = ocr_blocks_to_text(top_blocks)
    document["header_parse_status"] = "已完成"
    document["header_parse_error"] = ""


documents = []


def save_documents():
    tmp_path = pdf_text_path.with_suffix(".json.tmp")
    ordered = sorted(documents, key=lambda item: int(item.get("pdf_id") or 0))
    tmp_path.write_text(json.dumps(ordered, ensure_ascii=False), encoding="utf-8")
    tmp_path.replace(pdf_text_path)


def process_manifest_row(row):
    cached_previous = previous.get(row["url"])
    needs_ocr_refresh = bool(
        cached_previous and cached_previous.get("text_chars", 0) < 300 and OCR_ENABLED
        and (
            not cached_previous.get("ocr_attempted")
            or bool(cached_previous.get("ocr_error"))
        )
    )
    if cached_previous and row["status"] != "failed" and not needs_ocr_refresh:
        cached = {**previous[row["url"]], **row, "extract_error": previous[row["url"]].get("extract_error", "")}
        simplify_document_text(cached)
        ensure_report_header(cached)
        cached.update(extract_report_metadata(cached))
        return cached, (
            f"#{row['pdf_id']:04d} cached "
            f"header={cached.get('header_parse_status', '') or '-'} "
            f"chars={cached.get('text_chars', 0)} {cached.get('report_no', '')}"
        )

    document = {
        **row, "pages": [], "text_chars": 0, "native_text_chars": 0,
        "report_no": "", "institution": "", "extract_error": "",
        "text_source": "native", "ocr_attempted": False, "ocr_text_chars": 0, "ocr_error": "",
    }
    if row["status"] == "failed":
        return document, f"#{row['pdf_id']:04d} download_failed {document.get('error', '')}"
    try:
        with pdfplumber.open(row["path"]) as pdf:
            pages = []
            for page_number, page in enumerate(pdf.pages, start=1):
                text = page.extract_text(x_tolerance=2, y_tolerance=3) or ""
                pages.append({"page": page_number, "text": text})
                if page_number == 1:
                    header_crop = page.crop((0, 0, page.width, page.height * (1 - HEADER_TOP_MIN_Y)))
                    document["header_native_text"] = header_crop.extract_text(x_tolerance=2, y_tolerance=3) or ""
            full_text = "\n".join(page["text"] for page in pages)
            document["native_text_chars"] = len(full_text)
            if len(full_text) < 300 and OCR_ENABLED:
                document["ocr_attempted"] = True
                ocr_pages, ocr_error = run_vision_ocr(row["path"])
                document["ocr_error"] = ocr_error
                if ocr_pages:
                    pages = ocr_pages
                    full_text = "\n".join(page["text"] for page in pages)
                    document["ocr_text_chars"] = len(full_text)
                    document["text_source"] = "vision_ocr"
            document["pages"] = pages
            simplify_document_text(document)
            full_text = "\n".join(page["text"] for page in document["pages"])
            document["text_chars"] = len(full_text)
            document["report_no"] = guess_report_no(full_text, Path(row["path"]).stem)
            ensure_report_header(document)
            document.update(extract_report_metadata(document))
    except Exception as exc:
        document["extract_error"] = f"{type(exc).__name__}: {exc}"
    return document, f"#{row['pdf_id']:04d} chars={document['text_chars']} {document['report_no']}"


if EXTRACT_WORKERS == 1:
    for row in manifest:
        document, message = process_manifest_row(row)
        documents.append(document)
        print(f"[{len(documents)}/{len(manifest)}] {message}", flush=True)
        if len(documents) % 25 == 0 or len(documents) == len(manifest):
            save_documents()
else:
    with ThreadPoolExecutor(max_workers=EXTRACT_WORKERS) as executor:
        futures = {executor.submit(process_manifest_row, row): row for row in manifest}
        for completed, future in enumerate(as_completed(futures), start=1):
            document, message = future.result()
            documents.append(document)
            print(f"[{completed}/{len(manifest)}] {message}", flush=True)
            if completed % 25 == 0 or completed == len(manifest):
                save_documents()

if documents:
    save_documents()

print(json.dumps({
    "documents": len(documents),
    "download_failed": sum(row["status"] == "failed" for row in documents),
    "extract_failed": sum(bool(row["extract_error"]) for row in documents),
    "low_text": sum(0 < row["text_chars"] < 300 for row in documents),
    "no_text": sum(row["status"] != "failed" and row["text_chars"] == 0 for row in documents),
    "cache_reused": sum(row["url"] in previous and row["status"] != "failed" for row in manifest),
}, ensure_ascii=False, indent=2))
