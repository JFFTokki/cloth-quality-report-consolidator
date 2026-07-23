import os
import platform
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


def _find_macos_sdk():
    configured = os.environ.get("SDKROOT", "").strip()
    if configured and Path(configured).is_dir():
        return configured, ""
    xcrun = shutil.which("xcrun")
    if not xcrun:
        return "", "缺少 xcrun，无法定位 macOS SDK"
    result = subprocess.run(
        [xcrun, "--sdk", "macosx", "--show-sdk-path"],
        capture_output=True,
        text=True,
        timeout=15,
    )
    sdk = result.stdout.strip()
    if result.returncode != 0 or not sdk or not Path(sdk).is_dir():
        detail = result.stderr.strip() or sdk or "未返回 SDK 路径"
        return "", f"macOS SDK 定位失败：{detail}"
    return sdk, ""


def inspect_ocr_environment(ocr_helper):
    helper = Path(ocr_helper).resolve()
    tools = {
        "pdftoppm": shutil.which("pdftoppm") or "",
        "pdfinfo": shutil.which("pdfinfo") or "",
        "swift": shutil.which("swift") or "",
        "xcrun": shutil.which("xcrun") or "",
    }
    sdk, sdk_error = _find_macos_sdk()
    errors = []
    if platform.system() != "Darwin":
        errors.append(f"macOS Vision OCR 仅支持 macOS，当前系统为 {platform.system()}")
    for name in ("pdftoppm", "pdfinfo", "swift", "xcrun"):
        if not tools[name]:
            errors.append(f"缺少系统命令 {name}")
    if sdk_error:
        errors.append(sdk_error)
    if not helper.is_file():
        errors.append(f"缺少 Vision OCR 辅助脚本：{helper}")
    return {
        "ready": not errors,
        "errors": errors,
        "pdftoppm": tools["pdftoppm"],
        "pdfinfo": tools["pdfinfo"],
        "swift": tools["swift"],
        "xcrun": tools["xcrun"],
        "sdk": sdk,
        "helper": str(helper),
    }


def format_ocr_environment_error(status):
    errors = status.get("errors") or ["未知 OCR 环境错误"]
    return "；".join(errors)


def smoke_test_ocr(status):
    if not status.get("ready"):
        return {"ready": False, "error": format_ocr_environment_error(status)}
    from PIL import Image, ImageDraw

    with tempfile.TemporaryDirectory(prefix="qc_env_smoke_") as temp_dir:
        temp = Path(temp_dir)
        image_path = temp / "ocr-smoke.png"
        pdf_path = temp / "render-smoke.pdf"
        render_prefix = temp / "rendered"
        image = Image.new("RGB", (800, 240), "white")
        ImageDraw.Draw(image).text((40, 80), "CMA CNAS OCR SMOKE", fill="black")
        image.save(image_path)
        image.save(pdf_path, "PDF", resolution=150)
        render = subprocess.run(
            [status["pdftoppm"], "-f", "1", "-l", "1", "-png", str(pdf_path), str(render_prefix)],
            capture_output=True,
            text=True,
            timeout=45,
        )
        rendered_images = sorted(temp.glob("rendered-*.png"))
        if render.returncode != 0 or not rendered_images:
            detail = render.stderr.strip() or "没有生成 PNG"
            return {"ready": False, "error": f"pdftoppm 渲染自检失败：{detail}"}
        module_cache = temp / "swift_module_cache"
        module_cache.mkdir()
        vision = subprocess.run(
            [
                status["swift"], "-sdk", status["sdk"],
                "-module-cache-path", str(module_cache), status["helper"], str(rendered_images[0]),
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )
        if vision.returncode != 0:
            return {"ready": False, "error": f"macOS Vision OCR 自检失败：{vision.stderr.strip()}"}
        outputs = [line for line in vision.stdout.splitlines() if line.strip()]
        if not outputs:
            return {"ready": False, "error": "macOS Vision OCR 自检没有输出"}
        payload = None
        for line in outputs:
            try:
                candidate = json.loads(line)
            except json.JSONDecodeError:
                continue
            if isinstance(candidate, dict) and candidate.get("status"):
                payload = candidate
                break
        if payload is None:
            preview = " | ".join(outputs)[:300]
            return {"ready": False, "error": f"macOS Vision OCR 自检输出不是有效JSON：{preview}"}
        if payload.get("status") != "ok":
            detail = payload.get("error") or str(payload)
            if "Foundation._GenericObjCError" in detail or "CVPixelBuffer" in detail:
                detail = "当前进程受限沙箱阻止 macOS Vision 创建识别任务；请在普通终端或获准的非沙箱环境运行"
            return {"ready": False, "error": f"macOS Vision OCR 自检失败：{detail}"}
        return {"ready": True, "recognized_blocks": len(payload.get("blocks") or [])}
