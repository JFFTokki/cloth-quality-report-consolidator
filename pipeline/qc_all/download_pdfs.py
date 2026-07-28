import hashlib
import json
import os
import shutil
import ssl
import tempfile
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote, urlsplit, urlunsplit

import pdfplumber


DEFAULT_TIMEOUT_SECONDS = 90
DEFAULT_MAX_BYTES = 100 * 1024 * 1024
DOWNLOAD_VERSION = "qc-download-v2"
SUCCESS_STATUSES = {"cached", "downloaded", "ok"}


def safe_url(url):
    parts = urlsplit(url)
    if parts.scheme.lower() not in {"http", "https"} or not parts.netloc:
        raise ValueError("only absolute http/https URLs are allowed")
    path = quote(parts.path, safe="/%")
    query = quote(parts.query, safe="=&%")
    return urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))


def ssl_context(ca_file=""):
    return ssl.create_default_context(cafile=ca_file or None)


def output_path(pdf_dir, url):
    return pdf_dir / f"{hashlib.sha1(url.encode('utf-8')).hexdigest()}.pdf"


def file_sha256(path):
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def validate_pdf(path, max_bytes=DEFAULT_MAX_BYTES):
    size = path.stat().st_size
    if size <= 1024:
        raise ValueError(f"PDF too small: {size} bytes")
    if size > max_bytes:
        raise ValueError(f"PDF exceeds size limit: {size} > {max_bytes} bytes")
    with path.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise ValueError("response does not start with %PDF-")
    with pdfplumber.open(path) as pdf:
        if len(pdf.pages) < 1:
            raise ValueError("PDF has no pages")
    return size, file_sha256(path)


def copy_cached_pdf(source_path, target_path, max_bytes):
    target_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=target_path.parent, suffix=".part", delete=False) as temp:
        temp_path = Path(temp.name)
    try:
        shutil.copy2(source_path, temp_path)
        size, sha256 = validate_pdf(temp_path, max_bytes=max_bytes)
        os.replace(temp_path, target_path)
        return size, sha256
    finally:
        temp_path.unlink(missing_ok=True)


def download_to_temp(url, target_path, context, timeout_seconds, max_bytes, opener=urllib.request.urlopen):
    request = urllib.request.Request(safe_url(url), headers={"User-Agent": "Mozilla/5.0"})
    target_path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temp_name = tempfile.mkstemp(dir=target_path.parent, suffix=".part")
    temp_path = Path(temp_name)
    try:
        with os.fdopen(descriptor, "wb") as temp:
            with opener(request, timeout=timeout_seconds, context=context) as response:
                status = response.getcode()
                if status is not None and not 200 <= status < 300:
                    raise ValueError(f"HTTP status {status}")
                final_url = response.geturl()
                safe_url(final_url)
                content_length = response.headers.get("Content-Length")
                if content_length and int(content_length) > max_bytes:
                    raise ValueError(f"PDF exceeds size limit: {content_length} > {max_bytes} bytes")
                total = 0
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    total += len(chunk)
                    if total > max_bytes:
                        raise ValueError(f"PDF exceeds size limit: {total} > {max_bytes} bytes")
                    temp.write(chunk)
                response_headers = dict(response.headers)
            temp.flush()
            os.fsync(temp.fileno())
        size, sha256 = validate_pdf(temp_path, max_bytes=max_bytes)
        os.replace(temp_path, target_path)
        return {
            "final_url": final_url,
            "bytes": size,
            "sha256": sha256,
            "etag": response_headers.get("ETag", ""),
            "last_modified": response_headers.get("Last-Modified", ""),
        }
    finally:
        temp_path.unlink(missing_ok=True)


def valid_cached_file(path, manifest_row, max_bytes):
    try:
        size, sha256 = validate_pdf(path, max_bytes=max_bytes)
    except Exception:
        return None
    expected_hash = str((manifest_row or {}).get("sha256") or "")
    if expected_hash and sha256 != expected_hash:
        return None
    return size, sha256


def download_one(index, url, pdf_dir, old_manifest, context, timeout_seconds, max_bytes, opener=urllib.request.urlopen):
    output = output_path(pdf_dir, url)
    result = {
        "pdf_id": index,
        "url": url,
        "final_url": "",
        "path": str(output),
        "status": "failed",
        "bytes": 0,
        "sha256": "",
        "etag": "",
        "last_modified": "",
        "download_version": DOWNLOAD_VERSION,
        "error": "",
    }
    try:
        safe_url(url)
        cached = valid_cached_file(output, old_manifest.get(url), max_bytes) if output.exists() else None
        if cached:
            result.update(status="cached", bytes=cached[0], sha256=cached[1], final_url=url)
            return result

        old_row = old_manifest.get(url) or {}
        old_path_text = str(old_row.get("path") or "")
        old_path = Path(old_path_text) if old_path_text else None
        if old_path and old_path.is_file():
            try:
                size, sha256 = copy_cached_pdf(old_path, output, max_bytes)
                expected_hash = str(old_row.get("sha256") or "")
                if expected_hash and sha256 != expected_hash:
                    raise ValueError("cached PDF SHA-256 does not match manifest")
                result.update(status="cached", bytes=size, sha256=sha256, final_url=old_row.get("final_url") or url)
                return result
            except Exception:
                output.unlink(missing_ok=True)

        last_error = ""
        for _ in range(2):
            try:
                metadata = download_to_temp(
                    url, output, context, timeout_seconds, max_bytes, opener=opener,
                )
                result.update(status="downloaded", error="", **metadata)
                return result
            except Exception as exc:
                last_error = f"{type(exc).__name__}: {exc}"
        result["error"] = last_error
    except Exception as exc:
        result["error"] = f"{type(exc).__name__}: {exc}"
    if output.exists() and not valid_cached_file(output, {}, max_bytes):
        output.unlink(missing_ok=True)
    return result


def load_success_manifest(path):
    rows = {}
    if not path.exists():
        return rows
    for row in json.loads(path.read_text(encoding="utf-8")):
        if row.get("status") in SUCCESS_STATUSES and row.get("url"):
            rows[row["url"]] = row
    return rows


def main():
    root = Path(os.environ.get("QC_WORK_DIR", Path(__file__).resolve().parent)).resolve()
    source = json.loads((root / "source_records.json").read_text(encoding="utf-8"))
    pdf_dir = root / "pdfs_by_url"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = root / "download_manifest.json"
    cache_manifest_path = Path(os.environ.get("QC_CACHE_MANIFEST", manifest_path))
    old_manifest = load_success_manifest(cache_manifest_path)
    timeout_seconds = int(os.environ.get("QC_DOWNLOAD_TIMEOUT_SECONDS", DEFAULT_TIMEOUT_SECONDS))
    max_bytes = int(os.environ.get("QC_MAX_PDF_BYTES", DEFAULT_MAX_BYTES))
    workers = int(os.environ.get("QC_DOWNLOAD_WORKERS", "12"))
    context = ssl_context(os.environ.get("QC_CA_FILE", ""))

    manifest = []
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {
            executor.submit(
                download_one, index, url, pdf_dir, old_manifest, context, timeout_seconds, max_bytes,
            ): (index, url)
            for index, url in enumerate(source["selected_urls"], start=1)
        }
        for done, future in enumerate(as_completed(futures), start=1):
            row = future.result()
            manifest.append(row)
            print(f"[{done}/{len(source['selected_urls'])}] #{row['pdf_id']:04d} {row['status']}: {row['url'][-70:]}", flush=True)

    manifest.sort(key=lambda row: row["pdf_id"])
    temp_manifest = manifest_path.with_suffix(".json.tmp")
    temp_manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    temp_manifest.replace(manifest_path)
    print(json.dumps({
        "total": len(manifest),
        "success": sum(row["status"] != "failed" for row in manifest),
        "failed": sum(row["status"] == "failed" for row in manifest),
        "bytes": sum(row["bytes"] for row in manifest),
    }, ensure_ascii=False))


if __name__ == "__main__":
    main()
