import hashlib
import json
import os
import shutil
import ssl
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote, urlsplit, urlunsplit


ROOT = Path(os.environ.get("QC_WORK_DIR", Path(__file__).resolve().parent)).resolve()
source = json.loads((ROOT / "source_records.json").read_text(encoding="utf-8"))
pdf_dir = ROOT / "pdfs_by_url"
pdf_dir.mkdir(parents=True, exist_ok=True)
old_manifest_path = ROOT / "download_manifest.json"
old_manifest = {}
cache_manifest_path = Path(os.environ.get("QC_CACHE_MANIFEST", old_manifest_path))
if cache_manifest_path.exists():
    for row in json.loads(cache_manifest_path.read_text(encoding="utf-8")):
        if row.get("status") != "failed" and Path(row.get("path", "")).exists():
            old_manifest[row["url"]] = row

context = ssl._create_unverified_context()


def safe_url(url):
    parts = urlsplit(url)
    path = quote(parts.path, safe="/%")
    query = quote(parts.query, safe="=&%")
    return urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))


def output_path(url):
    return pdf_dir / f"{hashlib.sha1(url.encode('utf-8')).hexdigest()}.pdf"


def download_one(index, url):
    output = output_path(url)
    if output.exists() and output.stat().st_size > 1024:
        status = "cached"
        error = ""
    elif url in old_manifest:
        old_path = Path(old_manifest[url]["path"])
        shutil.copy2(old_path, output)
        status = "cached"
        error = ""
    else:
        status = "failed"
        error = ""
        for _ in range(2):
            try:
                request = urllib.request.Request(safe_url(url), headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(request, timeout=90, context=context) as response:
                    data = response.read()
                if data.startswith(b"%PDF") and len(data) > 1024:
                    output.write_bytes(data)
                    status = "downloaded"
                    error = ""
                    break
                error = f"not a pdf or too small: {len(data)} bytes"
            except Exception as exc:
                error = f"{type(exc).__name__}: {exc}"
        if status == "failed" and output.exists():
            output.unlink()
    return {
        "pdf_id": index,
        "url": url,
        "path": str(output),
        "status": status,
        "bytes": output.stat().st_size if output.exists() else 0,
        "error": error,
    }


manifest = []
with ThreadPoolExecutor(max_workers=12) as executor:
    futures = {
        executor.submit(download_one, index, url): (index, url)
        for index, url in enumerate(source["selected_urls"], start=1)
    }
    for done, future in enumerate(as_completed(futures), start=1):
        row = future.result()
        manifest.append(row)
        print(f"[{done}/{len(source['selected_urls'])}] #{row['pdf_id']:04d} {row['status']}: {row['url'][-70:]}", flush=True)

manifest.sort(key=lambda row: row["pdf_id"])
(ROOT / "download_manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
print(json.dumps({
    "total": len(manifest),
    "success": sum(row["status"] != "failed" for row in manifest),
    "failed": sum(row["status"] == "failed" for row in manifest),
    "bytes": sum(row["bytes"] for row in manifest),
}, ensure_ascii=False))
