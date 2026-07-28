import io
import ssl
import tempfile
import urllib.error
from pathlib import Path

from download_pdfs import (
    DEFAULT_MAX_BYTES,
    download_one,
    download_to_temp,
    file_sha256,
    output_path,
    safe_url,
    ssl_context,
)


def minimal_pdf():
    objects = [
        b"1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
        b"2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
        b"3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 72 72] >>\nendobj\n",
    ]
    data = bytearray(b"%PDF-1.4\n")
    offsets = [0]
    for obj in objects:
        offsets.append(len(data))
        data.extend(obj)
    xref = len(data)
    data.extend(f"xref\n0 {len(objects) + 1}\n".encode())
    data.extend(b"0000000000 65535 f \n")
    for offset in offsets[1:]:
        data.extend(f"{offset:010d} 00000 n \n".encode())
    data.extend(f"trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode())
    data.extend(b"%" + b"x" * 1024 + b"\n")
    return bytes(data)


class FakeResponse:
    def __init__(self, payload, *, headers=None, fail_after_first=False):
        self.stream = io.BytesIO(payload)
        self.headers = headers or {}
        self.fail_after_first = fail_after_first
        self.reads = 0

    def __enter__(self):
        return self

    def __exit__(self, *_):
        return False

    def getcode(self):
        return 200

    def geturl(self):
        return "https://example.test/final.pdf"

    def read(self, size):
        self.reads += 1
        if self.fail_after_first and self.reads > 1:
            raise ConnectionError("download interrupted")
        return self.stream.read(min(size, 64))


def opener_for(response):
    return lambda *args, **kwargs: response


def main():
    payload = minimal_pdf()
    assert safe_url("https://example.test/a file.pdf").startswith("https://")
    try:
        safe_url("file:///tmp/a.pdf")
    except ValueError:
        pass
    else:
        raise AssertionError("file URL must be rejected")
    assert ssl_context().verify_mode == ssl.CERT_REQUIRED
    assert ssl_context().check_hostname

    with tempfile.TemporaryDirectory() as temp_dir:
        root = Path(temp_dir)
        output = root / "ok.pdf"
        result = download_to_temp(
            "https://example.test/a.pdf",
            output,
            ssl_context(),
            timeout_seconds=1,
            max_bytes=DEFAULT_MAX_BYTES,
            opener=opener_for(FakeResponse(payload)),
        )
        assert output.exists()
        assert result["sha256"] == file_sha256(output)
        assert result["bytes"] == len(payload)

        for name, response, max_bytes in (
            ("not-pdf", FakeResponse(b"not a pdf" * 300), DEFAULT_MAX_BYTES),
            ("oversize", FakeResponse(payload, headers={"Content-Length": str(len(payload) + 1)}), len(payload)),
            ("interrupted", FakeResponse(payload, fail_after_first=True), DEFAULT_MAX_BYTES),
        ):
            try:
                download_to_temp(
                    "https://example.test/a.pdf",
                    root / f"{name}.pdf",
                    ssl_context(),
                    timeout_seconds=1,
                    max_bytes=max_bytes,
                    opener=opener_for(response),
                )
            except Exception:
                pass
            else:
                raise AssertionError(f"{name} must fail")

        def tls_failure(*args, **kwargs):
            raise urllib.error.URLError(ssl.SSLCertVerificationError(1, "certificate verify failed"))

        failed = download_one(
            1,
            "https://example.test/tls.pdf",
            root,
            {},
            ssl_context(),
            1,
            DEFAULT_MAX_BYTES,
            opener=tls_failure,
        )
        assert failed["status"] == "failed"
        assert "certificate verify failed" in failed["error"]

        cached_root = root / "cached"
        cached_root.mkdir()
        cached_url = "https://example.test/hash-change.pdf"
        cached_target = output_path(cached_root, cached_url)
        cached_target.write_bytes(payload)
        redownloaded = download_one(
            2,
            cached_url,
            cached_root,
            {cached_url: {"sha256": "0" * 64, "path": str(cached_target)}},
            ssl_context(),
            1,
            DEFAULT_MAX_BYTES,
            opener=opener_for(FakeResponse(payload)),
        )
        assert redownloaded["status"] == "downloaded"
        assert redownloaded["sha256"] != "0" * 64

    source = Path(__file__).with_name("download_pdfs.py").read_text(encoding="utf-8")
    assert "_create_unverified_context" not in source
    assert "CERT_NONE" not in source
    print("download safety contract ok")


if __name__ == "__main__":
    main()
