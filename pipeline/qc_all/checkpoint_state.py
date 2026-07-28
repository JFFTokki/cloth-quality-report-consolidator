SUCCESSFUL_STATES = {"succeeded", "permanent_failed"}
RETRYABLE_STATES = {"partial", "retryable_failed"}
IDENTITY_FIELDS = (
    "pdf_sha256",
    "text_extractor_version",
    "ocr_config_version",
    "header_parser_version",
)


def attempt_status(stats, row_count):
    if stats.get("download_failed") or stats.get("missing_file"):
        return "retryable_failed"
    has_parser_error = any(str(key).startswith("pdf_error:") for key in stats)
    if has_parser_error:
        return "partial" if row_count else "retryable_failed"
    return "succeeded" if row_count else "partial"


def should_attempt(entry, *, parser_version, input_identity=None, retry_failed=False):
    if not entry:
        return True
    if entry.get("parser_version") != parser_version:
        return True
    input_identity = input_identity or {}
    if any(entry.get(field, "") != input_identity.get(field, "") for field in IDENTITY_FIELDS):
        return True
    status = entry.get("status")
    if status in SUCCESSFUL_STATES:
        return False
    return retry_failed and status in RETRYABLE_STATES


def choose_result(previous, candidate):
    if not previous:
        return candidate
    if candidate.get("status") == "succeeded":
        return candidate
    previous_rows = len(previous.get("rows") or [])
    candidate_rows = len(candidate.get("rows") or [])
    if previous.get("status") == "succeeded":
        return previous
    if previous_rows > candidate_rows:
        return previous
    return candidate
