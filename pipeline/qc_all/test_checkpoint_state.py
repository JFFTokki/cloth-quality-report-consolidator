from checkpoint_state import attempt_status, choose_result, should_attempt


def main():
    version = "table-parser-test"
    assert attempt_status({}, 3) == "succeeded"
    assert attempt_status({}, 0) == "partial"
    assert attempt_status({"pdf_error:ValueError": 1}, 2) == "partial"
    assert attempt_status({"pdf_error:ValueError": 1}, 0) == "retryable_failed"
    assert attempt_status({"download_failed": 1}, 0) == "retryable_failed"

    assert not should_attempt({"status": "succeeded", "parser_version": version}, parser_version=version)
    assert not should_attempt({"status": "permanent_failed", "parser_version": version}, parser_version=version)
    assert not should_attempt({"status": "retryable_failed", "parser_version": version}, parser_version=version)
    assert should_attempt(
        {"status": "retryable_failed", "parser_version": version},
        parser_version=version,
        retry_failed=True,
    )
    assert should_attempt({"status": "succeeded", "parser_version": "old"}, parser_version=version)
    assert should_attempt(
        {"status": "succeeded", "parser_version": version, "pdf_sha256": "old"},
        parser_version=version,
        input_identity={"pdf_sha256": "new"},
    )
    previous = {"status": "partial", "rows": [{"item": "A"}, {"item": "B"}]}
    worse_retry = {"status": "retryable_failed", "rows": []}
    assert choose_result(previous, worse_retry) is previous
    better_retry = {"status": "succeeded", "rows": [{"item": "A"}]}
    assert choose_result(previous, better_retry) is better_retry
    print("checkpoint state contract ok")


if __name__ == "__main__":
    main()
