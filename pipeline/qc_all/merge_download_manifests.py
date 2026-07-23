import argparse
import json
from pathlib import Path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", required=True)
    parser.add_argument("manifests", nargs="+")
    args = parser.parse_args()

    by_url = {}
    for filename in args.manifests:
        for row in json.loads(Path(filename).read_text(encoding="utf-8")):
            current = by_url.get(row.get("url"))
            if current is None or (current.get("status") == "failed" and row.get("status") != "failed"):
                by_url[row.get("url")] = row
    rows = [row for url, row in by_url.items() if url]
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    Path(args.output).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"output": args.output, "urls": len(rows)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
