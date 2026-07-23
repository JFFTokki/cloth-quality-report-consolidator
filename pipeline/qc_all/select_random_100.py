import argparse
import copy
import json
import random
import re
import os
from collections import Counter
from pathlib import Path
from urllib.parse import urlparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", default="pipeline/qc_all/source_records.json")
    parser.add_argument("--work-dir", required=True)
    parser.add_argument("--count", type=int, default=100)
    parser.add_argument("--seed", default="2026Q4-20260715-latest")
    parser.add_argument("--exclude-host", default="")
    parser.add_argument(
        "--max-per-primary-sample-type",
        type=int,
        default=0,
        help="分层随机抽样时限制单一主要样品类型的款号上限；0表示不限制",
    )
    parser.add_argument(
        "--exclude-skus-file",
        action="append",
        default=[],
        help="JSON文件，支持{skus:[...]}、{selected_skus:[...]}或SKU数组；可重复指定",
    )
    parser.add_argument(
        "--include-sku",
        action="append",
        default=[],
        help="必须包含的12位货号；可重复指定，也可用逗号/空格/换行分隔",
    )
    parser.add_argument(
        "--require-cache-manifest",
        action="append",
        default=[],
        help="仅从全部PDF已存在于这些下载清单的款号中抽样；可重复指定",
    )
    args = parser.parse_args()

    source = json.loads(Path(args.source).read_text(encoding="utf-8"))

    # 源表的货号单元格可能同时写入多个货号（以 /、空格等分隔）。
    # 抽样单位必须是独立的纯数字货号，不能把整串文本当成一个款号。
    expanded_records = []
    for row in source["records"]:
        skus = list(dict.fromkeys(re.findall(r"(?<!\d)\d{12}(?!\d)", row.get("sku", ""))))
        for sku in skus:
            expanded = copy.deepcopy(row)
            expanded["sku_raw"] = row.get("sku", "")
            expanded["sku"] = sku
            expanded_records.append(expanded)
    expanded_selected = [row for row in expanded_records if row.get("is_selected")]
    records_by_sku = {}
    for row in expanded_selected:
        records_by_sku.setdefault(row["sku"], []).append(row)
    candidates = sorted(records_by_sku)
    excluded_skus = set()
    for filename in args.exclude_skus_file:
        payload = json.loads(Path(filename).read_text(encoding="utf-8"))
        values = payload if isinstance(payload, list) else payload.get("skus", payload.get("selected_skus", []))
        for value in values:
            excluded_skus.update(re.findall(r"(?<!\d)\d{12}(?!\d)", str(value)))
    if excluded_skus:
        candidates = [sku for sku in candidates if sku not in excluded_skus]
    cached_urls = set()
    for filename in args.require_cache_manifest:
        for row in json.loads(Path(filename).read_text(encoding="utf-8")):
            if row.get("status") != "failed" and os.path.exists(row.get("path", "")):
                cached_urls.add(row.get("url"))
    if args.require_cache_manifest:
        candidates = [
            sku for sku in candidates
            if all(url in cached_urls for row in records_by_sku[sku] for url in row.get("urls", []))
        ]
    if args.exclude_host:
        candidates = [
            sku for sku in candidates
            if records_by_sku.get(sku)
            and all(
                urlparse(url).hostname != args.exclude_host
                for row in records_by_sku[sku]
                for url in row.get("urls", [])
            )
        ]
    required_skus = []
    for value in args.include_sku:
        for sku in re.findall(r"(?<!\d)\d{12}(?!\d)", str(value)):
            if sku not in required_skus:
                required_skus.append(sku)
    missing_required = [sku for sku in required_skus if sku not in candidates]
    if missing_required:
        raise ValueError(f"指定必须包含的款号不在可选池中: {', '.join(missing_required)}")
    if len(required_skus) > args.count:
        raise ValueError(f"指定必须包含的款号数量{len(required_skus)}超过目标{args.count}个")
    if len(candidates) < args.count:
        raise ValueError(f"可选款号只有{len(candidates)}个，少于目标{args.count}个")

    rng = random.Random(args.seed)
    required_set = set(required_skus)
    if args.max_per_primary_sample_type:
        buckets = {}
        for sku in candidates:
            if sku in required_set:
                continue
            primary_type = next((row.get("sample_type") or "未标明" for row in records_by_sku[sku]), "未标明")
            buckets.setdefault(primary_type, []).append(sku)
        for values in buckets.values():
            rng.shuffle(values)
        sample_skus = list(required_skus)
        type_counts = Counter()
        for sku in required_skus:
            primary_type = next((row.get("sample_type") or "未标明" for row in records_by_sku[sku]), "未标明")
            type_counts[primary_type] += 1
        active_types = sorted(buckets)
        while active_types and len(sample_skus) < args.count:
            next_types = []
            for sample_type in active_types:
                if buckets[sample_type] and type_counts[sample_type] < args.max_per_primary_sample_type:
                    sample_skus.append(buckets[sample_type].pop())
                    type_counts[sample_type] += 1
                    if len(sample_skus) == args.count:
                        break
                if buckets[sample_type] and type_counts[sample_type] < args.max_per_primary_sample_type:
                    next_types.append(sample_type)
            active_types = next_types
        if len(sample_skus) < args.count:
            remaining = [sku for values in buckets.values() for sku in values if sku not in sample_skus]
            rng.shuffle(remaining)
            sample_skus.extend(remaining[:args.count - len(sample_skus)])
        sample_skus = sorted(sample_skus)
    else:
        remaining_candidates = [sku for sku in candidates if sku not in required_set]
        sample_skus = sorted(required_skus + rng.sample(remaining_candidates, args.count - len(required_skus)))
    sample_set = set(sample_skus)
    records = [row for row in expanded_records if row.get("sku") in sample_set]
    selected_records = [row for row in expanded_selected if row.get("sku") in sample_set]
    selected_urls = list(dict.fromkeys(url for row in selected_records for url in row.get("urls", [])))
    summary_keys = sorted(
        ({"sku": row["sku"], "color": color} for row in selected_records for color in row.get("selected_colors", [])),
        key=lambda row: (row["sku"], row["color"]),
    )
    seen = set()
    summary_keys = [
        row for row in summary_keys
        if (row["sku"], row["color"]) not in seen and not seen.add((row["sku"], row["color"]))
    ]

    payload = {
        "selected_skus": sample_skus,
        "sample_seed": args.seed,
        "records": records,
        "invalid_records": [],
        "selected_records": selected_records,
        "selected_urls": selected_urls,
        "summary_keys": summary_keys,
    }
    work_dir = Path(args.work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    (work_dir / "source_records.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (work_dir / "sample_skus.json").write_text(
        json.dumps({"seed": args.seed, "skus": sample_skus}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    sample_types = Counter(row.get("sample_type") or "未标明" for row in selected_records)
    print(json.dumps({
        "seed": args.seed,
        "candidate_pool": len(candidates),
        "excluded_skus": len(excluded_skus),
        "required_skus": required_skus,
        "required_skus_count": len(required_skus),
        "cached_urls": len(cached_urls),
        "excluded_host": args.exclude_host,
        "max_per_primary_sample_type": args.max_per_primary_sample_type,
        "sample_skus": len(sample_skus),
        "source_records": len(records),
        "selected_records": len(selected_records),
        "summary_rows": len(summary_keys),
        "pdf_urls": len(selected_urls),
        "multi_pdf_records": sum(len(row.get("urls", [])) > 1 for row in selected_records),
        "sample_types": dict(sample_types),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
