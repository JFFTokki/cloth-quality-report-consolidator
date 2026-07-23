import re


HEADER_TOP_MIN_Y = 0.68
INSTITUTION_PENDING = "机构名称待确认"
INSTITUTION_UNKNOWN = "未识别"

_MARK_SIDE_TOKENS = (
    "CMA", "CNAS", "中国认可", "国际互认", "资质认定", "TESTING", "CALIBRATION",
    "INSPECTION", "检测能力", "认可准则", "实验室认可",
)

_INSTITUTION_EXCLUSIONS = (
    "委托单位", "委托方", "送检单位", "受检单位", "被检单位", "申请单位", "申请人",
    "客户", "生产单位", "生产者", "标称生产企业", "供应商", "制造商",
    "附设机构", "分包机构", "分包实验室", "未经", "复制报告", "无效", "本中心批准",
)

_INSTITUTION_TERMS = ("检测", "检验", "测试", "试验", "实验室", "认证", "计量", "质量技术服务", "质检")
_INSTITUTION_SUFFIXES = (
    "检测实验室", "检验实验室", "质量中心检测部", "检验检测中心", "质量检验中心",
    "纤维检验中心", "检测中心", "检验中心", "测试中心", "研究院有限公司", "集团有限公司",
    "股份有限公司", "有限公司", "研究院", "实验室", "中心检测部",
)

_INSTITUTION_DISPLAY_ALIASES = (
    (r"cttc\.net\.cn|中纺标", "中纺标检验认证股份有限公司"),
    (r"ISTEST|IST-J003A|创标[（(]?北京[）)]?检测技术服务有限公司", "创标（北京）检测技术服务有限公司"),
    (r"ECT\s*远东正大|远东正大检验集团有限公司", "远东正大检验集团有限公司"),
)


def select_header_blocks(blocks, min_y=HEADER_TOP_MIN_Y):
    """Vision uses normalized bottom-left coordinates; keep the full-width top band."""
    return [
        block for block in (blocks or [])
        if float(block.get("mid_y") or 0) >= min_y
    ]


def _normalize_institution_text(value):
    translation = str.maketrans({
        "質": "质", "檢": "检", "測": "测", "驗": "验", "認": "认", "證": "证",
        "術": "术", "纖": "纤", "團": "团", "臺": "台", "灣": "湾",
    })
    return re.sub(r"\s+", "", str(value or "").translate(translation))


def _institution_candidate(segment):
    raw = str(segment or "").strip()
    if not raw or any(token in raw for token in _INSTITUTION_EXCLUSIONS):
        return ""
    value = _normalize_institution_text(raw)
    for token in _MARK_SIDE_TOKENS:
        value = re.sub(re.escape(token), "", value, flags=re.I)
    value = re.sub(r"^(?:L\s*)?\d{4,12}(?=[\u4e00-\u9fff])", "", value, flags=re.I)
    value = re.sub(r"^(?:MMA|CMA|MA)[A-Za-z]{0,3}(?=[\u4e00-\u9fff])", "", value, flags=re.I)
    value = re.sub(r"^[A-Z]{2,8}(?=[\u4e00-\u9fff])", "", value)
    value = re.sub(r"^[匦囗口]+(?=[\u4e00-\u9fff])", "", value)
    value = re.sub(r"^(?:Semir森[馬马]|森[馬马])(?=[\u4e00-\u9fff])", "", value, flags=re.I)
    value = re.sub(r"^(?:批准签发|报告签发|报告出具机构|签发机构|机构名称|实验室)[:：]?", "", value)
    value = re.split(r"(?:Tel\.?|Fax\.?|电话|地址|网址|https?://)", value, maxsplit=1, flags=re.I)[0]
    value = value.replace("：", "").replace(":", "").strip("；;，,。|/-")
    suffix_pattern = "|".join(re.escape(suffix) for suffix in sorted(_INSTITUTION_SUFFIXES, key=len, reverse=True))
    suffix_matches = list(re.finditer(rf"(?:{suffix_pattern})", value))
    if not suffix_matches:
        return ""
    end = max(match.end() for match in suffix_matches)
    parenthetical = re.match(r"[（(][^）)]{1,12}[）)]", value[end:])
    if parenthetical:
        end += parenthetical.end()
    value = value[:end]
    if not (4 <= len(value) <= 70):
        return ""
    if value.startswith(("经理", "主任", "本公司", "中心化学")):
        return ""
    if not any(token in value for token in _INSTITUTION_TERMS):
        return ""
    suffix_value = re.sub(r"[（(][^）)]{1,12}[）)]$", "", value)
    if suffix_value in _INSTITUTION_SUFFIXES:
        return ""
    if not suffix_value.endswith(_INSTITUTION_SUFFIXES):
        return ""
    return value.replace("(福建)", "（福建）")


def _institution_quality(value):
    suffix_value = re.sub(r"[（(][^）)]{1,12}[）)]$", "", value)
    if "有限公司" in suffix_value:
        legal_form = 4
    elif suffix_value.endswith("研究院"):
        legal_form = 3
    elif suffix_value.endswith(("检验检测中心", "质量检验中心", "纤维检验中心", "检测中心", "检验中心")):
        legal_form = 2
    else:
        legal_form = 1
    business_specificity = sum(token in value for token in ("计量", "检验检测", "质量监督检验", "检测认证", "检验认证"))
    return legal_form, business_specificity


def guess_institution(text):
    """Extract a legal issuing institution without relying on a layout or identity registry."""
    source = str(text or "")
    candidates = []
    lines = source.splitlines()
    previous_context = ""
    for line_no, line in enumerate(lines):
        if line.strip() == "__QC_SOURCE_BREAK__":
            previous_context = ""
            continue
        if not line.strip():
            continue
        line_context = _normalize_institution_text(line)
        if any(token in line_context for token in _INSTITUTION_EXCLUSIONS):
            previous_context = line_context
            continue
        if previous_context and len(previous_context) <= 24 and any(
            token in previous_context for token in _INSTITUTION_EXCLUSIONS
        ):
            previous_context = line_context
            continue
        if line_no + 1 < len(lines):
            next_context = _normalize_institution_text(lines[line_no + 1])
            if next_context in {"测试中心", "检测中心", "检验中心", "检测实验室", "检验实验室"}:
                combined_candidate = _institution_candidate(line_context + next_context)
                if combined_candidate:
                    candidates.append((combined_candidate, line_no, False))
        # OCR often places a legal name and CMA/CNAS wording on one visual row.
        # Evaluate tab-separated blocks independently so mark text does not
        # cause the legal-name block to be discarded.
        parts = [part for part in line.split("\t") if part.strip()]
        segments = [(line, False), *((part, True) for part in parts)]
        for segment, is_block in segments:
            candidate = _institution_candidate(segment)
            if candidate:
                candidates.append((candidate, line_no, is_block))
        previous_context = line_context
    if not candidates:
        return INSTITUTION_PENDING

    counts = {}
    block_counts = {}
    first_seen = {}
    for candidate, line_no, is_block in candidates:
        counts[candidate] = counts.get(candidate, 0) + 1
        block_counts[candidate] = block_counts.get(candidate, 0) + int(is_block)
        first_seen.setdefault(candidate, line_no)
    return max(
        counts,
        key=lambda value: (*_institution_quality(value), len(value), block_counts[value], counts[value], -first_seen[value]),
    )


def guess_institution_display(text):
    """Return a best-effort display value; this business field never requests manual confirmation."""
    source = str(text or "")
    strict = guess_institution(source)
    if strict != INSTITUTION_PENDING:
        return strict

    previous_context = ""
    for line in source.splitlines():
        if line.strip() == "__QC_SOURCE_BREAK__":
            previous_context = ""
            continue
        if not line.strip():
            continue
        line_context = _normalize_institution_text(line)
        excluded = any(token in line_context for token in _INSTITUTION_EXCLUSIONS)
        excluded = excluded or (
            previous_context and len(previous_context) <= 24
            and any(token in previous_context for token in _INSTITUTION_EXCLUSIONS)
        )
        if not excluded:
            for pattern, display_name in _INSTITUTION_DISPLAY_ALIASES:
                if re.search(pattern, line, re.I):
                    return display_name
        previous_context = line_context
    return INSTITUTION_UNKNOWN


def normalize_institution_display_value(value):
    text = str(value or "").strip()
    return INSTITUTION_UNKNOWN if not text or text == INSTITUTION_PENDING else text


def _normalized_cma_token(value):
    return re.sub(r"[^A-Z]", "", str(value or "").upper())


def _qualification_numbers(value):
    numbers = []
    for match in re.finditer(r"(?<!\d)(?:\d[\s._-]*){10,12}(?!\d)", str(value or "")):
        number = re.sub(r"\D", "", match.group(0))
        if len(number) == 12:
            numbers.append(number)
    return numbers


def _horizontally_aligned(mark_block, number_block):
    if mark_block.get("mid_x") is None or number_block.get("mid_x") is None:
        return False
    mark_x = float(mark_block["mid_x"])
    number_x = float(number_block["mid_x"])
    mark_width = float(mark_block.get("width") or 0)
    number_width = float(number_block.get("width") or 0)
    if mark_width > 0 and number_width > 0:
        mark_left, mark_right = mark_x - mark_width / 2, mark_x + mark_width / 2
        number_left, number_right = number_x - number_width / 2, number_x + number_width / 2
        overlap = max(0.0, min(mark_right, number_right) - max(mark_left, number_left))
        return overlap / min(mark_width, number_width) >= 0.25
    return abs(mark_x - number_x) <= 0.04


def find_cma_combination(blocks):
    """Accept a stylized CMA OCR token only with a clear aligned qualification number."""
    mark_blocks = []
    number_blocks = []
    for block in blocks or []:
        text = str(block.get("text") or "")
        if _normalized_cma_token(text) in {"CMA", "MA", "MMA"} and float(block.get("confidence") or 0) >= 0.25:
            mark_blocks.append(block)
        if float(block.get("confidence") or 0) >= 0.8:
            number_blocks.extend((block, number) for number in _qualification_numbers(text))
    for mark_block in mark_blocks:
        for number_block, number in number_blocks:
            if mark_block.get("mid_y") is None or number_block.get("mid_y") is None:
                continue
            mark_y = float(mark_block["mid_y"])
            number_y = float(number_block["mid_y"])
            if _horizontally_aligned(mark_block, number_block) and 0 <= mark_y - number_y <= 0.15:
                return number, str(mark_block.get("text") or "")
    return "", ""


def mark_patterns(mark):
    token = rf"(?<![A-Z0-9]){re.escape(mark)}(?![A-Z0-9])"
    return [rf"{token}\s*[A-Z]?\s*\d+", token]


def has_mark_text(text, mark):
    return any(re.search(pattern, text or "", re.I) for pattern in mark_patterns(mark))


def has_nonnegative_mark_text(text, mark):
    negative = re.compile(
        rf"(?:未|不|暂未).{{0,12}}{re.escape(mark)}|{re.escape(mark)}.{{0,12}}(?:未授权|不适用|未获得)",
        re.I,
    )
    token = re.compile(rf"(?<![A-Z0-9]){re.escape(mark)}(?![A-Z0-9])", re.I)
    source = str(text or "")
    for match in token.finditer(source):
        window = source[max(0, match.start() - 18):min(len(source), match.end() + 24)]
        if not negative.search(window):
            return True
    return False


def extract_labeled_values(text: str, labels, value_pattern: str):
    values = []
    for label in labels:
        for match in re.finditer(rf"{label}\s*[：:]?\s*({value_pattern})", text or "", re.I):
            value = re.sub(r"\s+", "", match.group(1)).strip("：:;；,，")
            if value and value not in values:
                values.append(value)
    return values


def extract_same_line_labeled_values(text: str, labels, stop_labels=()):
    values = []
    stop_pattern = "|".join(stop_labels)
    for label in labels:
        for match in re.finditer(rf"{label}\s*[：:]?\s*([^\u3400-\u9fff]*)", str(text or ""), re.I):
            value = match.group(1)
            if stop_pattern:
                value = re.split(rf"\s*(?:{stop_pattern})\s*[：:]?", value, maxsplit=1)[0]
            value = re.sub(r"\s+", "", value)
            if re.search(r"[A-Za-z0-9]", value) and value not in values:
                values.append(value)
    return values


CODE_VALUE_PATTERN = r"[A-Za-z0-9]+(?:\s*[-/_]\s*[A-Za-z0-9]+)*"


DATE_TOKEN = r"(?:20\s*\d{2}\s*[年./-]\s*\d{1,2}\s*[月./-]\s*\d{1,2}\s*日?)"


def normalize_date(value: str) -> str:
    match = re.search(r"(20\s*\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})", value or "")
    if not match:
        return ""
    year = int(re.sub(r"\s+", "", match.group(1)))
    month = int(match.group(2))
    day = int(match.group(3))
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return ""
    return f"{year:04d}-{month:02d}-{day:02d}"


def page_texts(document):
    return [(int(page.get("page") or index + 1), str(page.get("text") or "")) for index, page in enumerate(document.get("pages") or [])]


def candidate_windows(page_no: int, text: str, label_pattern: str, canonical_label: str):
    windows = []
    for match in re.finditer(label_pattern, text or "", re.I):
        start = max(0, match.start() - 30)
        end = min(len(text), match.end() + 100)
        window = text[start:end]
        for date_match in re.finditer(rf"({DATE_TOKEN})(?:\s*(?:至|到|~|～|-|—)\s*({DATE_TOKEN}))?", window, re.I):
            selected = normalize_date(date_match.group(2) or date_match.group(1))
            if not selected:
                continue
            original = re.sub(r"\s+", " ", window.strip())
            windows.append({
                "date": selected,
                "label": canonical_label,
                "page": page_no,
                "original": original[:220],
                "is_range": bool(date_match.group(2)),
            })
    return windows


def extract_report_issue_date(document):
    priorities = [
        ("报告签发日期", r"(?:报告)?签发日期"),
        ("出具日期", r"(?:出具日期|Date\s+of\s+Issue)"),
        ("报告日期", r"报告日期"),
        ("检测日期", r"(?:样品)?(?:检测日期|检验日期)"),
    ]
    pages = page_texts(document)
    first_page = pages[:1]
    for canonical_label, label_pattern in priorities:
        candidates = []
        for page_no, text in first_page:
            candidates.extend(candidate_windows(page_no, text, label_pattern, canonical_label))
        search_scope = "首页"
        if not candidates:
            for page_no, text in pages[1:]:
                candidates.extend(candidate_windows(page_no, text, label_pattern, canonical_label))
            search_scope = "全文"
        unique_dates = list(dict.fromkeys(candidate["date"] for candidate in candidates))
        if len(unique_dates) == 1:
            selected = unique_dates[0]
            selected_candidate = next(candidate for candidate in candidates if candidate["date"] == selected)
            reason = f"已识别｜{search_scope}第{selected_candidate['page']}页{canonical_label}：{selected_candidate['original']}"
            return selected, "已识别", canonical_label, selected_candidate["original"], reason, candidates
        if len(unique_dates) > 1:
            selected = max(unique_dates)
            selected_candidate = next(candidate for candidate in candidates if candidate["date"] == selected)
            originals = "；".join(candidate["original"] for candidate in candidates)
            reason = (
                f"已识别｜{canonical_label}同一优先级出现多个日期：{'; '.join(unique_dates)}；"
                f"按最新规则采用较晚日期{selected}（第{selected_candidate['page']}页）"
            )
            return selected, "已识别", canonical_label, originals, reason, candidates
    return "", "未发现", "", "", "未发现｜首页及全文未找到签发日期、出具日期、报告日期、检测日期", []
