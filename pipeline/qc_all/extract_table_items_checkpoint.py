import argparse
import hashlib
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path('tmp/vendor').resolve()))
import pdfplumber
from checkpoint_state import IDENTITY_FIELDS, attempt_status, choose_result, should_attempt
from pipeline_versions import TABLE_PARSER_VERSION
from qc_rules import to_simplified


def parse_requested_skus(env_name, available_skus=None):
    raw = os.environ.get(env_name)
    if raw is None:
        return []
    skus = list(dict.fromkeys(re.findall(r'(?<!\d)\d{12}(?!\d)', raw)))
    if not skus:
        raise ValueError(f'{env_name}已设置，但没有可用的12位款号；为避免误跑全量，已停止')
    if available_skus is not None:
        missing = [sku for sku in skus if sku not in available_skus]
        if missing:
            raise ValueError(f'{env_name}包含源表未选中的款号：{", ".join(missing)}；为避免误跑全量，已停止')
    return skus

ROOT = Path.cwd()
WORK = Path(os.environ.get('QC_WORK_DIR', ROOT / 'pipeline/qc_all')).resolve()
ROWS_PATH = WORK / 'table_items_rows.jsonl'
STATE_PATH = WORK / 'table_items_state.json'
RESULTS_DIR = WORK / 'table_items_results'
source = json.loads((WORK / 'source_records.json').read_text(encoding='utf-8'))
manifest = json.loads((WORK / 'download_manifest.json').read_text(encoding='utf-8'))
pdf_text_docs = json.loads((WORK / 'pdf_text.json').read_text(encoding='utf-8'))
pdf_text = {d['url']: d for d in pdf_text_docs}
available_skus = set(source.get('selected_skus') or [rec.get('sku') for rec in source.get('selected_records', [])])
only_skus = set(parse_requested_skus('QC_ONLY_SKUS', available_skus))

records_by_url = defaultdict(list)
for rec in source['selected_records']:
    if only_skus and rec.get('sku') not in only_skus:
        continue
    for url in rec.get('urls', []):
        records_by_url[url].append(rec)
if only_skus:
    selected_urls = set(records_by_url)
    manifest = [row for row in manifest if row.get('url') in selected_urls]
    (WORK / 'download_manifest.targeted.json').write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8',
    )

known_noise = {
    '序号', '检测项目', '检验项目', '检测结果', '单项判定', '单项评价', '项目描述', '标准要求', '实测值', '标准值', '单位', '检测方法',
    '备注', '结论', '检验结论', '检验结果', '样品信息', '产品信息', '判定依据', '检测类别', '委托检测', '委托送样',
}
noise_patterns = [
    r'^(序号|检测|检验)?项目$', r'^(标准|实测|检测|单项|评价|结果|备注).{0,8}$', r'.*报告.*', r'.*委托.*',
    r'.*样品.*', r'.*产品等级.*', r'.*安全类别.*', r'.*见下.*表.*', r'.*符合.*要求.*', r'.*本页.*', r'.*空白.*',
]

def load_state():
    if STATE_PATH.exists():
        state = json.loads(STATE_PATH.read_text(encoding='utf-8'))
        if 'url_states' not in state:
            state['url_states'] = {
                url: {'status': 'succeeded', 'attempts': 1, 'parser_version': ''}
                for url in state.get('processed_urls', [])
            }
        return state
    return {'processed_urls': [], 'url_states': {}, 'stats': {}}

def save_state(state):
    tmp = STATE_PATH.with_suffix('.json.tmp')
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding='utf-8')
    tmp.replace(STATE_PATH)

def result_path(url):
    return RESULTS_DIR / f"{hashlib.sha1(url.encode('utf-8')).hexdigest()}.json"

def input_identity(manifest_row):
    document = pdf_text.get(manifest_row.get('url')) or {}
    return {
        'pdf_sha256': manifest_row.get('sha256') or document.get('pdf_sha256') or '',
        'text_extractor_version': document.get('text_extractor_version') or '',
        'ocr_config_version': document.get('ocr_config_version') or '',
        'header_parser_version': document.get('header_parser_version') or '',
    }

def load_pdf_result(url):
    target = result_path(url)
    if not target.exists():
        return None
    try:
        return json.loads(target.read_text(encoding='utf-8'))
    except (OSError, json.JSONDecodeError):
        return None

def save_pdf_result(url, rows, status, stats, identity):
    RESULTS_DIR.mkdir(parents=True, exist_ok=True)
    target = result_path(url)
    candidate = {
        'url': url,
        'status': status,
        'parser_version': TABLE_PARSER_VERSION,
        **identity,
        'stats': dict(stats),
        'rows': rows,
    }
    previous = load_pdf_result(url)
    if previous and (
        previous.get('parser_version') != TABLE_PARSER_VERSION
        or any(previous.get(field, '') != identity.get(field, '') for field in IDENTITY_FIELDS)
    ):
        previous = None
    selected = choose_result(previous, candidate)
    temp = target.with_suffix('.json.tmp')
    temp.write_text(json.dumps(selected, ensure_ascii=False), encoding='utf-8')
    temp.replace(target)
    return selected

def recover_result_states(url_states):
    if not RESULTS_DIR.exists():
        return
    for path in RESULTS_DIR.glob('*.json'):
        try:
            payload = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, json.JSONDecodeError):
            continue
        url = payload.get('url')
        if not url or payload.get('parser_version') != TABLE_PARSER_VERSION:
            continue
        current = url_states.get(url) or {}
        url_states[url] = {
            **current,
            'status': payload.get('status', 'retryable_failed'),
            'parser_version': TABLE_PARSER_VERSION,
            **{field: payload.get(field, '') for field in IDENTITY_FIELDS},
        }

def aggregate_result_stats():
    total = Counter()
    if not RESULTS_DIR.exists():
        return total
    for path in RESULTS_DIR.glob('*.json'):
        try:
            payload = json.loads(path.read_text(encoding='utf-8'))
        except (OSError, json.JSONDecodeError):
            continue
        if payload.get('parser_version') == TABLE_PARSER_VERSION:
            total.update(payload.get('stats') or {})
    return total

def rebuild_rows_file():
    temp = ROWS_PATH.with_suffix('.jsonl.tmp')
    with temp.open('w', encoding='utf-8') as output:
        if RESULTS_DIR.exists():
            for path in sorted(RESULTS_DIR.glob('*.json')):
                try:
                    payload = json.loads(path.read_text(encoding='utf-8'))
                except (OSError, json.JSONDecodeError):
                    continue
                if payload.get('parser_version') != TABLE_PARSER_VERSION:
                    continue
                if payload.get('status') not in {'succeeded', 'partial'}:
                    continue
                for row in payload.get('rows') or []:
                    output.write(json.dumps(row, ensure_ascii=False) + '\n')
    temp.replace(ROWS_PATH)

def clean_cell(v):
    if v is None:
        return ''
    return re.sub(r'\s+', ' ', str(v).replace('\r', '\n')).strip()

def clean_cell_keep_lines(v):
    if v is None:
        return ''
    text = str(v).replace('\r', '\n')
    text = re.sub(r'[ \t]+', ' ', text)
    return text.strip()

def zh_norm(v):
    return to_simplified(v or '')

def compact(v):
    return re.sub(r'\s+', '', zh_norm(v or ''))

def page_has_keywords(url, page_no):
    doc = pdf_text.get(url, {})
    pages = doc.get('pages') or []
    if page_no - 1 >= len(pages):
        return True
    text = pages[page_no - 1].get('text', '')
    c = compact(text)
    if len(c) < 20:
        return False
    score = 0
    if any(token in c for token in ('检测项目', '检验项目', '测试项目', '检验检测项目')):
        score += 3
    if any(token in c for token in ('检测结果', '检验结果', '测试结果', '实测结果', '实测值')):
        score += 2
    if any(token in c for token in ('单项判定', '单项评价', '单项结论', '符合', '合格', '不符合', '不合格')):
        score += 2
    if any(token in c for token in ('GB/T', 'GB', 'FZ/T', 'QB/T', 'ISO', 'AATCC', 'ASTM', 'IDFB', '检测依据', '检验依据')):
        score += 2
    if re.search(r'[≤≥<>≦≧]\s*\d', c):
        score += 1
    # 至少需要两类证据。封面、说明页和纯照片页通常达不到4分。
    return score >= 4

def is_header_row(row):
    joined = ''.join(compact(c) for c in row)
    item_header = any(token in joined for token in ('检测项目', '检验项目', '测试项目', '检验检测项目'))
    result_header = any(token in joined for token in (
        '单项', '判定', '评价', '结论', '检测结果', '检验结果', '测试结果',
        '实测结果', '实测值', '标准值', '标准要求', '要求值',
    ))
    # 无网格报告经文本定位后，常见表头只有“标准（称）值 实测值 单项判定”，
    # 第一列没有重复印“检测项目”，仍应建立列结构。
    inferred_header = any(token in joined for token in ('实测值', '实测结果', '测试结果')) and any(
        token in joined for token in ('判定', '评价', '结论')
    )
    return (item_header and result_header) or inferred_header

def header_indices(row):
    idx = {
        'item': None, 'item_cols': [], 'detail_cols': [], 'method': None,
        'unit': None, 'requirement': None, 'result': None, 'verdict': None,
        'cas': None, 'detection_limit': None, 'inferred': False,
        'wrapped_item_cols': False,
    }
    compacted = [compact(cell) for cell in row]
    for i, cell in enumerate(row):
        c = compacted[i]
        pair_left = compacted[i - 1] + c if i else c
        pair_right = c + compacted[i + 1] if i + 1 < len(compacted) else c
        combined = pair_left + pair_right
        item_tokens = ('检测项目', '检验项目', '测试项目', '检验检测项目')
        if any(token in c for token in item_tokens):
            idx['item'] = i
            idx['item_cols'] = [i]
            idx['wrapped_item_cols'] = False
        elif any(token in pair_right for token in item_tokens):
            idx['item'] = i
            idx['item_cols'].append(i)
            idx['wrapped_item_cols'] = True
        if c in ('检测方法', '检验方法') or '检测方法' in c or '检验方法' in c:
            idx['method'] = i
        if c == '单位' or '单位' in c:
            idx['unit'] = i
        if c in {'CAS号', 'CASNO.', 'CASNO', 'CAS编号'} or re.fullmatch(r'CAS(?:号|NO\.?)?', c, re.I):
            idx['cas'] = i
        if any(token in c for token in ('报告限', '检出限', '检测限', '定量限')) or c in {'LOD', 'LOQ'}:
            idx['detection_limit'] = i
        requirement_tokens = ('标准要求', '标准值', '要求值', '技术要求')
        split_requirement_tokens = ('标准称值', '标准（称）值', '标准(称)值')
        if any(token in c for token in requirement_tokens):
            idx['requirement'] = i
        elif idx['requirement'] is None and any(token in pair_left for token in split_requirement_tokens):
            idx['requirement'] = i
        elif idx['requirement'] is None and any(token in pair_right for token in split_requirement_tokens):
            idx['requirement'] = i + 1
        if any(token in c for token in ('检测结果', '检验结果', '测试结果', '实测结果', '实测值', '测试值')):
            idx['result'] = i
        verdict_tokens = ('单项判定', '单项评价', '单项结论')
        if any(token in c for token in verdict_tokens) or c in {'判定', '评价', '结论'}:
            idx['verdict'] = i
        elif idx['verdict'] is None and any(token in pair_left for token in verdict_tokens):
            idx['verdict'] = i
        elif idx['verdict'] is None and any(token in pair_right for token in verdict_tokens):
            idx['verdict'] = i + 1
        if any(token in c for token in ('细项', '子项', '组分', '成分项', '项目描述')):
            idx['detail_cols'].append(i)
    if idx['verdict'] is None:
        for end in range(len(compacted)):
            window = ''.join(compacted[max(0, end - 2):end + 1])
            if any(token in window for token in ('单项判定', '单项评价', '单项结论')):
                idx['verdict'] = end
                break
    if idx['requirement'] is None:
        for end in range(len(compacted)):
            window = ''.join(compacted[max(0, end - 2):end + 1])
            if any(token in window for token in ('标准称值', '标准（称）值', '标准(称)值')):
                idx['requirement'] = end
                break
    if idx['item'] is not None:
        boundary_fields = [idx['method'], idx['unit'], idx['cas'], idx['detection_limit'], idx['requirement'], idx['result'], idx['verdict']]
        if idx['wrapped_item_cols']:
            boundary_fields.remove(idx['unit'])
        stop_candidates = [i for i in boundary_fields if i is not None and i > idx['item']]
        stop = min(stop_candidates) if stop_candidates else len(row)
        # Many reports split the "检测项目" area into parent item and sub-item columns.
        # Treat blank header columns between item and result/verdict as part of the item area.
        idx['item_cols'] = list(range(idx['item'], max(idx['item'] + 1, stop)))
    elif idx['result'] is not None or idx['verdict'] is not None:
        stop_candidates = [i for i in (idx['method'], idx['unit'], idx['cas'], idx['detection_limit'], idx['requirement'], idx['result'], idx['verdict']) if i is not None]
        stop = min(stop_candidates) if stop_candidates else len(row)
        start = 1 if compacted and compacted[0] in {'序号', '编号', 'No.', 'NO.'} else 0
        if stop > start:
            idx['item'] = start
            # 无“检测项目”表头时，中间空表头列常是文字定位产生的溢出区，
            # 其中可能落入标准要求或结果的断字，不能一律拼进项目名。
            idx['item_cols'] = [start]
            idx['inferred'] = True
    return idx


def infer_headerless_schema(table):
    """从没有表头的规则数据表中推断列结构，覆盖羽绒及机构续页。"""
    rows = [[clean_cell(c) for c in row] for row in table if row]
    if not rows:
        return None
    width = max(len(row) for row in rows)
    if width < 5:
        return None
    padded = [row + [''] * (width - len(row)) for row in rows]

    def score(col, predicate):
        return sum(bool(predicate(row[col])) for row in padded)

    single_data_row = len(padded) == 1
    min_hits = 1 if single_data_row else 2
    seq_candidates = [i for i in range(width) if score(i, lambda v: re.fullmatch(r'\d{1,3}', compact(v or ''))) >= min_hits]
    method_candidates = [i for i in range(width) if score(i, lambda v: bool(re.search(r'GB/?T|GB\s|FZ/?T|QB/?T|ISO|AATCC|ASTM|SN/?T|IDFB', v or '', re.I))) >= min_hits]
    verdict_candidates = [i for i in range(width) if score(i, lambda v: bool(re.search(r'不符合|不合格|符合|合格|通过|不适用', zh_norm(v or '')))) >= min_hits]
    requirement_candidates = [i for i in range(width) if score(i, lambda v: bool(re.search(r'[≤≥<>≦≧]\s*\d', v or ''))) >= min_hits]
    if not method_candidates or not verdict_candidates or not requirement_candidates:
        return None
    seq_col = min(seq_candidates) if seq_candidates else None
    method_col = max(method_candidates, key=lambda i: score(i, lambda v: bool(re.search(r'GB/?T|GB\s|FZ/?T|QB/?T|ISO|AATCC|ASTM|SN/?T|IDFB', v or '', re.I))))
    verdict_col = max(verdict_candidates)
    requirement_col = max(requirement_candidates, key=lambda i: score(i, lambda v: bool(re.search(r'[≤≥<>≦≧]\s*\d', v or ''))))
    result_col = requirement_col + 1 if requirement_col + 1 < verdict_col else None
    if result_col is None:
        return None
    start = (seq_col + 1) if seq_col is not None else 0
    if method_col <= start:
        return None
    # 多行无表头表格必须具有稳定的序号列，避免把说明/条款表误当检测结果。
    if not single_data_row and seq_col is None:
        return None
    item_cols = [start]
    detail_cols = [i for i in range(method_col + 1, requirement_col) if i != result_col]
    unit_col = detail_cols[-1] if detail_cols else None
    if unit_col is not None:
        detail_cols = detail_cols[:-1]
    return {
        'item': start,
        'item_cols': item_cols,
        'detail_cols': detail_cols,
        'method': method_col,
        'unit': unit_col,
        'requirement': requirement_col,
        'result': result_col,
        'verdict': verdict_col,
        'sequence': seq_col,
        'inferred': True,
    }


def ocr_tables_for_page(url, page_no):
    """利用Vision文字块坐标重建扫描PDF中的表格列。"""
    doc = pdf_text.get(url, {})
    pages = doc.get('pages') or []
    if page_no - 1 >= len(pages):
        return []
    blocks = pages[page_no - 1].get('ocr_blocks') or []
    if not blocks:
        return []
    ordered = sorted(blocks, key=lambda b: (-float(b.get('mid_y', 0)), float(b.get('x', 0))))
    visual_rows = []
    for block in ordered:
        text = clean_cell(block.get('text'))
        if not text:
            continue
        y = float(block.get('mid_y', 0))
        tolerance = max(0.008, float(block.get('height', 0)) * 0.45)
        if not visual_rows or abs(visual_rows[-1][0] - y) > tolerance:
            visual_rows.append([y, [block]])
        else:
            visual_rows[-1][1].append(block)
    tables = []
    for header_pos, (_, header_blocks) in enumerate(visual_rows):
        header_blocks = sorted(header_blocks, key=lambda b: float(b.get('x', 0)))
        joined = ''.join(compact(b.get('text')) for b in header_blocks)
        if not any(token in joined for token in ('检测项目', '检验项目', '测试项目', '检验检测项目')):
            continue
        if not any(token in joined for token in ('结果', '实测值', '判定', '单项')):
            continue
        anchors = [float(block.get('mid_x', 0)) for block in header_blocks]
        if len(anchors) < 3:
            continue
        header = [clean_cell(block.get('text')) for block in header_blocks]
        start_pos = header_pos + 1
        if header and compact(header[-1]) == '单项' and start_pos < len(visual_rows):
            next_blocks = visual_rows[start_pos][1]
            if len(next_blocks) == 1 and compact(next_blocks[0].get('text')) in {'判定', '评价', '结论'}:
                header[-1] += clean_cell(next_blocks[0].get('text'))
                start_pos += 1
        table = [header]
        pending = [''] * len(anchors)
        for _, row_blocks in visual_rows[start_pos:]:
            row_blocks = sorted(row_blocks, key=lambda b: float(b.get('x', 0)))
            row_joined = ''.join(compact(b.get('text')) for b in row_blocks)
            if any(token in row_joined for token in ('报告结束', '地址：', '电话：', '样品照片')):
                break
            cells = [''] * len(anchors)
            for block in row_blocks:
                mid_x = float(block.get('mid_x', 0))
                col = min(range(len(anchors)), key=lambda i: abs(mid_x - anchors[i]))
                value = clean_cell(block.get('text'))
                cells[col] = f"{cells[col]} {value}".strip() if cells[col] else value
            populated = [i for i, value in enumerate(cells) if value]
            if not populated:
                continue
            # OCR可能把同一单元格的第二行识别成独立视觉行，暂存并并入下一数据行。
            if len(populated) == 1 and not re.fullmatch(r'\d{1,3}', compact(cells[populated[0]])):
                col = populated[0]
                pending[col] = f"{pending[col]} {cells[col]}".strip() if pending[col] else cells[col]
                continue
            for col, value in enumerate(pending):
                if value:
                    cells[col] = f"{value} {cells[col]}".strip() if cells[col] else value
            pending = [''] * len(anchors)
            table.append(cells)
        if len(table) > 1:
            tables.append(table)
            break
    return tables

def split_item_cell(text):
    text = clean_cell(text)
    if not text:
        return []
    return [p.strip(' ：:/') for p in re.split(r'[,，;；、]\s*|\n+', text) if p.strip(' ：:/')]

def normalize_item_name(name):
    name = zh_norm(clean_cell(name))
    name = re.sub(r'^\d{1,3}[.)、]?\s*', '', name)
    name = re.sub(r'\((?:级|個|个|mg/kg|cm|mm|N|%)\)|（(?:级|個|个|mg/kg|cm|mm|N|%)）', '', name, flags=re.I)
    name = name.replace('(级)', '').replace('（级）', '').replace('(級)', '').replace('（級）', '')
    name = re.sub(r'\((?=-)', '', name)
    name = re.split(r'GB/T|GB |FZ/T|QB/T|ISO|AATCC|ASTM|SN/T|EN ', name, maxsplit=1)[0]
    name = re.split(r'客户要求|采购内控标准|Q/|Semir|≤|≥|<|>|=|＝|N\.D\.|未检出|mg/kg|cm|mm|级', name, maxsplit=1, flags=re.I)[0]
    name = re.sub(r'^[\d\s,，.、:：;；/／\\\-—－●•·▪■]+', '', name)
    name = re.sub(r'[\d\s,，.、:：;；/／\\\-—－●•·▪■_%(（]+$', '', name)
    name = re.sub(r'(^|[-—－])[/／\\]+(?=$|[-—－])', r'\1', name)
    name = re.sub(r'[-—－]{2,}', '-', name)
    name = name.strip('-—－/／\\,，.、:：;；●•·▪■ ')
    name = re.sub(r'\s+', '', name)
    return name

def normalize_item_part(name):
    name = zh_norm(clean_cell(name))
    name = re.sub(r'^\d{1,3}[.)、]?\s*', '', name)
    name = re.split(r'GB/T|GB |FZ/T|QB/T|ISO|AATCC|ASTM|SN/T|EN ', name, maxsplit=1)[0]
    name = re.split(r'客户要求|采购内控标准|Q/|Semir|≤|≥|<|>|=|＝|N\.D\.|未检出|mg/kg|cm|mm|级', name, maxsplit=1, flags=re.I)[0]
    name = re.sub(r'^[\d\s,，.、:：;；/／\\\-—－●•·▪■]+', '', name)
    name = re.sub(r'[\d\s,，.、:：;；/／\\\-—－●•·▪■_%(（]+$', '', name)
    name = name.strip('-—－/／\\,，.、:：;；●•·▪■ ')
    value = re.sub(r'\s+', '', name)
    return '' if value in {'/', '／', '\\', '-', '—', '－', ',', '，', '、', '●', '•', '·', '▪', '■'} else value

def is_subitem(part):
    part = compact(part)
    if not part:
        return False
    if re.fullmatch(r'(变色|沾色|沾色[（(［\[].+[）)］\]]|.*布|.*纤维|.*毛|.*棉|.*锦纶|.*腈纶|.*聚酯|.*醋酯|.*羊毛)', part):
        return True
    return False

def compose_item(parent, parts):
    parent = normalize_item_part(parent)
    cleaned_parts = [normalize_item_part(p) for p in parts if normalize_item_part(p)]
    if not cleaned_parts:
        return parent, ''
    if not parent:
        parent = cleaned_parts[0]
        cleaned_parts = cleaned_parts[1:]
    subparts = [p for p in cleaned_parts if p != parent]
    if subparts:
        return f"{parent}-{'-'.join(subparts)}", '；'.join(subparts)
    return parent, ''

def item_parts_from_row(row, item_cols, parent_item, force_child=False):
    values = []
    for col in item_cols:
        if col < len(row):
            value = normalize_item_part(row[col])
            if value and value not in known_noise:
                values.append(value)
    if not values:
        return '', '', parent_item
    first = values[0]
    if (is_subitem(first) or force_child) and parent_item:
        item, detail = compose_item(parent_item, values)
        return item, detail, parent_item
    item, detail = compose_item(first, values[1:])
    return item, detail, first or parent_item


def split_mixed_item_text(text):
    """把误落入项目单元格的要求、方法和方向限值拆回各自字段。"""
    text = clean_cell(text)
    if re.fullmatch(r'方法\s*[A-ZＡ-Ｚ](?:\s*[（(][^）)]{0,20}[）)])?', text, re.I):
        return '', '', text
    requirement_match = re.search(
        r'不\s*[-—－]?\s*应有|应\s*[-—－]?\s*是固定的|长度\s*[-—－]?\s*超过|应超\s*[-—－]?\s*过',
        text,
    )
    if requirement_match:
        prefix = text[:requirement_match.start()].rstrip('-—－,，:：;； ')
        requirement = text[requirement_match.start():].lstrip('-—－,，:：;； ')
        complete_item_suffix = r'(?:含量|牢度|性能|强度|强力|变化率|外观质量|测定|分析|要求)$'
        prose_prefix = r'检测时|测定时|试验时|要求如下|规定如下|不得|不应|应当|超过'
        if re.search(complete_item_suffix, prefix) and not re.search(prose_prefix, prefix):
            return prefix, requirement, ''
        return '', text, ''
    direction = re.match(r'^(直向|横向|纵向|经向|纬向)\s*[-—－]\s*([~～+＋\-－—].+)$', text)
    if direction:
        return direction.group(1), direction.group(2), ''
    split = re.split(r'(?=符装|符\s*[-—－]?\s*合(?:森马|GB|Q/|采购))', text, maxsplit=1, flags=re.I)
    if len(split) == 2:
        return split[0].strip(), split[1].strip(), ''
    if compact(text) in {'符', '合', '装'}:
        return '', text, ''
    return text, '', ''


def inferred_item_fields(row, schema):
    """拆出文本定位表中混在项目视觉列里的项目、要求和方法。"""
    start = schema['item']
    if start >= len(row) or re.fullmatch(r'\d{1,3}', compact(row[start])):
        return '', '', ''
    if schema.get('wrapped_item_cols'):
        stop = max(schema.get('item_cols') or [start]) + 1
    else:
        stop_candidates = [
            value for value in (
                schema.get('method'), schema.get('unit'), schema.get('cas'),
                schema.get('detection_limit'), schema.get('requirement'),
                schema.get('result'), schema.get('verdict'),
            )
            if value is not None and value > start
        ]
        stop = min(stop_candidates) if stop_candidates else len(row)
    parts = []
    requirement_parts = []
    for col in range(start, min(stop, len(row))):
        value = clean_cell(row[col])
        if not value:
            break
        if col > start and re.match(r'^(?:符装|符?合(?:GB|Q/|森马|采购|$)|装(?:采购|$))', compact(value), re.I):
            requirement_parts.append(value)
            break
        if schema.get('wrapped_item_cols'):
            value = re.sub(r'(^|\s)\d{1,3}(?=\s|$)', ' ', value).strip()
            value = re.sub(r'(?i)(?:in\s*\^?\s*3\s*/\s*30|mg\s*/\s*kg|cm|mm|%)', '', value)
            value = re.sub(r'(?<![A-Za-z])(?:N|g)(?![A-Za-z])', '', value).strip()
        parts.append(value)
    wrapped_item = ''.join(parts)
    # 文本定位表常把“项目名 - 颜色款”落入同一视觉列；款式描述不是检测子项。
    wrapped_item = re.sub(r'[-—－][^—－-]{0,60}款(?:.*)?$', '', wrapped_item).strip()
    wrapped_item, embedded_requirement, embedded_method = split_mixed_item_text(wrapped_item)
    if embedded_requirement:
        requirement_parts.insert(0, embedded_requirement)
    return wrapped_item, ''.join(requirement_parts), embedded_method


def item_parts_from_inferred_row(row, schema, parent_item, force_child=False):
    """拼回主项目列断字，并把要求/方法残片留给调用方回填。"""
    item_text, _, _ = inferred_item_fields(row, schema)
    return item_parts_from_row([item_text], [0], parent_item, force_child=force_child)


def join_field_text(left, right):
    left = clean_cell(left)
    right = clean_cell(right)
    if not left:
        return right
    if not right or right in left:
        return left
    if left.endswith('服') and right.startswith('装'):
        return left + right
    for size in range(min(len(left), len(right)), 0, -1):
        if left[-size:] == right[:size]:
            return left + right[size:]
    return f'{left}；{right}'


def requirement_with_left_fragment(row, requirement_col, requirement):
    if requirement_col is None or not requirement:
        return requirement
    prefix = ''
    for col in range(requirement_col - 1, max(-1, requirement_col - 4), -1):
        value = compact(row[col])
        if not value:
            continue
        if value in {'符', '装'}:
            prefix = value + prefix
        break
    return prefix + requirement


def attach_context_fragment(rows, parent_item, field, value, page, table):
    """把独立续行回填到同表中最近的父项/子项记录，不丢弃原文。"""
    parent = normalize_item_name(parent_item)
    source_row = None
    updated = False
    for existing in reversed(rows):
        if existing.get('page') != page or existing.get('table') != table:
            continue
        existing_item = normalize_item_name(existing.get('raw_item', ''))
        if not existing_item.startswith(parent):
            if source_row is not None:
                break
            continue
        if source_row is None:
            source_row = existing.get('raw_row')
        if existing.get('raw_row') != source_row:
            break
        existing[field] = join_field_text(existing.get(field, ''), value)
        updated = True
    return updated


def item_parts_from_schema(row, schema, parent_item):
    """按推断列读取跨行父项和子项，保证羽绒等复合检测项名称完整。"""
    parent_values = []
    for col in schema.get('item_cols') or []:
        if col < len(row):
            value = normalize_item_part(row[col])
            if value and value not in known_noise:
                parent_values.append(value)
    detail_values = []
    for col in schema.get('detail_cols') or []:
        if col < len(row):
            value = normalize_item_part(row[col])
            if value and value not in known_noise:
                detail_values.append(value)
    if parent_values:
        parent_item = parent_values[0]
    if not parent_item:
        return '', '', parent_item
    if detail_values:
        item, detail = compose_item(parent_item, detail_values)
        return item, detail, parent_item
    return parent_item, '', parent_item

def valid_item(name):
    if not name or name in known_noise:
        return False
    if len(name) > 40:
        return False
    if not re.search(r'[\u4e00-\u9fff]', name):
        return False
    if re.search(r'\d', name):
        return False
    if name.startswith(('#', '[#]', 'Requ', 'requ')) or name in {'不', '无', '其他', '内', '要求', '条款', '为不含荧'}:
        return False
    if re.fullmatch(r'方法\s*[A-ZＡ-Ｚ](?:\s*[（(][^）)]{0,20}[）)])?', name, re.I):
        return False
    if name.endswith(('。', '；')):
        return False
    if re.search(r'(^|[-—－])(合格|符合|不适用|无)$', name):
        return False
    if re.search(
        r'检测依据|检验依据|本方法的检出限|低于检出限|报告中的|表示无|'
        r'不应超过|不得|不允许|应超过|应超出|允许出现|已经凝固|全置于|'
        r'伸出的长度|绳圈的周长|两固定端的长度|测定低限|低限结果|'
        r'客户提供|测试部位|测试在基布|企业名称及联系方式|不应有|应是固定的|'
        r'长度[-—－]?超过|应超[-—－]?过|符装|符[-—－]?合(?:森马|GB|Q/|采购)',
        name,
    ):
        return False
    # 无网格文本定位偶尔会把标准条款、样品描述或上一列尾句切成“项目名”。
    # 这些片段没有完整的检测项语义，必须留在原PDF中而不能进入统计列。
    fragment_prefixes = (
        '上的', '且无', '以外', '何绳带', '带上', '带外', '尺寸时', '超出服装',
        '不-绳', 'm-m', 'T春', '主体面层材料', '本方法的',
    )
    if name.startswith(fragment_prefixes):
        return False
    if re.search(r'平摊至|自由末端|装饰性绳带|袖口处绳带|肩带上|连续固定|宜标注', name):
        return False
    for pat in noise_patterns:
        if re.fullmatch(pat, name):
            return False
    short_subitems = {'沾色', '变色'}
    if name in short_subitems:
        return False
    if len(name) == 1 and name not in {'铅', '镉', '汞', '砷', '铬', '镍'}:
        return False
    return True

def infer_verdict(result, verdict):
    text = compact(zh_norm(f'{result} {verdict}'))
    if re.search(r'不符合|不合格', text):
        return '不合格'
    if re.search(r'符合|合格|通过', text):
        return '合格'
    return verdict or ''


def complete_split_verdict(row, verdict_col, verdict):
    value = compact(verdict)
    if value not in {'合', '格'} or verdict_col is None:
        return verdict
    prefix = ''
    completed = ''
    for col in range(verdict_col - 1, max(-1, verdict_col - 4), -1):
        fragment = compact(row[col])
        if not fragment:
            continue
        if fragment not in {'不', '符', '合'}:
            break
        prefix = fragment + prefix
        candidate = prefix + value
        if candidate in {'符合', '不符合', '合格', '不合格'}:
            completed = candidate
    return completed or verdict


def sanitize_requirement_result(requirement, result):
    standard_ref = r'(?:GB|FZ|QB|SN)\s*/?\s*T?|ISO|AATCC|ASTM|EN\s*\d'
    combined = f'{requirement} {result}'
    if not re.search(standard_ref, combined, re.I):
        return requirement, result
    if re.search(standard_ref, requirement, re.I):
        requirement = ''
    if re.search(standard_ref, result, re.I) or re.fullmatch(r'\s*[-—－]?\s*\d{4}(?:\s*[、,，].*)?', result or ''):
        result = ''
    return requirement, result


def valid_cas_number(value):
    match = re.fullmatch(r'(\d{2,7})-(\d{2})-(\d)', value or '')
    if not match:
        return False
    digits = match.group(1) + match.group(2)
    expected = sum(int(digit) * weight for weight, digit in enumerate(reversed(digits), start=1)) % 10
    return expected == int(match.group(3))


def extract_cas_numbers(*values):
    found = []
    for value in values:
        normalized = re.sub(r'\s+', '', str(value or ''))
        for match in re.finditer(r'(?<!\d)(\d{2,7}-\d{2}-\d)(?!\d)', normalized):
            candidate = match.group(1)
            if valid_cas_number(candidate) and candidate not in found:
                found.append(candidate)
    return '\n'.join(found)


def extract_detection_limit(*values):
    found = []
    pattern = re.compile(
        r'(?:报告限|检出限|检测限|定量限|方法检出限|LOD|LOQ)\s*[：:=]?\s*'
        r'((?:[<>≤≥＜＞]?\s*)?\d+(?:\.\d+)?\s*(?:mg/kg|μg/kg|ug/kg|mg/L|μg/L|ug/L|%|％|ppm|ppb)?)',
        re.I,
    )
    for value in values:
        for match in pattern.finditer(str(value or '')):
            candidate = re.sub(r'\s+', '', match.group(1)).replace('％', '%')
            if candidate and candidate not in found:
                found.append(candidate)
    return '\n'.join(found)


def classify_processing_status(manifest_row, stats, document=None, parsed_count=0):
    """把零结果按真实失败阶段拆分，避免统一写成“未识别检测项”。"""
    document = document or {}
    if parsed_count:
        if document.get('text_source') == 'vision_ocr' and stats.get('ocr_documents_parsed'):
            return '已解析（OCR）', '扫描PDF已通过macOS Vision OCR及坐标列定位提取检测项目'
        page_errors = stats.get('pages_no_table', 0)
        pdf_errors = sum(value for key, value in stats.items() if str(key).startswith('pdf_error:'))
        if page_errors or pdf_errors:
            detail = []
            if page_errors:
                detail.append(f'{page_errors}个疑似结果页未能重建表格')
            if pdf_errors:
                detail.append(f'{pdf_errors}个PDF解析异常')
            return '部分解析', '；'.join(detail)
        return '已解析', ''
    if manifest_row.get('status') == 'failed' or stats.get('download_failed'):
        return '下载失败', manifest_row.get('error') or 'PDF下载失败'
    if stats.get('missing_file'):
        return '下载文件缺失', '下载清单存在，但本地PDF文件缺失'
    text_chars = document.get('text_chars', 0) or 0
    if text_chars < 300:
        if document.get('ocr_attempted') and document.get('ocr_text_chars', 0):
            return 'OCR后未识别检测项', '扫描件已执行OCR，但仍未可靠定位检测项目'
        if document.get('ocr_attempted') and document.get('ocr_error'):
            return 'OCR失败', document.get('ocr_error')
        return '扫描件需OCR', 'PDF可打开，但有效文本少于300字，需要OCR'
    if stats.get('pages_table_checked', 0) == 0:
        return '页面识别规则未覆盖', '报告有文本，但未定位到疑似检测结果页'
    if stats.get('pages_no_table', 0) >= stats.get('pages_table_checked', 0):
        return '无可识别表格结构', '疑似结果页可读，但网格和文本定位均未重建出表格'
    if stats.get('candidate_item_cells', 0) and stats.get('invalid_item_cell', 0) >= stats.get('candidate_item_cells', 0):
        return '检测项被有效性规则排除', '已定位候选检测项，但全部被名称有效性规则排除'
    if stats.get('tables_extracted', 0) and not stats.get('header_rows_recognized', 0) and not stats.get('tables_headerless_inferred', 0):
        return '表头字段未匹配', '已抽取表格，但表头/列结构尚未覆盖该机构格式'
    return '未识别检测项', '报告有文本和表格，但尚未可靠提取检测项目及结果'


STATUS_REASON_CODES = {
    '已解析': 'parsed', '已解析（OCR）': 'ocr_parsed', '部分解析': 'partial_parse',
    '下载失败': 'download_failed', '下载文件缺失': 'download_file_missing',
    'OCR待处理': 'ocr_pending', '扫描件需OCR': 'ocr_pending', 'OCR失败': 'ocr_failed',
    'OCR后未识别检测项': 'ocr_no_items', 'PDF文本异常': 'pdf_text_abnormal',
    '页面识别规则未覆盖': 'page_candidate_unmatched',
    '无可识别表格结构': 'no_table_structure', '表头字段未匹配': 'header_unmatched',
    '缺少结果或判定列': 'missing_result_column',
    '检测项被有效性规则排除': 'invalid_item_filtered',
    '项名不完整待复核': 'incomplete_item_name', '未识别检测项': 'unrecognized_items',
}


def processing_metadata(status, stats, document=None):
    document = document or {}
    if document.get('text_source') == 'vision_ocr':
        parse_method = 'OCR'
    elif stats.get('pages_text_table_fallback'):
        parse_method = '无网格文本定位'
    elif stats.get('tables_extracted'):
        parse_method = '原生网格表格'
    else:
        parse_method = ''
    metrics = {
        'native_text_chars': document.get('native_text_chars', document.get('text_chars', 0)),
        'ocr_text_chars': document.get('ocr_text_chars', 0),
        'tables': stats.get('tables_extracted', 0),
        'headers': stats.get('header_rows_recognized', 0),
        'candidate_items': stats.get('candidate_item_cells', 0) + stats.get('ocr_candidate_rows', 0),
        'invalid_items': stats.get('invalid_item_cell', 0) + stats.get('ocr_invalid_item_cell', 0),
        'rows': stats.get('rows', 0),
    }
    if document.get('ocr_error'):
        metrics['ocr_error'] = document.get('ocr_error')
    return {
        'reason_code': STATUS_REASON_CODES.get(status, 'unclassified'),
        'parse_method': parse_method,
        'needs_review': '否' if status in {'已解析', '已解析（OCR）'} else '是',
        'diagnostic_metrics': json.dumps(metrics, ensure_ascii=False, separators=(',', ':')),
    }

def result_subitems(result):
    text = zh_norm(clean_cell_keep_lines(result))
    if not text:
        return []
    # pdfplumber sometimes collapses multi-line result cells into one line.
    text = re.sub(r'\s+(?=(变色|沾色|干摩|湿摩|醋纤|棉|锦纶|聚酯纤维|腈纶|羊毛)\b)', '\n', text)
    text = re.sub(r'\s*[-－]\s*(?=(醋纤|棉|锦纶|聚酯|聚酯纤维|腈纶|羊毛)\b)', '\n', text)
    text = text.replace('聚酯 ', '聚酯纤维 ')
    tokens = []
    current_prefix = ''
    for line in re.split(r'\n+|；|;', text):
        line = clean_cell(line)
        if not line:
            continue
        if line in {'沾色', '变色'}:
            current_prefix = line
            tokens.append((line, line))
            continue
        match = re.match(r'^(?P<label>[A-Za-z\u4e00-\u9fff（）()\[\]［］+]{1,20})\s*(?P<value>(?:>|＜|<|≥|≤|≧|≦)?\d[\d.+\\-～~]*|N\\.?D\\.?|未检出|合格|符合)', line, re.I)
        if not match:
            continue
        label = normalize_item_part(match.group('label'))
        value = match.group('value')
        if not label or not re.search(r'[A-Za-z\u4e00-\u9fff]', label) or re.search(r'^\d', label):
            continue
        if current_prefix == '沾色' and label not in {'变色', '沾色'}:
            tokens.append((f'沾色[{label}]', f'{label} {value}'))
        else:
            tokens.append((label, f'{label} {value}'))
    return tokens

def get_report_no(url):
    d = pdf_text.get(url, {})
    return d.get('report_no') or Path(url.split('?')[0]).stem[:40]


def ocr_text_norm(value):
    text = zh_norm(clean_cell(value))
    corrections = {
        '千摩': '干摩', '汗溃': '汗渍', '汗清': '汗渍', 'PH值': 'pH值', 'ph值': 'pH值',
        '甲醒': '甲醛', '蓬松庋': '蓬松度',
    }
    for old, new in corrections.items():
        text = text.replace(old, new)
    return text


def ocr_header_key(text):
    c = compact(ocr_text_norm(text))
    if c in {'序号', '编号'}:
        return 'sequence'
    if any(token in c for token in ('检测项目', '检验项目', '测试项目')):
        return 'item'
    if any(token in c for token in ('检测方法', '检验方法', '测试方法')):
        return 'method'
    if any(token in c for token in ('项目描述', '检测内容', '细项', '子项')):
        return 'detail'
    if c == '单位':
        return 'unit'
    if c in {'CAS号', 'CASNO.', 'CASNO', 'CAS编号'} or re.fullmatch(r'CAS(?:号|NO\.?)?', c, re.I):
        return 'cas'
    if any(token in c for token in ('报告限', '检出限', '检测限', '定量限')) or c in {'LOD', 'LOQ'}:
        return 'detection_limit'
    if any(token in c for token in ('标准要求', '标准值', '要求值', '技术要求')):
        return 'requirement'
    if any(token in c for token in ('检测结果', '检验结果', '测试结果', '实测结果', '实测值')):
        return 'result'
    if any(token in c for token in ('单项判定', '单项评价', '单项结论')) or c in {'单项', '判定', '评价', '结论'}:
        return 'verdict'
    return ''


def process_ocr_document(url):
    """用Vision OCR坐标恢复扫描表格；保留父项+细项并输出可追溯结果。"""
    document = pdf_text.get(url, {})
    if document.get('text_source') != 'vision_ocr':
        return [], Counter()
    report_no = get_report_no(url)
    local_rows = []
    stats = Counter()
    meaningful_detail = re.compile(
        r'干摩|湿摩|变色|沾色|棉|羊毛|锦纶|腈纶|聚酯|醋酯|聚酰胺|聚丙烯腈|含量|杂质|陆禽毛|异色毛绒|绒丝|羽丝'
    )
    for page_data in document.get('pages') or []:
        blocks = page_data.get('ocr_blocks') or []
        if not blocks:
            continue
        header_candidates = []
        for block in blocks:
            key = ocr_header_key(block.get('text', ''))
            if not key:
                continue
            y = float(block.get('mid_y') or 0)
            header_candidates.append((key, block, y))
        clusters = []
        for _, _, candidate_y in header_candidates:
            members = [entry for entry in header_candidates if abs(entry[2] - candidate_y) <= 0.05]
            keys = {entry[0] for entry in members}
            score = len(keys) + (5 if {'item', 'result', 'verdict'}.issubset(keys) else 0)
            clusters.append((score, candidate_y, members))
        if clusters:
            _, _, selected_headers = max(clusters, key=lambda value: (value[0], value[1]))
        else:
            selected_headers = []
        headers = defaultdict(list)
        selected_y_values = sorted(entry[2] for entry in selected_headers)
        median_y = selected_y_values[len(selected_y_values) // 2] if selected_y_values else 0
        by_key = defaultdict(list)
        for key, block, value_y in selected_headers:
            by_key[key].append((block, value_y))
        for key, values in by_key.items():
            block, _ = min(values, key=lambda value: abs(value[1] - median_y))
            headers[key].append(block)
        header_y = max((float(values[0].get('mid_y') or 0) for values in headers.values()), default=None)
        anchors = {}
        for key, values in headers.items():
            anchors[key] = sum(float(v.get('mid_x') or 0) for v in values) / len(values)
        required = {'item', 'result', 'verdict'}
        if not required.issubset(anchors):
            stats['ocr_pages_no_full_header'] += 1
            continue
        if 'sequence' not in anchors:
            anchors['sequence'] = max(0.02, anchors['item'] - 0.09)
        ordered_columns = sorted(anchors, key=lambda key: anchors[key])
        data_blocks = [
            block for block in blocks
            if float(block.get('mid_y') or 0) < (header_y or 1) - 0.01
            and not ocr_header_key(block.get('text', ''))
        ]
        sequence_blocks = [
            block for block in data_blocks
            if re.fullmatch(r'\d{1,3}', compact(block.get('text', '')))
            and abs(float(block.get('mid_x') or 0) - anchors['sequence']) < 0.065
        ]
        sequence_blocks.sort(key=lambda block: float(block.get('mid_y') or 0), reverse=True)
        if not sequence_blocks:
            stats['ocr_pages_no_sequence'] += 1
            continue
        stats['ocr_pages_positioned'] += 1
        for index, sequence in enumerate(sequence_blocks):
            center_y = float(sequence.get('mid_y') or 0)
            upper = (header_y or 1) if index == 0 else (float(sequence_blocks[index - 1].get('mid_y') or 0) + center_y) / 2
            lower = 0 if index + 1 == len(sequence_blocks) else (center_y + float(sequence_blocks[index + 1].get('mid_y') or 0)) / 2
            segment = [block for block in data_blocks if lower < float(block.get('mid_y') or 0) <= upper]
            columns = defaultdict(list)
            for block in segment:
                x = float(block.get('mid_x') or 0)
                key = min(ordered_columns, key=lambda name: abs(x - anchors[name]))
                columns[key].append(block)
            for values in columns.values():
                values.sort(key=lambda block: float(block.get('mid_y') or 0), reverse=True)

            def texts(key):
                return [ocr_text_norm(block.get('text', '')) for block in columns.get(key, []) if ocr_text_norm(block.get('text', ''))]

            item_blocks = [block for block in columns.get('item', []) if not re.fullmatch(r'\d{1,3}', compact(block.get('text', '')))]
            if not item_blocks:
                continue
            item_groups = []
            for block in item_blocks:
                value = ocr_text_norm(block.get('text', ''))
                value_y = float(block.get('mid_y') or 0)
                if item_groups and abs(item_groups[-1][0] - value_y) < 0.028:
                    previous_y, previous_text = item_groups[-1]
                    item_groups[-1] = ((previous_y + value_y) / 2, previous_text + value)
                else:
                    item_groups.append((value_y, value))
            method = ' '.join(texts('method'))
            unit = ' '.join(texts('unit'))
            verdict = ' '.join(texts('verdict'))
            cas_column = ' '.join(texts('cas'))
            detection_limit_column = ' '.join(texts('detection_limit'))
            detail_blocks = columns.get('detail', [])
            result_blocks = columns.get('result', [])
            requirement_blocks = columns.get('requirement', [])
            if not result_blocks and verdict:
                result_blocks = [None]
            material_pattern = re.compile(r'^(醋酯纤维|棉|锦纶|聚酰胺纤维|聚酯纤维|聚丙烯腈纤维|腈纶|羊毛)\s*(.*)$')
            result_entries = []
            pending_material = ''
            for result_block in result_blocks:
                if result_block is None:
                    result_entries.append(('', center_y, ''))
                    continue
                value = ocr_text_norm(result_block.get('text', ''))
                value_y = float(result_block.get('mid_y') or center_y)
                material_match = material_pattern.match(value)
                if material_match:
                    pending_material = material_match.group(1)
                    trailing = material_match.group(2).strip()
                    if trailing:
                        result_entries.append((trailing, value_y, pending_material))
                        pending_material = ''
                    continue
                if result_entries and ('检出限' in value or value.startswith(('（', '(')) or value.endswith(('）', ')'))):
                    old_value, old_y, old_material = result_entries[-1]
                    result_entries[-1] = (old_value + value, old_y, old_material)
                    continue
                result_entries.append((value, value_y, pending_material))
                pending_material = ''
            for result, result_y, material in result_entries:
                parent_text = min(item_groups, key=lambda value: abs(value[0] - result_y))[1]
                parent_item = normalize_item_part(parent_text)
                if not parent_item:
                    continue
                detail_block = min(detail_blocks, key=lambda block: abs(float(block.get('mid_y') or 0) - result_y)) if detail_blocks else None
                requirement_block = min(requirement_blocks, key=lambda block: abs(float(block.get('mid_y') or 0) - result_y)) if requirement_blocks else None
                detail = ocr_text_norm(detail_block.get('text', '')) if detail_block else ''
                requirement = ocr_text_norm(requirement_block.get('text', '')) if requirement_block else ''
                item = parent_item
                item_detail = ''
                if detail and detail not in {'/', '-', '—'} and meaningful_detail.search(detail):
                    if material and '沾色' in detail:
                        detail = f'{detail}[{material}]'
                    item, item_detail = compose_item(parent_item, [detail])
                item = normalize_item_name(item)
                if not valid_item(item):
                    stats['ocr_invalid_item_cell'] += 1
                    continue
                for rec in records_by_url.get(url) or [{}]:
                    for color in rec.get('selected_colors') or ['未标明颜色']:
                        local_rows.append({
                            'item': item, 'raw_item': item, 'item_detail': item_detail,
                            'raw_row': page_data.get('original_text') or page_data.get('text', ''),
                            'result': result, 'result_detail': result if item_detail else '',
                            'verdict': infer_verdict(result, verdict), 'verdict_raw': verdict,
                            'unit': unit, 'requirement': requirement, 'method': method,
                            'cas_number': extract_cas_numbers(cas_column, parent_text, detail, result),
                            'detection_limit': detection_limit_column or extract_detection_limit(requirement, result),
                            'report_no': report_no, 'url': url,
                            'source_order_no': rec.get('order_no', ''), 'sku': rec.get('sku', ''),
                            'source_sheet': rec.get('source_sheet', ''),
                            'source_row': rec.get('source_row', ''),
                            'source_cell': rec.get('source_cell', ''),
                            'subcategory': rec.get('subcategory', ''),
                            'sample_type': rec.get('sample_type', ''),
                            'color': color, 'page': page_data.get('page', ''), 'table': 'OCR',
                        })
                stats['ocr_candidate_rows'] += 1
    deduped = []
    seen = set()
    for row in local_rows:
        key = (
            row['item'], row['result'], row['report_no'], row['source_order_no'],
            row.get('source_sheet', ''), row.get('source_row', ''), row.get('source_cell', ''),
            row.get('sample_type', ''), row['sku'], row['color'], row['page'],
        )
        if key not in seen:
            seen.add(key)
            deduped.append(row)
    stats['ocr_rows'] += len(deduped)
    return deduped, stats

def process_pdf(m):
    url = m['url']
    local_rows = []
    stats = Counter()
    if m.get('status') == 'failed' or not m.get('path'):
        stats['download_failed'] += 1
        return local_rows, stats
    path = Path(m['path'])
    if not path.exists():
        stats['missing_file'] += 1
        return local_rows, stats
    report_no = get_report_no(url)
    try:
        with pdfplumber.open(path) as pdf:
            for page_no, page in enumerate(pdf.pages, 1):
                if not page_has_keywords(url, page_no):
                    stats['pages_skipped_no_keyword'] += 1
                    continue
                tables = page.extract_tables() or []
                if not tables:
                    # 无完整网格线的机构报告使用文本位置重建表格。
                    text_settings = {
                        'vertical_strategy': 'text',
                        'horizontal_strategy': 'text',
                        'min_words_vertical': 2,
                        'min_words_horizontal': 1,
                        'snap_tolerance': 5,
                        'join_tolerance': 5,
                        'intersection_tolerance': 5,
                    }
                    tables = page.extract_tables(text_settings) or []
                    if tables:
                        stats['pages_text_table_fallback'] += 1
                    else:
                        stats['pages_no_table'] += 1
                stats['pages_table_checked'] += 1
                stats['tables_extracted'] += len(tables)
                for table_no, table in enumerate(tables, 1):
                    current = None
                    current_parent_item = ''
                    current_verdict = ''
                    current_method = ''
                    parent_has_merged_verdict = False
                    inferred_schema = infer_headerless_schema(table)
                    if inferred_schema:
                        stats['tables_headerless_inferred'] += 1
                    for raw_row in table:
                        original_row = [clean_cell(c) for c in raw_row]
                        row = [to_simplified(cell) for cell in original_row]
                        original_row_text = ' | '.join(cell for cell in original_row if cell)
                        if is_header_row(row):
                            stats['header_rows_recognized'] += 1
                            current = header_indices(row)
                            if current['item'] is None:
                                current = None
                            current_parent_item = ''
                            current_verdict = ''
                            current_method = ''
                            parent_has_merged_verdict = False
                            continue
                        if not current and inferred_schema:
                            current = inferred_schema
                        if not current:
                            continue
                        item_cols = current.get('item_cols') or [current['item']]
                        if min(item_cols) >= len(row):
                            continue
                        result = row[current['result']] if current.get('result') is not None and current['result'] < len(row) else ''
                        verdict = row[current['verdict']] if current.get('verdict') is not None and current['verdict'] < len(row) else ''
                        method = row[current['method']] if current.get('method') is not None and current['method'] < len(row) else ''
                        requirement = row[current['requirement']] if current.get('requirement') is not None and current['requirement'] < len(row) else ''
                        unit = row[current['unit']] if current.get('unit') is not None and current['unit'] < len(row) else ''
                        cas_column = row[current['cas']] if current.get('cas') is not None and current['cas'] < len(row) else ''
                        detection_limit_column = row[current['detection_limit']] if current.get('detection_limit') is not None and current['detection_limit'] < len(row) else ''
                        inferred_requirement = ''
                        inferred_method = ''
                        if current.get('inferred') or current.get('wrapped_item_cols'):
                            _, inferred_requirement, inferred_method = inferred_item_fields(row, current)
                        requirement = requirement_with_left_fragment(row, current.get('requirement'), requirement)
                        requirement = join_field_text(inferred_requirement, requirement)
                        method = join_field_text(inferred_method, method)
                        verdict = complete_split_verdict(row, current.get('verdict'), verdict)
                        requirement, result = sanitize_requirement_result(requirement, result)
                        candidate_parent_item = current_parent_item
                        if current.get('detail_cols'):
                            raw_item, item_detail, candidate_parent_item = item_parts_from_schema(row, current, current_parent_item)
                        elif current.get('inferred') or current.get('wrapped_item_cols'):
                            force_child = bool(current_parent_item and result and not verdict and parent_has_merged_verdict)
                            raw_item, item_detail, candidate_parent_item = item_parts_from_inferred_row(
                                row, current, current_parent_item, force_child=force_child,
                            )
                        else:
                            force_child = bool(current_parent_item and result and not verdict and parent_has_merged_verdict)
                            raw_item, item_detail, candidate_parent_item = item_parts_from_row(row, item_cols, current_parent_item, force_child=force_child)
                        original_raw_item = raw_item
                        raw_item, item_requirement, item_method = split_mixed_item_text(raw_item)
                        requirement = join_field_text(item_requirement, requirement)
                        method = join_field_text(item_method, method)
                        inferred_requirement = join_field_text(inferred_requirement, item_requirement)
                        inferred_method = join_field_text(inferred_method, item_method)
                        if raw_item and original_raw_item != raw_item and candidate_parent_item == original_raw_item:
                            candidate_parent_item = raw_item
                        direction_requirement = item_requirement or inferred_requirement
                        if raw_item in {'直向', '横向', '纵向', '经向', '纬向'} and direction_requirement and current_parent_item:
                            raw_item, item_detail = compose_item(current_parent_item, [raw_item])
                            candidate_parent_item = current_parent_item
                        has_item_text = bool(raw_item)
                        has_context_fragment = bool(inferred_requirement or inferred_method or requirement or method)
                        if not result and not verdict:
                            # 纯续行也必须进入回填；没有项目、要求、方法和继承上下文时才跳过。
                            if (not has_item_text and not has_context_fragment) or not (current_verdict or requirement or method or current_method):
                                continue
                        if not raw_item and current_parent_item and (requirement or method):
                            attached = False
                            context_only = False
                            if requirement:
                                attached = attach_context_fragment(
                                    local_rows, current_parent_item, 'requirement', requirement, page_no, table_no,
                                ) or attached
                            if method:
                                attached = attach_context_fragment(
                                    local_rows, current_parent_item, 'method', method, page_no, table_no,
                                ) or attached
                                current_method = join_field_text(current_method, method)
                            if attached:
                                continue
                            raw_item = current_parent_item
                            item_detail = ''
                            candidate_parent_item = current_parent_item
                            context_only = True
                        else:
                            context_only = False
                        parent_changed = bool(raw_item and candidate_parent_item != current_parent_item)
                        effective_verdict = verdict if parent_changed else (verdict or current_verdict)
                        effective_method = method if parent_changed else (method or current_method)
                        expanded_items = [(raw_item, item_detail, result)]
                        subitems = result_subitems(result)
                        if subitems and raw_item:
                            expanded_items = []
                            for subitem, sub_result in subitems:
                                if subitem == raw_item:
                                    expanded_items.append((raw_item, item_detail, sub_result))
                                else:
                                    expanded_items.append((f'{raw_item}-{subitem}', subitem, sub_result))
                        row_has_valid_item = False
                        for raw_part, raw_detail, raw_result in expanded_items:
                            stats['candidate_item_cells'] += 1
                            item = normalize_item_name(raw_part)
                            if not valid_item(item):
                                stats['invalid_item_cell'] += 1
                                continue
                            row_has_valid_item = True
                            linked_records = records_by_url.get(url) or [{}]
                            for rec in linked_records:
                                colors = rec.get('selected_colors') or ['未标明颜色']
                                for color in colors:
                                    local_rows.append({
                                        'item': item,
                                        'raw_item': raw_part,
                                        'raw_row': original_row_text,
                                        'item_detail': raw_detail,
                                        'result': raw_result,
                                        'result_detail': raw_result if raw_detail else '',
                                        'verdict': infer_verdict(result, effective_verdict),
                                        'verdict_raw': effective_verdict,
                                        'unit': unit,
                                        'cas_number': extract_cas_numbers(cas_column, raw_part, raw_detail),
                                        'detection_limit': detection_limit_column or extract_detection_limit(requirement, raw_result, raw_part),
                                        'requirement': requirement,
                                        'method': effective_method,
                                        'report_no': report_no,
                                        'url': url,
                                        'source_order_no': rec.get('order_no', ''),
                                        'source_sheet': rec.get('source_sheet', ''),
                                        'source_row': rec.get('source_row', ''),
                                        'source_cell': rec.get('source_cell', ''),
                                        'subcategory': rec.get('subcategory', ''),
                                        'sample_type': rec.get('sample_type', ''),
                                        'sku': rec.get('sku', ''),
                                        'color': color,
                                        'page': page_no,
                                        'table': table_no,
                                        'context_only': context_only,
                                    })
                        parent_changed = bool(row_has_valid_item and parent_changed)
                        if row_has_valid_item:
                            current_parent_item = candidate_parent_item
                        if row_has_valid_item or not has_item_text:
                            if parent_changed:
                                current_method = method
                                current_verdict = verdict
                            else:
                                if method:
                                    current_method = method
                                if verdict:
                                    current_verdict = verdict
                        if row_has_valid_item and verdict:
                                parent_has_merged_verdict = bool(not result)
        if pdf_text.get(url, {}).get('text_source') == 'vision_ocr':
            ocr_rows, ocr_stats = process_ocr_document(url)
            local_rows.extend(ocr_rows)
            stats.update(ocr_stats)
            if ocr_rows:
                stats['ocr_documents_parsed'] += 1
        deduped = []
        seen = set()
        for row in local_rows:
            key = (
                        row.get('item'), row.get('result'), row.get('verdict'), row.get('report_no'),
                row.get('source_order_no'), row.get('source_sheet'), row.get('source_row'), row.get('source_cell'),
                row.get('sample_type'), row.get('sku'), row.get('color'), row.get('page'),
            )
            if key in seen:
                continue
            seen.add(key)
            deduped.append(row)
        if len(deduped) != len(local_rows):
            stats['duplicate_rows_removed'] += len(local_rows) - len(deduped)
        local_rows = deduped
        stats['rows'] += len(local_rows)
    except Exception as exc:
        stats[f'pdf_error:{type(exc).__name__}'] += 1
    return local_rows, stats

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--limit', type=int, default=300)
    parser.add_argument('--reset', action='store_true')
    parser.add_argument('--retry-failed', action='store_true')
    args = parser.parse_args()
    if args.reset:
        ROWS_PATH.write_text('', encoding='utf-8')
        if RESULTS_DIR.exists():
            for path in RESULTS_DIR.glob('*.json'):
                path.unlink()
        save_state({'processed_urls': [], 'url_states': {}, 'stats': {}})
        print('reset checkpoint')
        return
    state = load_state()
    url_states = state.setdefault('url_states', {})
    recover_result_states(url_states)
    stats_total = aggregate_result_stats()
    pending = [
        m for m in manifest
        if should_attempt(
            url_states.get(m['url']),
            parser_version=TABLE_PARSER_VERSION,
            input_identity=input_identity(m),
            retry_failed=args.retry_failed,
        )
    ]
    batch = pending[:args.limit]
    start = time.time()
    batch_stats = Counter()
    batch_rows = 0
    for idx, m in enumerate(batch, 1):
        rows, stats = process_pdf(m)
        status = attempt_status(stats, len(rows))
        identity = input_identity(m)
        saved_result = save_pdf_result(m['url'], rows, status, stats, identity)
        saved_status = saved_result.get('status', status)
        batch_rows += len(saved_result.get('rows') or [])
        batch_stats.update(stats)
        stats_total = aggregate_result_stats()
        previous_entry = url_states.get(m['url']) or {}
        url_states[m['url']] = {
            'status': saved_status,
            'attempts': int(previous_entry.get('attempts') or 0) + 1,
            'parser_version': TABLE_PARSER_VERSION,
            **{field: saved_result.get(field, '') for field in IDENTITY_FIELDS},
            'last_error': '; '.join(key for key in stats if str(key).startswith('pdf_error:')),
        }
        processed = sorted(
            url for url, entry in url_states.items()
            if entry.get('status') in {'succeeded', 'permanent_failed'}
        )
        state = {'processed_urls': processed, 'url_states': url_states, 'stats': dict(stats_total)}
        save_state(state)
        rebuild_rows_file()
        if idx % 25 == 0 or idx == len(batch):
            print(f"progress batch={idx}/{len(batch)} terminal={len(processed)}/{len(manifest)} rows_batch={batch_rows} rows_total={stats_total.get('rows',0)} elapsed={time.time()-start:.1f}s", flush=True)
    processed = sorted(
        url for url, entry in url_states.items()
        if entry.get('status') in {'succeeded', 'permanent_failed'}
    )
    state = {'processed_urls': processed, 'url_states': url_states, 'stats': dict(stats_total)}
    save_state(state)
    rebuild_rows_file()
    retryable = sum(
        1 for entry in url_states.values()
        if entry.get('status') in {'partial', 'retryable_failed'}
    )
    print(json.dumps({
        'processed_total': len(processed),
        'manifest_total': len(manifest),
        'remaining_unattempted': sum(1 for m in manifest if m['url'] not in url_states),
        'retryable_or_partial': retryable,
        'batch_processed': len(batch),
        'batch_rows': batch_rows,
        'stats_total': dict(stats_total),
        'stats_batch': dict(batch_stats),
        'rows_path': str(ROWS_PATH),
        'state_path': str(STATE_PATH),
    }, ensure_ascii=False, indent=2))

if __name__ == '__main__':
    main()
