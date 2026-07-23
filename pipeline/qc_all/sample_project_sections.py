import json,re
from pathlib import Path

docs=json.loads(Path('pipeline/qc_all/pdf_text.json').read_text(encoding='utf-8'))
needles=['检验项目','检测项目','检验检测项目','测试项目','检测结果']
shown=0
for doc in docs:
    text='\n'.join(p.get('text','') for p in doc.get('pages',[]))
    if any(n in text for n in needles):
        lines=[re.sub(r'\s+',' ',l).strip() for l in text.splitlines() if l.strip()]
        for i,l in enumerate(lines):
            if any(n in l for n in needles):
                print('\n---', doc.get('report_no'), 'line', i, '---')
                for j in range(max(0,i-3), min(len(lines), i+18)):
                    print(j, lines[j])
                shown+=1
                break
    if shown>=8:
        break
