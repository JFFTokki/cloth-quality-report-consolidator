import json
from pathlib import Path
import sys
sys.path.insert(0, str(Path('tmp/vendor').resolve()))
import pdfplumber

manifest=json.loads(Path('pipeline/qc_all/download_manifest.json').read_text(encoding='utf-8'))
texts=json.loads(Path('pipeline/qc_all/pdf_text.json').read_text(encoding='utf-8'))
# pick PDFs whose text includes common table headers
picked=[]
for d in texts:
    full='\n'.join(p.get('text','') for p in d.get('pages',[])[:3])
    if '检测项目' in full and '检测结果' in full and '单项' in full and d.get('status')!='failed':
        picked.append(d)
    if len(picked)>=5: break
print('picked', len(picked))
for d in picked:
    print('\nURL', d['url'][:120], 'report', d.get('report_no'), 'path', d.get('path'))
    try:
        with pdfplumber.open(d['path']) as pdf:
            for pi,page in enumerate(pdf.pages[:2],1):
                tables=page.extract_tables() or []
                print(' page',pi,'tables',len(tables))
                for ti,t in enumerate(tables[:2],1):
                    print('  table',ti,'rows',len(t),'cols',max(len(r) for r in t if r))
                    for r in t[:8]: print('   ', r)
    except Exception as e:
        print('ERR',type(e).__name__,e)
