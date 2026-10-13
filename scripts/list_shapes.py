"""Enumerate top-level shapes of each target template slide.

Prints shape index -> current text, so the content file can address shapes
directly (needed on pages whose sibling shapes share the same placeholder
text, where substring replacement cannot tell them apart).
"""
import json
import re
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
content = json.loads((ROOT / 'input' / (sys.argv[1] if len(sys.argv) > 1 else 'xauat-content-final.json')).read_text(encoding='utf-8'))

with zipfile.ZipFile(ROOT / 'templates' / 'xauat.pptx') as z:
    for entry in content['slides']:
        xml = z.read('ppt/slides/slide%d.xml' % entry['src']).decode('utf-8')
        print('\n== template slide %d (%s)' % (entry['src'], entry.get('role')))
        n = 0
        for m in re.finditer(r'<p:(sp|graphicFrame|grpSp)>.*?</p:\1>', xml, re.S):
            n += 1
            sp = m.group(0)
            ts = re.findall(r'<a:t>(.*?)</a:t>', sp)
            if ts:
                print('  %2d %s' % (n, ts[:3]))
