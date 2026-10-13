import json, re, sys, zipfile
from pathlib import Path
from html import unescape

cfg = json.loads(Path(sys.argv[1] if len(sys.argv)>1 else 'deck.config.json').read_text(encoding='utf-8'))
pptx = Path(cfg['output'])
q = cfg['quality']
errors=[]
if not pptx.exists(): errors.append(f'missing output: {pptx}')
else:
    try:
        with zipfile.ZipFile(pptx) as z:
            names=set(z.namelist())
            if '[Content_Types].xml' not in names or 'ppt/presentation.xml' not in names:
                errors.append('invalid PPTX package')
            slides=sorted(n for n in names if re.fullmatch(r'ppt/slides/slide[^/]+\.xml',n))
            texts=[]
            for n in slides:
                texts += [unescape(x) for x in re.findall(r'<a:t>(.*?)</a:t>', z.read(n).decode('utf-8'))]
            joined=''.join(texts)
            if len(slides) < q['min_slides']: errors.append(f'slides {len(slides)} < {q["min_slides"]}')
            if len(joined) < q['min_text_chars']: errors.append(f'text chars {len(joined)} < {q["min_text_chars"]}')
            for term in q['required_terms']:
                if term not in joined: errors.append(f'missing term: {term}')
            media=len([n for n in names if n.startswith('ppt/media/')])
            print(json.dumps({'pptx':str(pptx),'slides':len(slides),'text_chars':len(joined),'media_files':media,'required_terms':q['required_terms'],'status':'PASS' if not errors else 'FAIL'}, ensure_ascii=False))
    except zipfile.BadZipFile:
        errors.append('bad zip / corrupt pptx')
for path_key in ['sources','notes']:
    p=Path(cfg[path_key])
    if not p.exists() or p.stat().st_size==0:
        errors.append(f'missing {path_key}: {p}')
    else:
        text=p.read_text(encoding='utf-8', errors='replace')
        if path_key == 'sources' and len(re.findall(r'https?://', text)) < 1:
            errors.append('sources has no URLs')
        if path_key == 'notes' and len(re.findall(r'^##?\s+', text, re.M)) < 2:
            errors.append('notes has too few headings')
if errors:
    print(json.dumps({'status':'FAIL','errors':errors}, ensure_ascii=False))
    sys.exit(1)
