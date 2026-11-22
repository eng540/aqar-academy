#!/usr/bin/env python3
"""Build the local reading/search catalog and optional reproducible course bundles."""
from pathlib import Path
import argparse, hashlib, json, re, zipfile
ROOT=Path(__file__).resolve().parents[1]
parser=argparse.ArgumentParser()
parser.add_argument('--bundles',action='store_true')
args=parser.parse_args()
def eligible(p):
 rel=p.relative_to(ROOT)
 return p.is_file() and not any(x in {'.git','__pycache__','node_modules'} for x in rel.parts) and p.suffix not in {'.pyc','.zip'} and str(rel) not in {'assets/catalog.json','bundle-manifest.json'}
files=sorted((p for p in ROOT.rglob('*') if eligible(p)),key=lambda p:str(p.relative_to(ROOT)))
documents=[]
for p in files:
 rel=p.relative_to(ROOT).as_posix()
 if p.suffix!='.md' or rel.startswith('assets/vendor/') or 'تقارير-التطوير/' in rel:continue
 text=p.read_text(encoding='utf-8')
 title=next((s.lstrip('# ').strip() for s in text.splitlines() if s.startswith('# ')),p.stem)
 documents.append({'path':rel,'title':title,'text':text})
catalog={'version':1,'files':sorted([p.relative_to(ROOT).as_posix() for p in files]+['الأكاديمية-العقارية-الكاملة.zip','أكاديمية-العقار-كاملة.zip','bundle-manifest.json']),'documents':documents}
(ROOT/'assets').mkdir(exist_ok=True)
(ROOT/'assets/catalog.json').write_text(json.dumps(catalog,ensure_ascii=False,separators=(',',':')),encoding='utf-8')
print('catalog:',len(documents),'searchable documents;',len(catalog['files']),'file paths')
if args.bundles:
 # Historical development reports and the independent media lab are not the course.
 selected=[p for p in files if not p.relative_to(ROOT).as_posix().startswith(('media-platform/','الحوكمة/تقارير-التطوير/','tests/','scripts/'))]
 selected.append(ROOT/'assets/catalog.json')
 included={p.relative_to(ROOT).as_posix() for p in selected}|{'bundle-manifest.json'}
 bundled_catalog={'version':1,'files':sorted(included),'documents':[d for d in documents if d['path'] in included]}
 bundled_bytes=json.dumps(bundled_catalog,ensure_ascii=False,separators=(',',':')).encode()
 manifest={'version':'2026-11-22','description':'Course reference and practice; media lab and historical reports remain separate.','files':{p.relative_to(ROOT).as_posix():hashlib.sha256(bundled_bytes if p.relative_to(ROOT).as_posix()=='assets/catalog.json' else p.read_bytes()).hexdigest() for p in selected}}
 (ROOT/'bundle-manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
 selected.append(ROOT/'bundle-manifest.json')
 for name in ['الأكاديمية-العقارية-الكاملة.zip','أكاديمية-العقار-كاملة.zip']:
  with zipfile.ZipFile(ROOT/name,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
   for p in sorted(selected,key=lambda p:str(p.relative_to(ROOT))):
    rel=p.relative_to(ROOT).as_posix();info=zipfile.ZipInfo(rel,date_time=(2026,11,22,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;info.external_attr=0o644<<16
    data=json.dumps(bundled_catalog,ensure_ascii=False,separators=(',',':')).encode() if rel=='assets/catalog.json' else p.read_bytes()
    z.writestr(info,data)
  print(name,len(selected),'files', (ROOT/name).stat().st_size,'bytes')
