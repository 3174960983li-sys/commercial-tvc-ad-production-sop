#!/usr/bin/env python3
"""T03-08 V1.1 static schema recheck. Uses Python standard library only.
Usage: python T03-08_静态接口验证器.py [V1.1 release directory]
Checks structural metadata only, not business permissions or real-world approvals.
"""
import json,re,sys,zipfile,posixpath
from pathlib import Path
from xml.etree import ElementTree as E
NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
GATE=['gate_id','scope_ids','reviewed_artifact_ids','reviewed_versions','reviewed_hashes','decision','findings','required_actions','owner','approver','approval_evidence','reviewed_at','validity_conditions']
def col_num(x):
 n=0
 for c in x: n=n*26+ord(c)-64
 return n
def xlsx_headers(path):
 with zipfile.ZipFile(path) as z:
  if z.testzip() is not None: raise RuntimeError('Bad XLSX ZIP: '+str(path))
  x=E.fromstring(z.read('xl/workbook.xml'))
  rel=E.fromstring(z.read('xl/_rels/workbook.xml.rels'))
  rr={r.attrib['Id']:r.attrib['Target'] for r in rel}
  ss=[]
  if 'xl/sharedStrings.xml' in z.namelist():
   st=E.fromstring(z.read('xl/sharedStrings.xml'))
   ss=[''.join(t.text or '' for t in q.iter('{%s}t'%NS)) for q in st]
  out={}
  for sh in x.find('{%s}sheets'%NS):
   name=sh.attrib['name']; target=rr[sh.attrib['{%s}id'%R]]
   target=target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/'+target)
   root=E.fromstring(z.read(target)); fields=[]
   for row in root.findall('.//{%s}sheetData/{%s}row'%(NS,NS)):
    if row.attrib.get('r')!='5': continue
    for c in row:
     m=re.fullmatch('([A-Z]+)5',c.attrib.get('r',''))
     if not m: continue
     v=c.find('{%s}v'%NS); inline=c.find('{%s}is'%NS)
     if v is not None and v.text is not None:
      val=ss[int(v.text)] if c.attrib.get('t')=='s' else v.text
     elif inline is not None: val=''.join(t.text or '' for t in inline.iter('{%s}t'%NS))
     else: continue
     fields.append((col_num(m.group(1)),val))
   out[name]=[v for _,v in sorted(fields)]
  table_errors=[]
  for file in z.namelist():
   if not re.fullmatch(r'xl/tables/table\d+\.xml',file):continue
   tr=E.fromstring(z.read(file)); cc=[i.attrib['name'] for i in tr.findall('.//{%s}tableColumn'%NS)]
   if not any(any(headers[i:i+len(cc)]==cc for i in range(max(0,len(headers)-len(cc)+1))) for headers in out.values()):
    table_errors.append(tr.attrib.get('name',file))
  return out,table_errors

def main(path):
 errors=[]; allsteps=[]; allgates=[]
 for i in range(1,7):
  tag=f'T03-{i:02d}';folder=path/tag
  js=next(folder.glob('*字段字典与机器接口.json'))
  schema=json.loads(js.read_text(encoding='utf-8'))
  xlsx=next(folder.glob('*V1.1.xlsx'))
  tabs,table_errors=xlsx_headers(xlsx)
  if table_errors:errors.append(f'{tag} table schema inconsistent: {table_errors}')
  if schema['interface_schema_version']!='1.1.0':errors.append(f'{tag} schema version')
  if set(tabs)!=set(schema['workbook_sheets']):errors.append(f'{tag} workbook tab mismatch')
  for name,data in schema['workbook_sheets'].items():
   if name not in tabs or tabs[name]!=data['columns']:errors.append(f'{tag}:{name} header mismatch')
  if not {'gate_record_id',*GATE}.issubset(tabs.get('V11_Gate记录',[])):
   errors.append(f'{tag} Gate minimum fields absent')
  if schema['canonical_contract']['gate_decisions']!=['PASS','CONDITIONAL','REWORK','HOLD']:
   errors.append(f'{tag} Gate enum is not S07')
  allsteps+=schema['scope_steps'];allgates += [g['gate_id'] for g in schema['gates']]
  for name,fields in tabs.items():
   if 'variant_id' in fields:errors.append(f'{tag}:{name} ambiguous variant_id')
 if len(allsteps)!=99 or len(set(allsteps))!=99:errors.append('99 formal steps not unique and complete')
 if set(allgates)!=set(['G0','G1','G2','G3','G4','G5','G6','G7a','G7b','G8','G9']):errors.append('Gate scope mismatch')
 if errors:
  print('FAIL:',len(errors));print('\n'.join(' - '+e for e in errors));return 1
 print('PASS_STATIC: 6 workbooks / 99 step IDs / G0-G9 + G7a/G7b / normalized Gate fields, enums, table headers and ID keys.')
 print('This does NOT validate real client permissions, media files, legal compliance, model output or campaign results.')
 return 0
if __name__=='__main__':sys.exit(main(Path(sys.argv[1]) if len(sys.argv)>1 else Path(__file__).resolve().parent))
