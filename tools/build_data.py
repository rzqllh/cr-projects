"""Build a read-only snapshot for the dashboard. No writes to the source workbook."""
import json, hashlib, re, sys
from pathlib import Path
from datetime import datetime, timedelta
from artifact_tool import Blob, SpreadsheetFile

ROOT=Path(__file__).resolve().parents[1]
SRC=Path(sys.argv[1]) if len(sys.argv)>1 else Path('/mnt/data/Progress CR TSEL-2.xlsx')
SCHEMAS={
 'CR Ultimate': {'header':2,'maxrow':96,'title':2,'id':1,'no':0,'status':3,'progress':4,'action':7,'note':6,'date':20,'year':None,'type':8,'segment':1,'cost':21,'month':23,'yearRekon':24,'recon':None,'due':9,'requestor':18},
 'CR Interim': {'header':2,'maxrow':463,'title':2,'id':1,'no':0,'status':3,'progress':6,'action':5,'note':4,'date':15,'year':None,'type':25,'segment':None,'cost':None,'month':None,'yearRekon':None,'recon':None,'due':7,'requestor':13},
 'Cost CR': {'header':1,'maxrow':338,'title':4,'id':1,'no':0,'status':5,'progress':None,'action':17,'note':None,'date':None,'year':2,'type':10,'segment':None,'cost':16,'month':22,'yearRekon':None,'recon':21,'due':None,'requestor':8},
 'UPCOMING CR': {'header':2,'maxrow':173,'title':2,'id':None,'no':0,'status':None,'progress':None,'action':12,'note':13,'date':1,'year':None,'type':3,'segment':None,'cost':None,'month':None,'yearRekon':None,'recon':None,'due':9,'requestor':5},
}
def clean(v):
 if v is None: return None
 if isinstance(v,float) and v.is_integer(): return int(v)
 if isinstance(v,(float,int,bool)):return v
 s=str(v).strip()
 return s if s else None

def yy(v):
 if v is None:return None
 s=str(v).strip()
 if isinstance(v,(int,float)):
  if 2020<=v<=2100: return int(v)
  if 40000<=v<90000:return (datetime(1899,12,30)+timedelta(days=v)).year
 m=re.search(r'(20[2-9][0-9])',s)
 return int(m.group(1)) if m else None

def normalized(s):return re.sub(r'\s+',' ',str(s or '').strip().lower())

def stable_id(sheet, rid, title, date, counter):
 # A title is used as a matching hint, never as permission to write without a fresh uniqueness check.
 h=hashlib.sha256(f'{sheet}|{rid or ""}|{normalized(title)}|{date or ""}'.encode()).hexdigest()[:16]
 return f'{sheet}:{h}:{counter}'

w=SpreadsheetFile.import_xlsx(Blob.load(str(SRC)))
out={'schemaVersion':3,'source':'Progress CR TSEL-2.xlsx','snapshotDate':'2026-10-28','sourceSha256':hashlib.sha256(SRC.read_bytes()).hexdigest(), 'mode':'snapshot','workstreams':[],'records':[],'schemas':SCHEMAS}
for sheet, config in SCHEMAS.items():
 s=w.worksheets.get_item(sheet)
 rows=s.get_range_by_indexes(0,0,config['maxrow'],26).values
 counts={}
 kept=0
 for i,raw in enumerate(rows):
  rowNum=i+1
  if rowNum<=config['header']:continue
  a=[clean(v) for v in raw]
  title=a[config['title']]
  if not title or normalized(title) in ('total','subtotal','description','note') or str(title).startswith('Note :'):
   continue
  # Avoid worksheet footer, empty decoration rows and notes.
  if sheet=='CR Ultimate' and (rowNum>80 or not (isinstance(a[0],(int,float)) or str(a[0] or '').isdigit())): continue
  no=a[config['no']] if config['no'] is not None else None
  if sheet!='CR Ultimate' and not (isinstance(no,(int,float)) or str(no or '').isdigit()):continue
  cid=a[config['id']] if config['id'] is not None else None
  rd=a[config['date']] if config['date'] is not None else None
  year=yy(rd) or (yy(a[config['year']]) if config['year'] is not None else None)
  if sheet=='CR Ultimate' and not year:
   # NDE Permintaan TSEL (col R) may contain explicit year.
   year=yy(a[17])
  month=a[config['month']] if config['month'] is not None else None
  if month in ('-',0,'0','N/A','n/a',''):month=None
  ry=yy(a[config['yearRekon']]) if config['yearRekon'] is not None else None
  state=a[config['recon']] if config['recon'] is not None else None
  title_norm=normalized(title)
  bkey=(str(cid or ''),title_norm,str(rd or ''))
  counts[bkey]=counts.get(bkey,0)+1
  key=stable_id(sheet,cid,title,rd,counts[bkey])
  out['records'].append({'key':key,'sheet':sheet,'rowHint':rowNum,'number':no,'crId':cid,
    'title':title,'status':a[config['status']] if config['status'] is not None else None,
    'progress':a[config['progress']] if config['progress'] is not None else None,
    'action':a[config['action']] if config['action'] is not None else None,
    'note':a[config['note']] if config['note'] is not None else None,
    'yearCR':year,'yearRekon':ry,'reconMonth':month,'reconRaw':state,
    'requestDate':rd,'dueDate':a[config['due']] if config['due'] is not None else None,
    'requestor':a[config['requestor']] if config['requestor'] is not None else None,
    'type':a[config['type']] if config['type'] is not None else None,
    'segment':a[config['segment']] if config['segment'] is not None else None,
    'cost':a[config['cost']] if config['cost'] is not None else None, 'cells':a})
  kept+=1
 out['workstreams'].append({'sheet':sheet,'count':kept,'headerRow':config['header']})
print('COUNTS',out['workstreams'])
print('SAMPLE',[(r['sheet'],r['key'],r['title'][:35]) for r in out['records'][:3]])
(ROOT/'dist'/'data.js').write_text('window.CR_WORKBOOK_SNAPSHOT = '+json.dumps(out,ensure_ascii=False,separators=(',',':')).replace('<',r'\u003c').replace('\u2028',r'\u2028').replace('\u2029',r'\u2029')+';\n',encoding='utf-8')
(ROOT/'snapshot_summary.json').write_text(json.dumps({'sourceSha256':out['sourceSha256'],'snapshotDate':out['snapshotDate'],'workstreams':out['workstreams'],'totalRecords':len(out['records'])},ensure_ascii=False,indent=2))
