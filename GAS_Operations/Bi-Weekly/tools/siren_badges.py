"""Add a linked 'SIREN 등록됨' badge to claim/review cards whose SKU + 인입사유 match a
registered row of the 26년 SIREN sheet. Link = the SIREN review deck found in Drive by the
row's C-column title (fallback: the sheet row)."""
import json, re, sys
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
PID='1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc'; SID='15Jh6ZFDBIbpv4OANVtD3g4wFBJxoof9SHWDUEU3GiXI'; GID=1840076165; EMU=12700
p='/Users/kevinkim/.config/gws_shim/token.json'
d=json.load(open(p)); c=Credentials.from_authorized_user_info(d); c.refresh(Request()); d['token']=c.token; json.dump(d,open(p,'w'))
svc=build('slides','v1',credentials=c).presentations(); ss=build('sheets','v4',credentials=c).spreadsheets(); drv=build('drive','v3',credentials=c).files()
STOP={'이슈','불량','문제','발생','현상'}
def norm(s): return re.sub(r'[\s\(\)\[\]\-_/,.:·]+','',str(s).lower()).replace('이슈','')
def tokens(s): return [t for t in re.split(r'[\s\(\)\[\]\-_/,.:·]+',str(s).lower()) if len(t)>=2 and t not in STOP]
def issue_match(a,b):
    na,nb=norm(a),norm(b)
    if not na or not nb: return False
    if na==nb or (len(na)>=2 and na in nb) or (len(nb)>=2 and nb in na): return True
    for t in tokens(a):
        if t in nb: return True
    for t in tokens(b):
        if t in na: return True
    bg=lambda s:{s[i:i+2] for i in range(len(s)-1)}
    A,B=bg(na),bg(nb); return bool(A and B) and 2*len(A&B)/(len(A)+len(B))>=0.5
rows=ss.values().get(spreadsheetId=SID,range="'26년 SIREN'!A19:K200").execute().get('values',[])
entries=[]
for i,r in enumerate(rows):
    g=lambda k: (r[k].strip() if k<len(r) else '')
    if not g(4): continue
    title=g(2); link=''
    for q in ["name = '%s' and trashed = false"%title.replace("'","\\'"), "name contains '%s' and trashed = false"%title[:30].replace("'","\\'")]:
        f=drv.list(q=q,fields='files(id,name,mimeType)',supportsAllDrives=True,includeItemsFromAllDrives=True,pageSize=1).execute().get('files',[])
        if f: link='https://docs.google.com/presentation/d/%s/edit'%f[0]['id'] if 'presentation' in f[0]['mimeType'] else 'https://drive.google.com/file/d/%s/view'%f[0]['id']; break
    if not link: link='https://docs.google.com/spreadsheets/d/%s/edit#gid=%d&range=C%d'%(SID,GID,19+i)
    entries.append({'row':19+i,'skus':[s.strip().upper() for s in re.split(r'[\s,/]+',g(4)) if s.strip()],'issue':g(5),'registered':g(7).upper()=='O','title':title,'link':link})
def find(sku,reason):
    for e in entries:
        if sku.upper() in e['skus'] and issue_match(reason,e['issue']): return e
pres=svc.get(presentationId=PID).execute()
def txt(e): return ''.join(te.get('textRun',{}).get('content','') for te in e.get('shape',{}).get('text',{}).get('textElements',[])).strip()
def at(e,L,T): t=e.get('transform',{}); return abs(t.get('translateX',0)/12700-L)<1 and abs(t.get('translateY',0)/12700-T)<1
reqs=[]; log=[]
for i,s in enumerate(pres['slides']):
    title=sku=reason=''; has=False
    for e in s.get('pageElements',[]):
        if 'shape' not in e: continue
        t=txt(e)
        if 'Claims / Reviews' in t: title=t
        if at(e,495.5,308.7): sku=t
        if at(e,494.8,217.9): reason=t
        if 'SIREN' in t: has=True
    if not ('Galaxy Z8' in title or 'Pixel 11' in title) or not sku: continue
    m=find(sku,reason)
    if not m: continue
    if not m['registered']: log.append('%d %s/%s: SIREN row %d matched but SIREN 등록 = X → no badge'%(i+1,sku,reason,m['row'])); continue
    if has: log.append('%d %s: badge already present'%(i+1,sku)); continue
    oid='siren_%s'%s['objectId'].replace('-','_')[:30]
    reqs.append({'createShape':{'objectId':oid,'shapeType':'ROUND_RECTANGLE','elementProperties':{'pageObjectId':s['objectId'],
        'size':{'width':{'magnitude':81*EMU,'unit':'EMU'},'height':{'magnitude':14*EMU,'unit':'EMU'}},
        'transform':{'scaleX':1,'scaleY':1,'translateX':391*EMU,'translateY':63.8*EMU,'unit':'EMU'}}}})
    reqs.append({'updateShapeProperties':{'objectId':oid,'shapeProperties':{
        'shapeBackgroundFill':{'solidFill':{'color':{'rgbColor':{'red':0.29,'green':0.08,'blue':0.15}},'alpha':1}},
        'outline':{'propertyState':'NOT_RENDERED'},'contentAlignment':'MIDDLE','link':{'url':m['link']}},
        'fields':'shapeBackgroundFill,outline,contentAlignment,link'}})
    reqs.append({'insertText':{'objectId':oid,'text':'SIREN 등록됨'}})
    reqs.append({'updateTextStyle':{'objectId':oid,'style':{'fontFamily':'Roboto Mono','weightedFontFamily':{'fontFamily':'Roboto Mono','weight':700},'bold':True,
        'fontSize':{'magnitude':7.5,'unit':'PT'},'foregroundColor':{'opaqueColor':{'rgbColor':{'red':1,'green':0.32,'blue':0.32}}}},'fields':'fontFamily,weightedFontFamily,bold,fontSize,foregroundColor'}})
    reqs.append({'updateParagraphStyle':{'objectId':oid,'style':{'alignment':'CENTER'},'fields':'alignment'}})
    log.append('%d %s / %s → SIREN row %d (%s) link=%s'%(i+1,sku,reason,m['row'],m['issue'],m['link'][:60]))
print('\n'.join(log)); print(len(reqs)//5,'badges to add')
if '--apply' in sys.argv and reqs:
    for k in range(0,len(reqs),50): svc.batchUpdate(presentationId=PID,body={'requests':reqs[k:k+50]}).execute()
    print('applied')
