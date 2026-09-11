"""Fill 아마존 리뷰 평점/갯수 (and blank ASIN) on the Slide-Maker cards from the DE sheets."""
import json, sys
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
PID='1qHQoYAOvmI-X1rQrRlbzxWtkFQtyNH1lqjpB2szG9vc'
SRC={'Galaxy Z8':'19OhswglYMx_dxSFFDtWI1WYPWq2jONJn6RK84KITwy4','Pixel 11':'12I6z_FFmDIMHa0rLanltKKFp7kI_yREQj3adkMamPgI'}
p='/Users/kevinkim/.config/gws_shim/token.json'
d=json.load(open(p)); c=Credentials.from_authorized_user_info(d); c.refresh(Request()); d['token']=c.token; json.dump(d,open(p,'w'))
svc=build('slides','v1',credentials=c).presentations(); ss=build('sheets','v4',credentials=c).spreadsheets()
def de_map(sid):
    v=ss.values().get(spreadsheetId=sid,range="'DE'!A1:J400").execute().get('values',[])
    hi=next(i for i,r in enumerate(v) if 'SKU' in r and 'ASIN' in r); h=v[hi]
    ix={k:h.index(k) for k in ('SKU','ASIN','Score','Global Ratings')}
    m={}
    for r in v[hi+1:]:
        g=lambda k: r[ix[k]].strip() if ix[k]<len(r) else ''
        if g('SKU') and g('SKU') not in m: m[g('SKU')]={'asin':g('ASIN'),'score':g('Score'),'count':g('Global Ratings')}
    return m
def rating_text(e):
    if not e or not e['score']: return ''
    try: score='%.1f'%float(e['score'])
    except: score=e['score']
    cnt=e['count'].replace(',','')
    try: cnt='{:,}'.format(int(float(cnt)))
    except: pass
    return '%s점  Global Ratings: %s'%(score,cnt) if cnt else '%s점'%score
maps={k:de_map(v) for k,v in SRC.items()}
pres=svc.get(presentationId=PID).execute()
def txt(e): return ''.join(te.get('textRun',{}).get('content','') for te in e.get('shape',{}).get('text',{}).get('textElements',[])).strip()
def style_of(e):
    for te in e['shape']['text']['textElements']:
        if 'textRun' in te: return te['textRun'].get('style',{})
    return {}
def box(e,L,T): t=e.get('transform',{}); return abs(t.get('translateX',0)/12700-L)<1 and abs(t.get('translateY',0)/12700-T)<1
# reference style from an original card's filled rating/ASIN box
ref={}
for s in pres['slides']:
    if s['objectId'].startswith('SLIDES_API'): continue
    for e in s.get('pageElements',[]):
        if 'shape' in e and box(e,495.9,355.0) and txt(e) and 'rating' not in ref: ref['rating']=style_of(e)
        if 'shape' in e and box(e,495.5,261.7) and txt(e) and 'asin' not in ref: ref['asin']=style_of(e)
keep=('fontFamily','fontSize','foregroundColor','bold','weightedFontFamily')
reqs=[]; log=[]
for i,s in enumerate(pres['slides']):
    if not s['objectId'].startswith('SLIDES_API'): continue
    title=''; sku=''; boxes={}
    for e in s.get('pageElements',[]):
        if 'shape' not in e: continue
        t=txt(e)
        if 'Claims / Reviews' in t: title=t
        if box(e,495.5,308.7): sku=t
        if box(e,495.9,355.0): boxes['rating']=e
        if box(e,495.5,261.7): boxes['asin']=e
    fam=next((f for f in SRC if f in title),None)
    if not fam or not sku: continue
    entry=maps[fam].get(sku)
    for field,val in (('rating',rating_text(entry)),('asin',entry['asin'] if entry else '')):
        e=boxes.get(field)
        if not e or not val or txt(e): continue
        st={k:v for k,v in ref.get(field,{}).items() if k in keep}
        reqs.append({'insertText':{'objectId':e['objectId'],'text':val,'insertionIndex':0}})
        if st: reqs.append({'updateTextStyle':{'objectId':e['objectId'],'style':st,'fields':','.join(st.keys())}})
        log.append('%d %s %s -> %s'%(i+1,sku,field,val))
    if not entry: log.append('%d %s: not in %s DE sheet'%(i+1,sku,fam))
print('\n'.join(log)); print(len(reqs),'requests')
if '--apply' in sys.argv and reqs:
    for k in range(0,len(reqs),60): svc.batchUpdate(presentationId=PID,body={'requests':reqs[k:k+60]}).execute()
    print('applied')
