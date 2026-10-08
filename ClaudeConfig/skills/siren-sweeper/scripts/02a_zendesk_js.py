"""Step 2a — print the browser JS that collects every claim ticket's customer messages + attachments.
Why the browser: there is no Zendesk API token on this Mac; the agent session (Chrome, logged in) is the only
access. Passing a Google token into the page is blocked by policy, and the page can't POST to localhost
(Private-Network-Access prompt), so results are read back with get_page_text (see SKILL.md step 2).
Usage: python3 02a_zendesk_js.py > zendesk_collect.js   (then paste into javascript_tool on an /api/v2/... page)"""
import json, common  # noqa
ids = sorted(int(x) for x in json.load(open('ids.json')))
prev, out = 0, []
for i in ids:
    out.append(format(i - prev, 'x') if prev else str(i)); prev = i
D = ','.join(out)   # delta-encoded so ~1000 ids fit in a few KB of JS
print("""const D='%s';
const parts=D.split(',');const ids=[];let p=0;parts.forEach((x,k)=>{p=k?p+parseInt(x,16):+x;ids.push(p)});
window.__p={done:0,err:0,errs:[],total:ids.length,fin:false};window.__recs={};
(async()=>{
 async function get(u){for(let k=0;k<6;k++){const r=await fetch(u,{credentials:'include'}); if(r.status===429){await new Promise(s=>setTimeout(s,(+r.headers.get('retry-after')||10)*1000));continue;} if(!r.ok) throw new Error(r.status); return r.json();} throw new Error('429x');}
 let i=0;
 async function worker(){while(i<ids.length){const id=ids[i++];try{
   const j=await get(`/api/v2/tickets/${id}/comments.json?include=users`);
   const cs=j.comments; const endUsers=new Set((j.users||[]).filter(u=>u.role==='end-user').map(u=>u.id));
   const mine=cs.filter(c=>endUsers.has(c.author_id));
   const src=mine.length?mine:cs.slice(0,1);
   const body=src.slice(0,3).map(c=>(c.plain_body||c.body||'').replace(/\\s+/g,' ').trim()).join(' || ').slice(0,700);
   const att=mine.flatMap(c=>c.attachments).filter(a=>/^(image|video)\\//.test(a.content_type)||/\\.(heic|heif|mov|mp4|jpe?g|png|webp)$/i.test(a.file_name)).slice(0,6).map(a=>{const m=a.content_url.match(/token\\/([^/]+)\\//);return [m?m[1]:a.content_url,(a.file_name||'').replace(/[\\s|]/g,'_').slice(-40),a.content_type,a.size]});
   window.__recs[id]={b:body,a:att};
   window.__p.done++;}catch(e){window.__p.err++;window.__p.errs.push(id+':'+e.message);}}}
 await Promise.all([worker(),worker(),worker(),worker()]); window.__p.fin=true;
})();
ids.length""" % D)
