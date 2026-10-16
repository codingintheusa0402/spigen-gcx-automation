"""Write temp AaMailDraft.js (GAS) that creates the Gmail draft from mail.html (+ deletes mail.json replace_draft)."""
import os, json
HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.join(HERE, '..', '..')
M = json.load(open(os.path.join(HERE, 'mail.json')))
CC = ['arrha@spigen.com','kisepark@spigen.com','ckchoi@sgp-korea.com','jmchoi@spigen.com','chpark@spigen.com','sjjung@spigen.com','swyoon@spigen.com','dklee@spigen.com','kjh3526@spigen.com','gina_jin@spigen.com','swyun@spigen.com','chpak@spigen.com','mhsong@spigen.com','kmson@spigen.com','skpark@spigen.com','kskwon@spigen.com','mgchoi@spigen.com','aya.chen@spigen.com','cqc@spigen.com','cq_production@spigen.com','kbkim@spigen.com','gwlee@spigen.com','swpark@spigen.com','parksm@spigen.com','jihyunlee@spigen.com','yangsr@spigen.com','wanqiong.yang@spigen.com','reda.zou@spigen.com','mjlee@spigen.com','dykim@spigen.com','ebhwang@spigen.com','andy.lee@spigen.com','parkse@spigen.com','jspark@spigen.com','yjlee@spigen.com','wikim@spigen.com','jsong@spigen.com','jokim@spigen.com','hychae@spigen.com','ynchoi@spigen.com','sungjun@spigen.com','jaeseoung@spigen.com','shwoo@spigen.com']
cc = M.get('cc', CC)
html = open(os.path.join(HERE, 'mail.html')).read()
js = f"""// TEMP (bi-weekly-builder): create the bi-weekly Gmail draft (live HTML). Delete after running.
function aaMailDraft() {{
  const old = {json.dumps(M.get('replace_draft', ''))};
  if (old) {{ try {{ GmailApp.getDraft(old).deleteDraft(); Logger.log('deleted old draft ' + old); }} catch (e) {{ Logger.log('old draft not found'); }} }}
  const html = {json.dumps(html, ensure_ascii=False)};
  const d = GmailApp.createDraft({json.dumps(M['to'])}, {json.dumps(M['subject'], ensure_ascii=False)},
    {json.dumps(M['title'] + ' — ' + M['subtitle'], ensure_ascii=False)}, {{ htmlBody: html, cc: {json.dumps(','.join(cc))} }});
  Logger.log('draft ' + d.getId());
}}
"""
open(os.path.join(ROOT, 'AaMailDraft.js'), 'w').write(js)
print('wrote AaMailDraft.js', len(js))
