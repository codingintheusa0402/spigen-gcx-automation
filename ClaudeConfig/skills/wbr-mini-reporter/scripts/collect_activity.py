#!/usr/bin/env python3
"""Collect the user's own prompts + first assistant summary lines from Claude Code
session transcripts in a date window, for drafting the WBR portion.
Usage: collect_activity.py SINCE(YYYY-MM-DD) [UNTIL(YYYY-MM-DD)] [--exclude SESSION_ID]"""
import json, glob, os, sys, datetime as dt
args=sys.argv[1:]; excl=set()
if '--exclude' in args:
    i=args.index('--exclude'); excl.add(args[i+1]); del args[i:i+2]
since=args[0]; until=args[1] if len(args)>1 else '9999-12-31'
root=os.path.expanduser('~/.claude/projects')
files=[f for f in glob.glob(root+'/*/*.jsonl')
       if 'scratchpad' not in f and os.path.basename(f)[:-6] not in excl
       and dt.datetime.fromtimestamp(os.path.getmtime(f)).strftime('%Y-%m-%d')>=since]
SKIP=('<command-','<local-command','<system-reminder','Caveat:','<task-notification','[Request interrupted')
for f in sorted(files, key=os.path.getmtime):
    out=[]
    for line in open(f, errors='ignore'):
        try: o=json.loads(line)
        except: continue
        if o.get('type')!='user' or o.get('isSidechain') or o.get('isMeta'): continue
        ts=(o.get('timestamp') or '')[:10]
        if not (since<=ts<=until): continue
        c=o.get('message',{}).get('content')
        if isinstance(c,list):
            c=' '.join(x.get('text','') for x in c if isinstance(x,dict) and x.get('type')=='text')
        if not isinstance(c,str): continue
        c=c.strip()
        if not c or c.startswith(SKIP) or 'tool_use_id' in c: continue
        out.append(f"  [{o['timestamp'][5:16]}] {c[:400].replace(chr(10),' ')}")
    if out:
        print(f"\n### {os.path.basename(f)[:8]}  ({len(out)} prompts)")
        print('\n'.join(out[:40]))
