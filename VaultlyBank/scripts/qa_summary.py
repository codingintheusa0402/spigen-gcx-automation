#!/usr/bin/env python3
"""Print one line of pass/fail counts (+ failing test names) for an .xcresult. Used by scripts/qa.sh."""
import json, subprocess, sys
try:
    raw = subprocess.run(["xcrun", "xcresulttool", "get", "test-results", "summary", "--path", sys.argv[1]],
                         capture_output=True, text=True, timeout=120).stdout
    d = json.loads(raw)
    print(f"{d.get('passedTests', 0)} passed · {d.get('failedTests', 0)} failed · {d.get('skippedTests', 0)} skipped")
    for f in d.get("testFailures", [])[:20]:
        print(f"  - ❌ {f.get('testName')}: {(f.get('failureText') or '')[:200]}")
except Exception:
    print("no test summary (build failed?)")
