#!/bin/bash
# Vaultly quality gate — run before EVERY push of a Vaultly change (the pre-push hook does it for you).
#   1. regenerate the Xcode project (XcodeGen) and the fictional-name table
#   2. personal build: unit tests + every UI walkthrough (all tabs, quests, mini games …) with screenshots
#   3. App Store build (STORE flag): the same UI walkthroughs with screenshots
#   4. export every screenshot to build/qa/<run>/<scheme>/ and write build/qa/<run>/REPORT.md
# Exit code ≠ 0 if anything failed — the hook then refuses the push.
# Usage: bash scripts/qa.sh [--fast]   (--fast: personal build only, for quick local loops; the hook never uses it)
set -uo pipefail
cd "$(dirname "$0")/.."
RUN="build/qa/$(date +%Y%m%d-%H%M%S)"; mkdir -p "$RUN"
REPORT="$RUN/REPORT.md"
DEST="${QA_DESTINATION:-platform=iOS Simulator,name=iPhone 16,OS=latest}"
SCHEMES=("Vaultly" "Vaultly-Store"); [ "${1:-}" = "--fast" ] && SCHEMES=("Vaultly")
fail=0
echo "# Vaultly QA — $(date '+%Y-%m-%d %H:%M') · $(git rev-parse --short HEAD 2>/dev/null)" > "$REPORT"

command -v xcodegen >/dev/null || { echo "❌ xcodegen missing (brew install xcodegen)"; exit 2; }
python3 scripts/store_names.py >/dev/null && echo "- fictional-name table regenerated" >> "$REPORT"
xcodegen generate --quiet || { echo "❌ xcodegen failed"; exit 2; }
if ! git diff --quiet -- Vaultly/App/StoreNames.swift Vaultly/Models/QuestLibrary.swift 2>/dev/null; then
  echo "- ⚠️ StoreNames.swift / QuestLibrary.swift were stale — regenerated; commit them" >> "$REPORT"; fail=1
fi

for S in "${SCHEMES[@]}"; do
  R="$RUN/$S.xcresult"
  echo "▶︎ $S: build + test on $DEST"
  xcodebuild test -project Vaultly.xcodeproj -scheme "$S" -destination "$DEST" -resultBundlePath "$R" \
      CODE_SIGNING_ALLOWED=NO -quiet > "$RUN/$S.log" 2>&1
  code=$?
  mkdir -p "$RUN/$S"
  xcrun xcresulttool export attachments --path "$R" --output-path "$RUN/$S" >/dev/null 2>&1 || true
  shots=$(find "$RUN/$S" -name '*.png' -o -name '*.jpeg' 2>/dev/null | wc -l | tr -d ' ')
  summary=$(xcrun xcresulttool get test-results summary --path "$R" 2>/dev/null | python3 -c '
import sys, json
try:
    d = json.load(sys.stdin); print(f"{d.get(\"passedTests\",0)} passed · {d.get(\"failedTests\",0)} failed · {d.get(\"skippedTests\",0)} skipped")
    for f in d.get("testFailures", [])[:20]: print(f"  - ❌ {f.get(\"testName\")}: {f.get(\"failureText\",\"\")[:200]}")
except Exception: print("no test summary (build failed?)")')
  if [ $code -eq 0 ]; then echo "## ✅ $S" >> "$REPORT"; else echo "## ❌ $S (exit $code)" >> "$REPORT"; fail=1
    grep -E "error:|failed" "$RUN/$S.log" | head -20 | sed 's/^/    /' >> "$REPORT"; fi
  echo "$summary" >> "$REPORT"
  echo "- screenshots: $shots in \`$RUN/$S/\`" >> "$REPORT"
done

[ $fail -eq 0 ] && echo -e "\n**PASS** — OK to push." >> "$REPORT" || echo -e "\n**FAIL** — do not push." >> "$REPORT"
cat "$REPORT"
ln -sfn "$(basename "$RUN")" build/qa/latest
exit $fail
