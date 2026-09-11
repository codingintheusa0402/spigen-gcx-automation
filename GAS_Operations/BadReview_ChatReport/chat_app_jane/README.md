# chat_app_jane — second identity of the BadReview Chat app

Same `Code.gs` as `../chat_app/` (copy — re-copy after editing the original:
`cp ../chat_app/Code.gs .` then `clasp push --force`). Only `appsscript.json`
differs: `addOns.common.name` / `logoUrl` = 나아름 Jane 글로벌CX전략팀.

Backed by GCP project `formats-uaox` (chosen 2026-09-11; Chat API had never been
enabled there). Follow `../chat_app/SETUP.md` → "Lessons that cost time".
