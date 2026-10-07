# chat_app_jane — second identity of the BadReview Chat app

Galaxy Z8 Series instance of the interactive BadReview Chat app (`../chat_app/` is the
Pixel 11 / Kevin instance). Shows a control card (시작일/종료일, 국가, 기종, [조회]) and a
배드리뷰 (1~3점) report card scoped to the chosen range + filters.

## What differs from `../chat_app/`

| File | Jane | Kevin (`../chat_app/`) |
|------|------|------------------------|
| `Config.gs` | `APP_PRODUCT = 'glxz8'` | `'pixel11'` |
| `appsscript.json` | `addOns.common.name` / `logoUrl` = 나아름 Jane 글로벌CX전략팀 | 김지우 Kevin 글로벌CX전략팀 |
| `Code.gs` | **identical copy** | original |

## Deploy

| | |
|---|---|
| Apps Script | `1Zhc91kpARwwlNctKsk2THxt8nIvj6SvuJMOFOaKzrJotbWsKobSxbFhI` |
| GCP project | `formats-uaox` (Chat API enabled + Internal OAuth consent, 2026-09-11) |
| Deployment ID | see `../chat_app/SETUP.md` → "Live deployment" |

```bash
cp ../chat_app/Code.gs .          # after every edit to the original
clasp push --force
clasp deploy -i <deploymentId>    # moves the live Chat app to the new version
```

Handlers (`onMessage`, `onAddToSpace`, `onRemoveFromSpace`, `refreshReport`) and the
add-on response envelope are documented in `../chat_app/SETUP.md` → "Lessons that cost
time" — read it before changing the manifest or GCP config.
