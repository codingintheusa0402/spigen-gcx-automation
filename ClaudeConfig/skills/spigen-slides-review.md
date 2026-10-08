---
name: spigen-slides-review
description: "Spigen Google Slides 덱 생성 후 사용자가 검증/검수/2번/검증 실행을 요청할 때만 발동. spigen-slides로 만든 완성본을 대상으로 자동 검증, 3종 에이전트 검수, gate status 확인을 수행한다."
license: MIT
metadata:
  category: productivity
  locale: ko-KR
  phase: v1.0.0
---

# spigen-slides-review

`spigen-slides`로 만든 덱을 사용자가 명시적으로 검증 요청했을 때만 실행한다.

## 절대 규칙

1. 이 스킬은 **사용자 요청 후에만** 실행한다.
2. 제작 스킬(`spigen-slides`) 안에서 자동 호출하지 않는다.
3. 검수 에이전트는 **기획자 / 디자이너 / 청중** 3종을 병렬 호출한다.
4. 검수 호출은 **1회만** 한다.
5. 검수 후 수정은 이 스킬에서 자동 수행하지 않는다. 수정이 필요하면 검수 결과를 요약하고 사용자에게 다음 제작 작업으로 넘긴다.

## 입력

검증 대상은 다음 중 하나로 받는다.

- Google Slides URL
- presentation ID
- 직전 `spigen-slides` 빌드 결과

URL이면 `/presentation/d/<ID>/edit`의 `<ID>`를 사용한다.

## 실행 흐름

```
Step 1. presentation ID 확인
Step 2. review gate init
Step 3. 자동 검증 결과 확인
Step 4. 에이전트 3종 병렬 검수
Step 5. gate record
Step 6. status --require-pass 확인
Step 7. 결과 보고
```

## Step 2. Review Gate Init

아래 명령을 실행한다.

```bash
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py init <PRESENTATION_ID> --audience "<청중>" --purpose "<목적>"
```

이 단계는 다음 파일을 생성한다.

- `/tmp/spigen_review_<PRESENTATION_ID>/manifest.json`
- `/tmp/spigen_review_<PRESENTATION_ID>/verify.txt`
- `/tmp/spigen_review_<PRESENTATION_ID>/presentation.json`
- `/tmp/spigen_review_<PRESENTATION_ID>/thumbnails/slide_<N>.png`
- `/tmp/spigen_review_<PRESENTATION_ID>/reports/*.md`

`verify.txt`가 FAIL이면 에이전트 검수 전에 자동 검증 실패로 보고한다. 단, 사용자가 계속 페르소나 의견까지 원하면 3종 검수를 진행한다.

## Step 4. 에이전트 3종 검수

반드시 `spawn_agent`를 사용한다.

검수 에이전트:

- `planner`: 메시지 구조, 소스 충실도, 구성 일관성
- `designer`: 규격 이탈, 겹침, 여백, 폰트, 컬러, 텍스트 넘침
- `audience`: 청중 이해도, 실무 적용성, 입력 자산 부족

각 에이전트에게 전달할 공통 정보:

- presentation ID
- review root path
- report path
- 덱 목적
- 청중
- 사용자가 승인한 아웃라인 또는 제작 요청
- 원본 소스 요약

각 에이전트는 리포트를 직접 저장해야 한다.

리포트 경로:

- `/tmp/spigen_review_<ID>/reports/planner.md`
- `/tmp/spigen_review_<ID>/reports/designer.md`
- `/tmp/spigen_review_<ID>/reports/audience.md`

## Step 5. Gate Record

각 에이전트 완료 후 `record`를 실행한다.

```bash
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py record <ID> --reviewer planner --status pass|fail --report /tmp/spigen_review_<ID>/reports/planner.md --summary "<요약>"
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py record <ID> --reviewer designer --status pass|fail --report /tmp/spigen_review_<ID>/reports/designer.md --summary "<요약>"
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py record <ID> --reviewer audience --status feedback --report /tmp/spigen_review_<ID>/reports/audience.md --summary "<요약>"
```

판정 기준:

- planner: 통과면 `pass`, 불통과면 `fail`
- designer: 통과면 `pass`, 불통과면 `fail`
- audience: 항상 `feedback`

## Step 6. Status

```bash
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py status <ID> --require-pass
```

통과 조건:

- 자동 검증 PASS
- `init` 이후 덱 내용 변경 없음
- 썸네일 생성 완료
- planner `pass`
- designer `pass`
- audience `feedback` 또는 `pass`

통과하면 cleanup을 실행한다.

```bash
python3 ~/.agents/skills/spigen-slides-review/spigen_review_gate.py cleanup <ID>
```

불통과하면 cleanup하지 않고 산출물 경로를 남긴다.

## 결과 보고

통과 시:

```text
검증 통과했습니다.
- 자동 검증: PASS
- 기획자: PASS
- 디자이너: PASS
- 청중: 피드백 완료
- Gate: PASS
```

불통과 시:

```text
검증 불통과입니다.
- 자동 검증: PASS/FAIL
- 기획자: PASS/FAIL
- 디자이너: PASS/FAIL
- 청중: 피드백 완료
- 산출물: /tmp/spigen_review_<ID>/

수정은 별도 제작 단계로 진행해야 합니다.
```

중요: 불통과 후 이 스킬 안에서 반복 수정/재검수를 시작하지 않는다.
