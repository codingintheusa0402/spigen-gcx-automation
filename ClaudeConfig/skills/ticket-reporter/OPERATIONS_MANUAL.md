# AI 티켓 리포트 운영 매뉴얼 (ticket-reporter)

관리자: 김지우 프로 (글로벌CX전략팀) · 최초 작성 2026-10-16
리포트 작성 규칙·판단 기준의 원본은 같은 폴더의 `SKILL.md`이며, 이 문서는 **운영(호스트·스케줄·장애·점검)** 기준을 다룬다.

---

## 1. 목적
Zendesk T2 에스컬레이션 티켓과 Lazada/Shopee TCT 로그를 감시해, TCK 전달 전 보고용 리포트를
"GCX T2 ESC. Ticket 보고" Chat 방(`spaces/AAQAdqYt1ro`)에 자동 게시한다. 최종 처리 결정은 항상
리더·프로 컨펌으로 이루어지며, AI 리포트는 컨펌을 돕는 초안이다.

## 2. 구성 요소
- `SKILL.md` — 리포트 양식·판단 기준·모니터 절차 (작성 규칙의 원본)
- `send.py` — cardsV2 카드 전송, `--ticket-id`로 TicketQueue 시트에 스레드 매핑 기록
- `siren.py` — 결함 패턴 조기경보(SIREN) 카운트·알림
- `scan_tct_log.py` — Lazada/Shopee TCT 로그 Esc T2 행 스캔
- `check_feedback.py` / `check_btw_queue.py` — /revision 피드백·/btw 요청 큐 확인
- `state/processed.json` — 처리 이력(dedup), 상태 추적, 틱 로그
- `kr_holidays.json` — 공휴일 게이트
- `~/.gcx-tasks/ticket_state_guard.py` — 기기 간 동기화 시 처리 이력 되돌림 방지 가드

## 3. 실행 호스트
- 기본: 24/7 서버(Windows + WSL2 Ubuntu, tmux `gcx:ticket-monitor`)
- 대체: 관리자 Mac (Chrome CDP 9222 프로필로 Zendesk 조회)
- **한 시점에 한 호스트만 실행한다.** 전환 시 이전 호스트 루프를 먼저 정지하고 `processed.json`을
  동기화한 뒤 새 호스트를 시작한다 (중복 게시 방지).

## 4. 운영 시간
- 평일 08:30–18:00 KST, 목요일은 16:00 종료, `kr_holidays.json` 공휴일 제외
- 업무 시간 외에는 리포트를 게시하지 않으며, 다음 업무 시작 시 누적 건을 처리한다.

## 5. 감시 주기
- Zendesk pending 뷰 2개: 5분 간격
- TCT 로그 스캔: 30분 간격 (`tct_log_last_scan` 기준)
- 매 틱 시작·종료 시 state guard 실행, 틱 결과를 `processed.json` 로그에 기록

## 6. 리포트 대상
- Status = Pending 이고 Category가 정확히 `4. Product Issue` 또는 `2. Delivery`인 티켓
- ESC. 사유가 MCF발송 / 출시·재입고 문의 / 유관부서 이관 요청이면 제외(skip 기록)
- TCT: Esc T2 행 (Esc T1 경유 후 T2 재인입 시 재보고)

## 7. 중복 방지 및 재인입
- 처리한 티켓은 `processed.json`에 기록해 재게시하지 않는다.
- 처리 후 Pending을 벗어났다가 다시 Pending이 된 티켓은 재인입으로 보고 새 코멘트 기준으로 재보고한다.
- 점검·장애로 감시가 중단됐던 경우, 재개 첫 틱에서 처리 이력 티켓의 최신 코멘트를 재확인한다.
  중단 중 이미 Solved된 티켓은 보고하지 않는다.

## 8. 리포트 품질 기준
`SKILL.md`의 HARD RULES·컨펌 판단 패턴을 따른다. 리포트 형식 변경은 /revision 피드백 또는
관리자 지시로만 반영하며, 반영 시 `SKILL.md`에 날짜와 근거를 남긴다.

## 9. SIREN 조기경보
신제품 라인업 대상 결함이 기준 건수 이상 누적되면 `siren.py`가 알림을 보낸다. 라인업 시트에 없는
제품은 자동 경보 대상이 아니므로, 반복 인입이 보이면 리포트 [비고]에 표기하고 관리자가 수동 SIREN 여부를 판단한다.

## 10. 알림 및 장애 대응
- Zendesk 세션 만료(API 401): 빈 결과로 처리하지 않고 해당 틱을 중단, `ticket_login_alert` 1회 발송.
  재로그인은 관리자가 직접 수행한다 (자동 로그인 금지).
- CDP(9222) 다운: 디버그 프로필 Chrome 재기동 후 재시도.
- 스크립트 오류: 해당 단계만 건너뛰고 틱 로그에 기록, 반복 시 관리자에게 보고.

## 11. 보안 원칙
- 웹훅·API 키·토큰은 코드/문서에 직접 기재하지 않는다 (`~/.config/gcx_webhooks.json` 등 별도 보관).
- Zendesk·Amazon·Seller Central 로그인은 사람만 수행한다.
- 테스트 전송은 비공개 테스트 방에서만 하고, 운영 Chat 방에는 하지 않는다.

## 12. 변경 관리
- 스크립트(.py) 변경은 Git에 커밋하고 서버와 동기화한다.
- 판단 기준 변경은 `SKILL.md`, 운영 절차 변경은 이 문서에 반영한다.

## 13. 정기 OS 업데이트 (Windows / Ubuntu)
- 서버의 Windows 및 WSL2 Ubuntu 업데이트는 **웹 보안 및 시스템 안정성** 확보를 위해 정기적으로 적용한다.
- **주요(major) 업데이트는 반드시 관리자에게 착수 승인을 받은 뒤 시작한다.** 자동 업데이트에 의한
  업무 시간 중 재부팅이 일어나지 않도록 한다.
- 착수 전 "GCX T2 ESC. Ticket 보고" 방에 점검 공지(점검 시간·리포트 중단 안내)를 게시한다.
- 점검 시간 동안 리포트 자동 전송을 중단하고, 완료 후 재부팅 자동 복구와 첫 틱 정상 동작을 확인한 뒤 재개한다.

## 14. 점검 시간 중 병행 작업
OS 업데이트가 진행되는 동안 서비스가 정지되어 있으므로, 다음 작업을 함께 수행한다.
- 누적 운영 기간 동안 쌓인 분석 결과를 바탕으로 한 **경미한 버그 수정 및 개선**
  (틱 로그·skip 기록·/revision 피드백에서 드러난 오류 패턴 정리)
- 유사 전례 데이터(`Previous_tickets.csv`)에 최근 컨펌 결과 반영
- 최근 리포트를 현행 판단 기준으로 재검증해 AI 예상답변과 실제 컨펌 간 일치율 점검
- 병행 작업의 변경 사항은 §12 변경 관리 절차에 따라 기록한다.
