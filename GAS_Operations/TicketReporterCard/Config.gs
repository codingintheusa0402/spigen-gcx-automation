/**
 * Ticket Reporter — interactive Chat app config.
 *
 * GCP project: gcx-zendesk-decision-maker (existing project, repurposed — its old
 * "Quickstart App" Apps Script binding was an inert Access-Denied placeholder).
 * Chat API Configuration → Connection settings → Apps Script → this project's deployment ID.
 *
 * Secrets (ZENDESK_EMAIL / ZENDESK_API_TOKEN) live in Script Properties, never in code —
 * set once via Apps Script editor: Project Settings → Script Properties, or run setupOnce_().
 *
 * `/btw` Q&A (see handleBtwQuestion_ in Code.gs) needs no Script Properties at all — by
 * design (사용자 지시 2026-10-07) it never calls an LLM API directly from Apps Script. It only
 * appends the question to the BtwQueue sheet tab; the ticket-reporter monitor Claude Code
 * session answers it on its next 5-minute tick and replies in-thread via send.py --thread.
 */

var ZENDESK_SUBDOMAIN = 'spigenhelp';                 // https://spigenhelp.zendesk.com
var QUEUE_SHEET_PROP = 'QUEUE_SHEET_ID';              // Script Property holding the queue Sheet's file ID
var QUEUE_TAB = 'TicketQueue';
var FEEDBACK_TAB = 'Feedback';                        // /revision submissions (see handleRevisionFeedback_)
var UNMAPPED_TAB = 'UnmappedThreads';                 // thread replies with no TicketQueue row (see logUnmappedThread_)
var MANUAL_TAB = 'Manual';                            // full SKILL.md text mirror for /manual (see handleManualRequest_)
var MANUAL_TAB_GID = '2076587794';                    // Manual tab's sheetId, for the direct #gid= deep link in /manual's reply
var BTW_TAB = 'BtwQueue';                             // /btw question queue, answered by the monitor session (see handleBtwQuestion_)

/**
 * Lazada/Shopee TCT log sheet — the second monitor source (사용자 지시 2026-09-18). A thread
 * reply for a ticket sourced from here (ticketId is NOT all-digit, e.g. "260915BQH9J0")
 * writes back into this sheet instead of calling the Zendesk API — see updateTctLogRow_.
 */
var TCT_LOG_SHEET_ID = '1HZ14uqTVeP7bGYZDu9v9Ve2C1xNY_m6dcSv-KMCoAKc';
var TCT_LOG_TABS = ['Lazada log', 'Shopee log'];
// Column letters (1-indexed) in both tabs — identical schema confirmed 2026-09-18.
var TCT_LOG_COL = { STATUS: 1, TICKET_ID: 2, VOUCHER: 20, MEMO: 21, GCX_STATUS: 22 };
// Exact strings required by each column's strict ONE_OF_LIST data validation — Sheets API/
// Apps Script writes bypass validation enforcement, but conditional formatting keyed on
// these exact strings (including the accidental trailing spaces baked into the dropdown
// itself) will only color correctly on an exact match. NEVER "fix" the whitespace.
var TCT_LOG_STATUS_ESC_T1 = 'Esc T1  ';                // two trailing spaces, verified 2026-09-18
var TCT_LOG_GCX_STATUS_ADVICE_GIVEN = 'Advice given';

var CONFIRMERS = ['KJW', 'YSR', 'NAR'];               // [GCX <이니셜> 컨펌] dropdown, KJW default

/**
 * Thread-reply direct-post path only (postThreadReplyAsNote_ in Code.gs): maps the Chat
 * user who sent the @Ticket Reporter mention to the confirmer code used in
 * "[GCX <code> 컨펌]" — independent of the CONFIRMERS dropdown above. Falls back to
 * CONFIRMERS[0] (KJW) for anyone not listed here (사용자 지시 2026-09-16).
 */
var CONFIRMER_BY_EMAIL = {
  'kjw@spigen.com': 'KJW',    // 김지우 Kevin
  'yangsr@spigen.com': 'L',   // 양숙랑 Grace
  'arrha@spigen.com': 'NAR'   // 나아름 Jane
};

/**
 * Single reference dropdown, in the user's original 1-43 numbering. It is a picker only —
 * the actual internal note body is the freeform "노트 내용" text box the user types/edits.
 * Selecting an item here just appends that phrase as a new line onto whatever is already
 * in the note box (see onRefChange in Code.gs); selecting it twice appends it twice.
 */
var REFERENCE_PHRASES = [
  '사진 확인 결과, 사용환경 및 외부 영향에 의한 파손으로 사료되며 제조상 결함은 확인되지 않음.', // 1
  '예외적인 50% 환불 또는 MCF 중 선택하여 회신 바랍니다.',                                        // 2
  '기기 낙하 등 제품 사용에 의한 파손은 워런티 대상이 아닙니다.',                                  // 3
  '도움 불가한 점 양해 부탁드립니다.',                                                              // 4
  '대체품 제공으로 고객 불편 해소 도움을 드린 바 있으나 재차 요청하셔서 예외적으로 환불 도움드립니다.', // 5
  '제품 사용 중 불편을 겪으신 점 유감입니다.',                                                       // 6
  '사용환경 및 외부 요인에 따른 이염·오염으로 확인되며 제조상 결함으로 보기 어렵습니다.',            // 7
  '고객 의견은 유관 부서에 전달 완료되었습니다.',                                                    // 8
  '고객 불편을 고려하여 환불 도움드립니다.',                                                         // 9
  '첨부된 제품 사진은 슈피겐 케이스와 상이합니다.',                                                  // 10
  '정품이 아닌 것으로 사료되어 폐기 사진 첨부는 불필요합니다.',                                     // 11
  '워런티 도움 불가합니다.',                                                                          // 12
  '고객 상황에 대해 유감스럽게 생각합니다.',                                                         // 13
  '황변의 경우, 형광등, 자외선 노출, 충전기 발열, 화장품 등 다양한 유저의 사용환경 및 외부요인에 의해 발생될 수 있음', // 14
  '재인입 시 NRN 처리 예정입니다.',                                                                   // 15
  '고객 불편에 대해 유감입니다.',                                                                     // 16
  '스크래치는 사용 중 마찰 등에 의한 외부 요인으로 워런티 대상이 아닙니다.',                        // 17
  '충성 고객을 고려하여 예외적 환불 도움드립니다.',                                                  // 18
  '고객님의 불편에 대해 유감이나, 첨부된 폐기 사진에서 디지털 편집이 감지되어 도움드리기 어렵습니다.', // 19
  '재인입 시 NRN 처리 바랍니다.',                                                                     // 20
  '제품 상태를 확인할 수 없어 불편한 부분의 영상 또는 사진을 첨부해 주시기 바랍니다.',              // 21
  '예: 케이스 전체가 보이도록 탈착 영상 첨부 요청드립니다.',                                        // 22
  '확인 후 빠른 도움을 드릴 예정입니다.',                                                             // 23
  '이전에 대체품이 제공되었으며, 수령일로부터 약 N개월 사용이 경과된 것으로 확인됩니다.',           // 24
  '제조상 결함은 확인되지 않았습니다.',                                                               // 25
  '사용환경 및 외부 영향에 따른 변형은 워런티 대상이 아닙니다.',                                    // 26
  '고객 상황에 대한 유감 표현드립니다.',                                                             // 27
  '낙하로 인한 파손은 제조상 결함이 아니므로 워런티 도움 불가합니다.',                              // 28
  '첨부된 사진에서 디지털 편집이 감지되었습니다.',                                                   // 29
  '워런티 도움 불가하며, 재인입 시에도 도움 불가합니다.',                                            // 30
  '해당 케이스는 관리대상으로 지정됩니다.',                                                           // 31
  '사진 확인 결과 제조상 결함은 확인되지 않았습니다.',                                               // 32
  '부착 방법 오류로 판단됩니다.',                                                                     // 33
  '가이드에 따라 재설치 바랍니다.',                                                                   // 34
  '사진 확인 결과 제조상 결함이 아니며 사용환경 및 외부 영향에 의한 파손으로 사료됩니다.',           // 35
  '도움 불가합니다.',                                                                                  // 36
  '확인 결과 제조상 결함은 발견되지 않았으나, 고객 불편을 고려하여 환불로 도움드립니다.',           // 37
  '영상 첨부 감사드립니다.',                                                                           // 38
  '주문 후 N년 N개월 경과된 점 확인됩니다.',                                                          // 39
  '사용환경 및 외부 영향에 따른 기능 저하는 워런티 대상이 아닙니다.',                                // 40
  '고객 불편에 대한 유감 표현드립니다.',                                                             // 41
  '예외적으로 환불 도움드립니다.',                                                                    // 42
  '앞으로도 슈피겐 제품에 많은 관심 부탁드립니다.'                                                   // 43
];
