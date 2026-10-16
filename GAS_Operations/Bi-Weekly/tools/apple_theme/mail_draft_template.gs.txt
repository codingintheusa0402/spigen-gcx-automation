// TEMP (bi-weekly-builder): Apple-themed Gmail draft for the bi-weekly mail. Delete after running.
function aaMailDraft() {
  const DECK = '1sJ49LsEqMFrFyykAsgpWIREO9SCnWa5GW-ILEBLbG8Q';
  const LINK = 'https://docs.google.com/presentation/d/' + DECK + '/edit?usp=sharing';
  const PAGES = ["g3f2ca1aded2_0_19", "SLIDES_API1905798936_0", "SLIDES_API1905798936_88", "SLIDES_API1905798936_176", "SLIDES_API1905798936_264", "SLIDES_API1905798936_352", "SLIDES_API1905798936_637", "SLIDES_API1905798936_922", "SLIDES_API1905798936_15996", "SLIDES_API1905798936_16294"];
  const TO = 'wbd@spigen.com';
  const CC = ['arrha@spigen.com','kisepark@spigen.com','ckchoi@sgp-korea.com','jmchoi@spigen.com','chpark@spigen.com','sjjung@spigen.com','swyoon@spigen.com','dklee@spigen.com','kjh3526@spigen.com','gina_jin@spigen.com','swyun@spigen.com','chpak@spigen.com','mhsong@spigen.com','kmson@spigen.com','skpark@spigen.com','kskwon@spigen.com','mgchoi@spigen.com','aya.chen@spigen.com','cqc@spigen.com','cq_production@spigen.com','kbkim@spigen.com','gwlee@spigen.com','swpark@spigen.com','parksm@spigen.com','jihyunlee@spigen.com','yangsr@spigen.com','wanqiong.yang@spigen.com','reda.zou@spigen.com','mjlee@spigen.com','dykim@spigen.com','ebhwang@spigen.com','andy.lee@spigen.com','parkse@spigen.com','jspark@spigen.com','yjlee@spigen.com','wikim@spigen.com','jsong@spigen.com','jokim@spigen.com','hychae@spigen.com','ynchoi@spigen.com','sungjun@spigen.com','jaeseoung@spigen.com','shwoo@spigen.com'].join(',');
  const SUBJECT = '3 of 41 Bi-Weekly 제품 이슈 Report, 26.10.02 ~ 26.10.15';

  const inline = {};
  PAGES.forEach(function(pid, i) {
    const t = Slides.Presentations.Pages.getThumbnail(DECK, pid, { 'thumbnailProperties.thumbnailSize': 'LARGE' });
    inline['s' + i] = UrlFetchApp.fetch(t.contentUrl).getBlob().setName('slide' + (i + 1) + '.png');
  });

  const F = "-apple-system,BlinkMacSystemFont,'SF Pro Text','Apple SD Gothic Neo','Noto Sans KR','Malgun Gothic',Arial,sans-serif";
  const INK = '#1D1D1F', GRAY = '#6E6E73', BLUE = '#0071E3', RED = '#FF3B30', CANVAS = '#F5F5F7';
  const tile = function(inner, pad) { return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#FFFFFF;border-radius:18px;margin:0 0 14px 0"><tr><td style="padding:' + (pad || '22px 26px') + '">' + inner + '</td></tr></table>'; };
  const label = function(t) { return '<div style="font:600 12px ' + F + ';color:' + GRAY + ';letter-spacing:.2px;margin:0 0 10px 0">' + t + '</div>'; };
  const img = function(i) { return '<img src="cid:s' + i + '" width="628" style="display:block;width:100%;max-width:628px;height:auto;border-radius:12px;border:1px solid #E5E5EA">'; };
  const stat = function(n, l, c) { return '<td width="33%" style="background:#FFFFFF;border-radius:18px;padding:18px 18px;text-align:left"><div style="font:600 26px ' + F + ';color:' + (c || INK) + ';line-height:1.1">' + n + '</div><div style="font:400 12px ' + F + ';color:' + GRAY + ';margin-top:6px">' + l + '</div></td>'; };
  const voc = function(prod, issue) { return '<tr><td style="padding:6px 0;font:400 14px ' + F + ';color:' + INK + '">• ' + prod + ' <b style="color:' + RED + '">' + issue + '</b></td></tr>'; };

  let html = '<div style="background:' + CANVAS + ';padding:28px 0;margin:0"><table role="presentation" align="center" width="680" cellpadding="0" cellspacing="0" style="width:680px;max-width:100%;font-family:' + F + '"><tr><td style="padding:0 16px">';
  // hero: report title + link on top (박기세님 피드백: 링크를 상단으로)
  html += tile(
    '<div style="font:600 12px ' + F + ';color:' + GRAY + ';letter-spacing:.3px">GCX Bi-weekly Report · 글로벌CX전략팀</div>' +
    '<div style="font:700 30px ' + F + ';color:' + INK + ';margin:8px 0 6px;line-height:1.15"><a href="' + LINK + '" style="color:' + INK + ';text-decoration:none">Bi-Weekly Report 26. 10. 16</a></div>' +
    '<div style="font:400 14px ' + F + ';color:' + GRAY + ';margin:0 0 18px">26.10.02 ~ 26.10.15 · iPhone 18 · Pixel 11 · Galaxy Z8</div>' +
    '<a href="' + LINK + '" style="display:inline-block;background:' + BLUE + ';color:#FFFFFF;font:600 14px ' + F + ';text-decoration:none;padding:11px 22px;border-radius:980px">리포트 슬라이드 열기 &rsaquo;</a>', '30px 30px');
  // greeting + summary
  html += tile(
    '<div style="font:400 14px ' + F + ';color:' + INK + ';line-height:1.7">안녕하세요.<br>글로벌CX전략팀 김지우입니다.<br><br>' +
    '지난번 보고에 이어 <b>26.10.02 ~ 26.10.15</b> 중 취합한, 새로 출시된 iPhone 18 Series, Galaxy Z8 Series, 그리고 Pixel 11 Series 제품들에 대한 이슈들 및 SIREN 등록한 이슈들 위주로 작성했습니다.</div>');
  // VOC highlights
  html += tile(label('이번 모니터링 기간 중 상당량 인입된 VOC') + '<table role="presentation" cellpadding="0" cellspacing="0">' +
    voc('iPhone 18 Series용 Glas.tR EZ Fit Slim', '글라스 깨짐 · 측면 들뜸 · 버블') +
    voc('iPhone 18 Series용 Glas.tR EZ Fit Pro', '측면 들뜸') +
    voc('Galaxy Z Fold 8용 Air Skin Aramid MagFit', '카메라링 변형') + '</table>');
  // SIREN this period
  html += tile(label('금주 신규 SIREN 등록') + '<table role="presentation" cellpadding="0" cellspacing="0">' +
    voc('10.13 · Tesla Model Y용 전면 카메라 보호 커버 (TO331)', '결로') +
    voc('10.14 · Galaxy Z Fold 8용 Air Skin Aramid MagFit', '카메라링 변형') + '</table>');
  // stat tiles
  html += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:0 0 14px 0;border-collapse:separate;border-spacing:0"><tr>' +
    stat('23,999', '2026 전제품 누적 클레임 (~10.14)') + '<td width="10"></td>' + stat('-2.01%', 'YoY', BLUE) + '<td width="10"></td>' + stat('42', '2026 SIREN 등록 (금주 +2)', RED) + '</tr></table>';
  // overview images
  let ov = label('Overview');
  for (let i = 0; i < 8; i++) ov += '<div style="margin:0 0 12px 0">' + img(i) + '</div>';
  html += tile(ov, '22px 26px 12px');
  let sr = label('GCX SIREN · 26년 하반기 등록 현황');
  for (let i = 8; i < 10; i++) sr += '<div style="margin:0 0 12px 0">' + img(i) + '</div>';
  html += tile(sr, '22px 26px 12px');
  // closing
  html += tile('<div style="font:400 14px ' + F + ';color:' + INK + ';line-height:1.7">자세한 내용은 상단 <a href="' + LINK + '" style="color:#0066CC;text-decoration:none;font-weight:600">리포트 슬라이드</a>에서 확인을 부탁 드립니다. 감사합니다!</div>');
  html += '</td></tr></table></div>';
  // signature (same as previous mail)
  html += '<div style="font-family:Arial,sans-serif;font-size:13px;color:#222;margin-top:12px">-- <br><b>김지우</b> 글로벌CX전략팀 | 담당 | <b>M</b> +82.10.2872.2066 | <b>E</b> <a href="mailto:kjw@spigen.com">kjw@spigen.com</a><br><img src="https://ci3.googleusercontent.com/mail-sig/AIorK4w-54OvB6nZ_GMuFrUspHQCAearahPbswha8CuA--8GlqZa_pDQBfoPfmf8VV4vi_6D3v88AkK-1htz"></div>';

  const plain = '안녕하세요.\n글로벌CX전략팀 김지우입니다.\n\nBi-Weekly Report 26. 10. 16 (' + LINK + ')\n지난번 보고에 이어 26.10.02 ~ 26.10.15 중 취합한 이슈 및 SIREN 등록 이슈 위주로 작성했습니다.';
  const d = GmailApp.createDraft(TO, SUBJECT, plain, { htmlBody: html, inlineImages: inline, cc: CC });
  Logger.log('draft ' + d.getId());
}
