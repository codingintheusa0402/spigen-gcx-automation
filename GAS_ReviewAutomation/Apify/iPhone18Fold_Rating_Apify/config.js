/***** ===== CONFIG (iPhone 18 / iPhone Fold / Apple ETC(26) / Pixel 11 — Amazon.de Product Rating) ===== *****/

/* ===== SHEETS ===== */
// All four sheets share the same layout: A=SKU, B=ASIN, C=제품 페이지, D=Review, E=Rating.
const SHEETS = [
  { name: 'iPhone 18', asinCol: 2, ratingCol: 5, dataStartRow: 6 },
  { name: 'iPhone Fold', asinCol: 2, ratingCol: 5, dataStartRow: 6 },
  { name: 'Apple ETC(26)', asinCol: 2, ratingCol: 5, dataStartRow: 6 },
  { name: 'Pixel 11', asinCol: 2, ratingCol: 5, dataStartRow: 6 }
];

/* ===== APIFY TASK ===== */
// "product-details-scraper-iphone18-fold" task (axesso_data/amazon-product-details-scraper
// actor — same actor as SKUSales_Rating_Apify), pre-loaded with amazon.de URLs for every
// ASIN across all sheets above. syncNewAsinsToApifyTask() (run before every rating run,
// plus the weekly trigger as a backstop) keeps this
// URL list in sync as ASINs are added to any of the sheets.
const TASK_ID = 'CLQ3G6Sokyr7AJtQi';

/* ===== BEHAVIOR ===== */
const CONFIG = {
  pollIntervalMinutes: 1, // ScriptApp.everyMinutes() only accepts 1/5/10/15/30
  pollMaxMinutes: 180,
  timezone: 'Asia/Seoul',
  dailyHour: 8, // kickoff hour, skipped on Sat/Sun — see dailyWeekdayKickoff()
  asinSyncWeekDay: 'SUNDAY', // ScriptApp.WeekDay name — must run before the next weekday kickoff
  asinSyncHour: 7
};

function getSpreadsheetId_() {
  return SpreadsheetApp.getActive().getId();
}

function _getToken() {
  const token = PropertiesService.getScriptProperties().getProperty('APIFY_TOKEN');
  if (!token) {
    throw new Error('APIFY_TOKEN is not set in Script Properties (Project Settings → Script Properties).');
  }
  return token;
}
