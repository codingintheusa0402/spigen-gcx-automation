/***** ===== CONFIG (Galaxy S26) ===== *****/

/* ===== SHEET ===== */
const UPLOAD_SHEET_ID = '1BpeGq5gIr4tNsPZmnHr19NNY6pQ6sb2_H-v3V9-It4E';
const UPLOAD_SHEET_NAME = '1-3점';

/***** ===== ENV ===== */
function _prop(key, fallback) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return (v == null || v === '') ? fallback : v;
}

// const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=PSHPzKIgMGy6kVu1yWvFiE332iPfDNZHaBPAB3MmcMs'; // Private test ❤️
const CHAT_WEBHOOK_URL = 'https://chat.googleapis.com/v1/spaces/AAQAFjxOPoY/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=MHegeUf86uuXeDIRL9E9zaFp2ASPwk-CVXhF3u3JJTo'; // TCK GCX Spigen


const CONFIG = {
  pollIntervalMinutes: 1,
  pollMaxMinutes: 180,
  timezone: 'Asia/Seoul'
};

function getSpreadsheetId_() {
  return SpreadsheetApp.getActive().getId();
}

// Preferred headers (used by _overwriteSheet)
const PREFERRED_HEADERS = [
  'country',
  'date',
  'variantAsin',
  'productAsin',
  'productOriginalAsin',
  'username',
  'ratingScore',
  'reviewTitle',
  'reviewDescription',
  'reviewId',
  'reviewImages/0',
  'reviewImages/1',
  'reviewImages/2',
  'reviewImages/3',
  'reviewImages/4',
  'reviewImages/5',
  'reviewImages/6',
  'reviewImages/7',
  'reviewImages/8',
  'reviewImages/9',
  'reviewImages/10',
  'reviewUrl',
  'product/productPageReviews/0/variant',
  'totalCategoryRatings',
  'totalCategoryReviews',
  'filterByRating',
  'variantAttributes',
  'averageCustomerReviews',
];

/* ===== GROUP MAPPING ===== */
const MODEL_HEADER_CANDIDATES = ['기종명', '모델명', 'Model', 'Model Name'];

const GROUP_TITLES = [
  'Galaxy S26',
  'Galaxy S26 Plus',
  'Galaxy S26 Ultra'
];