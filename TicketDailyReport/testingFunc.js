// // Test wrapper that injects a test date
// function test_fetchZendeskViewToKsheet(testDateStr) {
//   // Convert to Date
//   const testDate = new Date("2025-08-14"); // e.g., "2025-08-12"
//   if (isNaN(testDate)) {
//     throw new Error("Invalid date format. Use YYYY-MM-DD");
//   }
//   fetchZendeskViewToKsheet(testDate);
// }

// // Modified main function to accept an optional date parameter
// function fetchZendeskViewToKsheet(testDate) {
//   const email = 'kjw@spigen.com';
//   const token = 'QhM2AiBYwTZTSb04Qjor918PHtttxp8xAzCFfFsg';
//   const subdomain = 'spigenhelp';
//   const viewId = '49523632520985';

//   const sheet = SpreadsheetApp.openById('10VYnysCGztKWMXfvXIWBVcE2_zENnRxvXUr9nicHkpo')
//     .getSheetByName('K_시트');

//   const lastRow = sheet.getLastRow();
//   if (lastRow >= 5) {
//     sheet.getRange(5, 3, lastRow - 4, 5).clearContent();
//   }

//   const headers = {
//     "Authorization": "Basic " + Utilities.base64Encode(email + "/token:" + token)
//   };

//   const url = `https://${subdomain}.zendesk.com/api/v2/views/${viewId}/tickets.json`;
//   const response = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
//   const json = JSON.parse(response.getContentText());
//   const tickets = json.tickets;

//   if (!tickets || tickets.length === 0) {
//     Logger.log("No tickets found in view.");
//     return;
//   }

//   // Formatting helpers
//   const formatBrandCode = raw => ({
//     'spigen_case_': 'Spigen(CASE)',
//     'spigen_steinheil_': 'Spigen(Steinheil)',
//     'spigen_odm_': 'Spigen(ODM)',
//     'spigen_pacc._': 'Spigen(PAcc.)',
//     'spigen_new_biz_': 'Spigen(New Biz)',
//     'n/a': 'n/a'
//   }[raw] || raw);

//   const formatCategoryCode = raw => ({
//     '1._invoice': '1. Invoice',
//     '2._delivery': '2. Delivery',
//     '3._exchange': '3. Exchange',
//     '4._issue': '4. Product Issue',
//     '5._fbm': '5. FBM',
//     '6._product_inquiry': '6. Product Inquiry',
//     '7._other_inquiry': '7. Other Inquiry',
//     '8._문의_사항_파악_불가': '8. 문의 사항 파악 불가'
//   }[raw] || raw);

//   // PIC logic (accepts injected date)
//   const getPIC = (country, dateOverride) => {
//     const groupA = ['DE', 'FR', 'IT', 'UK', 'ES'];
//     const baseDate = dateOverride || new Date();
//     const weekNumber = Math.floor((baseDate - new Date(baseDate.getFullYear(), 0, 1)) / (7 * 24 * 60 * 60 * 1000));
//     console.log(baseDate,weekNumber);
//     const assignKJWThisWeek = weekNumber % 2 === 0;
//     if(groupA.includes(country)){
//       return assignKJWThisWeek ? 'LYS' : 'KJW'
//     }
//     return groupA.includes(country) ? (assignKJWThisWeek ? 'LYS' : 'KJW') : (assignKJWThisWeek ? 'KJW' : 'LYS');
//   };

//   // Build rows
//   const rawRows = tickets
//     .filter(t => t.status.toLowerCase() === 'pending')
//     .map(t => {
//       const customFields = t.custom_fields || [];
//       const brandRaw = customFields.find(f => f.id === 5495572594201)?.value || '';
//       const countryRaw = customFields.find(f => f.id === 4513936822297)?.value || '';
//       const categoryRaw = customFields.find(f => f.id === 900006613446)?.value || '';

//       const brand = formatBrandCode(brandRaw);
//       const country = countryRaw.toUpperCase();
//       const category = formatCategoryCode(categoryRaw);

//       return [country, brand, category];
//     });

//   // Group and count
//   const grouped = {};
//   rawRows.forEach(([country, brand, category]) => {
//     const key = `${country}|${brand}|${category}`;
//     grouped[key] = (grouped[key] || 0) + 1;
//   });

//   const finalRows = Object.entries(grouped).map(([key, count]) => {
//     const [country, brand, category] = key.split('|');
//     return [country, brand, category, count, getPIC(country, testDate)];
//   });

//   // Write to sheet
//   sheet.getRange(5, 3, finalRows.length, 5).setValues(finalRows);
//   sheet.getRange(5, 2, finalRows.length).setValues(finalRows.map((_, i) => [i + 1]));

//   // Update B3 with test date or today
//   const now = testDate || new Date();
//   const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd");
//   const dayName = now.toLocaleDateString('en-US', { weekday: 'short' });
//   sheet.getRange('B3').setValue(`${dateStr} (${dayName}) T2 업무시작 Pending Ticket 수`);
// }

// function testPIC(dateStr, country) {
//   const testDateKST = new Date(`${dateStr}T00:00:00+09:00`);
//   const groupA = ['DE', 'FR', 'IT', 'UK', 'ES'];
//   const baseDateKST = new Date('2025-08-11T00:00:00+09:00');
//   const msPerWeek = 7 * 24 * 60 * 60 * 1000;
//   const weekIndex = Math.floor((testDateKST - baseDateKST) / msPerWeek);
//   const assignKJW = weekIndex % 2 === 0;
//   const pic = groupA.includes(country) ? (assignKJW ? 'KJW' : 'LYS') : (assignKJW ? 'LYS' : 'KJW');
//   Logger.log(`${dateStr} → ${country} → ${pic}`);
// }

// testPIC('2025-08-27', 'DE')
