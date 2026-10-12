function createTriggers() {
  const endDate = new Date('2026-10-20');
  let date = new Date(); // start from today
  date.setHours(9, 0, 0, 0); // 9:00AM in KST

  // 🧹 Delete all existing runZendeskDailyJob triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'runZendeskDailyJob') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  let created = 0;
  while (date <= endDate) {
    const day = date.getDay(); // 0 = Sunday, 6 = Saturday
    if (day >= 1 && day <= 5) { // Only Mon–Fri
      ScriptApp.newTrigger('runZendeskDailyJob')
        .timeBased()
        .at(new Date(date)) // schedule trigger for this exact 9AM KST
        .create();
      Logger.log(`Trigger set for: ${date}`);
      created++;
    }

    // Move to next day at 9:00AM
    date.setDate(date.getDate() + 1);
    date.setHours(9, 0, 0, 0);
  }

  Logger.log(`Created ${created} weekday 9AM triggers until end of 2025.08.`);
}

/**
 * Self-perpetuating trigger scheduler for runZendeskDailyJob.
 * Run setupAutoExtendZendeskTrigger() once — after that this checks itself
 * daily and rolls the schedule forward 30 days whenever fewer than 3 days
 * remain, forever, until someone manually deletes the installed trigger.
 */
function autoExtendZendeskTriggers() {
  const HANDLER = 'runZendeskDailyJob';
  const PROP_KEY = 'ZENDESK_TRIGGER_SCHEDULE_END';
  const EXTEND_DAYS = 30;
  const THRESHOLD_DAYS = 3;

  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const storedEnd = props.getProperty(PROP_KEY);
  const currentEnd = storedEnd ? new Date(storedEnd) : null;

  if (currentEnd) {
    const daysLeft = (currentEnd - now) / 86400000;
    if (daysLeft >= THRESHOLD_DAYS) {
      Logger.log(`[autoExtendZendeskTriggers] ${daysLeft.toFixed(1)} days remain (schedule ends ${currentEnd}). No extension needed.`);
      return;
    }
  }

  const endDate = new Date(now);
  endDate.setDate(endDate.getDate() + EXTEND_DAYS);
  endDate.setHours(23, 59, 59, 999);

  // Wipe + fully rebuild the window so re-runs can never create duplicates.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === HANDLER) ScriptApp.deleteTrigger(t);
  });

  let created = 0;
  let date = new Date(now);
  date.setHours(9, 0, 0, 0);
  while (date <= endDate) {
    const day = date.getDay(); // 0 = Sunday, 6 = Saturday
    if (day >= 1 && day <= 5 && date > now) { // Mon-Fri, future only
      ScriptApp.newTrigger(HANDLER).timeBased().at(new Date(date)).create();
      created++;
    }
    date.setDate(date.getDate() + 1);
    date.setHours(9, 0, 0, 0);
  }

  props.setProperty(PROP_KEY, endDate.toISOString());
  Logger.log(`[autoExtendZendeskTriggers] Rebuilt schedule: ${created} weekday 9AM triggers through ${endDate}.`);
}

/**
 * Run this ONCE manually (or via GAS editor ▶ Run) to install the daily
 * self-check trigger. Safe to re-run — clears any prior copy first so no
 * duplicate installer trigger is ever created.
 */
function setupAutoExtendZendeskTrigger() {
  const HANDLER = 'autoExtendZendeskTriggers';
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === HANDLER) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(HANDLER).timeBased().everyDays(1).atHour(6).create();
  Logger.log('[setupAutoExtendZendeskTrigger] Daily 6AM self-check installed — runs forever until manually removed.');
}
