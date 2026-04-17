function createEscT2Triggers() {
  const functionName = 'sendDailyEscT2';

  // 1) Remove all existing triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 2) Base dates
  const tz = Session.getScriptTimeZone();
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate()); 
  const endDate = new Date('2026-4-26'); // adjust as needed

  // 3) Create new triggers 
  for (let d = new Date(today); d <= endDate; d.setDate(d.getDate() + 1)) {
    const day = d.getDay(); // 0 = Sun, 6 = Sat
    if (day === 0 || day === 6) continue; // skip weekends

    // Set trigger time
    const triggerTime = new Date(d);
    if (day === 4) {
      triggerTime.setHours(15, 30, 0, 0); // Thursday → 15:30
    } else {
      triggerTime.setHours(17, 30, 0, 0); // Other weekdays → 17:30
    }

    // Skip if time already passed today
    if (triggerTime <= now) continue;

    ScriptApp.newTrigger(functionName)
      .timeBased()
      .at(triggerTime)
      .create();
  }
}
