function createTriggers() {
  const endDate = new Date('2026-04-26');
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
