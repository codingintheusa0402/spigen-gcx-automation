// ---------------------------Trigger Maker ----------------------------------------//
function triggerGen() {
  var endDate = new Date('2026-01-30'); // 종료일(2주치만 셋팅하기) 
  var date = new Date(); // 시작일
  date.setHours(9, 0, 0, 0); // 9:00 AM KST

  // Remove ALL existing triggers
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  var created = 0;
  while (date <= endDate) {
    var day = date.getDay(); // 0 = Sunday, 6 = Saturday
    if (day >= 1 && day <= 5) { // Only Mon–Fri
      ScriptApp.newTrigger('MCFReporter')
        .timeBased()
        .at(new Date(date)) // schedule for this exact date/time
        .create();
      Logger.log("Trigger set for: " + date);
      created++;
    }

    // Move to next day at 9:00AM
    date.setDate(date.getDate() + 1);
    date.setHours(9, 0, 0, 0);
  }

  Logger.log("Created " + created + " weekday 9AM triggers until end of 2025.08.");
}


// ------------Tester ------------//
function triggerTester() {
  // Delete existing triggers for MCFReporter to avoid duplicates (optional)
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'MCFReporter') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Get current time and add 1 minute
  var now = new Date();
  var oneMinuteLater = new Date(now.getTime() + 60 * 1000); // add 60,000 ms

  // Create the time-based trigger for MCFReporter
  ScriptApp.newTrigger('MCFReporter')
    .timeBased()
    .at(oneMinuteLater)
    .create();

  Logger.log("Trigger set for: " + oneMinuteLater);
}
