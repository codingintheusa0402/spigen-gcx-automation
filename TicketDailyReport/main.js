const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQAdqYt1ro/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=FAP-OGDvupbed1cL0xFWP9Bq6aXqOxlJvBhH4vfF-b4'; // Ticket T2 Chatroom 
// const webhookUrl = 'https://chat.googleapis.com/v1/spaces/AAQAc9NQmJQ/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=b4zApCmKNq1pPBDmemgVv1Y8xoXm4h_w_eKccjtqCiI'; //Private Chatroom (for testing)

// https://hcti.io/v1/image  --> API Pic uploading site (Check Available Tokens Left)

function runZendeskDailyJob() {
  fetchZendeskViewToSheet();            // Step 1: Call Zendesk API and fill 'Zendesk_Daily'
  fetchZendeskViewToKsheet();    //Step 1.5: Call Zendesk API and fill 'K_시트' B5:Gn table
  Utilities.sleep(600);               // Wait 6 seconds
  appendZendeskDailyStatus();          // Step 2: Count & update 'All_Graph' + clear 'Zendesk_Daily'
  all_GraphChartToGoogleChat();     //Step 3: Send to Google Chat using API sent img
  collapseOldRowsIfNeeded();          // Step 3.5: collapse rows if exceed 4 weeks
  fetchZendeskViewToKsheet();          // Step 4: make up table for the 'P_시트' sheet
  kSheetToChat();                       //Step 5. Send IMG of 'K_시트' to Google Chat via Webhook + imgAPI extension
}
