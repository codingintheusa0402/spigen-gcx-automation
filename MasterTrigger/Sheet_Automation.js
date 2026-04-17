function dedupeSheetByReviewId_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (values.length < 2) return;

  const header = values[0].map(h => String(h).trim());
  const reviewIdCol = header.findIndex(h => h.toLowerCase() === 'reviewid');
  if (reviewIdCol === -1) return;

  const seen = new Set();
  const output = [values[0]];

  for (let i = 1; i < values.length; i++) {
    const rid = String(values[i][reviewIdCol] || '').trim();
    if (!rid || seen.has(rid)) continue;
    seen.add(rid);
    output.push(values[i]);
  }

  sheet.clearContents();
  sheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

  const TARGET_ID = "1tMbA_msRfCRY0KK40GnyZ_h1uNCldlnk9Cg-_MTcbsw";             
  const TARGET_SHEET = "tem";                                                   
   
  const COLUMN_MAP = [                                                          
    { col: "SDA",       srcId: "1sxapIqJgXcJdeqyCf9bAxCNXrVMsVjsZE9QWPwEm0R4",
  srcSheet: "1-3점" },                                                          
    { col: "Auto_Acc",  srcId: "1mEYb1b92D6BIOaSYkAnMit6THuw5ewtymhA-mSIVDfs",
  srcSheet: "1-3점" },
  { col: "유지훈P",  srcId: "1mEYb1b92D6BIOaSYkAnMit6THuw5ewtymhA-mSIVDfs",
  srcSheet: "1-3점" },                                      
    { col: "전략폰",    srcId: "1yo8CbLhJkuxrf3eXbAqZCb6qBejZhSR3YOt7nFv97fw",
  srcSheet: "1-3점" },                                                          
    { col: "Power_Acc", srcId: "1QC8Is6UvTnFXaOeXviKM_331i3Fo_CBIYx80VS696LI",
  srcSheet: "1-3점" },                                                          
    { col: "Pixel 10a", srcId: "1BpeGq5gIr4tNsPZmnHr19NNY6pQ6sb2_H-v3V9-It4E",
  srcSheet: "1-5점" },                                                          
    { col: "Glx26",     srcId: "1fpv9TEDPGR8D6QRRc0ll-WzF7sOkfxe9UNBCmdBSE9g",
  srcSheet: "1-5점" },                                                          
    { col: "iPh17e",    srcId: "16xRJHH7Ynii4erNOn_905ST4CZs6OLpOYTof4uqsGsQ",
  srcSheet: "1-5점" },                                                          
  ];                                                        
                                                                                
  function updateTemSheet() {                               
    const tgtSs = SpreadsheetApp.openById(TARGET_ID);
    const tgtWs = tgtSs.getSheetByName(TARGET_SHEET);
    const tgtHeaders = tgtWs.getRange(1, 1, 1,                                  
  tgtWs.getLastColumn()).getValues()[0];                                        
                                                                                
    for (const { col, srcId, srcSheet } of COLUMN_MAP) {                        
      try {                                                 
        const srcWs = SpreadsheetApp.openById(srcId).getSheetByName(srcSheet);
        const srcHeaders = srcWs.getRange(1, 1, 1,
  srcWs.getLastColumn()).getValues()[0];                                        
  
        const reviewColIdx = srcHeaders.indexOf("Review ID");                   
        if (reviewColIdx === -1) throw new Error("'Review ID' column not found");                                                                      
  
        const lastRow = srcWs.getLastRow();                                     
        if (lastRow < 2) { Logger.log(`${col}: no data rows`); continue; }
        const reviewIds = srcWs.getRange(2, reviewColIdx + 1, lastRow - 1,      
  1).getValues();
                                                                                
        const tgtColIdx = tgtHeaders.indexOf(col);                              
        if (tgtColIdx === -1) throw new Error(`'${col}' not found in tem`);
                                                                                
        // Clear old values then write new ones                                 
        const existingRows = tgtWs.getLastRow();
        if (existingRows > 1) {                                                 
          tgtWs.getRange(2, tgtColIdx + 1, existingRows - 1, 1).clearContent();
        }                                                                       
        tgtWs.getRange(2, tgtColIdx + 1, reviewIds.length,
  1).setValues(reviewIds);                                                      
                                                            
        Logger.log(`✓ ${col}: ${reviewIds.length} IDs`);                        
      } catch (e) {
        Logger.log(`✗ ${col}: ${e.message}`);                                   
      }                                                     
    }
    Logger.log("Done.");
  }                        