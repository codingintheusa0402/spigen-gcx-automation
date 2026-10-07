function diagCheckAug12PlusRows() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(BF_SHEET_NAME);
        var lastRow = sheet.getLastRow();
          var numRows = lastRow - BF_START_ROW + 1;

            var sents = sheet.getRange(BF_START_ROW, BF_COL_SENT, numRows, 1).getValues();
              var fees  = sheet.getRange(BF_START_ROW, BF_COL_FEE, numRows, 1).getValues();
                var notes = sheet.getRange(BF_START_ROW, BF_COL_FEE, numRows, 1).getNotes();
                  var orders = sheet.getRange(BF_START_ROW, BF_COL_ORDER, numRows, 1).getValues();
                    var track = sheet.getRange(BF_START_ROW, RETRY_R_COL, numRows, 1).getDisplayValues();

                      var cutoff = new Date('2026-08-12T00:00:00');
                        var out = [];
                          for (var i = 0; i < numRows; i++) {
                              var sentStr = String(sents[i][0]||'').trim();
                                  if (!sentStr) continue;
                                      var d = new Date(sentStr);
                                          if (isNaN(d.getTime()) || d < cutoff) continue;

                                              var feeStr = String(fees[i][0]===null||fees[i][0]===undefined?'':fees[i][0]).trim();
                                                  var isEst = _isFeeEstimateNote(notes[i][0]);
                                                      var hasTrack = String(track[i][0]||'').trim() !== '' && !_isErrorValue(String(track[i][0]||''));
                                                          out.push((BF_START_ROW+i) + '|' + orders[i][0] + '|sent=' + sentStr + '|fee=' + (feeStr||'(blank)') + '|est=' + isEst + '|hasTracking=' + hasTrack);
                                                            }
                                                              Logger.log('rows sent Aug 12+ (%s total):\n%s', out.length, out.join('\n'));
                                                              }
