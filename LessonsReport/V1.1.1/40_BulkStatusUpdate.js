/** BulkStatusUpdate: עדכון סטטוס חודשי ורישום לנעילת שורות */
function applyStatusForMonth(monthYear, newStatus, startRow = 2, batchSize = 50) {
  const runId = "RUN-BULK-" + new Date().getTime();
  const tz = 'Asia/Jerusalem';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דיווח שיעורים');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ הגיליון "דיווח שיעורים" לא נמצא.');
    LogSvc.error(runId, 'Sheet "דיווח שיעורים" not found', {}, []); // no logRows passed, but could be extended
    return 0;
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusIdx = headers.indexOf('סטטוס');
  const monthIdx  = headers.indexOf('חודש תשלום');
  const msgIdx    = headers.indexOf('הודעת מערכת');

  if (statusIdx < 0 || monthIdx < 0 || msgIdx < 0) {
    SpreadsheetApp.getUi().alert('⚠️ עמודות חסרות: סטטוס / חודש תשלום / הודעת מערכת');
    LogSvc.error(runId, 'Missing required columns', { statusIdx, monthIdx, msgIdx }, []);
    return 0;
  }

  const lastRow = sheet.getLastRow();
  let updatedCount = 0;

  // Normalize target month-year once
  const normalizedTarget = String(monthYear).replace(/\u00A0/g, ' ').trim();

  LogSvc.info(runId, 'Starting bulk status update', { monthYear: normalizedTarget, newStatus, lastRow, batchSize });

  for (let row = startRow; row <= lastRow; row += batchSize) {
    const endRow = Math.min(lastRow, row + batchSize - 1);
    const numRows = endRow - row + 1;

    const rangeMonth  = sheet.getRange(row, monthIdx  + 1, numRows);
    const rangeStatus = sheet.getRange(row, statusIdx + 1, numRows);
    const rangeMsg    = sheet.getRange(row, msgIdx    + 1, numRows);

    const monthVals  = rangeMonth.getValues();
    const statusVals = rangeStatus.getValues();
    const msgVals    = rangeMsg.getValues();

    for (let i = 0; i < numRows; i++) {
      const absoluteRow = row + i;
      const cellVal     = monthVals[i][0];
      let curMonth = '';

      if (cellVal instanceof Date) {
        const m = cellVal.getMonth() + 1;
        const y = cellVal.getFullYear();
        curMonth = (m < 10 ? '0' + m : m) + '-' + y;
      } else {
        curMonth = String(cellVal).replace(/\u00A0/g, ' ').trim();
      }

      if (curMonth === normalizedTarget) {
        statusVals[i][0] = newStatus;
        msgVals[i][0] = (newStatus === 'שולם - אסור לערוך שינויים'
                        ? '✅ שורה נעולה - שולם. 🔒 נעול לעריכה'
                        : (newStatus === 'הועבר לתשלום - אסור לערוך שינויים'
                          ? '✅ שורה נעולה - הועבר לתשלום. 🔒 נעול לעריכה'
                          : '') );

        ProtectSvc.toggleRowLockForRow(sheet, headers, absoluteRow);
        updatedCount++;
      }
    }

    rangeStatus.setValues(statusVals);
    rangeMsg.setValues(msgVals);
    SpreadsheetApp.flush();
    // Optional sleep removed for faster performance; keep only if necessary:
    // Utilities.sleep(300);
  }

  LogSvc.success(runId, 'Bulk update completed', { updatedCount });

  SpreadsheetApp.getUi().alert(`✅ בוצעו ${updatedCount} עדכונים.`);
  return updatedCount;
}
