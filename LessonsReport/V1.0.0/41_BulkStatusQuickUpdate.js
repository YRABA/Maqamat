/**
 * BulkStatusQuickUpdate: עדכון מהיר של סטטוס חודשי ורישום לנעילה / שחרור שורות
 */
/**
 * BulkStatusQuickUpdate: עדכון מהיר של סטטוס חודשי ורישום לנעילה / שחרור שורות
 */
function applyStatusForMonthQuick(monthYear, newStatus) {
  const runId = 'RUN-QUICK-' + new Date().getTime();
  const tz = 'Asia/Jerusalem';
  const now = new Date();
  const nowStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('דיווח שיעורים');
  const sheetStatus = ss.getSheetByName('סטטוס');
  const sheetLog = ss.getSheetByName('לוג ריצות');
  const logRows = [];

  if (!sheet) throw new Error('לא נמצא הגיליון "דיווח שיעורים"');

  logRows.push([runId, new Date(), 'INFO', 'START', { monthYear, newStatus }]);

  // 1. שחרור הגנות קיימות
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
       .filter(p => p.canEdit()).forEach(p => p.remove());
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
       .filter(p => p.canEdit()).forEach(p => p.remove());

  // 2. קריאת נתונים
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return 0;

  const headers = values[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);

  const monthCol = idx['חודש תשלום'];
  const statusCol = idx['סטטוס'];
  const msgCol = idx['הודעת מערכת'];

  if ([monthCol, statusCol, msgCol].includes(undefined)) {
    throw new Error('חסרה אחת מהעמודות: "חודש תשלום", "סטטוס", "הודעת מערכת"');
  }

  // 3. זיהוי שורות לעדכון
  const updates = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const currMonth = formatMonthYear(row[monthCol]);
    const currStatus = String(row[statusCol] || '').trim();
    if (currMonth === monthYear && currStatus !== newStatus) {
      updates.push(r + 1);
    }
  }

  if (updates.length === 0) {
    logRows.push([runId, new Date(), 'INFO', 'No updates needed', {}]);
    return 0;
  }

  // 4. קריאת טווחים
  const statusRange = sheet.getRange(2, statusCol + 1, values.length - 1, 1);
  const msgRange = sheet.getRange(2, msgCol + 1, values.length - 1, 1);
  const bgRange = sheet.getRange(2, 1, values.length - 1, sheet.getLastColumn());

  const statusVals = statusRange.getValues();
  const msgVals = msgRange.getValues();
  const bgColors = bgRange.getBackgrounds();

  // 5. עדכון בפועל
  updates.forEach(rowNum => {
    const i = rowNum - 2;
    statusVals[i][0] = newStatus;

    if (['שולם - אסור לערוך שינויים', 'הועבר לתשלום - אסור לערוך שינויים'].includes(newStatus)) {
      msgVals[i][0] = '🔒 נעול לעריכה';
      bgColors[i] = bgColors[i].map(() => '#e0e0e0');
    } else {
      msgVals[i][0] = '';
      bgColors[i] = bgColors[i].map(() => '#ffffff');
    }
  });

  // 6. כתיבה
  statusRange.setValues(statusVals);
  msgRange.setValues(msgVals);
  bgRange.setBackgrounds(bgColors);

  // 7. לוג לסטטוס
  if (sheetStatus) {
    SheetsSvc.ensureHeader(sheetStatus, ['Run ID','תאריך','חודש/שנה','סטטוס','שגיאה']);
    sheetStatus.appendRow([runId, nowStr, monthYear, 'הצלחה', `עודכנו ${updates.length} שורות`]);
    sheetStatus.hideSheet();
  }

  // 8. לוג מלא
  logRows.push([runId, new Date(), 'SUCCESS', 'Bulk status updated', {
    updatedRows: updates.length,
    newStatus,
    monthYear
  }]);

  if (sheetLog && logRows.length) {
    SheetsSvc.ensureHeader(sheetLog, ['Run ID','זמן','רמה','הודעה','נתונים']);
    const startRow = sheetLog.getLastRow() + 1;
    sheetLog.getRange(startRow, 1, logRows.length, 5).setValues(logRows);
    sheetLog.getRange(startRow, 2, logRows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sheetLog.hideSheet();
  }

  Logger.log(`✅ ${runId}: עודכנו ${updates.length} שורות.`);
  return updates.length;
}

/**
 * Helper: פורמט תאריך לתצורה MM-yyyy
 */
function formatMonthYear(val) {
  let d;
  if (val instanceof Date) {
    d = val;
  } else if (typeof val === 'number') {
    d = new Date(Math.round(val * 86400000) + new Date('1899-12-30').getTime());
  } else if (typeof val === 'string') {
    const clean = val.replace(/\u00A0/g, ' ').trim();
    const m = clean.match(/^(\d{1,2})[-\/](\d{4})$/);
    if (m) return `${m[1].padStart(2, '0')}-${m[2]}`;
    return clean;
  } else {
    return '';
  }
  return `${('0' + (d.getMonth() + 1)).slice(-2)}-${d.getFullYear()}`;
}
