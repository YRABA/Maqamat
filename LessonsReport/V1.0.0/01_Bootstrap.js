/** Bootstrap + UI glue **/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('עדכון דיווחים')
    .addItem('עדכן שיעורים', 'showSidebar')
    .addItem('עדכון סטטוס (גורף)', 'showBulkStatusSidebar')
    .addSeparator()
    .addToUi();

  ensureInstallableOnEditTrigger_();
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'דיווח שיעורים') return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idx = Utils.indexMap(headers);

    const startRow = e.range.getRow();
    const startCol = e.range.getColumn();
    const numRows = e.range.getNumRows();
    const numCols = e.range.getNumColumns();

    if (startRow < 2) return; // skip header

    const statusCol = idx['סטטוס'] + 1;
    const dateCol = idx['תאריך השיעור'] + 1;
    const msgCol = idx['הודעת מערכת'] + 1;

    const isStatusCol = (startCol === statusCol && numCols === 1);
    const isDateCol = (startCol === dateCol && numCols === 1);
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('לוג ריצות');
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // ✅ Status column logic (supports multi-row edits)
    if (isStatusCol) {
      const values = e.range.getValues();

      for (let i = 0; i < numRows; i++) {
        const row = startRow + i;
        const newVal = String(values[i][0] || '').trim();

        // Toggle locking
        ProtectSvc.toggleRowLockForRow(sheet, headers, row);

        if (newVal === 'דווח-טרם שולם') {
          // Set white background
          const bgRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
          bgRange.setBackgrounds([Array(sheet.getLastColumn()).fill('#ffffff')]);

          // ✅ Log unlocked row
          if (logSheet) {
            SheetsSvc.ensureHeader(logSheet, ['Timestamp', 'Row', 'Action', 'Status']);
            logSheet.appendRow([timestamp, row, 'Unlocked manually', newVal]);
          }
        } else if (ProtectSvc.isLockedStatus(newVal)) {
          // ✅ Log locked row
          if (logSheet) {
            SheetsSvc.ensureHeader(logSheet, ['Timestamp', 'Row', 'Action', 'Status']);
            logSheet.appendRow([timestamp, row, 'Locked manually', newVal]);
          }
        }
      }

      // Show one toast
      sheet.toast('סטטוס עודכן. שורות טופלו בהתאם.', 'הודעת מערכת', 3);
    }

    // ✅ תאריך השיעור logic: insert/remove warning
    if (isDateCol) {
      const dateValues = e.range.getValues();

      for (let i = 0; i < numRows; i++) {
        const row = startRow + i;
        const val = dateValues[i][0];
        const msgCell = sheet.getRange(row, msgCol);
        const currentMsg = String(msgCell.getValue() || '').trim();

        if (val instanceof Date) {
          // Remove warning if present
          if (currentMsg === 'יש לעדכן תאריך שיעור') {
            msgCell.setValue('');
          }
        } else {
          // Add warning if not already present
          if (currentMsg !== 'יש לעדכן תאריך שיעור') {
            msgCell.setValue('יש לעדכן תאריך שיעור');
          }
        }
      }
    }

    // ✅ Block changes in locked rows for non-status fields
    for (let r = 0; r < numRows; r++) {
      const row = startRow + r;
      for (let c = 0; c < numCols; c++) {
        const col = startCol + c;

        if (col !== statusCol) {
          const statusVal = String(sheet.getRange(row, statusCol).getValue() || '').trim();
          const isLocked = ProtectSvc.isLockedStatus(statusVal);
          if (isLocked) {
            const targetCell = e.range.getCell(r + 1, c + 1);

            // Try to restore the original value
            let originalValue;
            if (e.oldValues && Array.isArray(e.oldValues)) {
              originalValue = e.oldValues[r]?.[c];
            } else if (e.oldValue !== undefined) {
              originalValue = e.oldValue;
            }

            if (typeof originalValue !== 'undefined') {
              targetCell.setValue(originalValue);
            } else {
              targetCell.clearContent();
            }

            // Restore formatting
            if (col === dateCol)
              sheet.getRange(row, col).setNumberFormat('dd/MM/yyyy');
            if (col === idx['חודש תשלום'] + 1)
              sheet.getRange(row, col).setNumberFormat('MM-yyyy');
          }
        }
      }
    }

  } catch (err) {
    Logger.log('onEdit error: ' + err);
    LogSvc?.error?.(null, 'onEdit error', { error: err.message });
  }
}

function onEditShowLock(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (sheet.getName() !== 'דיווח שיעורים') return;

    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    const idx     = Utils.indexMap(headers);
    const row     = e.range.getRow(), col = e.range.getColumn();
    if (row < 2) return;

    const isStatusCol = (col === idx['סטטוס'] + 1);
    const statusVal  = String(sheet.getRange(row, idx['סטטוס'] + 1).getValue() || '').trim();
    const isLocked   = ProtectSvc.isLockedStatus(statusVal);

    if (!isStatusCol && isLocked) {
      showLockDialog(statusVal);
    }

  } catch (err) {
    Logger.log('onEditShowLock error: ' + err);
    LogSvc?.error?.(null, 'onEditShowLock error', { error: err.message });
  }
}

function ensureInstallableOnEditTrigger_() {
  const HANDLER = 'onEditShowLock';
  const triggers = ScriptApp.getProjectTriggers();
  const exists   = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === HANDLER);
  if (!exists) {
    ScriptApp.newTrigger(HANDLER)
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();
  }
}

function enableLockDialogs() {
  ensureInstallableOnEditTrigger_();
  showLockDialog('שולם - אסור לערוך שינויים');
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('עדכון שיעורים');
  SpreadsheetApp.getUi().showSidebar(html);
}

function showBulkStatusSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('BulkStatusSidebar')
    .setTitle('עדכון סטטוס (גורף)');
  SpreadsheetApp.getUi().showSidebar(html);
}

function processReport(monthYear, mode, runScope) {
  mode = mode || 'skip';
  runScope = runScope || 'both';
  PropertiesService.getUserProperties().setProperty('RUN_SCOPE_LAST', runScope);
  return App.processReport(monthYear, mode, runScope);
}

function getLastRunScope() {
  return PropertiesService.getUserProperties().getProperty('RUN_SCOPE_LAST') || 'both';
}

function showLockDialog(statusText) {
  const title   = 'רשומה נעולה';
  const message = 'לא ניתן לבצע שינויים ברשומה הנמצאת בשלבי תשלום!';
  const html    = `<!doctype html>
    <html dir="rtl" lang="he">
    <head>
      <meta charset="utf-8">
      <style>
        body {
          margin: 0;
          padding: 0;
          background: #0b0f19;
          color: #e5e7eb;
          font-family: "Rubik", "Segoe UI", system-ui, sans-serif;
        }
        .wrap {
          width: 100%;
          height: 100%;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
          box-sizing: border-box;
        }
        .card {
          background: #111827;
          border: 1px solid rgba(255,255,255,0.1);
          border-radius: 12px;
          box-shadow: 0 6px 24px rgba(0,0,0,0.4);
          max-width: 440px;
          width: 100%;
          text-align: center;
          padding: 24px 28px;
          position: relative;
        }
        .icon {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 56px;
          height: 56px;
          background: rgba(245,158,11,0.15);
          border-radius: 50%;
          margin: 0 auto 16px auto;
        }
        .icon svg {
          width: 28px;
          height: 28px;
          stroke: #fbbf24;
          fill: none;
          stroke-width: 2;
        }
        .title {
          font-size: 18px;
          font-weight: 700;
          margin-bottom: 6px;
        }
        .status {
          font-size: 14px;
          color: #9ca3af;
          margin-bottom: 12px;
        }
        .msg {
          font-size: 15px;
          line-height: 1.6;
          color: #e5e7eb;
          margin-bottom: 22px;
        }
        .btns {
          display: flex;
          justify-content: center;
          gap: 12px;
        }
        .btn {
          all: unset;
          padding: 10px 20px;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          text-align: center;
        }
        .btn.ok {
          background: #f59e0b;
          color: #1f2937;
        }
        .btn.close {
          background: transparent;
          color: #9ca3af;
        }
        .btn.close:hover {
          color: #ffffff;
        }
        .close-icon {
          position: absolute;
          top: 12px;
          right: 12px;
          width: 20px;
          height: 20px;
          cursor: pointer;
          color: #9ca3af;
        }
        .close-icon:hover {
          color: #fff;
        }
      </style>
    </head>
    <body>
      <div class="wrap">
        <div class="card">
          <div class="close-icon" onclick="google.script.host.close()">×</div>
          <div class="icon" title="נעול">
            <svg viewBox="0 0 24 24">
              <path d="M7 10V8a5 5 0 0 1 10 0v2" stroke-linecap="round"/>
              <rect x="5" y="10" width="14" height="11" rx="2.5"/>
              <circle cx="12" cy="15" r="1.6" fill="#fbbf24"/>
            </svg>
          </div>
          <div class="title">${ title }</div>
          <div class="status">סטטוס נוכחי: <strong>${ _escapeHtml(statusText || '—') }</strong></div>
          <div class="msg">${ message }</div>
          <div class="btns">
            <button class="btn ok" onclick="google.script.host.close()">הבנתי</button>
            <button class="btn close" onclick="google.script.host.close()">סגור</button>
          </div>
        </div>
      </div>
    </body>
    </html>`;
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(320), title);
}

/** Utility functions */
function _parseDDMMYYYY(s) {
  if (s instanceof Date) return s;
  if (typeof s !== 'string') return null;
  const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  d.setHours(0,0,0,0);
  return isNaN(d.getTime()) ? null : d;
}

function _parseMMYYYY(s) {
  if (s instanceof Date) return s;
  if (typeof s !== 'string') return null;
  const m = s.match(/^(\d{1,2})-(\d{4})$/);
  if (!m) return null;
  const d = new Date(Number(m[2]), Number(m[1]) - 1, 1);
  d.setHours(0,0,0,0);
  return isNaN(d.getTime()) ? null : d;
}

function _serialToDate(n) {
  if (typeof n !== 'number' || !isFinite(n)) return null;
  const ms = Math.round(n * 86400000) - 2209161600000;
  const d  = new Date(ms);
  d.setHours(0,0,0,0);
  return isNaN(d.getTime()) ? null : d;
}

function _restoreOldDateValue(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return _serialToDate(v);
  if (typeof v === 'string') return _parseDDMMYYYY(v) || v;
  return v;
}

function _restoreOldMonthYearValue(v) {
  if (v instanceof Date) return v;
  if (typeof v === 'number') return _serialToDate(v);
  if (typeof v === 'string') return _parseMMYYYY(v) || v;
  return v;
}

function _escapeHtml(s) {
  return String(s)
    .replace(/&/g,  '&amp;')
    .replace(/</g,  '&lt;')
    .replace(/>/g,  '&gt;')
    .replace(/"/g,  '&quot;')
    .replace(/'/g,  '&#039;');
}
