/** Orchestrator: זרימת הריצה הראשית */
const App = (() => {

  function ensureStatusValidation(sheet, headers) {
    const idx = {}; headers.forEach((h,i)=> idx[h] = i);
    const statusCol = idx['סטטוס'] + 1;
    const lastRow   = Math.max(sheet.getLastRow(), 2);

    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([
        'דווח-טרם שולם',
        'שולם - אסור לערוך שינויים',
        'הועבר לתשלום - אסור לערוך שינויים'
      ], true)
      .setAllowInvalid(false)
      .build();

    sheet.getRange(2, statusCol, lastRow - 1).setDataValidation(rule);
  }

  function processReport(monthYear, mode, runScope) {
    const runId = "RUN-" + new Date().getTime();
    const tz    = 'Asia/Jerusalem';
    const now   = new Date();
    const nowStr = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm:ss');

    const ss           = SpreadsheetApp.getActive();
    const sheetOut     = SheetsSvc.getOrCreateSheet(ss, 'דיווח שיעורים');
    const sheetCourses = ss.getSheetByName('רשימת קורסים-מערכת');
    const sheetPriv    = ss.getSheetByName('ריכוז שיעורים פרטיים');
    const sheetFilter  = SheetsSvc.getOrCreateSheet(ss, 'חריגים-קבוצתי');
    const sheetLog     = SheetsSvc.getOrCreateSheet(ss, 'לוג ריצות');
    const sheetStatus  = SheetsSvc.getOrCreateSheet(ss, 'סטטוס');

    const [yearStr, monthStr] = String(monthYear).split("-");
    const selectedYear  = parseInt(yearStr, 10);
    const selectedMonth = parseInt(monthStr, 10);

    const HMONTHS = ["ינואר","פברואר","מרץ","אפריל","מאי","יוני","יולי","אוגוסט","ספטמבר","אוקטובר","נובמבר","דצמבר"];
    const monthYearText = `${HMONTHS[selectedMonth - 1]} ${selectedYear}`;
    const monthPayCell  = ('0' + selectedMonth).slice(-2) + '-' + selectedYear;

    const HEADERS = [
      'שם המורה','סוג הדיווח','שם הקורס','שם התלמיד','שנה',
      'חודש תשלום','תאריך השיעור','כמות','סטטוס','הערות',
      'סך השיעורים שנותרו','מועד עדכון','הודעת מערכת'
    ];
    const FILTER_HEADERS = ['שם המורה','תאריך פעיל','תאריך לא פעיל מ','תאריך לא פעיל עד','הודעת מערכת'];

    SheetsSvc.ensureHeader(sheetOut, HEADERS);
    SheetsSvc.ensureHeader(sheetFilter, FILTER_HEADERS);
    SheetsSvc.alignSheetRight(sheetOut);
    SheetsSvc.alignSheetRight(sheetFilter);
    sheetOut.setRightToLeft(true);
    sheetFilter.setRightToLeft(true);

    const numRows = Math.max(1, sheetOut.getMaxRows() - 1);
    const formats = Array(numRows).fill([
      'General','General','General','General','General',
      'MM-yyyy','dd/MM/yyyy','General','General','General',
      'General','dd/MM/yyyy HH:mm','General'
    ]);
    sheetOut.getRange(2, 1, numRows, HEADERS.length).setNumberFormats(formats);

    ensureStatusValidation(sheetOut, HEADERS);

    const logRows     = [];
    const deletedCount = 0;

    let protection = null;
    if (mode === 'reset') {
      protection = sheetOut.protect().setDescription('Locked during run');
      try { protection.removeEditors(protection.getEditors()); } catch(e) {}
    }

    try {
      LogSvc.info(runId, 'START', { month: selectedMonth, year: selectedYear, mode, runScope }, logRows);

      if (!sheetCourses) throw new Error('לא נמצא גיליון "רשימת קורסים-מערכת".');

      let deleted = 0;
      if (mode === 'reset') {
        deleted = RowsSvc.deleteMonthRows(sheetOut, HEADERS, selectedYear, selectedMonth);
        LogSvc.info(runId, 'RESET deleted month rows', { deleted }, logRows);
      }

      const courses     = SheetsSvc.getRowsAsObjects(sheetCourses);
      const privRows    = sheetPriv ? SheetsSvc.getRowsAsObjects(sheetPriv) : [];
      const filter      = ExceptionsSvc.parseFilterSheet(sheetFilter, FILTER_HEADERS, logRows, runId);
      const existingMap = RowsSvc.getExistingMap(sheetOut, HEADERS);
      const rowsToInsert= [];
      const styleMarks  = { msgYellow: [], dateYellow: [] };
      const travelPerTeacherDate = new Set();

      if (runScope === 'both' || runScope === 'group') {
        LogSvc.startTimer('group');
        GroupSvc.processCourses({
          courses, filter, selectedYear, selectedMonth,
          existingMap, mode, rowsToInsert, travelPerTeacherDate,
          updateStamp: nowStr, monthPayCell, sheetOut, logRows, runId
        });
        GroupSvc.processExceptions({
          filter, courses, selectedYear, selectedMonth,
          existingMap, mode, rowsToInsert, travelPerTeacherDate,
          updateStamp: nowStr, monthPayCell, sheetOut, sheetFilter,
          FILTER_HEADERS, logRows, runId
        });
        LogSvc.endTimer('group', runId, 'Group processing completed', {}, logRows);
      }

      if ((runScope === 'both' || runScope === 'private') && sheetPriv && privRows.length > 0) {
        LogSvc.startTimer('private');
        PrivateSvc.processPrivateRows({
          sheetOut, HEADERS, privRows, selectedYear, selectedMonth,
          filter, rowsToInsert, existingMap, travelPerTeacherDate,
          styleMarks, monthYearCell: monthPayCell, updateStamp: nowStr,
          logRows, runId
        });
        LogSvc.endTimer('private', runId, 'Private processing completed', {}, logRows);
      }

      let insertedCount = 0;
      let startRow = null;

      if (rowsToInsert.length > 0) {
        startRow = sheetOut.getLastRow() + 1;

        rowsToInsert.forEach((row, i) => {
          const diff = HEADERS.length - row.length;
          if (diff > 0) row.push(...Array(diff).fill(''));
          else if (diff < 0) rowsToInsert[i] = row.slice(0, HEADERS.length);
        });

        sheetOut.getRange(startRow, 1, rowsToInsert.length, HEADERS.length).setValues(rowsToInsert);
        insertedCount = rowsToInsert.length;

        LogSvc.info(runId, 'INSERT', { inserted: insertedCount }, logRows);
      }

      if (insertedCount > 0) {
        const highlightRow = sheetOut.getLastRow() - insertedCount + 1;
        applyHighlights(sheetOut, HEADERS, highlightRow, styleMarks);
        PostProc.sortNewRows(sheetOut, HEADERS, highlightRow, insertedCount);
        ProtectSvc.applyLocksForNewRows(sheetOut, HEADERS, startRow, insertedCount);  // ✅ FIXED
      }

      PostProc.updateRemainingCounters(sheetOut, HEADERS, courses, privRows, logRows, runId);
      highlightFilterMessageCells(sheetFilter, FILTER_HEADERS);

      SheetsSvc.ensureHeader(sheetStatus, ['Run ID','תאריך','חודש/שנה','סטטוס','שגיאה']);
      sheetStatus.appendRow([runId, nowStr, monthYearText, 'הצלחה', '']);
      sheetStatus.hideSheet();
      sheetLog.hideSheet();

      LogSvc.success(runId, 'COMPLETE', { inserted: insertedCount, deleted }, logRows);
      return `העדכון בוצע בהצלחה ✅\nנוספו ${insertedCount} שורות${mode==='reset' ? `, נמחקו ${deleted}` : ''}.`;

    } catch (err) {
      const msg = `${err.message} (App.processReport)`;
      LogSvc.error(runId, msg, {}, logRows);
      sheetStatus.appendRow([runId, nowStr, monthYearText, 'שגיאה', err.message]);
      sheetStatus.showSheet();
      sheetLog.showSheet();
      return `❌ שגיאה: ${err.message}`;
    } finally {
      try { if (protection) protection.remove(); } catch(e) {}
      if (logRows.length > 0) {
        SheetsSvc.ensureHeader(sheetLog, ['Run ID','זמן','רמה','הודעה','נתונים']);
        const startLogRow = sheetLog.getLastRow() + 1;
        sheetLog.getRange(startLogRow, 1, logRows.length, logRows[0].length).setValues(logRows);
        sheetLog.getRange(startLogRow, 2, logRows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
      }
      Logger.log(`[App.processReport] Logging complete for run ${runId}`);
    }
  }

  function applyHighlights(sheetOut, HEADERS, startRow, styleMarks) {
    const idx = {};
    HEADERS.forEach((h,i)=> idx[h]=i);
    const dateCol    = idx['תאריך השיעור'] + 1;
    const msgCol     = idx['הודעת מערכת'] + 1;
    const counterCol = idx['סך השיעורים שנותרו'] + 1;

    (styleMarks.dateYellow || []).forEach(rel => {
      sheetOut.getRange(startRow + rel - 1, dateCol).setBackground('#FFFF00');
    });

    (styleMarks.msgYellow || []).forEach(rel => {
      sheetOut.getRange(startRow + rel - 1, msgCol).setBackground('#FFFF00');
    });

    const dateSet = new Set(styleMarks.dateYellow || []);
    const msgSet  = new Set(styleMarks.msgYellow  || []);
    for (const rel of dateSet) {
      if (msgSet.has(rel)) {
        sheetOut.getRange(startRow + rel - 1, counterCol).setBackground('#FFFF00');
      }
    }
  }

  function highlightFilterMessageCells(sheetFilter, FILTER_HEADERS) {
    if (!sheetFilter) return;
    const idx = {};
    FILTER_HEADERS.forEach((h,i)=> idx[h]=i);
    const msgCol  = idx['הודעת מערכת'] + 1;
    const lastRow = sheetFilter.getLastRow();
    if (lastRow <= 1 || msgCol <= 0) return;

    const values = sheetFilter.getRange(2, msgCol, lastRow - 1, 1).getValues();
    const bgs    = sheetFilter.getRange(2, msgCol, lastRow - 1, 1).getBackgrounds();

    const toYellow = [];
    for (let i = 0; i < values.length; i++) {
      const hasText  = String(values[i][0] ?? '').trim() !== '';
      const isYellow = String(bgs[i][0] || '').toUpperCase() === '#FFFF00';
      if (hasText && !isYellow) {
        toYellow.push(sheetFilter.getRange(2 + i, msgCol).getA1Notation());
      }
    }
    if (toYellow.length) {
      sheetFilter.getRangeList(toYellow).setBackground('#FFFF00');
    }
  }

  return { processReport };
})();
