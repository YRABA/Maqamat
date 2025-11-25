/** Post-processing: מיון + מחשב "סך השיעורים שנותרו" */
const PostProc = (() => {

  function buildGroupTargets_(courses) {
    const t = new Map();
    if (!courses) return t;
    for (const r of courses) {
      const teacher = Utils.str(r['שם המורה']);
      const course  = Utils.str(r['שם הקורס']);
      const weeksRaw = r["מס' שבועות"];
      const weeks = Number(weeksRaw);
      if (!teacher || !course || !isFinite(weeks) || weeks <= 0) continue;
      const k = `${teacher}|${course}`;
      t.set(k, Math.max(weeks, t.get(k) || 0));
    }
    return t;
  }

  function buildPrivateTargets_(privRows) {
    const t = new Map();
    if (!privRows) return t;
    for (const r of privRows) {
      const teacher = Utils.str(r['שם המורה']);
      const student = Utils.str(r['שם התלמיד']);
      const raw = (r['מספר שיעורים בשנה'] ?? r[' מספר שיעורים בשנה']);
      const total = Number(raw);
      if (!teacher || !student || !isFinite(total) || total <= 0) continue;
      const k = `${teacher}|${student}`;
      t.set(k, Math.max(total, t.get(k) || 0));
    }
    return t;
  }

  function updateRemainingCounters(sheetOut, headers, courses, privRows, logRows = [], runId = '') {
    if (!Array.isArray(logRows)) {
      logRows = [];
    }

    runId = runId || `RUN-${Date.now()}`;
    if (!Array.isArray(logRows)) logRows = [];

    LogSvc.info(runId, 'UPDATE_REMAIN_START', {
      sheet: sheetOut.getName(),
      totalRows: sheetOut.getLastRow() - 1,
      coursesCount: courses.length,
      privRowsCount: privRows.length
    }, logRows);

    const idx = Utils.indexMap(headers);
    const data = sheetOut.getDataRange().getValues();
    const totalRows = data.length - 1;
    if (totalRows < 1) {
      LogSvc.info(runId, 'UPDATE_REMAIN_NO_ROWS', { totalRows }, logRows);
      return;
    }

    const groupTargets   = buildGroupTargets_(courses);
    const privateTargets = buildPrivateTargets_(privRows);

    // Build minimal row info
    const rows = new Array(totalRows);
    for (let r = 1; r < data.length; r++) {
      const v = data[r];
      const type    = Utils.str(v[idx['סוג הדיווח']]);
      const teacher = Utils.str(v[idx['שם המורה']]);
      const course  = Utils.str(v[idx['שם הקורס']]);
      const student = Utils.str(v[idx['שם התלמיד']]);
      const date    = Utils.coerceToDate(v[idx['תאריך השיעור']]);
      const qty     = Number(v[idx['כמות']]) || 0;
      rows[r-1] = { rIdx: r, type, teacher, course, student, date, hasDate: !!date, qty };
    }

    rows.sort((a,b) => {
      // sort by date then teacher then key
      if (!a.hasDate && b.hasDate) return 1;
      if (a.hasDate && !b.hasDate) return -1;
      const aTime = a.date ? a.date.getTime() : 0;
      const bTime = b.date ? b.date.getTime() : 0;
      if (aTime !== bTime) return aTime - bTime;
      const kt = (a.teacher||'').localeCompare(b.teacher||'');
      if (kt !== 0) return kt;
      const aKey = a.type === 'פרטי' ? (a.student||'') : (a.course||'');
      const bKey = b.type === 'פרטי' ? (b.student||'') : (b.course||'');
      return aKey.localeCompare(bKey);
    });

    const outRemain = new Array(totalRows).fill('');
    const msgWrites = [];

    const groupCum   = new Map();
    const privateCum = new Map();

    for (const row of rows) {
      let remain = '';
      if (row.type === 'קבוצתי') {
        const key = `${row.teacher}|${row.course}`;
        const target = groupTargets.get(key);
        const nextCum = (groupCum.get(key) || 0) + (row.hasDate ? 1 : 0);
        groupCum.set(key, nextCum);
        if (isFinite(target)) remain = target - nextCum;
      } else if (row.type === 'פרטי') {
        const key = `${row.teacher}|${row.student}`;
        const target = privateTargets.get(key);
        const nextCum = (privateCum.get(key) || 0) + row.qty;
        privateCum.set(key, nextCum);
        if (isFinite(target)) remain = target - nextCum;
      }

      if (remain !== '') {
        outRemain[row.rIdx - 1] = remain;
        if (remain < 0) {
          msgWrites.push({
            row: row.rIdx + 1,
            msg: `❗ חריגה מהמכסה השנתית: עודף של ${Math.abs(remain)} שיעורים.`
          });
        }
      }
    }

    // Bulk write remain values
    const remainCol = idx['סך השיעורים שנותרו'];
    sheetOut.getRange(2, remainCol + 1, outRemain.length, 1)
        .setValues(outRemain.map(v => [v]));

    // Bulk apply msg writes & backgrounds
    const msgCol = idx['הודעת מערכת'];
    if (msgWrites.length) {
      const bgRange = sheetOut.getRange(2, 1, totalRows, headers.length);
      const bgColors = bgRange.getBackgrounds();
      for (const w of msgWrites) {
        const i = w.row - 2;
        // set message value
        sheetOut.getRange(w.row, msgCol + 1).setValue(w.msg);
        // paint background red-ish
        bgColors[i].fill('#f8d7da');
      }
      sheetOut.getRange(2, 1, totalRows, headers.length).setBackgrounds(bgColors);
    }
    Logger.log(`[${runId}] logRows type: ${typeof logRows}, isArray: ${Array.isArray(logRows)}, runId: ${runId}`);
    LogSvc.debug(runId, 'UPDATE_REMAIN_COMPLETE', {
      totalRows,
      rowsWithRemain: totalRows,
      overQuota: msgWrites.length
    }, logRows);

    // Optional: show log sheet or return counts
  }


  function markNearAnnualQuota(sheetOut, headers, privRows) {
    const idx = Utils.indexMap(headers);
    const TYPE = idx['סוג הדיווח'];
    const DATE = idx['תאריך השיעור'];
    const REM  = idx['סך השיעורים שנותרו'];
    const MSG  = idx['הודעת מערכת'];

    const values = sheetOut.getDataRange().getValues();
    if (values.length < 2) return;

    const msgWrites = [];
    const rangesToYellow = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (Utils.str(row[TYPE]) !== 'פרטי') continue;
      if (!Utils.coerceToDate(row[DATE])) continue;

      const remain = Number(row[REM]);
      if (!isFinite(remain)) continue;

      // Disabled on request: marking quota near
    }

    for (const u of msgWrites) sheetOut.getRange(u.rowIndex, MSG + 1).setValue(u.text);
    if (rangesToYellow.length) sheetOut.getRangeList(rangesToYellow).setBackground('#FFFF00');

    SpreadsheetApp.flush();
  }

  function sortReportSheet(sheetOut, headers) {
    const idx = Utils.indexMap(headers);
    let monthIdx = idx['חודש תשלום'];
    if (typeof monthIdx === 'undefined') monthIdx = idx['חודש-שנה'];
    const teacherIdx = idx['שם המורה'];
    const typeIdx    = idx['סוג הדיווח'];
    const dateIdx    = idx['תאריך השיעור'];

    const lastRow = sheetOut.getLastRow();
    const lastCol = headers.length;
    if (lastRow <= 1) return;

    if (typeof monthIdx !== 'undefined') sheetOut.getRange(2, monthIdx+1, lastRow-1, 1).setNumberFormat('MM-yyyy');
    if (typeof dateIdx  !== 'undefined') sheetOut.getRange(2, dateIdx+1,  lastRow-1, 1).setNumberFormat('dd/MM/yyyy');

    const range = sheetOut.getRange(2, 1, lastRow - 1, lastCol);
    const spec = [];
    if (typeof monthIdx  !== 'undefined') spec.push({column: monthIdx+1,  ascending: true});
    if (typeof teacherIdx!== 'undefined') spec.push({column: teacherIdx+1,ascending: true});
    if (typeof typeIdx   !== 'undefined') spec.push({column: typeIdx+1,   ascending: false});
    if (typeof dateIdx   !== 'undefined') spec.push({column: dateIdx+1,   ascending: true});
    if (spec.length) range.sort(spec);
  }

  function sortNewRows(sheetOut, headers, startRow, numRows) {
    if (!startRow || !numRows || numRows <= 1) return;

    const idx = Utils.indexMap(headers);
    let monthIdx = idx['חודש תשלום'];
    if (typeof monthIdx === 'undefined') monthIdx = idx['חודש-שנה'];
    const teacherIdx = idx['שם המורה'];
    const typeIdx    = idx['סוג הדיווח'];
    const dateIdx    = idx['תאריך השיעור'];

    const lastCol = headers.length;
    const range = sheetOut.getRange(startRow, 1, numRows, lastCol);

    if (typeof monthIdx !== 'undefined')
      sheetOut.getRange(startRow, monthIdx+1, numRows, 1).setNumberFormat('MM-yyyy');
    if (typeof dateIdx !== 'undefined')
      sheetOut.getRange(startRow, dateIdx+1,  numRows, 1).setNumberFormat('dd/MM/yyyy');

    const spec = [];
    if (typeof monthIdx  !== 'undefined') spec.push({column: monthIdx+1,  ascending: true});
    if (typeof teacherIdx!== 'undefined') spec.push({column: teacherIdx+1,ascending: true});
    if (typeof typeIdx   !== 'undefined') spec.push({column: typeIdx+1,   ascending: false});
    if (typeof dateIdx   !== 'undefined') spec.push({column: dateIdx+1,   ascending: true});
    if (spec.length) range.sort(spec);
  }

  function repaintPrivateBlankLessons(sheetOut, headers) {
    const idx = Utils.indexMap(headers);
    const typeCol    = idx['סוג הדיווח'] + 1;
    const dateCol    = idx['תאריך השיעור'] + 1;
    const msgCol     = idx['הודעת מערכת'] + 1;
    const counterCol = idx['סך השיעורים שנותרו'] + 1;

    const lastRow = sheetOut.getLastRow();
    if (lastRow <= 1) return;

    const vals = sheetOut.getRange(2, 1, lastRow - 1, headers.length).getValues();

    const toYellowDate   = [];
    const toYellowMsg    = [];
    const toYellowCount  = [];

    for (let i = 0; i < vals.length; i++) {
      const r = i + 2;
      const type = String(vals[i][typeCol-1] || '').trim();
      const dateV = vals[i][dateCol-1];
      const msgV  = String(vals[i][msgCol-1] || '').trim();
      const hasDate = !!Utils.coerceToDate(dateV);
      const isPrivate = (type === 'פרטי');
      const isException = (isPrivate && !hasDate && msgV === 'יש לעדכן תאריך שיעור');

      if (isException) {
        toYellowDate.push(sheetOut.getRange(r, dateCol).getA1Notation());
        toYellowMsg.push(sheetOut.getRange(r, msgCol).getA1Notation());
        toYellowCount.push(sheetOut.getRange(r, counterCol).getA1Notation());
      }
    }

    if (toYellowDate.length)  sheetOut.getRangeList(toYellowDate).setBackground('#FFFF00');
    if (toYellowMsg.length)   sheetOut.getRangeList(toYellowMsg).setBackground('#FFFF00');
    if (toYellowCount.length) sheetOut.getRangeList(toYellowCount).setBackground('#FFFF00');

    SpreadsheetApp.flush();
  }

  return {
    updateRemainingCounters,
    markNearAnnualQuota,
    sortReportSheet,
    sortNewRows,
    repaintPrivateBlankLessons,
  };
})();

function sortReportSheet(sheetOut, headers) {
  return PostProc.sortReportSheet(sheetOut, headers);
}
