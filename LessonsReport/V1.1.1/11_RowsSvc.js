/** RowsSvc: בניית/איתור/כתיבת רשומות */
if (typeof RowsSvc === 'undefined') {
  var RowsSvc = (function () {

    function keyFor(teacher, type, course, student, dateKey) {
      return `${teacher}|${type}|${course}|${student}|${dateKey || ''}`;
    }

    function getExistingMap(sheet, headers) {
      const map = new Map();
      const values = sheet.getDataRange().getValues();
      if (values.length < 2) return map;

      const idx = {}; headers.forEach((h, i) => idx[h] = i);
      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const teacher = Utils.str(row[idx['שם המורה']]);
        const type    = Utils.str(row[idx['סוג הדיווח']]);
        const course  = Utils.str(row[idx['שם הקורס']]);
        const student = Utils.str(row[idx['שם התלמיד']]);
        const d       = Utils.coerceToDate(row[idx['תאריך השיעור']]);
        const dateKey = d ? Utils.toDateKey(d) : '';
        const key = keyFor(teacher, type, course, student, dateKey);
        map.set(key, { rowIndex: r + 1, values: row });
      }

      Logger.log(`[RowsSvc.getExistingMap] Loaded ${map.size} rows`);
      return map;
    }

    function handleRow(sheet, existingMap, key, rowValues, mode, rowsToInsert) {
      if (existingMap.has(key)) {
        if (mode === 'overwrite' && sheet && existingMap.get(key).rowIndex) {
          const ex = existingMap.get(key);
          Logger.log(`[RowsSvc.handleRow] Overwriting row ${ex.rowIndex} for key ${key}`);
          sheet.getRange(ex.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
          existingMap.set(key, { rowIndex: ex.rowIndex, values: rowValues });
        }
        // skip/reset: don't duplicate
      } else {
        rowsToInsert.push(rowValues);
        existingMap.set(key, { rowIndex: null, values: rowValues });
        Logger.log(`[RowsSvc.handleRow] Queued new row for key ${key}`);
      }
    }

    function handleRowNoDate(sheet, existingMap, key, rowValues, rowsToInsert, mode) {
      if (existingMap && existingMap.has(key)) {
        if (mode === 'overwrite' && sheet && existingMap.get(key).rowIndex) {
          const ex = existingMap.get(key);
          Logger.log(`[RowsSvc.handleRowNoDate] Overwriting row ${ex.rowIndex} for key ${key}`);
          sheet.getRange(ex.rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
          existingMap.set(key, { rowIndex: ex.rowIndex, values: rowValues });
        }
      } else {
        rowsToInsert.push(rowValues);
        if (existingMap) {
          existingMap.set(key, { rowIndex: null, values: rowValues });
          Logger.log(`[RowsSvc.handleRowNoDate] Queued new row (no date) for key ${key}`);
        }
      }
    }

    function deleteMonthRows(sheetOut, headers, year, month) {
      const values = sheetOut.getDataRange().getValues();
      if (values.length < 2) return 0;

      const idx = {}; headers.forEach((h, i) => idx[h] = i);
      const dateCol = idx['תאריך השיעור'];
      if (dateCol == null) return 0;

      const rowsToDelete = [];
      for (let r = 1; r < values.length; r++) {
        const d = Utils.coerceToDate(values[r][dateCol]);
        if (!d) continue;
        if (d.getFullYear() === year && (d.getMonth() + 1) === month) {
          rowsToDelete.push(r + 1);
        }
      }

      rowsToDelete.sort((a, b) => b - a).forEach(rowIdx => sheetOut.deleteRow(rowIdx));
      Logger.log(`[RowsSvc.deleteMonthRows] Deleted ${rowsToDelete.length} rows for ${month}-${year}`);
      return rowsToDelete.length;
    }

    function getDatesInMonth(year, month, dayIndex) {
      const out = [];
      const first = new Date(year, month - 1, 1, 12, 0, 0);
      const last  = new Date(year, month, 0, 12, 0, 0);
      const d = new Date(first);
      while (d <= last) {
        if (d.getDay() === dayIndex) {
          out.push(new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12, 0, 0));
        }
        d.setDate(d.getDate() + 1);
        d.setHours(12, 0, 0, 0);
      }
      return out;
    }

    return {
      keyFor,
      getExistingMap,
      handleRow,
      handleRowNoDate,
      deleteMonthRows,
      getDatesInMonth
    };
  })();
}
