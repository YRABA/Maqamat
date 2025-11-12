const GroupSvc = (() => {
  const hebDayMap = { 'א': 0, 'ב': 1, 'ג': 2, 'ד': 3, 'ה': 4, 'ו': 5, 'ש': 6 };

  function processCourses(params) {
    const {
      courses, filter, selectedYear, selectedMonth,
      existingMap, mode, rowsToInsert, travelPerTeacherDate,
      updateStamp, monthPayCell, sheetOut,
      logRows, runId
    } = params;

    let lessonsInserted = 0;
    let travelsInserted = 0;

    Logger.log(`[GroupSvc.processCourses] Starting. Courses: ${courses.length}`);
    for (const rec of courses) {
      const teacher = Utils.str(rec['שם המורה']);
      const course = Utils.str(rec['שם הקורס']);
      const dayLet = Utils.str(rec['יום']);
      const hours = Number(rec['מס ש"ש'] || 1);
      const yearVal = Utils.str(rec['שנה']) || String(selectedYear);

      if (!teacher || !course || !dayLet || filter.teachersIgnoreWeekly.has(teacher)) continue;

      const dayIndex = hebDayMap[dayLet];
      if (dayIndex === undefined) continue;

      const dates = RowsSvc.getDatesInMonth(selectedYear, selectedMonth, dayIndex);
      for (const d of dates) {
        const dateKey = Utils.toDateKey(d);
        if (d.getMonth() + 1 !== selectedMonth || d.getFullYear() !== selectedYear) continue;
        if (ExceptionsSvc.isGloballyFiltered(dateKey, filter)) continue;
        if (ExceptionsSvc.isTeacherFiltered(teacher, dateKey, filter)) continue;

        const keyGroup = RowsSvc.keyFor(teacher, 'קבוצתי', course, '', dateKey);
        const rowGroup = [teacher, 'קבוצתי', course, '', yearVal, monthPayCell, d, hours, 'דווח-טרם שולם', '', '', updateStamp, ''];
        RowsSvc.handleRow(sheetOut, existingMap, keyGroup, rowGroup, mode, rowsToInsert);
        lessonsInserted++;

        const travelKey = `${teacher}|${dateKey}`;
        if (!travelPerTeacherDate.has(travelKey)) {
          const keyTravel = RowsSvc.keyFor(teacher, 'יום נסיעות', '', '', dateKey);
          const rowTravel = [teacher, 'יום נסיעות', '', '', '', monthPayCell, d, 1, 'דווח-טרם שולם', '', '', updateStamp, ''];
          RowsSvc.handleRow(sheetOut, existingMap, keyTravel, rowTravel, mode, rowsToInsert);
          travelPerTeacherDate.add(travelKey);
          travelsInserted++;
        }
      }
    }

    Logger.log(`[GroupSvc.processCourses] Lessons: ${lessonsInserted}, Travels: ${travelsInserted}`);
    if ((lessonsInserted + travelsInserted) > 0) {
      LogSvc.info(runId, 'Group lessons inserted', {
        lessonsInserted, travelsInserted
      }, logRows);
    }
  }

  function processExceptions(params) {
    const {
      filter, courses, selectedYear, selectedMonth,
      existingMap, mode, rowsToInsert, travelPerTeacherDate,
      updateStamp, monthPayCell, sheetOut,
      sheetFilter, FILTER_HEADERS, logRows, runId
    } = params;

    let added = 0, travels = 0, skipped = 0;

    Logger.log(`[GroupSvc.processExceptions] Rows: ${filter.activeRows.length}`);
    for (const ar of filter.activeRows) {
      const { rowIndex, teacher, date } = ar;
      if (!teacher || !date) continue;
      if (date.getFullYear() !== selectedYear || date.getMonth() + 1 !== selectedMonth) continue;

      const dateKey = Utils.toDateKey(date);
      if (ExceptionsSvc.isGloballyFiltered(dateKey, filter) || ExceptionsSvc.isTeacherFiltered(teacher, dateKey, filter)) {
        ExceptionsSvc.setFilterMessage(sheetFilter, FILTER_HEADERS, rowIndex, 'לא ניתן להוסיף רשומה מאחר והתאריך מסונן');
        skipped++;
        continue;
      }

      const teacherCourses = courses.filter(c => Utils.str(c['שם המורה']) === teacher);
      for (const tc of teacherCourses) {
        const courseName = Utils.str(tc['שם הקורס']);
        const yearVal = Utils.str(tc['שנה']) || String(selectedYear);
        const hours = Number(tc['מס ש"ש'] || 1);
        const keyGroup = RowsSvc.keyFor(teacher, 'קבוצתי', courseName, '', dateKey);
        const rowGroup = [teacher, 'קבוצתי', courseName, '', yearVal, monthPayCell, date, hours, 'דווח-טרם שולם', '', '', updateStamp, ''];
        RowsSvc.handleRow(sheetOut, existingMap, keyGroup, rowGroup, mode, rowsToInsert);
        added++;
      }

      const travelKey = `${teacher}|${dateKey}`;
      if (!travelPerTeacherDate.has(travelKey)) {
        const keyTravel = RowsSvc.keyFor(teacher, 'יום נסיעות', '', '', dateKey);
        const rowTravel = [teacher, 'יום נסיעות', '', '', '', monthPayCell, date, 1, 'דווח-טרם שולם', '', '', updateStamp, ''];
        RowsSvc.handleRow(sheetOut, existingMap, keyTravel, rowTravel, mode, rowsToInsert);
        travelPerTeacherDate.add(travelKey);
        travels++;
      }
    }

    Logger.log(`[GroupSvc.processExceptions] Added: ${added}, Travels: ${travels}, Skipped: ${skipped}`);
    if ((added + travels) > 0) {
      LogSvc.info(runId, 'Exception group lessons inserted', {
        added, travels, skipped
      }, logRows);
    }
  }

  return { processCourses, processExceptions };
})();
