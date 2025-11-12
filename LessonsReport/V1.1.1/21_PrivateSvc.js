/** PrivateSvc: טיפול בשיעורים פרטיים (גיליון "ריכוז שיעורים פרטיים") */
if (typeof PrivateSvc === 'undefined') {
  var PrivateSvc = (() => {

    function monthColName(headers) {
      return headers.indexOf('חודש תשלום') !== -1 ? 'חודש תשלום' : 'חודש-שנה';
    }

    function mmYYYY(v) {
      const d = Utils.coerceToDate(v);
      if (d) return ('0' + (d.getMonth() + 1)).slice(-2) + '-' + d.getFullYear();
      const s = (v == null) ? '' : String(v).trim();
      const m = s.match(/^(\d{1,2})-(\d{4})$/);
      if (m) return ('0' + m[1]).slice(-2) + '-' + m[2];
      return s;
    }

    function keyForPrivate(teacher, student, dateKey) {
      return `${teacher}|פרטי||${student}|${dateKey}`;
    }

    function countExistingAndPendingBlankPrivateByMonth(sheetOut, HEADERS, teacher, student, monthCell, rowsToInsert) {
      const idx = {}; HEADERS.forEach((h,i)=> idx[h]=i);
      const monthHeader = monthColName(HEADERS);
      const wantMY     = mmYYYY(monthCell);
      let cnt = 0;

      const values = sheetOut.getDataRange().getValues();
      if (values.length > 1) {
        for (let r = 1; r < values.length; r++) {
          const v = values[r];
          const isPrivate   = String(v[idx['סוג הדיווח']]||'').trim() === 'פרטי';
          const sameTeacher = String(v[idx['שם המורה']]||'').trim() === teacher;
          const sameStudent = String(v[idx['שם התלמיד']]||'').trim() === student;
          const sameMY      = mmYYYY(v[idx[monthHeader]]) === wantMY;
          const dateV       = v[idx['תאריך השיעור']];
          const isDateEmpty  = (dateV === '' || dateV == null || (Utils.coerceToDate && !Utils.coerceToDate(dateV)));
          if (isPrivate && sameTeacher && sameStudent && sameMY && isDateEmpty) cnt++;
        }
      }

      if (rowsToInsert && rowsToInsert.length) {
        for (const row of rowsToInsert) {
          const isPrivate   = String(row[idx['סוג הדיווח']]||'').trim() === 'פרטי';
          const sameTeacher = String(row[idx['שם המורה']]||'').trim() === teacher;
          const sameStudent = String(row[idx['שם התלמיד']]||'').trim() === student;
          const sameMY      = mmYYYY(row[idx[monthHeader]]) === wantMY;
          const dateV       = row[idx['תאריך השיעור']];
          const isDateEmpty  = (dateV === '' || dateV == null || (Utils.coerceToDate && !Utils.coerceToDate(dateV)));
          if (isPrivate && sameTeacher && sameStudent && sameMY && isDateEmpty) cnt++;
        }
      }
      return cnt;
    }

    function pushBlankPrivate(rowsToInsert, teacher, student, yearVal, qty, notes, monthCell, updateStamp, HEADERS) {
      const row = [
        teacher, 'פרטי', '', student, yearVal,
        mmYYYY(monthCell), '', qty,
        'דווח-טרם שולם', notes || '',
        '', updateStamp || '', 'יש לעדכן תאריך שיעור'
      ];
      rowsToInsert.push(row);
    }

    /**
     * תוספת חדשה: טיפול בחריגים לפרטי
     */
    /**
     * תוספת חדשה: טיפול בחריגים לפרטי עם fallback ללוגיקת תדירות אם אין תאריכים פעילים
     */
    function handlePrivateExceptions(normalized, sheetFilter, FILTER_HEADERS, sheetOut, existingMap, rowsToInsert, monthYearCell, updateStamp, filter, logRows, runId) {
      const exceptionTeachers = new Map();
      const handledWithDates  = new Set(); // חדש - כדי לזהות חריגים שטופלו ע"י תאריכים פעילים

      // 1. שלוף את כל המורים+תלמידים שיש להם חריג
      normalized.forEach(item => {
        if (String(item['חריג']).trim() === 'חריג' && item.teacher && item.student) {
          const key = `${item.teacher}|${item.student}`;
          exceptionTeachers.set(key, item);
        }
      });

      if (exceptionTeachers.size === 0) return handledWithDates;

      // 2. שלוף את גיליון "חריגים-כללי"
      const ss = SpreadsheetApp.getActive();
      const sheetGroupEx = ss.getSheetByName('חריגים-כללי');
      if (!sheetGroupEx) {
        Logger.log('[PrivateSvc.handlePrivateExceptions] גיליון חריגים-כללי לא נמצא');
        return handledWithDates;
      }

      const headers = sheetGroupEx.getRange(1,1,1,sheetGroupEx.getLastColumn()).getValues()[0];
      const idx = Utils.indexMap(headers);
      const values = sheetGroupEx.getDataRange().getValues();

      // 3. צור מפה של תאריכים פעילים לכל מורה
      const activeDatesByTeacher = new Map();
      for (let r=1; r<values.length; r++) {
        const teacher = Utils.str(values[r][idx['שם המורה']]);
        const active  = Utils.coerceToDate(values[r][idx['תאריך פעיל']]);
        if (!teacher || !active) continue;
        const arr = activeDatesByTeacher.get(teacher) || [];
        arr.push(active);
        activeDatesByTeacher.set(teacher, arr);
      }

      // 4. צור רשומות חדשות ב"דיווח שיעורים" לפי תאריכים פעילים (אם יש)
      let added = 0;
      exceptionTeachers.forEach(item => {
        const teacher = item.teacher;
        const student = item.student;
        const qty     = Number(item.qty || 1);
        const yearVal = item.yearVal;
        const notes   = item.notes || '';
        const dates   = activeDatesByTeacher.get(teacher) || [];

        if (dates.length === 0) return; // אין תאריכים פעילים – נטפל בלוגיקת תדירות בהמשך

        dates.forEach(d => {
          const dateKey = Utils.toDateKey(d);
          if (ExceptionsSvc.isGloballyFiltered(dateKey, filter)) return;
          if (ExceptionsSvc.isTeacherFiltered(teacher, dateKey, filter)) return;

          const rowPriv = [
            teacher, 'פרטי', '', student, yearVal,
            PrivateSvc.mmYYYY(monthYearCell), d, qty,
            'דווח-טרם שולם', notes, '', updateStamp, ''
          ];
          const k = PrivateSvc.keyForPrivate(teacher, student, dateKey);
          RowsSvc.handleRow(null, existingMap, k, rowPriv, 'skip', rowsToInsert);
          added++;
        });

        // סמן את המורה+תלמיד שטופלו ע"י תאריכים פעילים
        if (dates.length > 0) {
          handledWithDates.add(`${teacher}|${student}`);
        }
      });

      LogSvc.info(runId, 'Private exceptions processed', { added }, logRows);
      return handledWithDates;
    }

    function processPrivateRows(ctx) {
      const {
        sheetOut, HEADERS, privRows,
        selectedYear, selectedMonth, filter,
        rowsToInsert, existingMap, travelPerTeacherDate, styleMarks,
        monthYearCell, updateStamp,
        logRows, runId, sheetFilter, FILTER_HEADERS
      } = ctx;

      Logger.log(`[PrivateSvc.processPrivateRows] Start processing ${privRows.length} private rows`);
      let blanksAdded   = 0;
      let datedAdded    = 0;
      let travelsAdded  = 0;

      const monthHeader    = monthColName(HEADERS);
      const hebDayMap      = {'א':0,'ב':1,'ג':2,'ד':3,'ה':4,'ו':5,'ש':6};
      const travelTeacherByDay = new Set();

      const normalized = privRows.map(r => {
        const teacher   = Utils.str(r['שם המורה']);
        const student   = Utils.str(r['שם התלמיד']);
        const yearVal   = Utils.str(r['שנה']) || String(selectedYear);
        const dayLet    = Utils.str(r['יום בשבוע']);
        const qty       = Number(r['כמות שיעורים'] || 1);
        const freqRaw   = Utils.str(r['תדירות']);
        const freq      = freqRaw ? freqRaw.trim() : '';
        const travelYes = Utils.str(r['תשלום נסיעות למורה']).toLowerCase();
        const totalYear = Number(r['מספר שיעורים בשנה'] || r[' מספר שיעורים בשנה'] || 0);
        const notes     = Utils.str(r['הערות']);
        const dayIndex  = (dayLet in hebDayMap) ? hebDayMap[dayLet] : null;
        const exFlag    = Utils.str(r['חריג']);

        if (teacher && dayIndex !== null &&
            (travelYes === 'כן' || travelYes === 'yes' || travelYes === 'true')) {
          travelTeacherByDay.add(`${teacher}|${dayIndex}`);
        }
        return { teacher, student, yearVal, dayLet, dayIndex, qty, freq, notes, totalYear, חריג: exFlag };
      });

      // 🔸 שלב חדש – טיפול בחריגים לפרטי
      handlePrivateExceptions(normalized, sheetFilter, FILTER_HEADERS, sheetOut, existingMap, rowsToInsert,
                              monthYearCell, updateStamp, filter, logRows, runId);

      const handledTwoWeeks = new Set();
      const handledNoDay    = new Set();

      normalized.forEach(item => {
        const { teacher, student, yearVal, qty, notes, freq, חריג } = item;
        if (!teacher || !student) return;

        // 🔸 דלג על לוגיקת שבועיים אם יש חריג
        if (String(חריג).trim() === 'חריג') return;

        const freqNorm        = (freq || '').replace(/\s+/g,'').toLowerCase();
        const isEveryTwoWeeks = freqNorm.includes('שבועיים') || freqNorm === 'פעמבשבועיים';
        if (!isEveryTwoWeeks) return;

        const key = `${teacher}|${student}`;
        if (handledTwoWeeks.has(key)) { item._skipDates = true; return; }

        const exist = countExistingAndPendingBlankPrivateByMonth(sheetOut, HEADERS, teacher, student, monthYearCell, rowsToInsert);
        const target = 2;
        const toAdd  = Math.max(0, target - exist);
        for (let i = 0; i < toAdd; i++) {
          pushBlankPrivate(rowsToInsert, teacher, student, yearVal, qty, notes, monthYearCell, updateStamp, HEADERS);
          styleMarks.dateYellow.push(rowsToInsert.length);
          styleMarks.msgYellow.push(rowsToInsert.length);
          blanksAdded++;
        }
        handledTwoWeeks.add(key);
        item._skipDates = true;
      });

      normalized.forEach(item => {
        const { teacher, student, yearVal, qty, notes, dayIndex, freq, חריג } = item;
        if (!teacher || !student) return;
        if (String(חריג).trim() === 'חריג') return; // דלג על חריג

        const noDay  = (dayIndex === null || dayIndex === undefined);
        const noFreq = !freq;
        if (!(noDay && noFreq)) return;

        const key = `${teacher}|${student}`;
        if (handledNoDay.has(key)) { item._skipDates = true; return; }

        const exist = countExistingAndPendingBlankPrivateByMonth(sheetOut, HEADERS, teacher, student, monthYearCell, rowsToInsert);
        const target = 4;
        const toAdd  = Math.max(0, target - exist);
        for (let i = 0; i < toAdd; i++) {
          pushBlankPrivate(rowsToInsert, teacher, student, yearVal, qty, notes, monthYearCell, updateStamp, HEADERS);
          styleMarks.dateYellow.push(rowsToInsert.length);
          styleMarks.msgYellow.push(rowsToInsert.length);
          blanksAdded++;
        }
        handledNoDay.add(key);
        item._skipDates = true;
      });

      normalized.forEach(item => {
        const { teacher, student, yearVal, dayIndex, qty, notes, _skipDates, חריג } = item;
        if (_skipDates) return;
        if (!teacher || !student || dayIndex === null || dayIndex === undefined) return;
        if (String(חריג).trim() === 'חריג') return; // דלג על חריג

        const dates = RowsSvc.getDatesInMonth(selectedYear, selectedMonth, dayIndex);
        dates.forEach(d => {
          if (d.getMonth()+1 !== selectedMonth || d.getFullYear() !== selectedYear) return;

          const dateKey = Utils.toDateKey(d);
          if (ExceptionsSvc.isGloballyFiltered(dateKey, filter)) return;
          if (ExceptionsSvc.isTeacherFiltered(teacher, dateKey, filter)) return;

          const rowPriv = [
            teacher, 'פרטי', '', student, yearVal,
            mmYYYY(monthYearCell), d, qty,
            'דווח-טרם שולם', notes || '', '', updateStamp || '', ''
          ];
          const k = keyForPrivate(teacher, student, dateKey);
          RowsSvc.handleRow(null, existingMap, k, rowPriv, 'skip', rowsToInsert);
          datedAdded++;

          const wantsTravel   = travelTeacherByDay.has(`${teacher}|${dayIndex}`);
          const travelOnceKey = `${teacher}|${dateKey}`;
          if (wantsTravel && !travelPerTeacherDate.has(travelOnceKey)) {
            const rowTravel = [
              teacher, 'יום נסיעות', '', '', '', mmYYYY(monthYearCell),
              d, 1, 'דווח-טרם שולם', '', '', updateStamp || '', ''
            ];
            const kT = `${teacher}|יום נסיעות||${''}|${dateKey}`;
            RowsSvc.handleRow(null, existingMap, kT, rowTravel, 'skip', rowsToInsert);
            travelsAdded++;
            travelPerTeacherDate.add(travelOnceKey);
          }
        });
      });

      Logger.log(`[PrivateSvc.processPrivateRows] Added: blanks=${blanksAdded}, dated=${datedAdded}, travels=${travelsAdded}`);
      LogSvc.info(runId, 'Private lessons processed', {
        blanksAdded, datedAdded, travelsAdded, totalPrivate: privRows.length
      }, logRows);
    }

    return { processPrivateRows, mmYYYY, keyForPrivate };
  })();
}
