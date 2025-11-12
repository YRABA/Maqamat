/** ExceptionsSvc: ניהול חריגים / סינונים */
const ExceptionsSvc = (() => {

  /**
   * Build index map for headers for quick lookup.
   */
  function indexMap(headers) {
    const idx = {};
    headers.forEach((h, i) => idx[h] = i);
    return idx;
  }

  /**
   * Clear filter messages (currently disabled to preserve user messages).
   */
  function clearFilterMessages(sheetFilter, headers) {
    // intentionally left blank — preserving user notes
    Logger.log(`[ExceptionsSvc.clearFilterMessages] Skipped clearing user messages`);
  }

  /**
   * Write a message to the 'הודעת מערכת' column and highlight it.
   */
  function setFilterMessage(sheetFilter, headers, rowIndex, msg) {
    const idx = indexMap(headers);
    const cell = sheetFilter.getRange(rowIndex, idx['הודעת מערכת'] + 1);
    cell.setValue(msg);
    if (String(msg).trim()) {
      cell.setBackground('#FFFF00'); // emphasize message
    }
    Logger.log(`[ExceptionsSvc.setFilterMessage] Row ${rowIndex}: ${msg}`);
  }

  /**
   * Parse the exceptions sheet into structured filter sets.
   * Handles global filters, teacher-specific ranges, and active (exception) rows.
   */
  function parseFilterSheet(sheetFilter, headers, logRows, runId) {
    const idx = indexMap(headers);
    const values = sheetFilter.getDataRange().getValues();
    const out = {
      globalSingles: new Set(),
      globalRanges: [],
      teacherSingles: new Map(),
      teacherRanges: new Map(),
      activeRows: [],
      teachersIgnoreWeekly: new Set()
    };

    if (values.length < 2) {
      Logger.log(`[ExceptionsSvc.parseFilterSheet] Empty or header-only sheet`);
      LogSvc.info(runId, 'Filter sheet parsed', { activeRows: 0 }, logRows);
      return out;
    }

    let activeCount = 0, teacherFilters = 0, globalFilters = 0;

    for (let r = 1; r < values.length; r++) {
      const rowIndex = r + 1;
      const teacher  = Utils.str(values[r][idx['שם המורה']]);
      const active   = Utils.coerceToDate(values[r][idx['תאריך פעיל']]);
      const fromD    = Utils.coerceToDate(values[r][idx['תאריך לא פעיל מ']]);
      const toD      = Utils.coerceToDate(values[r][idx['תאריך לא פעיל עד']]);

      const hasTeacher = !!teacher;
      const hasActive  = !!active;
      const hasFrom    = !!fromD;
      const hasTo      = !!toD;

      // Active row → always collected
      if (hasTeacher && hasActive) {
        out.activeRows.push({ rowIndex, teacher, date: active });
        activeCount++;
      }

      // Mark teachers to ignore weekly generation
      if (hasTeacher && !hasFrom && !hasTo) {
        out.teachersIgnoreWeekly.add(teacher);
      }

      // Global single
      if (!hasTeacher && hasFrom && !hasTo) {
        out.globalSingles.add(Utils.toDateKey(fromD));
        globalFilters++;
      }

      // Global range
      if (!hasTeacher && hasFrom && hasTo) {
        out.globalRanges.push({ fromKey: Utils.toDateKey(fromD), toKey: Utils.toDateKey(toD) });
        globalFilters++;
      }

      // Teacher single
      if (hasTeacher && hasFrom && !hasTo) {
        const set = out.teacherSingles.get(teacher) || new Set();
        set.add(Utils.toDateKey(fromD));
        out.teacherSingles.set(teacher, set);
        teacherFilters++;
      }

      // Teacher range
      if (hasTeacher && hasFrom && hasTo) {
        const arr = out.teacherRanges.get(teacher) || [];
        arr.push({ fromKey: Utils.toDateKey(fromD), toKey: Utils.toDateKey(toD) });
        out.teacherRanges.set(teacher, arr);
        teacherFilters++;
      }
    }

    Logger.log(`[ExceptionsSvc.parseFilterSheet] Active=${activeCount}, TeacherFilters=${teacherFilters}, GlobalFilters=${globalFilters}`);
    LogSvc.info(runId, 'Filter sheet parsed', {
      activeRows: activeCount,
      teacherFilters,
      globalFilters
    }, logRows);

    return out;
  }

  /**
   * Checks whether a date is globally filtered.
   */
  function isGloballyFiltered(dateKey, filter) {
    if (filter.globalSingles.has(dateKey)) return true;
    for (const r of filter.globalRanges)
      if (dateKey >= r.fromKey && dateKey <= r.toKey) return true;
    return false;
  }

  /**
   * Checks whether a date for a given teacher is filtered.
   */
  function isTeacherFiltered(teacher, dateKey, filter) {
    const s = filter.teacherSingles.get(teacher);
    if (s && s.has(dateKey)) return true;
    const arr = filter.teacherRanges.get(teacher) || [];
    for (const r of arr)
      if (dateKey >= r.fromKey && dateKey <= r.toKey) return true;
    return false;
  }

  return {
    clearFilterMessages,
    setFilterMessage,
    parseFilterSheet,
    isGloballyFiltered,
    isTeacherFiltered
  };
})();
