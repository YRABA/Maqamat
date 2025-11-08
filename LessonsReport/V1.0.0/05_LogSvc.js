/**
 * LogSvc: שירות רישום אירועים / אבחון עבור הרצה מרכזית
 * רישום טבלאי בגיליון "לוג ריצות" + שימוש ב־Logger/console מפורט.
 */
const LogSvc = (() => {

  const TZ = 'Asia/Jerusalem';

  /**
   * Internal: format timestamp.
   */
  function timestamp() {
    return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm:ss');
  }

  /**
   * Internal: ensure `logRows` is array and runId string is set.
   */
  function normalize(runId, logRows) {
    if (typeof runId !== 'string' || runId.trim() === '') {
      runId = `RUN‑UNKNOWN‑${Date.now()}`;
    }
    if (!Array.isArray(logRows)) {
      logRows = [];
    }
    return { runId, logRows };
  }

  /**
   * Log an INFO event.
   * @param {string} runId 
   * @param {string} message 
   * @param {object=} data Additional data object (optional).
   * @param {Array=} logRows Array to push into.
   */
  function info(runId, message, data = {}, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const entry = {
      runId,
      time: timestamp(),
      level: 'INFO',
      message,
      data: (data && typeof data === 'object') ? data : {}
    };
    Logger.log(`${runId} INFO ${message} ‑ ${JSON.stringify(entry.data)}`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  /**
   * Log a DEBUG event.
   */
  function debug(runId, message, data = {}, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const entry = {
      runId,
      time: timestamp(),
      level: 'DEBUG',
      message,
      data: (data && typeof data === 'object') ? data : {}
    };
    Logger.log(`${runId} DEBUG ${message} ‑ ${JSON.stringify(entry.data)}`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  /**
   * Log a SUCCESS event.
   */
  function success(runId, message, data = {}, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const entry = {
      runId,
      time: timestamp(),
      level: 'SUCCESS',
      message,
      data: (data && typeof data === 'object') ? data : {}
    };
    Logger.log(`${runId} SUCCESS ${message} ‑ ${JSON.stringify(entry.data)}`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  /**
   * Log an ERROR event.
   */
  function error(runId, message, data = {}, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const entry = {
      runId,
      time: timestamp(),
      level: 'ERROR',
      message,
      data: (data && typeof data === 'object') ? data : {}
    };
    Logger.log(`${runId} ERROR ${message} ‑ ${JSON.stringify(entry.data)}`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  function startTimer(runId, label, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const data = { label, time: timestamp() };
    const entry = { runId, time: timestamp(), level: 'TIMER-START', message: `Start timer: ${label}`, data };
    Logger.log(`${runId} TIMER-START ${label}`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  function endTimer(runId, label, elapsedMs, logRows = []) {
    const norm = normalize(runId, logRows);
    runId = norm.runId;
    logRows = norm.logRows;

    const data = { label, elapsedMs, time: timestamp() };
    const entry = { runId, time: timestamp(), level: 'TIMER-END', message: `End timer: ${label}`, data };
    Logger.log(`${runId} TIMER-END ${label} elapsed: ${elapsedMs} ms`);
    logRows.push(Object.values(entry));
    return logRows;
  }

  return {
    info,
    debug,
    success,
    error,
    startTimer,
    endTimer
  };
})();
