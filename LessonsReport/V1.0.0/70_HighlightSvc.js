/** HighlightSvc: הדגשות ייעודיות לגיליון "חריגים-קבוצתי" */
var HighlightSvc = (() => {

  /**
   * צובע בצהוב את כל התאים הלא-ריקים בעמודה "הודעת מערכת" בגיליון "חריגים-קבוצתי".
   * לא מוחק צבעים קיימים, ולא מנקה תאים ריקים.
   * @param {{sheetName?:string, header?:string, color?:string, logRows?:any[], runId?:string}} opts
   */
  function highlightNonEmptyMessages(opts = {}) {
    const SHEET_NAME = opts.sheetName || 'חריגים-קבוצתי';
    const HEADER_TXT = opts.header    || 'הודעת מערכת';
    const COLOR      = opts.color     || '#FFF59D';

    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) {
      LogSvc.error('HighlightSvc.highlightNonEmptyMessages', 'לא נמצא גיליון: ' + SHEET_NAME, opts.runId);
      return;
    }

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return;

    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const colIdx = headers.indexOf(HEADER_TXT);
    if (colIdx === -1) return;

    const rowCount = lastRow - 1;
    const col = colIdx + 1;

    const vals = sh.getRange(2, col, rowCount, 1).getValues();
    const bgs  = sh.getRange(2, col, rowCount, 1).getBackgrounds();

    const toPaint = [];
    for (let i = 0; i < rowCount; i++) {
      const v = String(vals[i][0] ?? '').trim();
      if (v.length === 0) continue;
      if (bgs[i][0] === COLOR) continue;
      toPaint.push(2 + i);
    }

    if (toPaint.length) {
      const list = toPaint.map(r => sh.getRange(r, col).getA1Notation());
      sh.getRangeList(list).setBackground(COLOR);
    }

    LogSvc.debug('HighlightSvc.highlightNonEmptyMessages', `ניצבעו ${toPaint.length} תאים בגיליון "${SHEET_NAME}"`, opts.runId);
    if (opts.logRows && opts.runId) {
      LogSvc.log(opts.logRows, opts.runId, 'HighlightSvc.highlightNonEmptyMessages', `ניצבעו ${toPaint.length} תאים`);
    }
  }

  return { highlightNonEmptyMessages };
})();
