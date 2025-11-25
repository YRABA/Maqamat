/** ProtectSvc: נעילת שורות לפי סטטוס + עיצוב (אייקון/רקע) */
if (typeof ProtectSvc === 'undefined') {
  var ProtectSvc = (() => {
    // מזהה ב‑description כדי לאתר הגנות שנוצרו ע״י השירות
    const LOCK_PREFIX     = 'ROW_LOCK::';
    const LOCKED_VALUES   = new Set([
      'שולם - אסור לערוך שינויים',
      'הועבר לתשלום - אסור לערוך שינויים'
    ]);
    const OPEN_VALUE      = 'דווח-טרם שולם';
    const LOCKED_BG       = '#EDEFF2';
    const YELLOW          = '#FFFF00';

    // עזרי השוואה/איתור
    const idxFromHeaders = headers => {
      const m = {};
      headers.forEach((h,i) => m[h] = i);
      return m;
    };
    const isLockedStatus = v => LOCKED_VALUES.has(String(v||'').trim());
    const isYellowBg      = c => String(c || '').toUpperCase() === YELLOW;
    const isLockedGrayBg  = c => String(c || '').toUpperCase() === LOCKED_BG.toUpperCase();

    /**
     * מבצע נעילה של שורות שהוכנסו בשליטה (startRow, count) בלבד.
     * שורות קיימות נשארות במצבן.
     * @param {Sheet} sheet   – גיליון "דיווח שיעורים"
     * @param {Array} headers – כותרות העמודות
     * @param {number} startRow – מספר שורה ראשוני של שורות חדשות
     * @param {number} numRows – מספר שורות חדשות שהוכנסו
     */
    function applyLocksForNewRows(sheet, headers, startRow, numRows) {
      if (!sheet || numRows <= 0) return;
      const idx = idxFromHeaders(headers);
      const statusColIdx  = idx['סטטוס'];
      if (statusColIdx < 0) return;
      const statusCol = statusColIdx + 1;
      const lastCol   = headers.length;

      // Remove any existing protections created by this service — for new rows only
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => {
        if ((p.getDescription()||'').startsWith(LOCK_PREFIX)) {
          p.remove();
        }
      });

      // Build sheet‑level protection
      const prot = sheet.protect();
      prot.setDescription(LOCK_PREFIX + `newRows_${startRow}_${numRows}`);
      prot.setWarningOnly(false);
      try {
        prot.removeEditors(prot.getEditors());
        prot.setDomainEdit(false);
      } catch (e) {
        Logger.log(`⚠️ Protection editor cleanup failed: ${e}`);
      }


      // Determine ranges to leave unprotected: status cells of new rows
      const unprotectedRanges = [];
      for (let r = startRow; r < startRow + numRows; r++) {
        unprotectedRanges.push(sheet.getRange(r, statusCol, 1, 1));
      }
      prot.setUnprotectedRanges(unprotectedRanges);

      // For aesthetic: color the newly locked rows background if status indicates locked
      const values = sheet.getRange(startRow, statusCol, numRows, 1).getValues();
      const fullRange = sheet.getRange(startRow, 1, numRows, lastCol);
      const bgs = fullRange.getBackgrounds();
      for (let i = 0; i < numRows; i++) {
        const v = String(values[i][0] || '').trim();
        if (isLockedStatus(v)) {
          for (let j = 0; j < lastCol; j++) {
            if (!isYellowBg(bgs[i][j])) {
              bgs[i][j] = LOCKED_BG;
            }
          }
        }
      }
      fullRange.setBackgrounds(bgs);
    }

    /**
     * פעולת toggle לשורה בודדת (בעת שינוי סטטוס).
     * רק שורה שלא ב‑startRow/newRows‹ scenario.
     */
    function toggleRowLockForRow(sheet, headers, row) {
      if (row <= 1) return;
      const idx = idxFromHeaders(headers);
      const statusCol = idx['סטטוס'] + 1;
      const lastCol   = headers.length;

      const statusVal = String(sheet.getRange(row, statusCol).getValue() || '').trim();

      if (isLockedStatus(statusVal)) {
        _lockSingleRow(sheet, headers, row, statusCol, lastCol);
      } else {
        _unlockSingleRow(sheet, headers, row, lastCol);
      }
    }

    /**
     * נעילה של שורה בודדת.
     */
    function _lockSingleRow(sheet, headers, row, statusCol, lastCol) {
      const rowRange   = sheet.getRange(row, 1, 1, lastCol);
      const statusCell = sheet.getRange(row, statusCol, 1, 1);

      Logger.log(`🔧 Locking row ${row} based on status`);

      // Remove older protections on that sheet (to avoid stacking)
      const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(p => {
        if ((p.getDescription() || '').startsWith(LOCK_PREFIX + `row_${row}`)) {
          p.remove();
        }
      });

      // Apply protection with status cell unprotected
      try {
        const prot = sheet.protect();
        prot.setDescription(LOCK_PREFIX + `row_${row}`);
        prot.setWarningOnly(false);
        prot.setUnprotectedRanges([statusCell]);
        prot.removeEditors(prot.getEditors());
        prot.setDomainEdit(false);
      } catch (e) {
        Logger.log(`⚠️ Failed to apply protection on row ${row}: ${e}`);
      }

      // Paint row grey if needed
      const bgColors = rowRange.getBackgrounds()[0];
      for (let j = 0; j < bgColors.length; j++) {
        if (!isYellowBg(bgColors[j])) bgColors[j] = LOCKED_BG;
      }
      rowRange.setBackgrounds([bgColors]);

      Logger.log(`✅ Row ${row} locked.`);
    }

    /**
     * שחרור שורה בודדת.
     */
    function _unlockSingleRow(sheet, headers, row, lastCol) {
      const rowRange = sheet.getRange(row, 1, 1, lastCol);
      Logger.log(`🔧 Unlocking row ${row} based on status`);

      // Remove protections for this row
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)
           .forEach(p => {
             if ((p.getDescription()||'').startsWith(LOCK_PREFIX + `row_${row}`)) {
               p.remove();
             }
           });

      // Clear grey background if it was applied by our lock
      const bgs = rowRange.getBackgrounds()[0];
      for (let j = 0; j < bgs.length; j++) {
        if (isLockedGrayBg(bgs[j])) bgs[j] = null;
      }
      rowRange.setBackgrounds([bgs]);

      Logger.log(`✅ Row ${row} unlocked.`);
    }

    function isLockedStatusPublic(v) {
      return isLockedStatus(v);
    }

    return {
      applyLocksForNewRows,
      toggleRowLockForRow,
      isLockedStatus: isLockedStatusPublic,
      OPEN_VALUE
    };
  })();
}
