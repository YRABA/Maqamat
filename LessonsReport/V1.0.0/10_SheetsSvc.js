/** SheetsSvc: שירותי עבודה עם Sheets */
const SheetsSvc = (() => {

  function getOrCreateSheet(ss, name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      Logger.log(`[SheetsSvc.getOrCreateSheet] Found sheet "${name}"`);
      return sheet;
    } else {
      const created = ss.insertSheet(name);
      Logger.log(`[SheetsSvc.getOrCreateSheet] Created sheet "${name}"`);
      return created;
    }
  }

  function ensureHeader(sheet, headers) {
    const existingRows = sheet.getLastRow();
    const colCount = headers.length;
    const existingCols = sheet.getMaxColumns();

    if (existingCols < colCount) {
      sheet.insertColumnsAfter(existingCols, colCount - existingCols);
    }

    sheet.getRange(1, 1, 1, colCount).setValues([headers]);
    sheet.setFrozenRows(1);

    const headerRange = sheet.getRange(1, 1, 1, colCount);
    headerRange.setFontWeight("bold").setBackground("#ffffff").setHorizontalAlignment("center");

    const maxRows = Math.max(sheet.getMaxRows(), 2);
    const dataRange = sheet.getRange(1, 1, maxRows, colCount);
    
    if (sheet.getFilter()) {
      try {
        sheet.getFilter().remove();
      } catch (e) {
        Logger.log(`[SheetsSvc.ensureHeader] Failed to remove filter: ${e}`);
      }
    }

    dataRange.createFilter();
    Logger.log(`[SheetsSvc.ensureHeader] Headers ensured for sheet "${sheet.getName()}"`);
  }

  function alignSheetRight(sheet) {
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    if (maxRows > 0 && maxCols > 0) {
      sheet.getRange(1, 1, maxRows, maxCols).setHorizontalAlignment("right");
      Logger.log(`[SheetsSvc.alignSheetRight] Aligned content to right on sheet "${sheet.getName()}"`);
    }
  }

  function getRowsAsObjects(sheet) {
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      Logger.log(`[SheetsSvc.getRowsAsObjects] No data found in "${sheet.getName()}"`);
      return [];
    }

    const headers = values[0].map(h => Utils.str(h));
    const result = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (row.every(v => v === '' || v === null)) continue;

      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      result.push(obj);
    }

    Logger.log(`[SheetsSvc.getRowsAsObjects] Loaded ${result.length} row objects from "${sheet.getName()}"`);
    return result;
  }

  return {
    getOrCreateSheet,
    ensureHeader,
    alignSheetRight,
    getRowsAsObjects
  };
})();
