/**
 * @file SheetService.gs
 * Sheet lifecycle (restructure) + dataset reading.
 *
 * Key feature:
 * - Column A is reserved for "To generate" (checkboxes).
 * - Slide generation uses ONLY checked rows.
 */

/**
 * @typedef {Object} Dataset
 * @property {string[]} headers
 * @property {Array<Array<*>>} rows
 * @property {number} totalRows       Rows present in the range (from startRow to last data row)
 * @property {number} checkedRows     Rows checked in "To generate"
 */

class SheetService {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
   */
  constructor(spreadsheet) {
    this.ss = spreadsheet;
  }

  /**
   * Restructure action:
   * - Deletes Configuration + database sheets entirely (if they exist)
   * - Recreates them from scratch with formatting and examples
   */
  setup() {
    const targets = [APP.CONFIG_SHEET, APP.DEFAULT_DB_SHEET];

    // Ensure we never delete all sheets: create a temporary keeper if needed.
    const allSheets = this.ss.getSheets();
    const nonTargetSheets = allSheets.filter((s) => !targets.includes(s.getName()));
    let keeperSheet = null;

    if (nonTargetSheets.length === 0) {
      keeperSheet = this.ss.insertSheet('__KEEPER__');
    }

    // Delete target sheets if they exist.
    targets.forEach((name) => {
      const sh = this.ss.getSheetByName(name);
      if (sh) this.ss.deleteSheet(sh);
    });

    // Recreate from scratch.
    this._createConfiguration_();
    this._createDatabase_();

    // Remove keeper if we created it.
    if (keeperSheet) this.ss.deleteSheet(keeperSheet);

    SpreadsheetApp.getUi().alert(
      `Sheets recreated: "${APP.CONFIG_SHEET}" and "${APP.DEFAULT_DB_SHEET}"`
    );
  }

  /**
   * Ensures "To generate" exists in column A with centered checkboxes.
   *
   * Behavior:
   * - If "To generate" exists in another column, that column is moved to A.
   * - Otherwise, a new column is inserted before A.
   *
   * @param {string} sheetName
   * @param {number} startRow 1-based first row of data
   */
  ensureGenerateColumn(sheetName, startRow) {
    const sheet = this.ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Source sheet not found: "${sheetName}"`);

    const lastCol = Math.max(1, sheet.getLastColumn());
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((v) => String(v || '').trim());
    const foundIdx = headers.indexOf(DB.GENERATE_HEADER); // 0-based

    if (foundIdx === -1) {
      // Not found: insert a new A column.
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue(DB.GENERATE_HEADER);
    } else if (foundIdx !== 0) {
      // Found but not in column A: move it to column A.
      const sourceCol = foundIdx + 1; // 1-based
      const rangeToMove = sheet.getRange(1, sourceCol, sheet.getMaxRows(), 1);
      sheet.moveColumns(rangeToMove, 1);
    }
    // Now the generate column is column A.

    // Style A1.
    sheet.getRange(1, 1)
      .setBackground(STYLES.DB_HEADER_BG)
      .setFontColor(STYLES.DB_HEADER_FG)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

    sheet.setColumnWidth(1, DB.GENERATE_COL_WIDTH);

    // Insert / refresh checkboxes and center them.
    const lastRow = sheet.getLastRow();
    if (lastRow >= startRow) {
      const cbRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, 1);
      cbRange.insertCheckboxes();
      cbRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }
  }

  /**
   * Reads headers + checked rows ONLY (checkbox-driven).
   * Column "To generate" is excluded from headers/rows.
   *
   * @param {string} sheetName
   * @param {number} startRow 1-based
   * @return {Dataset}
   */
  readDataset(sheetName, startRow) {
    const sheet = this.ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Source sheet not found: "${sheetName}"`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      throw new Error(`No data found in "${sheetName}" (needs headers + at least one row).`);
    }

    const headerRow = data[0].map((h) => String(h || '').trim());
    const genIdx = headerRow.indexOf(DB.GENERATE_HEADER);

    if (genIdx === -1) {
      throw new Error(
        `Missing "${DB.GENERATE_HEADER}" header. ` +
        `Run "Generate slides" once (it will add it) or "Restructure the template".`
      );
    }

    // Determine which columns we keep in the dataset (exclude "To generate" and exclude empty headers).
    const keepIdx = headerRow
      .map((h, i) => ({ h, i }))
      .filter((o) => o.h && o.i !== genIdx)
      .map((o) => o.i);

    const headers = keepIdx.map((i) => headerRow[i]);
    if (!headers.length) {
      throw new Error(`No usable headers found in row 1 of "${sheetName}".`);
    }

    const rowsAll = data.slice(startRow - 1);
    const totalRows = rowsAll.length;

    // Keep only checked rows (strictly true).
    const checkedRowsAll = rowsAll.filter((r) => r[genIdx] === true);
    const checkedRows = checkedRowsAll.length;

    // Strip the "To generate" column from each row.
    const rows = checkedRowsAll.map((r) => keepIdx.map((i) => r[i]));

    return { headers, rows, totalRows, checkedRows };
  }

  /**
   * Creates and formats the Configuration sheet.
   * @private
   */
  _createConfiguration_() {
    const sheet = this.ss.insertSheet(APP.CONFIG_SHEET);

    const values = [
      ['Setting', 'Value', 'Help'],
      [CONFIG_KEYS.TEMPLATE_SLIDES_ID, '', 'Google Slides TEMPLATE file ID (from the URL)'],
      [CONFIG_KEYS.SOURCE_SHEET_NAME, APP.DEFAULT_DB_SHEET, 'Sheet name to read rows from (e.g., database)'],
      [CONFIG_KEYS.OUTPUT_FOLDER_ID, '', 'Google Drive folder ID where the generated Slides file will be saved'],
      [CONFIG_KEYS.OUTPUT_FILE_NAME, DEFAULTS.OUTPUT_FILE_NAME, 'Output file name. {{date}} is replaced automatically'],
      [CONFIG_KEYS.START_ROW, String(DEFAULTS.START_ROW), 'First data row (2 = right after headers)'],
      [
        CONFIG_KEYS.TEMPLATE_SLIDE_INDEX,
        String(DEFAULTS.TEMPLATE_SLIDE_INDEX),
        '1-based index of the slide INSIDE the template deck used as the blueprint for each generated slide. ' +
          'Example: 1 = first slide, 2 = second slide. Put ALL placeholders on that slide (e.g., {{firstName}}). ' +
          'Note: the generator uses ONLY this slide as a blueprint.',
      ],
    ];

    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    sheet.setFrozenRows(1);

    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 420);
    sheet.setColumnWidth(3, 620);

    sheet.setRowHeight(1, 34);
    for (let r = 2; r <= values.length; r++) sheet.setRowHeight(r, 52);

    sheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');

    sheet.getRange(1, 1, values.length, 3)
      .setWrap(true)
      .setVerticalAlignment('middle');
  }

  /**
   * Creates and formats the database sheet (checkbox in column A).
   * @private
   */
  _createDatabase_() {
    const db = this.ss.insertSheet(APP.DEFAULT_DB_SHEET);

    // Column A: checkbox-driven generation
    db.getRange(1, 1, 1, 5).setValues([[DB.GENERATE_HEADER, 'firstName', 'lastName', 'city', 'company']]);

    // Example rows: unchecked by default
    db.getRange(2, 1, 2, 5).setValues([
      [false, 'Alice', 'Johnson', 'Paris', 'ACME Inc.'],
      [false, 'Bob', 'Martin', 'Lyon', 'Globex Corp.'],
    ]);

    db.setFrozenRows(1);

    // Header styling
    const lastCol = db.getLastColumn();
    db.getRange(1, 1, 1, lastCol)
      .setBackground(STYLES.DB_HEADER_BG)
      .setFontColor(STYLES.DB_HEADER_FG)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');

    // Center only the checkbox column header
    db.getRange(1, 1).setHorizontalAlignment('center').setVerticalAlignment('middle');

    // Checkbox column width + checkboxes centered
    db.setColumnWidth(1, DB.GENERATE_COL_WIDTH);
    const lastRow = db.getLastRow();
    if (lastRow >= 2) {
      const cbRange = db.getRange(2, 1, lastRow - 1, 1);
      cbRange.insertCheckboxes();
      cbRange.setHorizontalAlignment('center').setVerticalAlignment('middle');
    }

    db.setRowHeight(1, 32);
    db.autoResizeColumns(2, Math.max(1, lastCol - 1));
  }
}
