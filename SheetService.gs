/**
 * @file SheetService.gs
 * Deletes and recreates sheets (Configuration + database), and reads the dataset.
 */

/**
 * @typedef {Object} Dataset
 * @property {string[]} headers
 * @property {Array<Array<*>>} rows
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
    const nonTargetSheets = allSheets.filter(s => !targets.includes(s.getName()));
    let keeperSheet = null;

    if (nonTargetSheets.length === 0) {
      keeperSheet = this.ss.insertSheet('__KEEPER__');
    }

    // Delete target sheets if they exist
    targets.forEach(name => {
      const sh = this.ss.getSheetByName(name);
      if (sh) this.ss.deleteSheet(sh);
    });

    // Recreate from scratch
    this._createConfiguration_();
    this._createDatabase_();

    // Remove keeper if we created it
    if (keeperSheet) {
      // After recreation, we definitely have other sheets, so it's safe to delete.
      this.ss.deleteSheet(keeperSheet);
    }

    SpreadsheetApp.getUi().alert(
      `Sheets recreated: "${APP.CONFIG_SHEET}" and "${APP.DEFAULT_DB_SHEET}"`
    );
  }

  /**
   * Reads headers + rows.
   * @param {string} sheetName
   * @param {number} startRow 1-based
   * @return {Dataset}
   */
  readDataset(sheetName, startRow) {
    const sheet = this.ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Source sheet not found: "${sheetName}"`);

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error(`No data found in "${sheetName}" (needs headers + at least one row).`);

    const headers = data[0].map(h => String(h).trim()).filter(Boolean);
    if (!headers.length) throw new Error(`No headers found in row 1 of "${sheetName}".`);

    const rows = data.slice(startRow - 1);
    return { headers, rows };
  }

  /**
   * Creates and formats the Configuration sheet.
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
          'Note: the generator uses ONLY this slide as a blueprint.'
      ],
    ];

    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);

    // Larger cells / better readability
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
   * Creates and formats the database sheet with English examples and cornflower-blue headers.
   */
  _createDatabase_() {
    const db = this.ss.insertSheet(APP.DEFAULT_DB_SHEET);

    // English example headers + data
    db.getRange(1, 1, 1, 4).setValues([['firstName', 'lastName', 'city', 'company']]);
    db.getRange(2, 1, 2, 4).setValues([
      ['Alice', 'Johnson', 'Paris', 'ACME Inc.'],
      ['Bob', 'Martin', 'Lyon', 'Globex Corp.'],
    ]);

    db.setFrozenRows(1);

    // Cornflower blue header row
    const lastCol = db.getLastColumn();
    db.getRange(1, 1, 1, lastCol)
      .setBackground(STYLES.DB_HEADER_BG)
      .setFontColor(STYLES.DB_HEADER_FG)
      .setFontWeight('bold')
      .setVerticalAlignment('middle');

    db.setRowHeight(1, 32);
    db.autoResizeColumns(1, Math.max(4, lastCol));
  }
}
