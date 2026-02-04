/**
 * @file ConfigService.gs
 * Reads, validates, and normalizes the Configuration sheet.
 */

/**
 * @typedef {Object} ImageConfig
 * @property {number} index
 * @property {string} field            Column name in database (ex: photoId)
 * @property {string} source           DRIVE_ID | URL
 * @property {string} fit              CONTAIN | STRETCH
 */

/**
 * @typedef {Object} AppConfig
 * @property {string} templateSlidesId
 * @property {string} outputFolderId
 * @property {string} sourceSheetName
 * @property {string} outputFileNamePattern
 * @property {number} startRow
 * @property {number} templateSlideIndex
 * @property {ImageConfig[]} images
 */

class ConfigService {
  /**
   * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
   */
  constructor(spreadsheet) {
    this.ss = spreadsheet;
  }

  /**
   * @return {AppConfig}
   */
  getConfig() {
    const sheet = this.ss.getSheetByName(APP.CONFIG_SHEET);
    if (!sheet) {
      throw new Error(`Missing sheet: "${APP.CONFIG_SHEET}". Use "Restructure the template" first.`);
    }

    const values = sheet.getDataRange().getValues();
    const raw = {};

    // Expected format: [Setting, Value, Help], header in row 1.
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || '').trim();
      const val = String(values[i][1] || '').trim();
      if (key) raw[key] = val;
    }

    const required = [
      CONFIG_KEYS.TEMPLATE_SLIDES_ID,
      CONFIG_KEYS.SOURCE_SHEET_NAME,
      CONFIG_KEYS.OUTPUT_FOLDER_ID,
    ];

    const missing = required.filter((k) => !raw[k]);
    if (missing.length) {
      throw new Error(
        `Missing configuration values: ${missing.join(', ')}.\n` +
        `Open "${APP.CONFIG_SHEET}" and fill them in.`
      );
    }

    const images = this._parseImages_(raw);

    /** @type {AppConfig} */
    const config = {
      templateSlidesId: raw[CONFIG_KEYS.TEMPLATE_SLIDES_ID],
      outputFolderId: raw[CONFIG_KEYS.OUTPUT_FOLDER_ID],
      sourceSheetName: raw[CONFIG_KEYS.SOURCE_SHEET_NAME] || APP.DEFAULT_DB_SHEET,
      outputFileNamePattern: raw[CONFIG_KEYS.OUTPUT_FILE_NAME] || DEFAULTS.OUTPUT_FILE_NAME,
      startRow: Utils.toNumber_(raw[CONFIG_KEYS.START_ROW], DEFAULTS.START_ROW),
      templateSlideIndex: Utils.toNumber_(raw[CONFIG_KEYS.TEMPLATE_SLIDE_INDEX], DEFAULTS.TEMPLATE_SLIDE_INDEX),
      images,
    };

    if (config.startRow < 2) {
      throw new Error(`START_ROW must be >= 2 (row 1 is the header row). Current: ${config.startRow}`);
    }
    if (config.templateSlideIndex < 1) {
      throw new Error(`TEMPLATE_SLIDE_INDEX must be >= 1. Current: ${config.templateSlideIndex}`);
    }

    // Validate image configs
    config.images.forEach((img) => {
      if (!img.field) throw new Error(`Invalid image config: IMAGE_${img.index}_FIELD is empty.`);
      if (![IMAGES.SOURCE.DRIVE_ID, IMAGES.SOURCE.URL].includes(img.source)) {
        throw new Error(`Invalid image config: IMAGE_${img.index}_SOURCE must be DRIVE_ID | URL. Got: ${img.source}`);
      }
      if (![IMAGES.FIT.CONTAIN, IMAGES.FIT.COVER, IMAGES.FIT.STRETCH].includes(img.fit)) {
        throw new Error(`Invalid image config: IMAGE_${img.index}_FIT must be CONTAIN | COVER | STRETCH. Got: ${img.fit}`);
      }
    });


    return config;
  }

  /**
   * Parse IMAGE_N_* blocks from the configuration key/value map.
   * @param {Object<string,string>} raw
   * @return {ImageConfig[]}
   * @private
   */
  _parseImages_(raw) {
    let maxN = 0;
    Object.keys(raw).forEach((k) => {
      const m = /^IMAGE_(\d+)_FIELD$/.exec(k);
      if (m) maxN = Math.max(maxN, Number(m[1]));
    });

    /** @type {ImageConfig[]} */
    const images = [];

    for (let n = 1; n <= maxN; n++) {
      const field = String(raw[`IMAGE_${n}_FIELD`] || '').trim();

      const hasAny =
        field ||
        raw[`IMAGE_${n}_SOURCE`] ||
        raw[`IMAGE_${n}_FIT`];

      if (!hasAny) continue;
      if (!field) continue;

      const source = String(raw[`IMAGE_${n}_SOURCE`] || IMAGES.SOURCE.DRIVE_ID).trim().toUpperCase();
      const fit = String(raw[`IMAGE_${n}_FIT`] || IMAGES.FIT.CONTAIN).trim().toUpperCase();

      images.push({
        index: n,
        field,
        source,
        fit,
      });
    }

    return images;
  }
}
