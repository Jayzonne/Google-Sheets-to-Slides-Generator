/**
 * @file Utils.gs
 * Generic helpers.
 */

class Utils {
  /**
   * @return {string} Formatted current datetime in script timezone.
   */
  static formatNow_() {
    const tz = Session.getScriptTimeZone();
    return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  }

  /**
   * Replaces {{date}} token in file name patterns.
   * @param {string} namePattern
   * @return {string}
   */
  static applyFileNameTokens_(namePattern) {
    return String(namePattern || '').replace('{{date}}', Utils.formatNow_());
  }

  /**
   * True if a row is empty (all cells empty/whitespace).
   * @param {Array<*>} row
   * @return {boolean}
   */
  static isEmptyRow_(row) {
    return row.every((cell) => String(cell ?? '').trim() === '');
  }

  /**
   * Safe number conversion with fallback.
   * @param {string|number} value
   * @param {number} fallback
   * @return {number}
   */
  static toNumber_(value, fallback) {
    const n = Number(value);
    return Number.isFinite(n) ? n : fallback;
  }
}
