/**
 * @file Constants.gs
 * Centralized constants for sheet names, config keys, and default values.
 */

const APP = Object.freeze({
  MENU_NAME: 'Manage the sheet',
  CONFIG_SHEET: 'Configuration',
  DEFAULT_DB_SHEET: 'database',
});

const CONFIG_KEYS = Object.freeze({
  TEMPLATE_SLIDES_ID: 'TEMPLATE_SLIDES_ID',
  SOURCE_SHEET_NAME: 'SOURCE_SHEET_NAME',
  OUTPUT_FOLDER_ID: 'OUTPUT_FOLDER_ID',
  OUTPUT_FILE_NAME: 'OUTPUT_FILE_NAME',
  START_ROW: 'START_ROW',
  TEMPLATE_SLIDE_INDEX: 'TEMPLATE_SLIDE_INDEX',
});

const DEFAULTS = Object.freeze({
  OUTPUT_FILE_NAME: 'Generated Slides - {{date}}',
  START_ROW: 2,              // 1-based (row 1 is headers)
  TEMPLATE_SLIDE_INDEX: 1,   // 1-based
});

const STYLES = Object.freeze({
  DB_HEADER_BG: '#6495ED',   // cornflower blue
  DB_HEADER_FG: '#FFFFFF',
});
