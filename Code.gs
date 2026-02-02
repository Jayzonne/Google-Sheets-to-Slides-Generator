/**
 * @file Code.gs
 * Global entry points (onOpen + menu actions).
 * Note: menu callbacks must be globally-scoped functions.
 */

function onOpen() {
  MenuController.onOpen();
}

/**
 * Menu action: Restructure the template (create/format sheets).
 */
function restructureTemplate() {
  new SheetService(SpreadsheetApp.getActiveSpreadsheet()).setup();
}

/**
 * Menu action: Generate slides from the database sheet using the config.
 */
function generateSlides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const config = new ConfigService(ss).getConfig();
  const sheetService = new SheetService(ss);

  const dataset = sheetService.readDataset(config.sourceSheetName, config.startRow);

  const totalRowsInRange = dataset.rows.length;
  const nonEmptyRows = dataset.rows.filter(r => !Utils.isEmptyRow_(r)).length;

  const confirmMessage =
    `You are about to generate Google Slides.\n\n` +
    `Source sheet: ${config.sourceSheetName}\n` +
    `Start row: ${config.startRow}\n` +
    `Rows found (from start row): ${totalRowsInRange}\n` +
    `Non-empty rows to generate: ${nonEmptyRows}\n` +
    `Template slide index (blueprint): ${config.templateSlideIndex}\n\n` +
    `Proceed?`;

  const choice = ui.alert('Confirm generation', confirmMessage, ui.ButtonSet.OK_CANCEL);
  if (choice !== ui.Button.OK) return;

  const generator = new SlidesGenerator(config);
  const result = generator.generate(dataset);

  ui.alert(
    `Done ✅\nSlides generated: ${result.slidesGenerated}\nFile name: ${result.fileName}\nFile ID: ${result.fileId}`
  );
}
