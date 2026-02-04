/**
 * @file Code.gs
 * Global entry points (onOpen + menu actions).
 * Menu callbacks must be globally-scoped functions.
 */

function onOpen() {
  MenuController.onOpen();
}

/**
 * Menu action: Restructure the template (create/format sheets).
 */
function restructureTemplate() {
  const ui = SpreadsheetApp.getUi();
  try {
    new SheetService(SpreadsheetApp.getActiveSpreadsheet()).setup();
  } catch (err) {
    ui.alert(`Restructure failed:\n\n${err && err.message ? err.message : String(err)}`);
  }
}

/**
 * Menu action: Append a new IMAGE_N_* block to the Configuration sheet.
 */
function addImageConfig() {
  const ui = SpreadsheetApp.getUi();
  try {
    new SheetService(SpreadsheetApp.getActiveSpreadsheet()).appendImageConfigBlock();
  } catch (err) {
    ui.alert(`Add image failed:\n\n${err && err.message ? err.message : String(err)}`);
  }
}

/**
 * Menu action: Generate slides from checked rows in the source sheet.
 * Only rows with "To generate" checked are generated.
 */
function generateSlides() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const config = new ConfigService(ss).getConfig();
    const sheetService = new SheetService(ss);

    // Ensure column A is "To generate" with centered checkboxes (adds/moves if needed).
    sheetService.ensureGenerateColumn(config.sourceSheetName, config.startRow);

    // Dataset contains ONLY checked rows and excludes the checkbox column from headers/rows.
    const dataset = sheetService.readDataset(config.sourceSheetName, config.startRow);

    const totalRowsInRange = dataset.totalRows;
    const checkedRows = dataset.checkedRows;

    const nonEmptyCheckedRows = dataset.rows.filter((r) => !Utils.isEmptyRow_(r)).length;

    if (nonEmptyCheckedRows === 0) {
      ui.alert(
        `No rows to generate.\n\n` +
        `➡️ Please check at least one box in the "${DB.GENERATE_HEADER}" column (and make sure the row is not empty).`
      );
      return;
    }

    const confirmMessage =
      `You are about to generate Google Slides.\n\n` +
      `Source sheet: ${config.sourceSheetName}\n` +
      `Start row: ${config.startRow}\n` +
      `Rows found (from start row): ${totalRowsInRange}\n` +
      `Rows checked ("${DB.GENERATE_HEADER}"): ${checkedRows}\n` +
      `Template slide index (blueprint): ${config.templateSlideIndex}\n\n` +
      `Proceed?`;

    const choice = ui.alert('Confirm generation', confirmMessage, ui.ButtonSet.OK_CANCEL);
    if (choice !== ui.Button.OK) return;

    const generator = new SlidesGenerator(config);
    const result = generator.generate(dataset);

    ui.alert(
      `Done ✅\nSlides generated: ${result.slidesGenerated}\nFile name: ${result.fileName}\nFile ID: ${result.fileId}`
    );
  } catch (err) {
    ui.alert(`Generation failed:\n\n${err && err.message ? err.message : String(err)}`);
  }
}
