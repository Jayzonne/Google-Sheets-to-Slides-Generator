/**
 * @file Menu.gs
 * Google Sheets UI menu.
 */

class MenuController {
  static onOpen() {
    SpreadsheetApp.getUi()
      .createMenu(APP.MENU_NAME)
      // Requested order: Generate first, then Restructure
      .addItem('Generate slides', 'generateSlides')
      .addSeparator()
      .addItem('Add image', 'addImageConfig')
      .addSeparator()
      .addItem('Restruct the template', 'restructureTemplate')
      .addToUi();
  }
}
