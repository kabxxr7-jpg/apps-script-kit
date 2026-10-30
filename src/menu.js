/** Adds a "Sheet tools" menu when the spreadsheet opens (simple trigger; only runs in a bound script). */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sheet tools')
    .addItem('Remove duplicate rows', 'removeDuplicateRows')
    .addItem('Clean this sheet into "Cleaned"', 'cleanActiveSheet')
    .addToUi();
}
