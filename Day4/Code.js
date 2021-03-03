/**
 * Add custom menu to sheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Sheet Sizer')
      .addItem('Open Sheet Sizer', 'showSidebar')
      .addToUi();
}

/**
 * function to show sidebar
 */
function showSidebar() {
  
  // create sidebar with HTML Service
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('Sheet Sizer');
  
  // add sidebar to spreadsheet UI 
  SpreadsheetApp.getUi().showSidebar(html);

}

/**
* Get size data for a given sheet url
*/
function auditSheet() {

  // get Sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  // get sheet name
  const name = sheet.getName();

  // get current sheet dimensions
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const totalCells = maxRows * maxCols;

  // output
  return totalCells;

}