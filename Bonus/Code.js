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
function auditSheet(sheet) {

  // get spreadsheet object
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // get sheet name
  const name = sheet.getName();

  // get current sheet dimensions
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();
  const totalCells = maxRows * maxCols;

  // put variables into object
  const sheetSize = {
    name: name,
    rows: maxRows,
    cols: maxCols,
    total: totalCells
  }

  // return object to function that called it
  return sheetSize;

}

/**
* Audits all Sheets and passes full data back to sidebar
*/
function auditAllSheets() {

  // get spreadsheet object
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // declare variables
  let output = '';
  let grandTotal = 0;

  // loop over sheets and get data for each
  sheets.forEach(sheet => {

    // get sheet results for the sheet
    const results = auditSheet(sheet);
    
    // create output string from results
    output = output + '<br><hr><br>Sheet: ' + results.name +
      '<br>Row count: ' + results.rows + 
      '<br>Column count: ' + results.cols +
      '<br>Total cells: ' + results.total + '<br>';

    // add results to grand total
    grandTotal = grandTotal + results.total;

  });

  // add grand total calculation to the output string
  output = output + '<br><hr><br>' + 
    'You have used ' + ((grandTotal / 5000000)*100).toFixed(2) + '% of your 5 million cell limit.';

  // pass results back to sidebar
  return output;

}