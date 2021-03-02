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