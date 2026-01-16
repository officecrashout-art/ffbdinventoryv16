function dimensionsShowUI() {
  const html = HtmlService.createTemplateFromFile('dimensions').evaluate()
      .setTitle('Manage Categories & Brands')
      .setWidth(1000).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Dimension Manager');
}

function dimGetAllData() {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEDIMENSIONS');
  const values = range.getValues();
  const headers = values[0];
  const data = {};
  
  headers.forEach((h, colIdx) => {
    data[h] = values.slice(1).map(row => row[colIdx]).filter(v => v !== "" && v !== null);
  });
  
  return data;
}
// Adds a menu item to Google Sheets
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🚀 FashionFizz Tools')
      .addItem('Manage Inventory', 'itemShowInventoryUI')
      .addItem('Manage Sales', 'soShowSalesUI')
      .addItem('Category & Brand Manager', 'dimensionsShowUI')
      .addSeparator()
      .addItem('⚠️ SYSTEM RESET (Fresh Start)', 'systemShowResetUI') // Added this line
      .addToUi();
}

// Function to actually show the modal
function dimensionsShowUI() {
  const html = HtmlService.createTemplateFromFile('dimensions').evaluate()
      .setTitle('Dimension Manager')
      .setWidth(600).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function dimUpdateList(type, newList) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEDIMENSIONS');
  const sheet = range.getSheet();
  const headers = range.getValues()[0];
  const colIdx = headers.indexOf(type);
  
  if (colIdx === -1) return;
  
  // Clear the existing column
  sheet.getRange(range.getRow() + 1, range.getColumn() + colIdx, sheet.getMaxRows(), 1).clearContent();
  
  // Set new values
  if (newList.length > 0) {
    const output = newList.map(val => [val]);
    sheet.getRange(range.getRow() + 1, range.getColumn() + colIdx, output.length, 1).setValues(output);
  }
  
  return { success: true };
}