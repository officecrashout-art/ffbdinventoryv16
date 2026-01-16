/**
 * FASHION FIZZ BD - GLOBAL CONFIG & ROUTING
 */

function doGet(e) {
  const page = e.parameter.page || 'dashboard';
  
  const pageTemplates = {
    'dashboard': 'index',
    'inventory': 'inventory',
    'suppliers': 'suppliers',
    'customers': 'customers',
    'purchases': 'purchases',
    'sales': 'sales',
    'sales_details': 'sales_details',
    'receipts': 'receipts',
    'payments': 'payments',
    'reports': 'reports', // <--- ADDED COMMA HERE
    'returns': 'returns'
  };
  
  const contentTemplate = pageTemplates[page] || 'index';
  const template = HtmlService.createTemplateFromFile('template');
  
  template.contentTemplate = contentTemplate;
  template.getScriptUrl = getScriptUrl;
  template.currentPage = page;
  
  return template.evaluate()
    .setTitle('FashionFizzBD Inventory App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * SHARED UTILITY: GET RANGE DATA AS OBJECTS
 * This is used by all modules to fetch data quickly.
 */
function soGetRangeDataAsObjects(rangeName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getRangeByName(rangeName);
  if (!range) throw new Error(`Named range "${rangeName}" not found.`);
  
  const values = range.getValues();
  if (values.length < 2) return [];
  
  const headers = values[0];
  const rows = values.slice(1).filter(r => r.some(cell => cell !== '' && cell !== null));

  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = r[i];
      // Auto-format dates for the UI
      if (val instanceof Date) {
        val = Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), 'MM/dd/yyyy');
      }
      obj[h] = val;
    });
    return obj;
  });
}

/**
 * SHARED UTILITY: UNIQUE DIMENSION VALUES
 * Used for dropdowns like State, City, Item Type, Category, etc.
 */
function _getUniqueDimension(colHeader) {
  const data = soGetRangeDataAsObjects('RANGEDIMENSIONS');
  return [...new Set(data.map(r => r[colHeader]).filter(v => v && v.trim() !== ""))];
}

/**
 * SHARED UTILITY: ADD DIMENSION VALUE
 * Dynamically adds a new value to your Dimensions sheet.
 */
function _addDimensionValue(colHeader, value) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEDIMENSIONS');
  const sheet = range.getSheet();
  const headers = range.getValues()[0];
  const colIdx = headers.indexOf(colHeader);
  if (colIdx === -1) return;
  
  // Find first empty cell in that specific column to keep the sheet clean
  const colData = sheet.getRange(range.getRow(), range.getColumn() + colIdx, sheet.getLastRow()).getValues();
  let firstEmptyRow = range.getRow() + colData.length;
  for (let i = 0; i < colData.length; i++) {
    if (colData[i][0] === "") {
      firstEmptyRow = range.getRow() + i;
      break;
    }
  }
  
  sheet.getRange(firstEmptyRow, range.getColumn() + colIdx).setValue(value);
}