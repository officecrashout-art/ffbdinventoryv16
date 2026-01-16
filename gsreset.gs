/**
 * FASHION FIZZ BD - SYSTEM RESET MODULE
 * Clears all data rows but preserves headers.
 */

function systemShowResetUI() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '☢️ SYSTEM RESET REQUESTED',
    'This will DELETE all Sales, Purchases, Customers, Suppliers, and Inventory Items. \n\nAre you absolutely sure you want to start from scratch?',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    _performSystemReset();
    ui.alert('System Reset Complete. Your database is now fresh and empty.');
  }
}

/**
 * Internal function to wipe sheets
 * Iterates through all named ranges and clears rows 2 downwards
 */
function _performSystemReset() {
  const ss = SpreadsheetApp.getActive();
  
  // List of all Named Ranges that contain data
  const dataRanges = [
    'RANGESD',            // Sales Details
    'RANGESO',            // Sales Orders
    'RANGEPD',            // Purchase Details
    'RANGEPO',            // Purchase Orders
    'RANGECUSTOMERS',     // Customers
    'RANGESUPPLIERS',    // Suppliers
    'RANGEINVENTORYITEMS',// Inventory
    'RANGEPAYMENTS',      // Payments to Suppliers
    'RANGERECEIPTS'       // Receipts from Customers
  ];

  dataRanges.forEach(rangeName => {
    try {
      const range = ss.getRangeByName(rangeName);
      if (range) {
        const sheet = range.getSheet();
        const startRow = range.getRow();
        const lastRow = sheet.getLastRow();
        
        // If there is data below the header (Row 1)
        if (lastRow > startRow) {
          // Clear everything from the row after the header to the bottom of the sheet
          sheet.getRange(startRow + 1, 1, lastRow - startRow, sheet.getLastColumn()).clearContent();
          
          // Optional: Reset "0" values for specific columns if needed, 
          // but clearContent is usually enough to "zero out" the system.
        }
      }
    } catch (e) {
      Logger.log('Could not reset range: ' + rangeName + ' - ' + e.toString());
    }
  });

  // Optional: Reset Dimensions? 
  // Usually, you want to KEEP your Brands and Categories even if you delete products.
  // If you want to wipe those too, uncomment the lines below:
  /*
  const dimRange = ss.getRangeByName('RANGEDIMENSIONS');
  if (dimRange) {
    const dimSheet = dimRange.getSheet();
    dimSheet.getRange(2, 1, dimSheet.getLastRow(), dimSheet.getLastColumn()).clearContent();
  }
  */

  SpreadsheetApp.flush(); // Ensure all changes are applied immediately
}