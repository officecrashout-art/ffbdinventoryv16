/**
 * FASHION FIZZ BD - RETURNS & EXCHANGES MODULE
 */

function retShowUI() {
  const html = HtmlService.createTemplateFromFile('returns')
    .evaluate()
    .setTitle('Returns & Exchanges')
    .setWidth(1250)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function retGetStartupData() {
  return {
    orders: soGetRangeDataAsObjects('RANGESO'),
    details: soGetRangeDataAsObjects('RANGESD'),
    inventory: soGetRangeDataAsObjects('RANGEINVENTORYITEMS')
  };
}

/**
 * PROCESS RETURN OR EXCHANGE
 */
/**
 * PROCESS RETURN OR EXCHANGE
 */
function retProcessTransaction(data) {
  const ss = SpreadsheetApp.getActive();
  const soSheet = ss.getSheetByName('SalesOrders');
  const sdSheet = ss.getSheetByName('SalesDetails');
  const returnsSheet = ss.getSheetByName('Returns'); // New recommended sheet
  const invData = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');
  const timestamp = new Date();
  
  try {
    // 1. HANDLE RETURN (Add Stock back)
    _syncSalesStockSafe(data.returnItemId, data.returnSize, -data.returnQty, invData);
    
    // Log Return in Sales Details (Aligning with your CSV structure)
    // Structure: Date, SO ID, Detail ID, Cust ID, Cust Name, State, City, Invoice, Item ID, Type, Cat, Subcat, Name, Size, Qty, Price, Ship, Total
    sdSheet.appendRow([
      timestamp, data.soId, "RET-" + Date.now(), data.custId, data.custName, 
      "", "", data.invoice, data.returnItemId, "Return", "", "", 
      data.returnItemName + " (Returned)", data.returnSize, -data.returnQty, 
      data.returnPrice, 0, -(data.returnQty * data.returnPrice)
    ]);

    // Optional: Log to a dedicated 'Returns' sheet if it exists
    if (returnsSheet) {
      returnsSheet.appendRow([timestamp, data.soId, data.custName, data.returnItemName, data.type, data.returnQty, data.returnPrice, data.reason]);
    }

    // 2. HANDLE EXCHANGE (Deduct New Stock)
    let exchangeTotal = 0;
    if (data.type === 'Exchange' && data.newItemId) {
      exchangeTotal = data.exchangeQty * data.newItemPrice;
      _syncSalesStockSafe(data.newItemId, data.newSize, data.exchangeQty, invData);

      sdSheet.appendRow([
        timestamp, data.soId, "EXC-" + Date.now(), data.custId, data.custName, 
        "", "", data.invoice, data.newItemId, "Exchange", "", "", 
        data.newItemName + " (Exchange)", data.newSize, data.exchangeQty, 
        data.newItemPrice, 0, exchangeTotal
      ]);
    }

    // 3. CALCULATE NET CHANGE & UPDATE FINANCIALS
    const returnValue = data.returnQty * data.returnPrice;
    const netChange = exchangeTotal - returnValue; 

    // Update Sales Order Header
    const soRows = soSheet.getDataRange().getValues();
    const headers = soRows[0];
    const rowIdx = soRows.findIndex(r => r[headers.indexOf('SO ID')] === data.soId);
    if (rowIdx > 0) {
      const r = rowIdx + 1;
      const amountCol = headers.indexOf('Total SO Amount') + 1;
      const balCol = headers.indexOf('SO Balance') + 1;
      const currentTotal = Number(soRows[rowIdx][headers.indexOf('Total SO Amount')]);
      const currentBal = Number(soRows[rowIdx][headers.indexOf('SO Balance')]);
      
      soSheet.getRange(r, amountCol).setValue(currentTotal + netChange);
      soSheet.getRange(r, balCol).setValue(currentBal + netChange);
    }
    
    // Update Customer Balance
    custUpdateCustomerFinancials(data.custId, netChange);
    
    SpreadsheetApp.flush();
    return { success: true };

  } catch (e) {
    return { error: e.message };
  }
}