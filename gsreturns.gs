

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

/**
 * Updates InventoryItems stock safely for one item+size.
 *
 * qtyDelta meaning:
 *   +qtyDelta  => deduct stock (like selling / exchange)
 *   -qtyDelta  => add stock back (like return)
 */
function _syncSalesStockSafe(itemId, size, qtyDelta) {
  const ss = SpreadsheetApp.getActive();
  const invSheet = ss.getSheetByName('InventoryItems');
  if (!invSheet) throw new Error('InventoryItems sheet not found');

  const data = invSheet.getDataRange().getValues();
  const headers = data[0];

  const idIdx   = headers.indexOf('Item ID');
  const sizeIdx = headers.indexOf('Size');
  const soldIdx = headers.indexOf('QTY Sold');
  const remIdx  = headers.indexOf('Remaining QTY');

  if (idIdx === -1)   throw new Error('InventoryItems missing column: Item ID');
  if (sizeIdx === -1) throw new Error('InventoryItems missing column: Size');
  if (soldIdx === -1) throw new Error('InventoryItems missing column: QTY Sold');
  if (remIdx === -1)  throw new Error('InventoryItems missing column: Remaining QTY');

  const rowIndex = data.findIndex((r, i) => i > 0 && String(r[idIdx]) === String(itemId));
  if (rowIndex === -1) throw new Error('Item not found in Inventory: ' + itemId);

  const sheetRow = rowIndex + 1;
  const row = data[rowIndex];

  // Parse "S:10, M:5" -> {S:10, M:5}
  const currentSizeStr = String(row[sizeIdx] || '');
  const sizeMap = {};
  if (currentSizeStr.trim()) {
    currentSizeStr.split(',').forEach(part => {
      const [kRaw, vRaw] = part.split(':');
      const k = (kRaw || '').trim();
      if (!k) return;
      sizeMap[k] = Number((vRaw || '0').trim()) || 0;
    });
  }

  if (!size || !sizeMap.hasOwnProperty(size)) {
    throw new Error(`Size "${size}" not found for item ${itemId}. Check InventoryItems "Size" format.`);
  }

  const q = Number(qtyDelta) || 0;

  // If deducting stock, ensure enough is available
  if (q > 0 && sizeMap[size] < q) {
    throw new Error(`Insufficient stock for item ${itemId} (Size: ${size}). Available: ${sizeMap[size]}`);
  }

  // Apply change: deduct when +, add back when -
  sizeMap[size] = sizeMap[size] - q;

  // Rebuild
  const newSizeStr = Object.entries(sizeMap).map(([k, v]) => `${k}:${v}`).join(', ');
  const newRemaining = Object.values(sizeMap).reduce((a, b) => a + b, 0);

  const currentSold = Number(row[soldIdx]) || 0;
  const newSold = currentSold + q; // selling/exchange increases sold, return decreases sold
  if (newSold < 0) {
    throw new Error(`Sold qty would go negative for item ${itemId}. Check return qty vs sold.`);
  }

  invSheet.getRange(sheetRow, sizeIdx + 1).setValue(newSizeStr);
  invSheet.getRange(sheetRow, soldIdx + 1).setValue(newSold);
  invSheet.getRange(sheetRow, remIdx + 1).setValue(newRemaining);
}
