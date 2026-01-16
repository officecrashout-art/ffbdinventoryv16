/**
 * FASHION FIZZ BD - SUPPLIER MANAGEMENT (ENHANCED)
 */

function supShowUI() {
  const html = HtmlService.createTemplateFromFile('suppliers')
    .evaluate()
    .setTitle('Supplier Directory')
    .setWidth(1250)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function supGetSuppliers() {
  return soGetRangeDataAsObjects('RANGESUPPLIERS') || [];
}

/**
 * FETCH SUPPLIER PRODUCTS (Aggregated from Purchase History)
 */
function supGetSupplierItems(supId) {
  // 1. Get all PO Details for this supplier
  const pd = soGetRangeDataAsObjects('RANGEPD').filter(r => r['Supplier ID'] === supId);
  
  // 2. Get all Inventory Items (to get current details like image, stock)
  const inventory = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');
  
  // 3. Map to unique items
  const itemMap = {};
  pd.forEach(r => {
    const id = r['Item ID'];
    if(!itemMap[id]) {
      // Find current inventory details to ensure we show latest data
      const invItem = inventory.find(i => i['Item ID'] === id) || {};
      
      itemMap[id] = {
        id: id,
        name: invItem['Item Name'] || r['Item Name'], // Prefer Inventory Name
        category: invItem['Item Category'] || r['Item Category'],
        brand: invItem['Brands'] || '',
        image: invItem['Image URL'] || '',
        totalPurchased: 0,
        costSum: 0
      };
    }
    // Aggregate Purchase Stats
    itemMap[id].totalPurchased += Number(r['QTY Purchased'] || 0);
    itemMap[id].costSum += Number(r['Total Purchase Price'] || 0);
  });
  
  // Calculate Avg Cost and Return array
  return Object.values(itemMap).map(i => ({
    ...i,
    avgCost: i.totalPurchased ? (i.costSum / i.totalPurchased).toFixed(2) : 0
  }));
}

/**
 * SAFE ITEM UPDATE
 * Only updates descriptive fields (Name, Brand, etc.)
 * Does NOT touch Stock/Size to prevent data loss.
 */
function supUpdateItemDetails(item) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idCol = headers.indexOf('Item ID');
  const rowIdx = data.findIndex(r => r[idCol] === item.id);
  
  if (rowIdx > 0) {
    const r = rowIdx + 1;
    // Helper to set value if column exists
    const set = (header, val) => {
      const col = headers.indexOf(header);
      if(col > -1) sheet.getRange(r, col+1).setValue(val);
    };

    if(item.name) set('Item Name', item.name);
    if(item.brand) set('Brands', item.brand);
    if(item.category) set('Item Category', item.category);
    if(item.image) set('Image URL', item.image);
    
    return { success: true };
  } else {
    throw new Error("Item not found in Inventory");
  }
}

// ... (Existing Functions: supSaveSupplier, supGetHistory, supUpdateFinancials remain unchanged)
function supSaveSupplier(data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Suppliers');
  if(data.isNew) {
    sheet.appendRow([
      "S" + Math.floor(10000 + Math.random() * 90000), data.name, data.contact, data.email, 
      data.state, data.city, data.address, 0, 0, 0
    ]);
  }
}

function supGetHistory(supId) {
  const pos = soGetRangeDataAsObjects('RANGEPO').filter(p => p['Supplier ID'] === supId);
  const pmts = soGetRangeDataAsObjects('RANGEPAYMENTS').filter(p => p['Supplier ID'] === supId);
  const history = [
    ...pos.map(x => ({ date: x['Date'], type: 'Purchase', ref: x['PO ID'], amount: x['Total Amount'], bal: x['PO Balance'] })),
    ...pmts.map(x => ({ date: x['Trx Date'], type: 'Payment', ref: x['Trx ID'], amount: x['Amount Paid'], bal: '-' }))
  ];
  return history.sort((a,b) => new Date(b.date) - new Date(a.date));
}

function supUpdateFinancials(supId, purchaseAmt, paidAmt) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Suppliers');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIdx = data.findIndex(r => r[headers.indexOf('Supplier ID')] === supId);
  if(rowIdx > 0) {
    const r = rowIdx + 1;
    const purCol = headers.indexOf('Total Purchases')+1;
    const paidCol = headers.indexOf('Total Payments')+1;
    const balCol = headers.indexOf('Balance Payable')+1;
    
    const curPur = Number(data[rowIdx][purCol-1]||0) + (purchaseAmt||0);
    const curPaid = Number(data[rowIdx][paidCol-1]||0) + (paidAmt||0);
    
    sheet.getRange(r, purCol).setValue(curPur);
    sheet.getRange(r, paidCol).setValue(curPaid);
    sheet.getRange(r, balCol).setValue(curPur - curPaid);
  }
}