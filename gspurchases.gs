/**
 * FASHION FIZZ BD - PURCHASES (WITH DIRECT SUPPLIER CREATION)
 */

function poShowUI() {
  const html = HtmlService.createTemplateFromFile('purchases')
    .evaluate().setTitle('Purchase Orders').setWidth(1250).setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getPurchaseStartupData() {
  return {
    suppliers: soGetRangeDataAsObjects('RANGESUPPLIERS') || [],
    items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS') || [],
    pos: soGetRangeDataAsObjects('RANGEPO') || []
  };
}

/**
 * MASTER SAVE: Handles New Supplier Creation + PO Saving
 */
function soSaveOrUpdatePO(poData, items, supplier) {
  const ss = SpreadsheetApp.getActive();
  const poSheet = ss.getSheetByName('PurchaseOrders');
  const pdSheet = ss.getSheetByName('PurchaseDetails');
  
  // 1. Handle Supplier (Create if New)
  let finalSupId = poData.supId;
  let finalSupName = poData.supName;

  if (supplier && supplier.isNew) {
    finalSupId = _poAddNewSupplier(supplier);
    finalSupName = supplier.name;
  }

  // 2. Save PO Header
  // Columns: [Date, PO ID, Supplier ID, Supplier Name, Bill Num, State, City, Total Amount, Total Paid, PO Balance, PMT Status, Status]
  poSheet.appendRow([
    new Date(), 
    poData.id, 
    finalSupId, 
    finalSupName, 
    poData.billNum, 
    supplier.state || "", 
    supplier.city || "", 
    poData.total, 
    0, // Paid
    poData.total, // Balance
    "Unpaid", 
    "Pending"
  ]);

  // FORCE SAVE: Ensure header is recorded before processing details
  SpreadsheetApp.flush();

  // 3. Save Details & Sync Stock (BATCHED OPTIMIZATION)
  const detailRows = [];

  items.forEach(item => {
    try {
      // A. Prepare row data
      // Note: We use the exact column structure from your original function
      const row = [
        new Date(), 
        poData.id, 
        "D-" + Date.now() + Math.floor(Math.random()*100), 
        finalSupId, 
        finalSupName,
        supplier.state || "", 
        supplier.city || "", 
        poData.billNum, 
        item.id, 
        "", // Type
        item.category, 
        "", // Subcat
        item.name,
        item.qty, 
        item.cost, 
        item.total, 
        0, // Tax Rate
        0, // Tax Total
        item.cost, // Cost Incl Tax
        0, // Shipping
        item.total
      ];
      
      detailRows.push(row); // Add to container
      
      // B. Update Inventory Stock
      // We call this individually to ensure the Lock Service safeguards each specific item update
      _syncPurchaseStock(item.id, item.size, item.qty);

    } catch (err) {
      console.error("Error processing PO item " + item.name + ": " + err.message);
    }
  });

  // 4. WRITE ALL DETAILS IN ONE SHOT
  if (detailRows.length > 0) {
    const lastRow = pdSheet.getLastRow();
    pdSheet.getRange(lastRow + 1, 1, detailRows.length, detailRows[0].length)
           .setValues(detailRows);
  }

  // 5. Update Supplier Financials
  try {
    supUpdateFinancials(finalSupId, poData.total, 0);
  } catch (e) {
    console.error("Error updating supplier financials: " + e.message);
  }

  return { success: true, message: "PO Saved & Supplier Created!" };
}

/**
 * HELPER: Creates a new supplier row
 */
function _poAddNewSupplier(s) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Suppliers');
  
  // Generate ID
  const newId = "S" + Math.floor(10000 + Math.random() * 90000);
  
  sheet.appendRow([
    newId, 
    s.name, 
    s.contact, 
    s.email || "", 
    s.state || "", 
    s.city || "", 
    s.address || "", 
    0, 0, 0 // Financials
  ]);
  
  return newId;
}

/**
 * STOCK SYNC
 */
/**
 * STOCK SYNC WITH LOCK SERVICE
 */
function _syncPurchaseStock(itemId, sizeName, qtyPurchased) {
  const lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000); // Wait up to 30 seconds
  } catch (e) {
    throw new Error('Could not update stock (Server Busy). Try again.');
  }

  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('InventoryItems');
    // Important: We must re-fetch data inside the lock to get the latest version
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const rowIdx = data.findIndex(r => r[headers.indexOf('Item ID')] === itemId);
    if (rowIdx === -1) return;

    const row = rowIdx + 1;
    const sizeCol = headers.indexOf('Size') + 1;
    const purCol = headers.indexOf('QTY Purchased') + 1;
    const remCol = headers.indexOf('Remaining QTY') + 1;

    const currentStr = data[rowIdx][headers.indexOf('Size')] || "";
    let sizeMap = {};
    
    // Parse "S:10, M:5"
    if(currentStr) {
      currentStr.split(',').forEach(p => {
        let [k, v] = p.split(':');
        if(k) sizeMap[k.trim()] = Number(v || 0);
      });
    }

    // Add Stock
    if(sizeName) {
      sizeMap[sizeName] = (sizeMap[sizeName] || 0) + Number(qtyPurchased);
    }

    // Rebuild String
    const newStr = Object.entries(sizeMap).map(([k,v]) => `${k}:${v}`).join(', ');
    
    // Write Data
    sheet.getRange(row, sizeCol).setValue(newStr);
    
    // Update Totals
    const curPur = Number(data[rowIdx][headers.indexOf('QTY Purchased')] || 0);
    const curRem = Number(data[rowIdx][headers.indexOf('Remaining QTY')] || 0);
    
    sheet.getRange(row, purCol).setValue(curPur + Number(qtyPurchased));
    sheet.getRange(row, remCol).setValue(curRem + Number(qtyPurchased));

    SpreadsheetApp.flush();

  } finally {
    lock.releaseLock();
  }
}

function poCreateInventoryItem(item) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems');
  const newId = "P" + Math.floor(Math.random() * 100000);
  sheet.appendRow([newId, "", item.category, "", "", item.brand, item.name, 0, 0, 0, 5, "No", ""]);
  return { id: newId, name: item.name };
}