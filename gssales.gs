/**
 * FASHION FIZZ BD - SALES ENGINE (TRANSACTIONAL & OPTIMIZED)
 * Fixes: Ghost Orders, Script Timeouts, Race Conditions
 */

function soShowSalesUI() {
  const html = HtmlService.createTemplateFromFile('sales')
    .evaluate()
    .setTitle('New Sales Order')
    .setWidth(1250)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getSalesStartupData() {
  try {
    return {
      customers: soGetRangeDataAsObjects('RANGECUSTOMERS') || [],
      items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS') || [], 
      sales: soGetRangeDataAsObjects('RANGESO') || [],
      cities: _getUniqueDimension('City') || []
    };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * MASTER SAVE FUNCTION - OPTIMIZED
 */
function soSaveOrder(soData, items, customer) {
  // 1. GLOBAL LOCK: Acquire ONE lock for the entire batch
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Wait up to 30s
  } catch (e) {
    throw new Error('System is busy processing another order. Please try again.');
  }

  try {
    const ss = SpreadsheetApp.getActive();
    const soSheet = ss.getSheetByName('SalesOrders');
    const sdSheet = ss.getSheetByName('SalesDetails');
    const invSheet = ss.getSheetByName('InventoryItems');
    
    // 2. Handle Customer (Create if New)
    let custId = customer.id;
    if (customer.isNew) {
       // Note: This helper should ideally be inside the lock or safe to duplicate
       custId = custAddNewCustomer({
        name: customer.name,
        contact: customer.contact,
        city: customer.city,
        address: customer.address
      });
    }

    // 3. READ INVENTORY ONCE (Single Read)
    const invData = invSheet.getDataRange().getValues();
    const headers = invData[0];
    const idIdx = headers.indexOf('Item ID');
    const sizeIdx = headers.indexOf('Size');
    const soldIdx = headers.indexOf('QTY Sold');
    const remIdx = headers.indexOf('Remaining QTY');
    
    const detailsRows = [];
    const inventoryUpdates = []; // Store {row: x, col: y, val: z}

    // 4. VALIDATE & PREPARE ALL ITEMS
    for (let item of items) {
      // Find item in loaded data
      const rowIndex = invData.findIndex(r => r[idIdx] === item.id);
      if (rowIndex === -1) throw new Error(`Item ${item.name} not found in inventory.`);
      
      const currentRow = invData[rowIndex];
      const sheetRowNum = rowIndex + 1;
      
      // Parse Size String
      const currentSizeStr = String(currentRow[sizeIdx]);
      let sizeMap = {};
      if (currentSizeStr) {
        currentSizeStr.split(',').forEach(part => {
          let [sName, sQty] = part.split(':').map(x => x ? x.trim() : "");
          if(sName) sizeMap[sName] = Number(sQty || 0);
        });
      }

      // Check Stock Availability
      if (item.size && sizeMap.hasOwnProperty(item.size)) {
         if (sizeMap[item.size] < Number(item.qty)) {
           throw new Error(`Insufficient stock for ${item.name} (Size: ${item.size}). Available: ${sizeMap[item.size]}`);
         }
         // Deduct from memory map
         sizeMap[item.size] -= Number(item.qty);
      } else {
         // Fallback for items without specific size tracking if needed, or throw error
         // Assuming all items must have valid size if size is selected
      }

      // Rebuild Strings & Values
      const newSizeStr = Object.entries(sizeMap).map(([k,v]) => `${k}:${v}`).join(', ');
      const newTotalStock = Object.values(sizeMap).reduce((a, b) => a + b, 0);
      const currentSold = Number(currentRow[soldIdx] || 0);
      const newSold = currentSold + Number(item.qty);

      // Queue Inventory Updates (We will write these later)
      inventoryUpdates.push({ row: sheetRowNum, col: sizeIdx + 1, val: newSizeStr });
      inventoryUpdates.push({ row: sheetRowNum, col: soldIdx + 1, val: newSold });
      inventoryUpdates.push({ row: sheetRowNum, col: remIdx + 1, val: newTotalStock });
      
      // Update the local invData array in case multiple items in this order use the same product!
      invData[rowIndex][sizeIdx] = newSizeStr;
      invData[rowIndex][soldIdx] = newSold;
      invData[rowIndex][remIdx] = newTotalStock;

      // Prepare Sales Detail Row
      detailsRows.push([
        new Date(), soData.id, "SD-" + Utilities.getUuid(), custId, customer.name,
        customer.state || "", customer.city, soData.invoice, item.id, item.category || "",
        item.category || "", item.subcategory || "", item.name, item.size || "",
        item.qty, item.price, item.price, 0, 0, item.price, item.ship, item.total
      ]);
    }

    // 5. WRITE EVERYTHING (Transaction Commit)
    
    // A. Update Inventory
    inventoryUpdates.forEach(u => {
      invSheet.getRange(u.row, u.col).setValue(u.val);
    });

    // B. Write Header
    soSheet.appendRow([
      new Date(), soData.id, custId, customer.name, soData.invoice, 
      customer.state || "", customer.city, soData.totalAmount, 0, 
      soData.totalAmount, "Unpaid", "Pending"
    ]);

    // C. Write Details
    if (detailsRows.length > 0) {
      sdSheet.getRange(sdSheet.getLastRow() + 1, 1, detailsRows.length, detailsRows[0].length)
             .setValues(detailsRows);
    }
    
    // D. Update Customer Financials
    custUpdateCustomerFinancials(custId, soData.totalAmount);
    
    SpreadsheetApp.flush(); // Force save before releasing lock

  } catch (err) {
    console.error("Order Failed: " + err.message);
    throw err; // Re-throw to alert user
  } finally {
    lock.releaseLock();
  }

  return { success: true, message: "Order " + soData.id + " saved successfully!" };
}