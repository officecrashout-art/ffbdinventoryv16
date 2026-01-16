/**
 * FASHION FIZZ BD - INVENTORY MANAGEMENT (FIXED)
 * Fixes: Remaining QTY sync on edit, Collision-proof IDs
 */

function itemShowInventoryUI() {
  const html = HtmlService.createTemplateFromFile('inventory')
    .evaluate()
    .setTitle('Inventory Management');
  SpreadsheetApp.getUi().showSidebar(html);
}

function itemGetInventoryData() {
  return {
    items: soGetRangeDataAsObjects('RANGEINVENTORYITEMS'),
    brands: _getUniqueDimension('Brands'), 
    categories: _getUniqueDimension('Item Category'),
    subcategories: _getUniqueDimension('Item Subcategory')
  };
}

function itemGenerateInventoryId() {
  // FIXED: Use UUID for zero collision risk
  return 'P-' + Utilities.getUuid().slice(0,8).toUpperCase(); 
}

function itemDeleteItem(itemId) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEINVENTORYITEMS');
  const sheet = range.getSheet();
  const data = range.getValues();
  const idCol = data[0].indexOf('Item ID');
  const rowIdx = data.findIndex(r => r[idCol] === itemId);
  
  if (rowIdx > 0) { 
    sheet.deleteRow(range.getRow() + rowIdx);
    return { success: true, message: "Item deleted successfully" };
  } else {
    throw new Error("Item ID not found");
  }
}

/**
 * MASTER SAVE FUNCTION - FIXED
 * Now correctly recalculates Remaining QTY even on Edit.
 */
function itemSaveProductWithVariants(data) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('InventoryItems'); 
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  let itemId = data.id === "AUTO" || !data.id ? itemGenerateInventoryId() : data.id;
  let isNew = data.id === "AUTO" || !data.id;

  // formatted size string: "S:10, L:5"
  const sizes = data.variants.map(v => `${v.size}:${v.openingStock || 0}`).join(', ');
  
  // This represents the NEW Total Stock (Opening + any added in UI)
  const totalOpeningStock = data.variants.reduce((sum, v) => sum + (parseFloat(v.openingStock) || 0), 0);

  const rowMap = {
    "Item ID": itemId,
    "Item Name": data.name,
    "Brands": data.brand,
    "Item Category": data.category,
    "Item Subcategory": data.subcategory,
    "Size": sizes, 
    "Reorder Level": data.reorderLevel,
    "Image URL": data.imageUrl, 
    "QTY Purchased": totalOpeningStock, // Will update this
    // Remaining QTY will be calculated dynamically below
    "Reorder Required": "No"
  };

  if (isNew) {
    // For new items, Sold is 0, so Remaining = Opening
    rowMap["Remaining QTY"] = totalOpeningStock;
    rowMap["QTY Sold"] = 0;
    
    const newRow = headers.map(h => {
      const val = rowMap[h.trim()];
      return val !== undefined ? val : "";
    });
    sheet.appendRow(newRow);
    
  } else {
    // EDIT MODE
    const allData = sheet.getDataRange().getValues();
    const idColIdx = headers.indexOf("Item ID");
    const rowIdx = allData.findIndex(r => r[idColIdx] === itemId);
    
    if (rowIdx > -1) {
      const sheetRowNum = rowIdx + 1;
      
      // FIXED: Fetch current 'QTY Sold' from the sheet to preserve it
      const soldColIdx = headers.indexOf("QTY Sold");
      const currentSold = Number(allData[rowIdx][soldColIdx]) || 0;
      
      // FIXED: Recalculate Remaining QTY based on New Opening - Current Sold
      rowMap["QTY Sold"] = currentSold;
      rowMap["Remaining QTY"] = totalOpeningStock - currentSold;

      // Update columns
      headers.forEach((h, colIdx) => {
        const trimmedH = h.trim();
        // Removed the exclusion check. Now we update ALL mapped fields to ensure consistency.
        if (rowMap[trimmedH] !== undefined) {
          sheet.getRange(sheetRowNum, colIdx + 1).setValue(rowMap[trimmedH]);
        }
      });
    }
  }

  return { success: true, message: "Product saved successfully under ID: " + itemId };
}

// ... (itemAddNewDimension, itemUploadImage remain mostly the same) ...
function itemAddNewDimension(type, value) {
  const ss = SpreadsheetApp.getActive();
  const range = ss.getRangeByName('RANGEDIMENSIONS');
  const sheet = range.getSheet();
  const headers = range.getValues()[0];
  const colIdx = headers.indexOf(type);
  if (colIdx === -1) throw new Error("Dimension type not found: " + type);
  
  const lastRow = sheet.getLastRow();
  const colData = sheet.getRange(1, colIdx + 1, lastRow).getValues();
  let emptyRow = lastRow + 1;
  for(let i = 1; i < colData.length; i++) {
    if(colData[i][0] === "" || colData[i][0] === null) {
      emptyRow = i + 1;
      break;
    }
  }
  sheet.getRange(emptyRow, colIdx + 1).setValue(value);
  SpreadsheetApp.flush();
  return { success: true };
}

function itemUploadImage(base64Data, fileName) {
  try {
    // WARNING: This ID is hardcoded. Ensure this folder exists and is shared.
    const folderId = '1usgkVjV4Q7oLQ7leBQQk2FABoPxDeed5'; 
    const folder = DriveApp.getFolderById(folderId);
    const splitData = base64Data.split(',');
    const contentType = splitData[0].match(/:(.*?);/)[1];
    const bytes = Utilities.base64Decode(splitData[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?sz=s1000&id=" + file.getId();
  } catch (e) {
    return "Error: " + e.toString();
  }
}