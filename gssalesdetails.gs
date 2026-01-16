/**
 * FASHION FIZZ BD - SALES HISTORY & ORDER MANAGEMENT
 * This file handles fetching orders, updating statuses, and retrieving 
 * detailed item information including Return/Exchange data.
 */

/**
 * Opens the Sales Management UI
 */
function sdShowSalesDetailsUI() {
  const html = HtmlService.createTemplateFromFile('sales_details')
    .evaluate()
    .setTitle('Sales Management')
    .setWidth(1250)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Fetches Sales Orders list for the main table based on time filters
 */
function sdGetOrders(filterType, startDate, endDate) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('SalesOrders');
    if (!sheet) return { error: "Sheet 'SalesOrders' not found" };

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const col = {};
    headers.forEach((h, i) => col[h.trim()] = i);

    const now = new Date();
    const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");

    let filtered = data.filter(r => r[col['SO ID']]);

    if (filterType !== 'all') {
      filtered = filtered.filter(r => {
        const orderDate = new Date(r[col['SO Date']]);
        const orderDateStr = Utilities.formatDate(orderDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        if (filterType === 'day') return orderDateStr === todayStr;
        if (filterType === 'month') return orderDate.getMonth() === now.getMonth() && orderDate.getFullYear() === now.getFullYear();
        if (filterType === 'year') return orderDate.getFullYear() === now.getFullYear();
        if (filterType === 'range' && startDate && endDate) {
          const s = new Date(startDate);
          const e = new Date(endDate);
          e.setHours(23, 59, 59); 
          return orderDate >= s && orderDate <= e;
        }
        return true;
      });
    }

    return filtered.map(r => ({
      date: r[col['SO Date']] instanceof Date ? Utilities.formatDate(r[col['SO Date']], Session.getScriptTimeZone(), "dd/MM/yyyy") : r[col['SO Date']],
      id: r[col['SO ID']],
      customer: r[col['Customer Name']],
      invoice: r[col['Invoice Num']],
      total: parseFloat(r[col['Total SO Amount']]) || 0,
      status: r[col['Receipt Status']] || 'Unpaid',
      shipping: r[col['Shipping Status']] || 'Pending'
    })).reverse();
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * Updates status fields (Payment or Shipping) in the SalesOrders sheet
 */
function sdUpdateStatus(soId, fieldName, newValue) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('SalesOrders');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIdx = headers.indexOf('SO ID');
    const fieldIdx = headers.indexOf(fieldName);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === soId) {
        sheet.getRange(i + 1, fieldIdx + 1).setValue(newValue);
        return { success: true };
      }
    }
    return { error: "Order ID not found" };
  } catch (e) {
    return { error: e.message };
  }
}

/**
 * REFINED: Fetches item-level details for a specific order.
 * This function maps the sheet data to the object keys used in your HTML
 * and looks up return reasons from the 'Returns' sheet.
 */
function sdGetOrderItems(soId) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sdSheet = ss.getSheetByName('SalesDetails');
    const returnsSheet = ss.getSheetByName('Returns');
    
    if (!sdSheet) return [];

    const sdData = sdSheet.getDataRange().getValues();
    const sdHeaders = sdData[0].map(h => h.toString().trim());
    const col = {};
    sdHeaders.forEach((h, i) => col[h] = i);

    const orderLines = sdData.filter(row => row[col['SO ID']] === soId);

    // Load Reasons from Returns Sheet
    let returnReasons = {};
    if (returnsSheet) {
      const retData = returnsSheet.getDataRange().getValues();
      const retHeaders = retData[0].map(h => h.toString().trim());
      const rCol = {};
      retHeaders.forEach((h, i) => rCol[h] = i);
      
      const orderReturns = retData.filter(row => row[rCol['Order ID']] === soId);
      orderReturns.forEach(row => {
        // Key is the Item Name
        const itemName = row[rCol['Item Name']].toString().trim();
        const reason = row[rCol['Reason']].toString().trim();
        returnReasons[itemName] = reason;
      });
    }

    return orderLines.map(row => {
      const fullName = row[col['Item Name']].toString().trim();
      const type = row[col['Item Type']];
      
      // CLEANING STEP: Remove " (Returned)" or " (Exchange)" to match the Returns sheet
      let baseName = fullName.replace(/\s\(Returned\)$|\s\(Exchange\)$/, "").trim();
      
      let foundReason = "";
      if (type === 'Return' || type === 'Exchange') {
        // Try matching with the cleaned baseName first, then the fullName
        foundReason = returnReasons[baseName] || returnReasons[fullName] || "Reason not found in sheet";
      }

      return {
        name: fullName,
        size: row[col['Size']] || '-',
        qty: Math.abs(parseFloat(row[col['QTY Sold']])) || 0,
        price: parseFloat(row[col['Unit Price']]) || 0,
        total: parseFloat(row[col['Total Sales Price']]) || 0,
        type: type, 
        Reason: foundReason 
      };
    });
  } catch (e) {
    return { error: e.message };
  }
}