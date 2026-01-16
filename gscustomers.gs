/**
 * FASHION FIZZ BD - CUSTOMER MANAGEMENT
 */

function custShowCustomersUI() {
  const html = HtmlService.createTemplateFromFile('customers')
    .evaluate()
    .setTitle('Customer Directory')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function custGetCustomers() {
  return soGetRangeDataAsObjects('RANGECUSTOMERS') || [];
}

/**
 * Creates a new customer row.
 */
function custAddNewCustomer(c) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Customers');
  
  // Generate ID
  const existing = sh.getRange("A2:A").getValues().flat();
  let id = "C" + Math.floor(10000 + Math.random() * 90000);
  while(existing.includes(id)) { id = "C" + Math.floor(10000 + Math.random() * 90000); }

  sh.appendRow([
    id,
    c.name,
    c.contact,
    c.email || "",
    c.state || "",
    c.city || "",
    c.address || "",
    0, // Sales
    0, // Receipts
    0  // Balance
  ]);
  
  return id;
}

/**
 * Updates Customer Financials (Sales & Balance)
 */
function custUpdateCustomerFinancials(custId, saleAmount) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('Customers');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idCol = headers.indexOf('Customer ID');
  const saleCol = headers.indexOf('Total Sales');
  const balCol = headers.indexOf('Balance Receivable');
  
  const rowIndex = data.findIndex(r => r[idCol] === custId);
  
  if (rowIndex > 0) {
    const r = rowIndex + 1;
    const currSales = Number(data[rowIndex][saleCol] || 0);
    const currBal = Number(data[rowIndex][balCol] || 0);
    
    sheet.getRange(r, saleCol + 1).setValue(currSales + Number(saleAmount));
    sheet.getRange(r, balCol + 1).setValue(currBal + Number(saleAmount));
  }
}

/**
 * FETCH HISTORY for Customer Modal
 */
function custGetHistory(custId) {
  // Get Sales Orders for this customer
  const allOrders = soGetRangeDataAsObjects('RANGESO');
  const customerOrders = allOrders.filter(o => o['Customer ID'] === custId);
  // Sort Newest First
  return customerOrders.sort((a,b) => new Date(b['SO Date']) - new Date(a['SO Date']));
}