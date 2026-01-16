/**
 * FASHION FIZZ BD - RECEIPTS & PAYMENT PROCESSING
 * Handles recording payments and updating SO/Customer balances.
 */

function rcShowUI() {
  const html = HtmlService.createTemplateFromFile('receipts')
    .evaluate()
    .setTitle('Receipts Management')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * Fetches data for the Receipts UI
 */
function getReceiptStartupData() {
  const ss = SpreadsheetApp.getActive();
  // Ensure we get fresh data
  return {
    customers: soGetRangeDataAsObjects('RANGECUSTOMERS'),
    orders: soGetRangeDataAsObjects('RANGESO'),
    receipts: soGetRangeDataAsObjects('RANGERECEIPTS'),
    modes: _getUniqueDimension('PMT Mode') || ['Cash', 'Bank Transfer', 'Mobile Money', 'Cheque']
  };
}

/**
 * GENERATE UNIQUE TRANSACTION ID
 */
function rcGenerateTrxID() {
  return "REC-" + Math.floor(100000 + Math.random() * 900000);
}

/**
 * MASTER SAVE FUNCTION
 * 1. Logs Receipt
 * 2. Updates Sales Order (Status: Paid/Partial)
 * 3. Updates Customer (Balance)
 */
function rcSaveNewReceipt(rec) {
  const ss = SpreadsheetApp.getActive();
  const recSheet = ss.getSheetByName('Receipts');
  const soSheet = ss.getSheetByName('SalesOrders');
  const custSheet = ss.getSheetByName('Customers');
  
  if (!recSheet || !soSheet || !custSheet) throw new Error("Missing Sheets");

  // --- 1. LOG RECEIPT ---
  recSheet.appendRow([
    new Date(rec.date), 
    rec.trxId, 
    rec.custId, 
    rec.custName, 
    "", // State (optional)
    "", // City (optional) 
    rec.soId, 
    rec.invoice, 
    rec.mode, 
    rec.amount
  ]);

  // --- 2. UPDATE SALES ORDER ---
  const soData = soSheet.getDataRange().getValues();
  const soHeaders = soData[0];
  const soIdCol = soHeaders.indexOf('SO ID');
  const soRecCol = soHeaders.indexOf('Total Received');
  const soBalCol = soHeaders.indexOf('SO Balance');
  const soStatCol = soHeaders.indexOf('Receipt Status');
  const soTotalCol = soHeaders.indexOf('Total SO Amount');
  
  const soRowIdx = soData.findIndex(r => r[soIdCol] === rec.soId);
  
  if (soRowIdx > 0) {
    const r = soRowIdx + 1;
    const currentRec = Number(soData[soRowIdx][soRecCol] || 0);
    const totalAmt = Number(soData[soRowIdx][soTotalCol] || 0);
    
    const newRec = currentRec + Number(rec.amount);
    const newBal = totalAmt - newRec;
    
    // Determine Status
    let status = "Unpaid";
    if (newBal <= 0) status = "Paid";
    else if (newRec > 0) status = "Partial";
    
    // Write Updates
    soSheet.getRange(r, soRecCol + 1).setValue(newRec);
    soSheet.getRange(r, soBalCol + 1).setValue(newBal > 0 ? newBal : 0); // Prevent negative
    soSheet.getRange(r, soStatCol + 1).setValue(status);
  }

  // --- 3. UPDATE CUSTOMER BALANCE ---
  const custData = custSheet.getDataRange().getValues();
  const cHeaders = custData[0];
  const cIdCol = cHeaders.indexOf('Customer ID');
  const cRecCol = cHeaders.indexOf('Total Receipts');
  const cBalCol = cHeaders.indexOf('Balance Receivable');
  
  const cRowIdx = custData.findIndex(r => r[cIdCol] === rec.custId);
  
  if (cRowIdx > 0) {
    const r = cRowIdx + 1;
    const currentRec = Number(custData[cRowIdx][cRecCol] || 0);
    const currentBal = Number(custData[cRowIdx][cBalCol] || 0);
    
    custSheet.getRange(r, cRecCol + 1).setValue(currentRec + Number(rec.amount));
    custSheet.getRange(r, cBalCol + 1).setValue(currentBal - Number(rec.amount));
  }

  return { success: true, message: "Receipt Saved & Balances Updated!" };
}