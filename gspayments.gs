/**
 * FASHION FIZZ BD - PAYMENTS (INTEGRATED)
 */

function ptShowUI() {
  const html = HtmlService.createTemplateFromFile('payments')
    .evaluate().setTitle('Vendor Payments').setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

function getPaymentStartupData() {
  return {
    suppliers: soGetRangeDataAsObjects('RANGESUPPLIERS'),
    pos: soGetRangeDataAsObjects('RANGEPO'),
    payments: soGetRangeDataAsObjects('RANGEPAYMENTS')
  };
}

function ptSaveNewPayment(pay) {
  const ss = SpreadsheetApp.getActive();
  const ptSheet = ss.getSheetByName('Payments');
  const poSheet = ss.getSheetByName('PurchaseOrders');
  
  // FIX 1: Generate Date on Server for consistency
  const timestamp = new Date(); 

  // FIX 2: Ensure amount is a number
  const amountPaid = Number(pay.amount);

  // 1. Log Payment
  ptSheet.appendRow([
    timestamp,    // Use server timestamp
    pay.id, 
    pay.supId, 
    pay.supName,
    "",           // Placeholder (Invoice?)
    "",           // Placeholder
    pay.poId, 
    pay.bill, 
    pay.mode, 
    amountPaid    // Use parsed number
  ]);
  
  // 2. Update PO Balance & Status
  const poData = poSheet.getDataRange().getValues();
  const h = poData[0];
  const rIdx = poData.findIndex(r => r[h.indexOf('PO ID')] == pay.poId); // Use == for loose string/number matching
  
  if(rIdx > 0) {
    const r = rIdx + 1;
    const paidCol = h.indexOf('Total Paid') + 1;
    const balCol = h.indexOf('PO Balance') + 1;
    const statCol = h.indexOf('PMT Status') + 1;
    
    const curPaid = Number(poData[rIdx][h.indexOf('Total Paid')] || 0);
    const total = Number(poData[rIdx][h.indexOf('Total Amount')] || 0);
    
    const newPaid = curPaid + amountPaid;
    const newBal = total - newPaid;
    
    poSheet.getRange(r, paidCol).setValue(newPaid);
    poSheet.getRange(r, balCol).setValue(newBal > 0 ? newBal : 0);
    poSheet.getRange(r, statCol).setValue(newBal <= 0 ? 'Paid' : 'Partial');
  }

  // 3. Update Supplier Financials
  // FIX 3: Wrap this in a try-catch so the script finishes even if this function is missing or buggy
  try {
    if (typeof supUpdateFinancials === 'function') {
      supUpdateFinancials(pay.supId, 0, amountPaid);
    } else {
      console.warn("supUpdateFinancials function is missing.");
    }
  } catch (e) {
    console.error("Error updating supplier financials: " + e.message);
  }

  return { success: true };
}