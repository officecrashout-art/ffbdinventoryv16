/**
 * FASHION FIZZ BD - DASHBOARD BACKEND
 * - Inventory value uses weighted avg purchase cost (fallback: last unit cost)
 * - Real profit uses SalesDetails revenue - COGS (supports returns)
 * - Flags missing cost items
 */
function getDashboardData() {
  const soData  = soGetRangeDataAsObjects('RANGESO') || [];
  const invData = soGetRangeDataAsObjects('RANGEINVENTORYITEMS') || [];
  const pdData  = soGetRangeDataAsObjects('RANGEPD') || []; // PurchaseDetails

  // --- Helpers ---
  const parseVal = (v) => {
    if (typeof v === 'number') return v;
    if (v === null || v === undefined || v === '') return 0;
    const clean = v.toString().replace(/[^0-9.-]/g, '');
    return parseFloat(clean) || 0;
  };

  const parseDateISO = (v) => {
    if (!v) return null;
    const d = new Date(v);
    if (isNaN(d)) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };

  const parseDateMs = (v) => {
    const d = new Date(v);
    const ms = d.getTime();
    return isNaN(ms) ? 0 : ms;
  };

  // Safe SalesDetails fetch (named range OR sheet fallback)
  const getSalesDetailsRows = () => {
    const tryRanges = ['RANGESD', 'RANGESALESDETAILS', 'RANGESODETAILS'];
    for (const rName of tryRanges) {
      try {
        const rows = soGetRangeDataAsObjects(rName);
        if (rows && rows.length) return rows;
      } catch (e) {}
    }
    try {
      const ss = SpreadsheetApp.getActive();
      const sh = ss.getSheetByName('SalesDetails');
      if (!sh) return [];
      const values = sh.getDataRange().getValues();
      if (values.length < 2) return [];
      const headers = values.shift().map(h => h.toString().trim());
      return values.map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      });
    } catch (e) {
      return [];
    }
  };

  const sdData = getSalesDetailsRows();

  // 1) Basic KPIs
  const totalSales = soData.reduce((sum, r) => sum + parseVal(r['Total SO Amount']), 0);
  const totalReceived = soData.reduce((sum, r) => sum + parseVal(r['Total Received']), 0);
  const totalDue = soData.reduce((sum, r) => sum + parseVal(r['SO Balance']), 0);

  // 2) Build cost maps from PurchaseDetails
  // Weighted avg cost: SUM(Total Purchase Price) / SUM(QTY Purchased)
  // Fallback last unit cost: latest Date -> Unit Cost (or Total Purchase Price / QTY Purchased)
  const avgCostMap = {};   // { id: { qty, total, avg } }
  const lastCostMap = {};  // { id: { ms, unit } }

  pdData.forEach(r => {
    const id = (r['Item ID'] || '').toString().trim();
    if (!id) return;

    const qtyPurchased = parseVal(r['QTY Purchased']);
    const totalPurchase = parseVal(r['Total Purchase Price']);
    const unitCostRaw = parseVal(r['Unit Cost']);
    const ms = parseDateMs(r['Date'] || r['PO Date'] || r['Created At'] || r['Timestamp']);

    // weighted avg
    if (qtyPurchased > 0) {
      if (!avgCostMap[id]) avgCostMap[id] = { qty: 0, total: 0, avg: 0 };
      avgCostMap[id].qty += qtyPurchased;
      avgCostMap[id].total += totalPurchase;
    }

    // last unit cost fallback (prefer explicit Unit Cost, else derive)
    let derivedUnit = 0;
    if (unitCostRaw > 0) derivedUnit = unitCostRaw;
    else if (qtyPurchased > 0 && totalPurchase > 0) derivedUnit = totalPurchase / qtyPurchased;

    if (derivedUnit > 0) {
      const prev = lastCostMap[id];
      if (!prev || ms >= prev.ms) {
        lastCostMap[id] = { ms: ms, unit: derivedUnit };
      }
    }
  });

  Object.keys(avgCostMap).forEach(id => {
    avgCostMap[id].avg = avgCostMap[id].qty > 0 ? (avgCostMap[id].total / avgCostMap[id].qty) : 0;
  });

  // For “missing costs” reporting (unique items)
  const missingCostSet = new Set();
  const missingCostItems = []; // {id, name, where}

  const noteMissing = (itemId, itemName, where) => {
    const id = (itemId || '').toString().trim();
    if (!id) return;
    const key = `${id}||${where}`;
    if (missingCostSet.has(key)) return;
    missingCostSet.add(key);
    missingCostItems.push({ id, name: itemName || id, where });
  };

  const getUnitCost = (itemId, itemName, whereForMissing) => {
    const id = (itemId || '').toString().trim();
    if (!id) return 0;

    // 1) weighted avg cost
    if (avgCostMap[id] && avgCostMap[id].avg > 0) return avgCostMap[id].avg;

    // 2) fallback: last unit cost
    if (lastCostMap[id] && lastCostMap[id].unit > 0) return lastCostMap[id].unit;

    // 3) still missing
    noteMissing(id, itemName, whereForMissing);
    return 0;
  };

  // 3) Inventory Value FIXED: Remaining QTY × unitCost (avg or last fallback)
  let stockValue = 0;
  const categoryMap = {};
  const lowStockItems = [];

  invData.forEach(item => {
    const itemId = (item['Item ID'] || '').toString().trim();
    const itemName = item['Item Name'] || itemId;

    const qty = parseVal(item['Remaining QTY'] || item['Stock'] || item['Qty'] || item['Quantity']);
    const unitCost = getUnitCost(itemId, itemName, 'inventory');

    stockValue += qty * unitCost;

    if (qty > 0) {
      const cat = item['Item Category'] || item['Category'] || 'Uncategorized';
      categoryMap[cat] = (categoryMap[cat] || 0) + qty;
    }

    const reorder = parseVal(item['Reorder Level']);
    if (qty === 0 || (reorder > 0 && qty <= reorder)) {
      lowStockItems.push({
        name: itemName,
        qty: qty,
        level: reorder > 0 ? reorder : 'N/A',
        id: itemId
      });
    }
  });

  // 4) REAL Profit from SalesDetails: profit = Total Sales Price - (QTY Sold * unitCost)
  const profitBySO = {}; // { [soId]: number }

  sdData.forEach(line => {
    const soId = line['SO ID'];
    if (!soId) return;

    const itemId = (line['Item ID'] || '').toString().trim();
    const itemName = line['Item Name'] || itemId;

    const qtySold = parseVal(line['QTY Sold']);              // can be negative for returns
    const revenue = parseVal(line['Total Sales Price']);     // can be negative for returns
    const unitCost = getUnitCost(itemId, itemName, 'sales'); // avg/last or missing

    const cogs = qtySold * unitCost; // negative qty -> negative cogs (works for returns)
    const profit = revenue - cogs;

    profitBySO[soId] = (profitBySO[soId] || 0) + profit;
  });

  const totalRealProfit = Object.values(profitBySO).reduce((a, b) => a + b, 0);

  // 5) Top Cities
  const cityMap = {};
  soData.forEach(r => {
    const amt = parseVal(r['Total SO Amount']);
    if (r['City']) cityMap[r['City']] = (cityMap[r['City']] || 0) + amt;
  });

  const getTop5 = (map, key) => Object.entries(map)
    .map(([k, v]) => ({ [key]: k, total: v }))
    .sort((a, b) => b.total - a.total)
    .slice(0, 5);

  // 6) Raw sales points for chart (frontend groups Day/Month/Year)
  const rawSales = soData
    .map(r => {
      const iso = parseDateISO(r['SO Date']);
      if (!iso) return null;
      const soId = r['SO ID'];
      return {
        date: iso,
        revenue: parseVal(r['Total SO Amount']),
        profit: parseVal(profitBySO[soId] || 0)
      };
    })
    .filter(Boolean);

  // 7) Recent Orders
  const recent = soData.slice().reverse().slice(0, 6).map(r => ({
    id: r['SO ID'],
    name: r['Customer Name'],
    amount: parseVal(r['Total SO Amount']),
    status: r['Receipt Status'] || 'Unpaid'
  }));

  return {
    kpi: {
      sales: totalSales,
      received: totalReceived,
      due: totalDue,
      stock: stockValue,
      profit: totalRealProfit
    },
    pie: { labels: Object.keys(categoryMap), values: Object.values(categoryMap) },
    lowStock: lowStockItems.slice(0, 10),
    topCities: getTop5(cityMap, 'city'),
    recent: recent,
    rawSales: rawSales,
    // ✅ NEW: show which items are missing cost data
    missingCosts: missingCostItems.slice(0, 30)
  };
}
