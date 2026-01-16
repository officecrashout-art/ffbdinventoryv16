/**
 * FASHION FIZZ BD - MASTER DASHBOARD BACKEND
 * Aggregates data from Sales, Purchases, Inventory, Customers, and Suppliers.
 */

function dashShowUI() {
  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard | Fashion Fizz BD')
    .setMode(HtmlService.SandboxMode.IFRAME);
  return html;
}

/**
 * Main entrypoint: returns all KPI values, alerts, and chart data.
 */
function getDashboardData() {
  // 1. FETCH ALL DATA
  const sales = soGetRangeDataAsObjects('RANGESO');        // Sales Orders
  const purchases = soGetRangeDataAsObjects('RANGEPO');    // Purchase Orders
  const customers = soGetRangeDataAsObjects('RANGECUSTOMERS');
  const suppliers = soGetRangeDataAsObjects('RANGESUPPLIERS');
  const inventory = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');
  const salesDetails = soGetRangeDataAsObjects('RANGESD'); // Sales Details for Top Items

  // 2. FINANCIAL KPIS
  const totalRevenue = sales.reduce((sum, r) => sum + (parseFloat(r['Total SO Amount']) || 0), 0);
  const totalExpense = purchases.reduce((sum, r) => sum + (parseFloat(r['Total Amount']) || 0), 0);
  const netProfit = totalRevenue - totalExpense;

  // 3. CASH FLOW KPIS
  const totalReceivables = customers.reduce((sum, r) => sum + (parseFloat(r['Balance Receivable']) || 0), 0);
  const totalPayables = suppliers.reduce((sum, r) => sum + (parseFloat(r['Balance Payable']) || 0), 0);

  // 4. INVENTORY ALERTS (Low Stock)
  // Filter items where Remaining QTY <= Reorder Level
  const lowStockItems = inventory.filter(i => {
    const stock = parseFloat(i['Remaining QTY']) || 0;
    const level = parseFloat(i['Reorder Level']) || 0;
    return stock <= level;
  }).map(i => ({
    name: i['Item Name'],
    stock: i['Remaining QTY'],
    level: i['Reorder Level']
  }));

  // 5. RECENT ORDERS (Last 5)
  // Sort by Date Descending
  const recentOrders = sales.sort((a,b) => new Date(b['SO Date']) - new Date(a['SO Date'])).slice(0, 5).map(o => ({
    id: o['SO ID'],
    customer: o['Customer Name'],
    amount: o['Total SO Amount'],
    status: o['Receipt Status'],
    date: new Date(o['SO Date']).toLocaleDateString()
  }));

  // 6. CHART DATA: SALES TREND (Last 6 Months)
  // Group sales by Month-Year
  const trendMap = {};
  sales.forEach(s => {
    if (!s['SO Date']) return;
    const d = new Date(s['SO Date']);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM yyyy');
    trendMap[key] = (trendMap[key] || 0) + (parseFloat(s['Total SO Amount']) || 0);
  });
  // Convert to array (you might want to sort these keys chronologically in a real app)
  const salesTrend = Object.keys(trendMap).map(k => ({ month: k, total: trendMap[k] }));

  // 7. CHART DATA: TOP 5 ITEMS
  const itemMap = {};
  salesDetails.forEach(d => {
    const name = d['Item Name'];
    itemMap[name] = (itemMap[name] || 0) + (parseFloat(d['QTY Sold']) || 0);
  });
  const topItems = Object.entries(itemMap)
    .sort((a,b) => b[1] - a[1])
    .slice(0, 5)
    .map(([name, qty]) => ({ name: name, sold: qty }));

  // 8. CHART DATA: CATEGORY DISTRIBUTION
  const catMap = {};
  inventory.forEach(i => {
    const cat = i['Item Category'] || 'Uncategorized';
    catMap[cat] = (catMap[cat] || 0) + (parseFloat(i['Remaining QTY']) || 0);
  });
  const categoryData = Object.entries(catMap).map(([cat, qty]) => ({ category: cat, stock: qty }));

  return {
    kpi: {
      revenue: totalRevenue,
      expense: totalExpense,
      profit: netProfit,
      receivable: totalReceivables,
      payable: totalPayables,
      ordersCount: sales.length
    },
    alerts: lowStockItems,
    recent: recentOrders,
    charts: {
      trend: salesTrend,
      topItems: topItems,
      categories: categoryData
    }
  };
}