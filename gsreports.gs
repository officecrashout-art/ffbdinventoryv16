/**
 * FASHION FIZZ BD - ADVANCED REPORTING ENGINE
 */

function rptShowUI() {
  const html = HtmlService.createTemplateFromFile('reports')
    .evaluate()
    .setTitle('Business Reports')
    .setWidth(1350)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * 1. INVENTORY REPORT DATA
 * Returns: Category, Name, Sold Qty, Available Stock
 */
function rptGetInventoryReport() {
  const items = soGetRangeDataAsObjects('RANGEINVENTORYITEMS');
  
  // Map to simple structure
  return items.map(i => ({
    id: i['Item ID'],
    category: i['Item Category'],
    name: i['Item Name'],
    brand: i['Brands'],
    sold: Number(i['QTY Sold'] || 0),
    stock: Number(i['Remaining QTY'] || 0)
  }));
}

/**
 * 2. SALES DATA (TRENDS)
 * Period: 'day', 'month', 'year'
 */
function rptGetSalesData(period) {
  const orders = soGetRangeDataAsObjects('RANGESO');
  const summary = {};

  orders.forEach(o => {
    if (!o['SO Date']) return;
    const d = new Date(o['SO Date']);
    
    let key;
    if (period === 'day') {
      key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else if (period === 'month') {
      key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM');
    } else {
      key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy');
    }

    if (!summary[key]) summary[key] = { date: key, revenue: 0, orders: 0 };
    
    summary[key].revenue += Number(o['Total SO Amount'] || 0);
    summary[key].orders += 1;
  });

  // Convert to array and sort descending (newest first)
  return Object.values(summary).sort((a, b) => b.date.localeCompare(a.date));
}

/**
 * 3. ANALYTICS (Top Products, Cities, Customers)
 */
function rptGetAnalytics() {
  const details = soGetRangeDataAsObjects('RANGESD');
  const orders = soGetRangeDataAsObjects('RANGESO');

  // A. MOST SOLD PRODUCTS (Name + Size)
  const prodMap = {};
  details.forEach(d => {
    // Combine Name + Size (e.g., "T-Shirt - M")
    const key = `${d['Item Name']} (${d['Size'] || d['Item Subcategory'] || '-'})`;
    if (!prodMap[key]) prodMap[key] = { name: key, qty: 0 };
    prodMap[key].qty += Number(d['QTY Sold'] || 0);
  });
  const topProducts = Object.values(prodMap).sort((a,b) => b.qty - a.qty).slice(0, 10);

  // B. TOP CITIES (By Revenue)
  const cityMap = {};
  orders.forEach(o => {
    const city = o['City'] || 'Unknown';
    if (!cityMap[city]) cityMap[city] = { city: city, revenue: 0 };
    cityMap[city].revenue += Number(o['Total SO Amount'] || 0);
  });
  const topCities = Object.values(cityMap).sort((a,b) => b.revenue - a.revenue).slice(0, 10);

  // C. MOST FREQUENT CUSTOMERS
  const custMap = {};
  orders.forEach(o => {
    const name = o['Customer Name'];
    if (!custMap[name]) custMap[name] = { name: name, count: 0, revenue: 0 };
    custMap[name].count += 1;
    custMap[name].revenue += Number(o['Total SO Amount'] || 0);
  });
  const topCustomers = Object.values(custMap).sort((a,b) => b.count - a.count).slice(0, 10);

  return { topProducts, topCities, topCustomers };
}