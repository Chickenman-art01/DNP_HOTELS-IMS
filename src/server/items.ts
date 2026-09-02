/**
 * DNP HOTELS - Product Master, Catalog & Dashboard Metrics
 */

function getItemsInternal(ss) {
  const sheet = ss.getSheetByName('Product_Master');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    const rate = Number(row[5]) || 0;
    const tax = Number(row[6]) || 0;
    const minStock = Number(row[7]) || 0;
    const stockS001 = Number(row[8]) || 0;
    const stockS002 = Number(row[9]) || 0;
    const centralStock = Number(row[10]) || 0;
    const totalStock = stockS001 + stockS002 + centralStock;
    const totalValue = totalStock * rate;

    items.push({
      id: String(row[0]),
      code: String(row[0]),
      sku: String(row[0]),
      name: String(row[1]),
      description: String(row[1]),
      category: String(row[2]),
      categoryCode: String(row[3] || ''),
      unit: String(row[4] || 'Pcs'),
      uom: String(row[4] || 'Pcs'),
      rate: rate,
      unitCost: rate,
      taxPercent: tax,
      minStock: minStock,
      stockS001: stockS001,
      stockS002: stockS002,
      centralStock: centralStock,
      totalStock: totalStock,
      totalValue: totalValue,
      supplierCode: String(row[13] || ''),
      supplier: String(row[13] || ''),
      status: String(row[14] || 'Active'),
      lastUpdated: row[15] ? new Date(row[15]).toISOString() : ''
    });
  }
  return items;
}

function getItems() {
  const ss = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
  return getItemsInternal(ss);
}

function saveItem(itemData) {
  const ss = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
  const sheet = ss.getSheetByName('Product_Master');
  if (!sheet) throw new Error('Product_Master sheet missing');

  const data = sheet.getDataRange().getValues();
  let itemCode = itemData.code || itemData.sku || itemData.id ? String(itemData.code || itemData.sku || itemData.id).trim().toUpperCase() : '';
  let targetRow = -1;

  if (itemCode) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === itemCode) {
        targetRow = i + 1;
        break;
      }
    }
  }

  const rate = Number(itemData.rate || itemData.unitCost) || 0;
  const tax = Number(itemData.taxPercent) || 0;
  const minStock = Number(itemData.minStock) || 0;
  const stockS001 = Number(itemData.stockS001 || itemData.denebStock) || 0;
  const stockS002 = Number(itemData.stockS002 || itemData.polluxStock) || 0;
  const centralStock = Number(itemData.centralStock) || 0;
  const totalStock = stockS001 + stockS002 + centralStock;
  const totalVal = totalStock * rate;
  const nowIso = new Date().toISOString();

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 16).setValues([[
      itemCode,
      itemData.name || itemData.description,
      itemData.category,
      itemData.categoryCode || ('CAT_' + String(itemData.category).slice(0, 3).toUpperCase()),
      itemData.unit || itemData.uom || 'Pcs',
      rate,
      tax,
      minStock,
      stockS001,
      stockS002,
      centralStock,
      totalStock,
      totalVal,
      itemData.supplierCode || itemData.supplier || '',
      itemData.status || 'Active',
      nowIso
    ]]);
  } else {
    if (!itemCode) {
      itemCode = 'ITM_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([
      itemCode,
      itemData.name || itemData.description,
      itemData.category,
      itemData.categoryCode || ('CAT_' + String(itemData.category).slice(0, 3).toUpperCase()),
      itemData.unit || itemData.uom || 'Pcs',
      rate,
      tax,
      minStock,
      stockS001,
      stockS002,
      centralStock,
      totalStock,
      totalVal,
      itemData.supplierCode || itemData.supplier || '',
      itemData.status || 'Active',
      nowIso
    ]);
  }

  return { success: true, itemCode: itemCode };
}

function deleteItem(itemCode) {
  const ss = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
  const sheet = ss.getSheetByName('Product_Master');
  if (!sheet) throw new Error('Product_Master sheet missing');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(itemCode).toUpperCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found' };
}

function calculateDashboardMetrics(items, supplierTxns, issuanceTxns) {
  let totalValuation = 0;
  let totalUnits = 0;
  let totalSkus = items.length;
  let lowStockCount = 0;
  let outOfStockCount = 0;

  let s001Units = 0;
  let s001Valuation = 0;
  let s002Units = 0;
  let s002Valuation = 0;
  let centralUnits = 0;
  let centralValuation = 0;

  const categoryMap = {};
  const lowStockItems = [];

  items.forEach(item => {
    const total = Number(item.totalStock) || 0;
    const min = Number(item.minStock) || 0;
    const rate = Number(item.rate) || 0;
    const itemVal = total * rate;

    totalUnits += total;
    totalValuation += itemVal;

    s001Units += Number(item.stockS001) || 0;
    s001Valuation += (Number(item.stockS001) || 0) * rate;

    s002Units += Number(item.stockS002) || 0;
    s002Valuation += (Number(item.stockS002) || 0) * rate;

    centralUnits += Number(item.centralStock) || 0;
    centralValuation += (Number(item.centralStock) || 0) * rate;

    if (total === 0) {
      outOfStockCount++;
      lowStockItems.push(item);
    } else if (total <= min) {
      lowStockCount++;
      lowStockItems.push(item);
    }

    const cat = item.category || 'General';
    if (!categoryMap[cat]) {
      categoryMap[cat] = { count: 0, units: 0, value: 0 };
    }
    categoryMap[cat].count += 1;
    categoryMap[cat].units += total;
    categoryMap[cat].value += itemVal;
  });

  const todayStr = new Date().toISOString().slice(0, 10);
  let todayPurchases = 0;
  let todayPurchasesVal = 0;
  let todayIssues = 0;
  let todayIssuesVal = 0;

  supplierTxns.forEach(t => {
    if (t.timestamp && t.timestamp.startsWith(todayStr)) {
      todayPurchases += Number(t.quantity) || 0;
      todayPurchasesVal += Number(t.totalAmount) || 0;
    }
  });

  issuanceTxns.forEach(t => {
    if (t.timestamp && t.timestamp.startsWith(todayStr)) {
      todayIssues += Number(t.quantity) || 0;
      todayIssuesVal += Number(t.totalValue) || 0;
    }
  });

  return {
    totalSkus: totalSkus,
    totalValuation: Math.round(totalValuation),
    totalUnits: totalUnits,
    lowStockCount: lowStockCount,
    outOfStockCount: outOfStockCount,
    stores: {
      s001: { name: '21 GUN SOLUTE GGN SEC 29', code: 'S_001', units: s001Units, valuation: Math.round(s001Valuation) },
      s002: { name: 'PAHLE CHAI GGN Sec 27', code: 'S_002', units: s002Units, valuation: Math.round(s002Valuation) },
      central: { name: 'Central Depot Warehouse', code: 'S_000', units: centralUnits, valuation: Math.round(centralValuation) }
    },
    categoryBreakdown: categoryMap,
    lowStockItems: lowStockItems.slice(0, 10),
    todaySummary: {
      txnCount: supplierTxns.length + issuanceTxns.length,
      stockIn: todayPurchases,
      stockInValue: todayPurchasesVal,
      stockOut: todayIssues,
      stockOutValue: todayIssuesVal
    }
  };
}

function getInitialData() {
  // Ensure all 6 workbooks are initialized
  initAllWorkbooks();

  const workbooksInfo = getWorkbooksInfo();
  const prodWb = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
  const items = getItemsInternal(prodWb);
  const suppliers = getSuppliers();
  const stores = getStores();
  const sellingPoints = getSellingPoints();
  const users = getUsers();
  const settings = getSettings();
  const supplierTxns = getSupplierTransactions(50);
  const issuanceTxns = getIssuanceTransactions(50);
  const metrics = calculateDashboardMetrics(items, supplierTxns, issuanceTxns);

  return {
    success: true,
    workbooksInfo: workbooksInfo,
    metrics: metrics,
    items: items,
    suppliers: suppliers,
    stores: stores,
    sellingPoints: sellingPoints,
    users: users,
    settings: settings,
    supplierTransactions: supplierTxns,
    issuanceTransactions: issuanceTxns,
    recentTransactions: issuanceTxns
  };
}
