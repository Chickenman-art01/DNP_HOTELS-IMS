/**
 * DNP HOTELS - Pure Sheet-Driven Product Master & Metrics
 * All calculations and listings are read 100% strictly from Google Sheets
 */

function getItemsInternal(ss) {
  const sheet = ss.getSheetByName('Product_Master') || ss.getSheets()[0];
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const codeIdx = headers.findIndex(h => h.includes('item code') || h.includes('code') || h.includes('sku'));
  const descIdx = headers.findIndex(h => h.includes('description') || h.includes('name') || h.includes('item'));
  const catIdx = headers.findIndex(h => h === 'category' || (h.includes('category') && !h.includes('code')));
  const catCodeIdx = headers.findIndex(h => h.includes('category code'));
  const uomIdx = headers.findIndex(h => h.includes('uom') || h.includes('unit'));
  const rateIdx = headers.findIndex(h => h.includes('rate') || h.includes('price') || h.includes('cost'));
  const taxIdx = headers.findIndex(h => h.includes('tax'));
  const minIdx = headers.findIndex(h => h.includes('min'));
  const s001Idx = headers.findIndex(h => h.includes('s_001') || h.includes('sec 29') || h.includes('deneb'));
  const s002Idx = headers.findIndex(h => h.includes('s_002') || h.includes('sec 27') || h.includes('pollux'));
  const centralIdx = headers.findIndex(h => h.includes('central') || h.includes('depot'));
  const totalStockIdx = headers.findIndex(h => h.includes('total stock') || h.includes('current stock') || h.includes('stock'));
  const supIdx = headers.findIndex(h => h.includes('supplier'));
  const statusIdx = headers.findIndex(h => h.includes('status'));

  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const code = codeIdx !== -1 ? String(row[codeIdx] || '').trim() : String(row[0] || '').trim();
    const desc = descIdx !== -1 ? String(row[descIdx] || '').trim() : String(row[1] || '').trim();
    if (!code && !desc) continue;

    const rate = rateIdx !== -1 ? Number(row[rateIdx]) || 0 : 0;
    const tax = taxIdx !== -1 ? Number(row[taxIdx]) || 0 : 0;
    const minStock = minIdx !== -1 ? Number(row[minIdx]) || 0 : 0;
    const stockS001 = s001Idx !== -1 ? Number(row[s001Idx]) || 0 : 0;
    const stockS002 = s002Idx !== -1 ? Number(row[s002Idx]) || 0 : 0;
    const centralStock = centralIdx !== -1 ? Number(row[centralIdx]) || 0 : 0;

    let totalStock = 0;
    if (s001Idx !== -1 || s002Idx !== -1 || centralIdx !== -1) {
      totalStock = stockS001 + stockS002 + centralStock;
    } else if (totalStockIdx !== -1) {
      totalStock = Number(row[totalStockIdx]) || 0;
    }

    const totalValue = totalStock * rate;

    items.push({
      id: code,
      code: code,
      sku: code,
      name: desc,
      description: desc,
      category: catIdx !== -1 ? String(row[catIdx] || 'General') : 'General',
      categoryCode: catCodeIdx !== -1 ? String(row[catCodeIdx] || '') : '',
      unit: uomIdx !== -1 ? String(row[uomIdx] || 'Pcs') : 'Pcs',
      uom: uomIdx !== -1 ? String(row[uomIdx] || 'Pcs') : 'Pcs',
      rate: rate,
      unitCost: rate,
      taxPercent: tax,
      minStock: minStock,
      stockS001: stockS001,
      stockS002: stockS002,
      centralStock: centralStock,
      totalStock: totalStock,
      totalValue: totalValue,
      supplierCode: supIdx !== -1 ? String(row[supIdx] || '') : '',
      supplier: supIdx !== -1 ? String(row[supIdx] || '') : '',
      status: statusIdx !== -1 ? String(row[statusIdx] || 'Active') : 'Active'
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
  const sheet = ss.getSheetByName('Product_Master') || ss.getSheets()[0];
  if (!sheet) throw new Error('Product_Master sheet missing');

  const data = sheet.getDataRange().getValues();
  let itemCode = itemData.code || itemData.sku || itemData.id ? String(itemData.code || itemData.sku || itemData.id).trim().toUpperCase() : '';
  let targetRow = -1;

  if (itemCode && data.length > 1) {
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
  const sheet = ss.getSheetByName('Product_Master') || ss.getSheets()[0];
  if (!sheet) throw new Error('Product_Master sheet missing');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(itemCode).toUpperCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Item not found in sheet' };
}

function calculateDashboardMetrics(items, supplierTxns, issuanceTxns, storesList) {
  let totalValuation = 0;
  let totalUnits = 0;
  let totalSkus = items.length;
  let lowStockCount = 0;
  let outOfStockCount = 0;

  const categoryMap = {};
  const lowStockItems = [];

  // Initialize store units and valuation purely from the registered stores
  const storeMetrics = {};
  (storesList || []).forEach(st => {
    storeMetrics[st.code] = {
      code: st.code,
      name: st.name,
      units: 0,
      valuation: 0
    };
  });

  items.forEach(item => {
    const total = Number(item.totalStock) || 0;
    const min = Number(item.minStock) || 0;
    const rate = Number(item.rate) || 0;
    const itemVal = total * rate;

    totalUnits += total;
    totalValuation += itemVal;

    // Distribute to registered stores if matching codes exist
    if (storeMetrics['S_001'] && item.stockS001) {
      storeMetrics['S_001'].units += item.stockS001;
      storeMetrics['S_001'].valuation += item.stockS001 * rate;
    }
    if (storeMetrics['S_002'] && item.stockS002) {
      storeMetrics['S_002'].units += item.stockS002;
      storeMetrics['S_002'].valuation += item.stockS002 * rate;
    }
    if (storeMetrics['S_000'] && item.centralStock) {
      storeMetrics['S_000'].units += item.centralStock;
      storeMetrics['S_000'].valuation += item.centralStock * rate;
    }

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
    stores: storeMetrics,
    categoryBreakdown: categoryMap,
    lowStockItems: lowStockItems.slice(0, 15),
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
  const metrics = calculateDashboardMetrics(items, supplierTxns, issuanceTxns, stores);

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
