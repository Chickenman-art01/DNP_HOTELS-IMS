/**
 * DNP HOTELS - Items Catalog & Dashboard Metrics
 */

function getItemsInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[2]) continue;

    const central = Number(row[6]) || 0;
    const deneb = Number(row[7]) || 0;
    const pollux = Number(row[8]) || 0;
    const total = central + deneb + pollux;
    const unitCost = Number(row[10]) || 0;
    const totalValue = total * unitCost;

    items.push({
      rowIndex: i + 1,
      id: String(row[0] || ''),
      sku: String(row[1] || ''),
      name: String(row[2] || ''),
      category: String(row[3] || 'General'),
      unit: String(row[4] || 'Pcs'),
      minStock: Number(row[5]) || 0,
      centralStock: central,
      denebStock: deneb,
      polluxStock: pollux,
      totalStock: total,
      unitCost: unitCost,
      totalValue: totalValue,
      supplier: String(row[12] || ''),
      storageLocation: String(row[13] || ''),
      status: String(row[14] || 'Active'),
      lastUpdated: row[15] ? new Date(row[15]).toISOString() : ''
    });
  }
  return items;
}

function getItems() {
  const ss = getSpreadsheet();
  return getItemsInternal(ss);
}

function saveItem(itemData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Server is busy. Please retry in a few moments.');
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
    if (!sheet) throw new Error('Items sheet not found.');

    const data = sheet.getDataRange().getValues();
    const nowIso = new Date().toISOString();

    let targetRow = -1;
    let itemId = itemData.id ? String(itemData.id).trim() : '';

    if (itemId) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === itemId) {
          targetRow = i + 1;
          break;
        }
      }
    }

    const minStock = Number(itemData.minStock) || 0;
    const centralStock = Number(itemData.centralStock) || 0;
    const denebStock = Number(itemData.denebStock) || 0;
    const polluxStock = Number(itemData.polluxStock) || 0;
    const totalStock = centralStock + denebStock + polluxStock;
    const unitCost = Number(itemData.unitCost) || 0;
    const totalValue = totalStock * unitCost;

    let sku = itemData.sku ? String(itemData.sku).trim() : '';
    if (!sku) {
      const catPrefix = (itemData.category || 'GEN').substring(0, 2).toUpperCase();
      sku = 'DP-' + catPrefix + '-' + Math.floor(1000 + Math.random() * 9000);
    }

    if (targetRow > 1) {
      sheet.getRange(targetRow, 1, 1, 16).setValues([[
        itemId,
        sku,
        itemData.name,
        itemData.category,
        itemData.unit || 'Pcs',
        minStock,
        centralStock,
        denebStock,
        polluxStock,
        totalStock,
        unitCost,
        totalValue,
        itemData.supplier || '',
        itemData.storageLocation || '',
        itemData.status || 'Active',
        nowIso
      ]]);
    } else {
      if (!itemId) {
        itemId = 'ITM-' + (1000 + data.length);
      }

      sheet.appendRow([
        itemId,
        sku,
        itemData.name,
        itemData.category,
        itemData.unit || 'Pcs',
        minStock,
        centralStock,
        denebStock,
        polluxStock,
        totalStock,
        unitCost,
        totalValue,
        itemData.supplier || '',
        itemData.storageLocation || '',
        itemData.status || 'Active',
        nowIso
      ]);
    }

    return {
      success: true,
      item: {
        id: itemId,
        sku: sku,
        name: itemData.name,
        category: itemData.category,
        unit: itemData.unit,
        minStock: minStock,
        centralStock: centralStock,
        denebStock: denebStock,
        polluxStock: polluxStock,
        totalStock: totalStock,
        unitCost: unitCost,
        totalValue: totalValue,
        supplier: itemData.supplier,
        storageLocation: itemData.storageLocation,
        status: itemData.status || 'Active',
        lastUpdated: nowIso
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function deleteItem(itemId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Server is busy. Please retry.');
  }

  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
    if (!sheet) throw new Error('Items sheet not found.');

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(itemId)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Item deleted successfully.' };
      }
    }
    return { success: false, error: 'Item ID not found.' };
  } finally {
    lock.releaseLock();
  }
}

function calculateDashboardMetrics(items, transactions) {
  let totalSkus = items.length;
  let totalValuation = 0;
  let totalUnits = 0;
  let lowStockCount = 0;
  let outOfStockCount = 0;

  let centralStockUnits = 0;
  let centralValuation = 0;
  let denebStockUnits = 0;
  let denebValuation = 0;
  let polluxStockUnits = 0;
  let polluxValuation = 0;

  const categoryMap = {};
  const lowStockItems = [];

  items.forEach(item => {
    const totalStock = Number(item.totalStock) || 0;
    const minStock = Number(item.minStock) || 0;
    const unitCost = Number(item.unitCost) || 0;
    const centralStock = Number(item.centralStock) || 0;
    const denebStock = Number(item.denebStock) || 0;
    const polluxStock = Number(item.polluxStock) || 0;

    const itemValuation = totalStock * unitCost;
    totalValuation += itemValuation;
    totalUnits += totalStock;

    centralStockUnits += centralStock;
    centralValuation += centralStock * unitCost;

    denebStockUnits += denebStock;
    denebValuation += denebStock * unitCost;

    polluxStockUnits += polluxStock;
    polluxValuation += polluxStock * unitCost;

    if (totalStock === 0) {
      outOfStockCount++;
      lowStockItems.push(item);
    } else if (totalStock <= minStock) {
      lowStockCount++;
      lowStockItems.push(item);
    }

    const cat = item.category || 'Uncategorized';
    if (!categoryMap[cat]) {
      categoryMap[cat] = { count: 0, value: 0 };
    }
    categoryMap[cat].count++;
    categoryMap[cat].value += itemValuation;
  });

  const todayStr = new Date().toISOString().slice(0, 10);
  let todayTxnCount = 0;
  let todayStockIn = 0;
  let todayStockOut = 0;
  let todayTransfers = 0;

  transactions.forEach(t => {
    if (t.timestamp && t.timestamp.startsWith(todayStr)) {
      todayTxnCount++;
      if (t.type === 'STOCK_IN') todayStockIn += Number(t.quantity) || 0;
      if (t.type === 'STOCK_OUT') todayStockOut += Number(t.quantity) || 0;
      if (t.type === 'TRANSFER') todayTransfers += Number(t.quantity) || 0;
    }
  });

  return {
    totalSkus: totalSkus,
    totalValuation: Math.round(totalValuation),
    totalUnits: totalUnits,
    lowStockCount: lowStockCount,
    outOfStockCount: outOfStockCount,
    locations: {
      central: { units: centralStockUnits, valuation: Math.round(centralValuation) },
      deneb: { units: denebStockUnits, valuation: Math.round(denebValuation) },
      pollux: { units: polluxStockUnits, valuation: Math.round(polluxValuation) }
    },
    categoryBreakdown: categoryMap,
    lowStockItems: lowStockItems.slice(0, 10),
    todaySummary: {
      txnCount: todayTxnCount,
      stockIn: todayStockIn,
      stockOut: todayStockOut,
      transfers: todayTransfers
    }
  };
}

function getInitialData() {
  const ss = getSpreadsheet();
  const dbInfo = {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl()
  };

  const items = getItemsInternal(ss);
  const suppliers = getSuppliersInternal(ss);
  const departments = getDepartmentsInternal(ss);
  const locations = getLocationsInternal(ss);
  const settings = getSettingsInternal(ss);
  const transactions = getTransactionsInternal(ss, 25);
  const metrics = calculateDashboardMetrics(items, transactions);

  return {
    success: true,
    dbInfo: dbInfo,
    metrics: metrics,
    items: items,
    suppliers: suppliers,
    departments: departments,
    locations: locations,
    settings: settings,
    recentTransactions: transactions
  };
}
