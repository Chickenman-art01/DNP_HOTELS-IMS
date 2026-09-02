/**
 * DNP HOTELS - Atomic Stock Movements & Transactions
 */

function processStockTransaction(txn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Server busy with concurrent transaction. Please retry in a moment.');
  }

  try {
    const ss = getSpreadsheet();
    const itemsSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
    const txnSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);

    if (!itemsSheet || !txnSheet) throw new Error('Required sheets missing in database.');

    const itemsData = itemsSheet.getDataRange().getValues();
    let targetRow = -1;
    let itemRowData = null;

    for (let i = 1; i < itemsData.length; i++) {
      if (String(itemsData[i][0]) === String(txn.itemId) || String(itemsData[i][1]) === String(txn.sku)) {
        targetRow = i + 1;
        itemRowData = itemsData[i];
        break;
      }
    }

    if (targetRow === -1 || !itemRowData) {
      throw new Error('Item not found for ID / SKU: ' + (txn.itemId || txn.sku));
    }

    const itemId = String(itemRowData[0]);
    const sku = String(itemRowData[1]);
    const itemName = String(itemRowData[2]);
    const category = String(itemRowData[3]);
    const unit = String(itemRowData[4]);
    let minStock = Number(itemRowData[5]) || 0;
    let centralStock = Number(itemRowData[6]) || 0;
    let denebStock = Number(itemRowData[7]) || 0;
    let polluxStock = Number(itemRowData[8]) || 0;
    let unitCost = Number(txn.unitPrice) || Number(itemRowData[10]) || 0;

    const qty = Math.abs(Number(txn.quantity));
    if (isNaN(qty) || qty <= 0) {
      throw new Error('Invalid quantity specified. Must be greater than 0.');
    }

    const type = String(txn.type).toUpperCase();
    const source = String(txn.sourceLocation || '').trim();
    const dest = String(txn.destLocation || '').trim();

    function getStockByLoc(locName) {
      const l = locName.toUpperCase();
      if (l.includes('CENTRAL')) return centralStock;
      if (l.includes('DENEB')) return denebStock;
      if (l.includes('POLLUX')) return polluxStock;
      return null;
    }

    function updateStockByLoc(locName, delta) {
      const l = locName.toUpperCase();
      if (l.includes('CENTRAL')) centralStock += delta;
      else if (l.includes('DENEB')) denebStock += delta;
      else if (l.includes('POLLUX')) polluxStock += delta;
      else {
        throw new Error('Unrecognized property location: ' + locName + '. Must be Central Warehouse, Deneb Hotel, or Pollux Hotel.');
      }
    }

    if (type === 'STOCK_IN') {
      updateStockByLoc(dest, qty);
      if (txn.unitPrice && Number(txn.unitPrice) > 0) {
        unitCost = Number(txn.unitPrice);
      }
    } else if (type === 'STOCK_OUT') {
      const currentLocStock = getStockByLoc(source);
      if (currentLocStock === null) {
        throw new Error('Invalid source location for Stock Out: ' + source);
      }
      if (currentLocStock < qty) {
        throw new Error('Insufficient stock in ' + source + '. Available: ' + currentLocStock + ', Requested: ' + qty);
      }
      updateStockByLoc(source, -qty);
    } else if (type === 'TRANSFER') {
      const currentLocStock = getStockByLoc(source);
      if (currentLocStock === null) {
        throw new Error('Invalid source location for Transfer: ' + source);
      }
      if (currentLocStock < qty) {
        throw new Error('Insufficient stock in ' + source + ' to transfer. Available: ' + currentLocStock + ', Transfer: ' + qty);
      }
      updateStockByLoc(source, -qty);
      updateStockByLoc(dest, qty);
    } else if (type === 'ADJUSTMENT') {
      const targetLoc = source || dest;
      const currentLocStock = getStockByLoc(targetLoc);
      if (currentLocStock === null) {
        throw new Error('Invalid location for Stock Adjustment: ' + targetLoc);
      }
      const countedStock = Number(txn.countedQuantity);
      if (isNaN(countedStock) || countedStock < 0) {
        throw new Error('Counted physical stock must be a non-negative number.');
      }
      const diff = countedStock - currentLocStock;
      updateStockByLoc(targetLoc, diff);
    } else {
      throw new Error('Invalid transaction type: ' + type);
    }

    const totalStock = centralStock + denebStock + polluxStock;
    const totalValue = totalStock * unitCost;
    const nowIso = new Date().toISOString();

    itemsSheet.getRange(targetRow, 7, 1, 6).setValues([[
      centralStock,
      denebStock,
      polluxStock,
      totalStock,
      unitCost,
      totalValue
    ]]);
    itemsSheet.getRange(targetRow, 16).setValue(nowIso);

    const dateCode = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
    const txnId = 'TXN-' + dateCode + '-' + Math.floor(1000 + Math.random() * 9000);
    const lineTotalCost = qty * unitCost;

    txnSheet.appendRow([
      txnId,
      nowIso,
      type,
      itemId,
      sku,
      itemName,
      category,
      qty,
      unit,
      source,
      dest,
      unitCost,
      lineTotalCost,
      txn.reference || '',
      txn.performedBy || 'Staff',
      txn.notes || ''
    ]);

    return {
      success: true,
      transactionId: txnId,
      updatedItem: {
        id: itemId,
        sku: sku,
        name: itemName,
        centralStock: centralStock,
        denebStock: denebStock,
        polluxStock: polluxStock,
        totalStock: totalStock,
        unitCost: unitCost,
        totalValue: totalValue,
        lastUpdated: nowIso
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function getTransactionsInternal(ss, limit) {
  const sheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const txns = [];
  const max = limit || 100;
  for (let i = data.length - 1; i >= 1 && txns.length < max; i--) {
    const row = data[i];
    if (!row[0]) continue;

    txns.push({
      id: String(row[0]),
      timestamp: row[1] ? new Date(row[1]).toISOString() : '',
      type: String(row[2]),
      itemId: String(row[3]),
      sku: String(row[4]),
      itemName: String(row[5]),
      category: String(row[6]),
      quantity: Number(row[7]) || 0,
      unit: String(row[8]),
      sourceLocation: String(row[9]),
      destLocation: String(row[10]),
      unitPrice: Number(row[11]) || 0,
      totalCost: Number(row[12]) || 0,
      reference: String(row[13] || ''),
      performedBy: String(row[14] || ''),
      notes: String(row[15] || '')
    });
  }
  return txns;
}

function getTransactions(limit) {
  const ss = getSpreadsheet();
  return getTransactionsInternal(ss, limit || 100);
}
