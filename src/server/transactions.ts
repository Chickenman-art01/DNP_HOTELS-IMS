/**
 * DNP HOTELS - Supplier & Issuance Transactions Engine
 */

// ==========================================
// 1. SUPPLIER PURCHASES & STOCK IN
// ==========================================
function processSupplierPurchase(txn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Server busy with concurrent transaction. Please retry.');
  }

  try {
    const prodWb = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
    const prodSheet = prodWb.getSheetByName('Product_Master');
    const supTxnWb = getWorkbook(WORKBOOKS.SUPPLIER_TXNS);
    const supTxnSheet = supTxnWb.getSheetByName('Supplier_Transactions');

    const itemCode = String(txn.itemCode || txn.itemId || txn.sku).trim().toUpperCase();
    const prodData = prodSheet.getDataRange().getValues();
    let targetRow = -1;
    let itemRow = null;

    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).toUpperCase() === itemCode) {
        targetRow = i + 1;
        itemRow = prodData[i];
        break;
      }
    }

    if (targetRow === -1 || !itemRow) {
      throw new Error('Product not found for Item Code: ' + itemCode);
    }

    const itemDesc = String(itemRow[1]);
    const category = String(itemRow[2]);
    const uom = String(itemRow[4] || 'Pcs');
    let rate = Number(txn.rate) || Number(itemRow[5]) || 0;
    let taxPct = Number(txn.taxPercent) || Number(itemRow[6]) || 0;
    const minStock = Number(itemRow[7]) || 0;
    let stockS001 = Number(itemRow[8]) || 0;
    let stockS002 = Number(itemRow[9]) || 0;
    let centralStock = Number(itemRow[10]) || 0;

    const qty = Math.abs(Number(txn.quantity));
    if (isNaN(qty) || qty <= 0) throw new Error('Quantity must be greater than 0');

    const storeCode = String(txn.storeCode || txn.destLocation || 'S_001').toUpperCase();
    let storeName = txn.storeName || 'Store';

    if (storeCode.includes('S_001') || storeCode.includes('21 GUN') || storeCode.includes('DENEB')) {
      stockS001 += qty;
      storeName = '21 GUN SOLUTE GGN SEC 29';
    } else if (storeCode.includes('S_002') || storeCode.includes('PAHLE') || storeCode.includes('POLLUX')) {
      stockS002 += qty;
      storeName = 'PAHLE CHAI GGN Sec 27';
    } else {
      centralStock += qty;
      storeName = 'Central Depot Warehouse';
    }

    const totalStock = stockS001 + stockS002 + centralStock;
    const totalVal = totalStock * rate;
    const nowIso = new Date().toISOString();

    // Update Product_Master row
    prodSheet.getRange(targetRow, 6, 1, 8).setValues([[
      rate,
      taxPct,
      minStock,
      stockS001,
      stockS002,
      centralStock,
      totalStock,
      totalVal
    ]]);
    prodSheet.getRange(targetRow, 16).setValue(nowIso);

    // Append to Supplier_Transactions
    const dateCode = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
    const txnId = 'TXN_SUP_' + dateCode + '_' + Math.floor(1000 + Math.random() * 9000);
    const taxAmt = (qty * rate * taxPct) / 100;
    const totalAmt = (qty * rate) + taxAmt;

    supTxnSheet.appendRow([
      txnId,
      nowIso,
      txn.supplierCode || '',
      txn.supplierName || '',
      itemCode,
      itemDesc,
      category,
      qty,
      uom,
      rate,
      taxPct,
      totalAmt,
      storeCode,
      storeName,
      txn.reference || txn.invoiceRef || '',
      txn.performedBy || txn.receivedBy || 'Store Keeper',
      txn.notes || ''
    ]);

    return {
      success: true,
      transactionId: txnId,
      itemCode: itemCode,
      updatedTotalStock: totalStock,
      updatedValuation: totalVal
    };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// 2. ISSUANCE TO SELLING POINTS & DEPARTMENTS
// ==========================================
function processStockIssuance(txn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('Server busy with concurrent transaction. Please retry.');
  }

  try {
    const prodWb = getWorkbook(WORKBOOKS.PRODUCT_MASTER);
    const prodSheet = prodWb.getSheetByName('Product_Master');
    const issueWb = getWorkbook(WORKBOOKS.ISSUANCE_TXNS);
    const issueSheet = issueWb.getSheetByName('Issuance_Transactions');

    const itemCode = String(txn.itemCode || txn.itemId || txn.sku).trim().toUpperCase();
    const prodData = prodSheet.getDataRange().getValues();
    let targetRow = -1;
    let itemRow = null;

    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][0]).toUpperCase() === itemCode) {
        targetRow = i + 1;
        itemRow = prodData[i];
        break;
      }
    }

    if (targetRow === -1 || !itemRow) {
      throw new Error('Product not found for Item Code: ' + itemCode);
    }

    const itemDesc = String(itemRow[1]);
    const uom = String(itemRow[4] || 'Pcs');
    const rate = Number(itemRow[5]) || 0;
    let stockS001 = Number(itemRow[8]) || 0;
    let stockS002 = Number(itemRow[9]) || 0;
    let centralStock = Number(itemRow[10]) || 0;

    const qty = Math.abs(Number(txn.quantity));
    if (isNaN(qty) || qty <= 0) throw new Error('Quantity must be greater than 0');

    const fromStore = String(txn.sourceLocation || txn.fromStore || 'S_001').toUpperCase();
    const toSp = String(txn.destLocation || txn.toSellingPoint || 'SP_001');

    // Validate available stock in source store
    if (fromStore.includes('S_001') || fromStore.includes('21 GUN')) {
      if (stockS001 < qty) throw new Error(`Insufficient stock in 21 GUN SOLUTE GGN SEC 29 (Available: ${stockS001}, Requested: ${qty})`);
      stockS001 -= qty;
    } else if (fromStore.includes('S_002') || fromStore.includes('PAHLE')) {
      if (stockS002 < qty) throw new Error(`Insufficient stock in PAHLE CHAI GGN Sec 27 (Available: ${stockS002}, Requested: ${qty})`);
      stockS002 -= qty;
    } else {
      if (centralStock < qty) throw new Error(`Insufficient stock in Central Depot (Available: ${centralStock}, Requested: ${qty})`);
      centralStock -= qty;
    }

    // If it's a transfer between stores, increase stock in destination
    const type = String(txn.type || 'DISBURSEMENT').toUpperCase();
    if (type === 'TRANSFER') {
      const dest = toSp.toUpperCase();
      if (dest.includes('S_001') || dest.includes('21 GUN')) {
        stockS001 += qty;
      } else if (dest.includes('S_002') || dest.includes('PAHLE')) {
        stockS002 += qty;
      } else {
        centralStock += qty;
      }
    }

    const totalStock = stockS001 + stockS002 + centralStock;
    const totalVal = totalStock * rate;
    const nowIso = new Date().toISOString();

    // Update Product_Master
    prodSheet.getRange(targetRow, 9, 1, 5).setValues([[
      stockS001,
      stockS002,
      centralStock,
      totalStock,
      totalVal
    ]]);
    prodSheet.getRange(targetRow, 16).setValue(nowIso);

    // Append to Issuance_Transactions
    const dateCode = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
    const issueId = 'ISS_' + dateCode + '_' + Math.floor(1000 + Math.random() * 9000);
    const lineTotal = qty * rate;

    issueSheet.appendRow([
      issueId,
      nowIso,
      type,
      itemCode,
      itemDesc,
      qty,
      uom,
      fromStore,
      txn.fromStoreName || fromStore,
      toSp,
      txn.toSellingPointName || toSp,
      rate,
      lineTotal,
      txn.reference || txn.requisitionRef || '',
      txn.performedBy || txn.issuedBy || 'Store Incharge',
      'Approved',
      txn.notes || ''
    ]);

    return {
      success: true,
      issuanceId: issueId,
      itemCode: itemCode,
      updatedTotalStock: totalStock,
      updatedValuation: totalVal
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Universal dispatcher for frontend modal calls
 */
function processStockTransaction(txn) {
  const type = String(txn.type).toUpperCase();
  if (type === 'STOCK_IN') {
    return processSupplierPurchase(txn);
  } else {
    return processStockIssuance(txn);
  }
}

// ==========================================
// 3. READ TRANSACTIONS
// ==========================================
function getSupplierTransactions(limit) {
  try {
    const ss = getWorkbook(WORKBOOKS.SUPPLIER_TXNS);
    const sheet = ss.getSheetByName('Supplier_Transactions');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const list = [];
    const max = limit || 100;
    for (let i = data.length - 1; i >= 1 && list.length < max; i--) {
      const r = data[i];
      if (!r[0]) continue;
      list.push({
        id: String(r[0]),
        timestamp: r[1] ? new Date(r[1]).toISOString() : '',
        supplierCode: String(r[2]),
        supplierName: String(r[3]),
        itemCode: String(r[4]),
        itemName: String(r[5]),
        category: String(r[6]),
        quantity: Number(r[7]) || 0,
        unit: String(r[8]),
        rate: Number(r[9]) || 0,
        taxPercent: Number(r[10]) || 0,
        totalAmount: Number(r[11]) || 0,
        storeCode: String(r[12]),
        storeName: String(r[13]),
        reference: String(r[14] || ''),
        performedBy: String(r[15] || ''),
        notes: String(r[16] || '')
      });
    }
    return list;
  } catch (e) {
    return [];
  }
}

function getIssuanceTransactions(limit) {
  try {
    const ss = getWorkbook(WORKBOOKS.ISSUANCE_TXNS);
    const sheet = ss.getSheetByName('Issuance_Transactions');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const list = [];
    const max = limit || 100;
    for (let i = data.length - 1; i >= 1 && list.length < max; i--) {
      const r = data[i];
      if (!r[0]) continue;
      list.push({
        id: String(r[0]),
        timestamp: r[1] ? new Date(r[1]).toISOString() : '',
        type: String(r[2]),
        itemCode: String(r[3]),
        itemName: String(r[4]),
        sku: String(r[3]),
        quantity: Number(r[5]) || 0,
        unit: String(r[6]),
        fromStoreCode: String(r[7]),
        sourceLocation: String(r[8] || r[7]),
        toSellingPointCode: String(r[9]),
        destLocation: String(r[10] || r[9]),
        rate: Number(r[11]) || 0,
        unitPrice: Number(r[11]) || 0,
        totalCost: Number(r[12]) || 0,
        reference: String(r[13] || ''),
        performedBy: String(r[14] || ''),
        status: String(r[15] || 'Approved'),
        notes: String(r[16] || '')
      });
    }
    return list;
  } catch (e) {
    return [];
  }
}

function getTransactions(limit) {
  return getIssuanceTransactions(limit || 100);
}
