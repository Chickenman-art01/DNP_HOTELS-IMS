/**
 * DNP HOTELS - Supplier Master & Location Master Handlers
 */

// ==========================================
// 1. SUPPLIERS (Supplier_Master Workbook)
// ==========================================
function getSuppliers() {
  const ss = getWorkbook(WORKBOOKS.SUPPLIER_MASTER);
  const sheet = ss.getSheetByName('Supplier_Master');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const suppliers = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    suppliers.push({
      id: String(row[0]), // e.g. SUP_001
      code: String(row[0]),
      name: String(row[1]),
      category: String(row[2] || 'General'),
      contactPerson: String(row[3] || ''),
      phone: String(row[4] || ''),
      email: String(row[5] || ''),
      status: String(row[6] || 'Active')
    });
  }
  return suppliers;
}

function saveSupplier(supData) {
  const ss = getWorkbook(WORKBOOKS.SUPPLIER_MASTER);
  const sheet = ss.getSheetByName('Supplier_Master');
  if (!sheet) throw new Error('Supplier_Master sheet not found');

  const data = sheet.getDataRange().getValues();
  let supCode = supData.code || supData.id ? String(supData.code || supData.id).trim() : '';
  let targetRow = -1;

  if (supCode) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === supCode.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      supCode,
      supData.name,
      supData.category || 'General',
      supData.contactPerson || '',
      supData.phone || '',
      supData.email || '',
      supData.status || 'Active'
    ]]);
  } else {
    if (!supCode) {
      supCode = 'SUP_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([
      supCode,
      supData.name,
      supData.category || 'General',
      supData.contactPerson || '',
      supData.phone || '',
      supData.email || '',
      supData.status || 'Active'
    ]);
  }

  return { success: true, supplierCode: supCode };
}

function deleteSupplier(supCode) {
  const ss = getWorkbook(WORKBOOKS.SUPPLIER_MASTER);
  const sheet = ss.getSheetByName('Supplier_Master');
  if (!sheet) throw new Error('Supplier_Master sheet not found');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(supCode)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Supplier not found' };
}

// ==========================================
// 2. LOCATIONS: STORES & SELLING POINTS (Location_Master Workbook)
// ==========================================
function getStores() {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Store');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const stores = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    stores.push({
      code: String(row[0]),
      name: String(row[1]),
      type: String(row[2] || 'Store'),
      status: String(row[3] || 'Active'),
      description: String(row[4] || '')
    });
  }
  return stores;
}

function saveStore(storeData) {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Store');
  if (!sheet) throw new Error('Store sheet missing in Location_Master');

  const data = sheet.getDataRange().getValues();
  let code = storeData.code ? String(storeData.code).trim() : '';
  let targetRow = -1;

  if (code) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === code.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 5).setValues([[
      code,
      storeData.name,
      storeData.type || 'Store',
      storeData.status || 'Active',
      storeData.description || ''
    ]]);
  } else {
    if (!code) {
      code = 'S_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([
      code,
      storeData.name,
      storeData.type || 'Store',
      storeData.status || 'Active',
      storeData.description || ''
    ]);
  }
  return { success: true, code: code };
}

function getSellingPoints() {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Selling_Point');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const sps = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    sps.push({
      code: String(row[0]),
      name: String(row[1]),
      storeCode: String(row[2] || 'S_001'),
      type: String(row[3] || 'Counter'),
      status: String(row[4] || 'Active')
    });
  }
  return sps;
}

function saveSellingPoint(spData) {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Selling_Point');
  if (!sheet) throw new Error('Selling_Point sheet missing in Location_Master');

  const data = sheet.getDataRange().getValues();
  let code = spData.code ? String(spData.code).trim() : '';
  let targetRow = -1;

  if (code) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).toUpperCase() === code.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 5).setValues([[
      code,
      spData.name,
      spData.storeCode || 'S_001',
      spData.type || 'Counter',
      spData.status || 'Active'
    ]]);
  } else {
    if (!code) {
      code = 'SP_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([
      code,
      spData.name,
      spData.storeCode || 'S_001',
      spData.type || 'Counter',
      spData.status || 'Active'
    ]);
  }
  return { success: true, code: code };
}

// ==========================================
// 3. USERS (Users_and_Settings Workbook)
// ==========================================
function getUsers() {
  const ss = getWorkbook(WORKBOOKS.USERS_SETTINGS);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    users.push({
      id: String(row[0]),
      name: String(row[1]),
      role: String(row[2] || 'Staff'),
      email: String(row[3] || ''),
      assignedStore: String(row[4] || 'ALL'),
      status: String(row[5] || 'Active')
    });
  }
  return users;
}
