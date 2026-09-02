/**
 * DNP HOTELS - Pure Google Sheets Master Data Handler
 * Reads and writes 100% strictly from Google Sheets in Drive folder
 */

// ==========================================
// 1. SUPPLIER MASTER (Supplier_Master Workbook)
// ==========================================
function getSuppliers() {
  const ss = getWorkbook(WORKBOOKS.SUPPLIER_MASTER);
  const sheet = ss.getSheetByName('Supplier_Master') || ss.getSheets()[0];
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const codeIdx = headers.findIndex(h => h.includes('code') || h.includes('id'));
  const nameIdx = headers.findIndex(h => h.includes('name') || h.includes('supplier'));
  const catIdx = headers.findIndex(h => h.includes('category'));
  const contactIdx = headers.findIndex(h => h.includes('contact') || h.includes('person'));
  const phoneIdx = headers.findIndex(h => h.includes('phone') || h.includes('mobile'));
  const emailIdx = headers.findIndex(h => h.includes('email') || h.includes('mail'));
  const statusIdx = headers.findIndex(h => h.includes('status'));

  const suppliers = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : String(row[0] || '').trim();
    const code = codeIdx !== -1 ? String(row[codeIdx] || '').trim() : String(row[1] || '').trim();
    if (!name && !code) continue;

    suppliers.push({
      id: code || ('SUP_' + String(i).padStart(3, '0')),
      code: code || ('SUP_' + String(i).padStart(3, '0')),
      name: name,
      category: catIdx !== -1 ? String(row[catIdx] || 'General') : 'General',
      contactPerson: contactIdx !== -1 ? String(row[contactIdx] || '') : '',
      phone: phoneIdx !== -1 ? String(row[phoneIdx] || '') : '',
      email: emailIdx !== -1 ? String(row[emailIdx] || '') : '',
      status: statusIdx !== -1 ? String(row[statusIdx] || 'Active') : 'Active'
    });
  }
  return suppliers;
}

function saveSupplier(supData) {
  const ss = getWorkbook(WORKBOOKS.SUPPLIER_MASTER);
  const sheet = ss.getSheetByName('Supplier_Master') || ss.getSheets()[0];
  if (!sheet) throw new Error('Supplier_Master sheet not found');

  const data = sheet.getDataRange().getValues();
  let supCode = supData.code || supData.id ? String(supData.code || supData.id).trim() : '';
  let targetRow = -1;

  if (supCode && data.length > 1) {
    const headers = data[0].map(h => String(h).trim().toLowerCase());
    const codeIdx = headers.findIndex(h => h.includes('code') || h.includes('id'));
    const searchCol = codeIdx !== -1 ? codeIdx : 1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][searchCol]).toUpperCase() === supCode.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      supData.name,
      supCode,
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
      supData.name,
      supCode,
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
  const sheet = ss.getSheetByName('Supplier_Master') || ss.getSheets()[0];
  if (!sheet) throw new Error('Supplier_Master sheet not found');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toUpperCase() === String(supCode).toUpperCase() || 
        String(data[i][1]).toUpperCase() === String(supCode).toUpperCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Supplier not found in sheet' };
}

// ==========================================
// 2. LOCATION MASTER (Store & Selling_Point)
// ==========================================
function getStores() {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Store') || ss.getSheets()[0];
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const nameIdx = headers.findIndex(h => h.includes('name'));
  const codeIdx = headers.findIndex(h => h.includes('code'));
  const statusIdx = headers.findIndex(h => h.includes('status'));

  const stores = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : String(row[0] || '').trim();
    const code = codeIdx !== -1 ? String(row[codeIdx] || '').trim() : String(row[1] || '').trim();
    if (!name && !code) continue;

    stores.push({
      code: code || ('S_' + String(i).padStart(3, '0')),
      name: name || code,
      type: 'Store',
      status: statusIdx !== -1 ? String(row[statusIdx] || 'Active') : 'Active'
    });
  }
  return stores;
}

function saveStore(storeData) {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  let sheet = ss.getSheetByName('Store');
  if (!sheet) sheet = ss.insertSheet('Store');

  const data = sheet.getDataRange().getValues();
  let code = storeData.code ? String(storeData.code).trim() : '';
  let targetRow = -1;

  if (code && data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toUpperCase() === code.toUpperCase() || String(data[i][0]).toUpperCase() === code.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 3).setValues([[
      storeData.name,
      code,
      storeData.status || 'Active'
    ]]);
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Store Name', 'Store Code', 'Status']);
      formatHeaderRow(sheet, 3);
    }
    if (!code) {
      code = 'S_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([storeData.name, code, storeData.status || 'Active']);
  }
  return { success: true, code: code };
}

function getSellingPoints() {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  const sheet = ss.getSheetByName('Selling_Point');
  if (!sheet) return [];

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const nameIdx = headers.findIndex(h => h.includes('name'));
  const codeIdx = headers.findIndex(h => h.includes('code'));
  const storeIdx = headers.findIndex(h => h.includes('store'));

  const sps = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : String(row[0] || '').trim();
    const code = codeIdx !== -1 ? String(row[codeIdx] || '').trim() : String(row[1] || '').trim();
    if (!name && !code) continue;

    sps.push({
      code: code || ('SP_' + String(i).padStart(3, '0')),
      name: name || code,
      storeCode: storeIdx !== -1 ? String(row[storeIdx] || 'S_001') : 'S_001',
      type: 'Selling Point',
      status: 'Active'
    });
  }
  return sps;
}

function saveSellingPoint(spData) {
  const ss = getWorkbook(WORKBOOKS.LOCATION_MASTER);
  let sheet = ss.getSheetByName('Selling_Point');
  if (!sheet) sheet = ss.insertSheet('Selling_Point');

  const data = sheet.getDataRange().getValues();
  let code = spData.code ? String(spData.code).trim() : '';
  let targetRow = -1;

  if (code && data.length > 1) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toUpperCase() === code.toUpperCase() || String(data[i][0]).toUpperCase() === code.toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 3).setValues([[
      spData.name,
      code,
      spData.status || 'Active'
    ]]);
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Selling Point Name', 'Selling Point Code', 'Status']);
      formatHeaderRow(sheet, 3);
    }
    if (!code) {
      code = 'SP_' + String(data.length).padStart(3, '0');
    }
    sheet.appendRow([spData.name, code, spData.status || 'Active']);
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

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const idIdx = headers.findIndex(h => h.includes('id') || h.includes('code'));
  const nameIdx = headers.findIndex(h => h.includes('name'));
  const roleIdx = headers.findIndex(h => h.includes('role'));
  const emailIdx = headers.findIndex(h => h.includes('email'));
  const statusIdx = headers.findIndex(h => h.includes('status'));

  const users = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : String(row[1] || '').trim();
    if (!name) continue;

    users.push({
      id: idIdx !== -1 ? String(row[idIdx] || '') : ('USR_' + String(i).padStart(3, '0')),
      name: name,
      role: roleIdx !== -1 ? String(row[roleIdx] || 'Staff') : 'Staff',
      email: emailIdx !== -1 ? String(row[emailIdx] || '') : '',
      status: statusIdx !== -1 ? String(row[statusIdx] || 'Active') : 'Active'
    });
  }
  return users;
}
