/**
 * DNP HOTELS - Suppliers, Departments & Locations Management
 */

function getSuppliersInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const suppliers = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;

    suppliers.push({
      id: String(row[0]),
      name: String(row[1]),
      contactPerson: String(row[2] || ''),
      email: String(row[3] || ''),
      phone: String(row[4] || ''),
      categorySupplied: String(row[5] || ''),
      address: String(row[6] || ''),
      status: String(row[7] || 'Active'),
      notes: String(row[8] || '')
    });
  }
  return suppliers;
}

function getSuppliers() {
  const ss = getSpreadsheet();
  return getSuppliersInternal(ss);
}

function saveSupplier(supData) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (!sheet) throw new Error('Suppliers sheet missing');

  const data = sheet.getDataRange().getValues();
  let supId = supData.id ? String(supData.id).trim() : '';
  let targetRow = -1;

  if (supId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === supId) {
        targetRow = i + 1;
        break;
      }
    }
  }

  if (targetRow > 1) {
    sheet.getRange(targetRow, 1, 1, 9).setValues([[
      supId,
      supData.name,
      supData.contactPerson || '',
      supData.email || '',
      supData.phone || '',
      supData.categorySupplied || '',
      supData.address || '',
      supData.status || 'Active',
      supData.notes || ''
    ]]);
  } else {
    if (!supId) {
      supId = 'SUP-' + (100 + data.length);
    }
    sheet.appendRow([
      supId,
      supData.name,
      supData.contactPerson || '',
      supData.email || '',
      supData.phone || '',
      supData.categorySupplied || '',
      supData.address || '',
      supData.status || 'Active',
      supData.notes || ''
    ]);
  }

  return { success: true, supplierId: supId };
}

function deleteSupplier(supId) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (!sheet) throw new Error('Suppliers sheet missing');

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(supId)) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Supplier not found' };
}

function getDepartmentsInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const depts = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    depts.push({
      id: String(row[0]),
      name: String(row[1]),
      propertyScope: String(row[2] || 'Both Properties'),
      headOfDepartment: String(row[3] || ''),
      notes: String(row[4] || '')
    });
  }
  return depts;
}

function getDepartments() {
  const ss = getSpreadsheet();
  return getDepartmentsInternal(ss);
}

function getLocationsInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.LOCATIONS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const locs = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1]) continue;
    locs.push({
      id: String(row[0]),
      name: String(row[1]),
      code: String(row[2]),
      type: String(row[3]),
      description: String(row[4] || '')
    });
  }
  return locs;
}

function getLocations() {
  const ss = getSpreadsheet();
  return getLocationsInternal(ss);
}
