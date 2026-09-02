/**
 * DNP HOTELS - Multi-Workbook Database Architecture
 * Manages bifurcated workbooks inside Google Drive Folder: 1lkSx36mqaqnF8gfqNswdSPb0zqY4lvOx
 */

const DEFAULT_DRIVE_FOLDER_ID = '1lkSx36mqaqnF8gfqNswdSPb0zqY4lvOx';

const WORKBOOKS = {
  LOCATION_MASTER: 'Location_Master',
  PRODUCT_MASTER: 'Product_Master',
  SUPPLIER_MASTER: 'Supplier_Master',
  SUPPLIER_TXNS: 'Supplier_Transactions',
  ISSUANCE_TXNS: 'Issuance_Transactions',
  USERS_SETTINGS: 'Users_and_Settings'
};

/**
 * Gets the designated Google Drive Folder
 */
function getDriveFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID;
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    console.warn('Could not open Drive folder: ' + folderId, err);
    return null;
  }
}

/**
 * Formats table header row with luxury styling
 */
function formatHeaderRow(sheet, numCols) {
  const range = sheet.getRange(1, 1, 1, numCols);
  range.setBackground('#11131b');
  range.setFontColor('#b4c5ff');
  range.setFontWeight('bold');
  range.setFontFamily('Roboto');
  range.setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  try {
    for (let c = 1; c <= numCols; c++) {
      sheet.autoResizeColumn(c);
    }
  } catch (e) {}
}

/**
 * Gets or creates a specific workbook inside the Google Drive folder
 */
function getWorkbook(name) {
  const props = PropertiesService.getScriptProperties();
  const cachedId = props.getProperty('WB_ID_' + name);
  if (cachedId) {
    try {
      return SpreadsheetApp.openById(cachedId);
    } catch (e) {
      // Cached ID was deleted or inaccessible, will re-fetch
    }
  }

  const folder = getDriveFolder();
  if (folder) {
    const files = folder.getFilesByName(name);
    if (files.hasNext()) {
      const file = files.next();
      const ss = SpreadsheetApp.openById(file.getId());
      props.setProperty('WB_ID_' + name, ss.getId());
      return ss;
    }
  }

  // Not found in folder, create and initialize it
  return createAndInitWorkbook(name);
}

/**
 * Creates and sets up initial schema & sample data for a workbook
 */
function createAndInitWorkbook(name) {
  const folder = getDriveFolder();
  const ss = SpreadsheetApp.create(name);
  const ssId = ss.getId();
  PropertiesService.getScriptProperties().setProperty('WB_ID_' + name, ssId);

  // Move into Drive folder
  if (folder) {
    try {
      const file = DriveApp.getFileById(ssId);
      file.moveTo(folder);
    } catch (err) {
      console.warn('Could not move workbook to Drive folder', err);
    }
  }

  // Schema initialization based on workbook type
  if (name === WORKBOOKS.LOCATION_MASTER) {
    initLocationMaster(ss);
  } else if (name === WORKBOOKS.PRODUCT_MASTER) {
    initProductMaster(ss);
  } else if (name === WORKBOOKS.SUPPLIER_MASTER) {
    initSupplierMaster(ss);
  } else if (name === WORKBOOKS.SUPPLIER_TXNS) {
    initSupplierTransactions(ss);
  } else if (name === WORKBOOKS.ISSUANCE_TXNS) {
    initIssuanceTransactions(ss);
  } else if (name === WORKBOOKS.USERS_SETTINGS) {
    initUsersAndSettings(ss);
  }

  return ss;
}

/**
 * 1. Location_Master Workbook (Store & Selling_Point sheets)
 */
function initLocationMaster(ss) {
  // Sheet 1: Store
  let storeSheet = ss.getSheetByName('Store');
  if (!storeSheet) {
    storeSheet = ss.getSheets()[0];
    storeSheet.setName('Store');
  }
  if (storeSheet.getLastRow() === 0) {
    storeSheet.appendRow(['Store Code', 'Store Name', 'Type', 'Status', 'Description']);
    formatHeaderRow(storeSheet, 5);
    storeSheet.appendRow(['S_001', '21 GUN SOLUTE GGN SEC 29', 'Main Outlet Store', 'Active', 'Sector 29 Gurgaon Property']);
    storeSheet.appendRow(['S_002', 'PAHLE CHAI GGN Sec 27', 'Outlet Store', 'Active', 'Sector 27 Gurgaon Outlet']);
    storeSheet.appendRow(['S_000', 'Central Depot Warehouse', 'Central Depot', 'Active', 'Central Replenishment Warehouse']);
  }

  // Sheet 2: Selling_Point
  let spSheet = ss.getSheetByName('Selling_Point');
  if (!spSheet) {
    spSheet = ss.insertSheet('Selling_Point');
  }
  if (spSheet.getLastRow() === 0) {
    spSheet.appendRow(['Selling Point Code', 'Selling Point Name', 'Assigned Store Code', 'Type', 'Status']);
    formatHeaderRow(spSheet, 5);
    spSheet.appendRow(['SP_001', '21 GUN SOLUTE GGN SEC 29', 'S_001', 'Dining & Bar', 'Active']);
    spSheet.appendRow(['SP_002', 'PAHLE CHAI GGN Sec 27', 'S_002', 'Chai & Cafe Counter', 'Active']);
    spSheet.appendRow(['SP_003', '21 Gun Salute - Kitchen Store', 'S_001', 'F&B Production', 'Active']);
    spSheet.appendRow(['SP_004', '21 Gun Salute - Bar & Lounge', 'S_001', 'Beverage Counter', 'Active']);
  }
}

/**
 * 2. Product_Master Workbook (Product_Master sheet)
 */
function initProductMaster(ss) {
  let prodSheet = ss.getSheetByName('Product_Master');
  if (!prodSheet) {
    prodSheet = ss.getSheets()[0];
    prodSheet.setName('Product_Master');
  }

  if (prodSheet.getLastRow() === 0) {
    prodSheet.appendRow([
      'Item Code',
      'Item Description',
      'Category',
      'Category Code',
      'UOM',
      'Rate',
      'TAX %',
      'Min Stock',
      'Stock S_001',
      'Stock S_002',
      'Central Stock',
      'Total Stock',
      'Total Valuation',
      'Preferred Supplier Code',
      'Status',
      'Last Updated'
    ]);
    formatHeaderRow(prodSheet, 16);

    const nowIso = new Date().toISOString();
    const demoItems = [
      ['ITM_001', 'Special Assam Orthodox Tea Leaves', 'Tea & Beverages', 'CAT_TEA', 'Kg', 650, 5, 10, 25, 15, 40, 80, 52000, 'SUP_021', 'Active', nowIso],
      ['ITM_002', 'Fresh Cow Milk 1L Pack', 'Dairy & Fresh', 'CAT_DAI', 'Pack', 65, 0, 30, 40, 30, 50, 120, 7800, 'SUP_004', 'Active', nowIso],
      ['ITM_003', 'Fresh Malai Paneer Block', 'Dairy & Fresh', 'CAT_DAI', 'Kg', 380, 0, 15, 15, 10, 20, 45, 17100, 'SUP_004', 'Active', nowIso],
      ['ITM_004', 'Artisan Multigrain Bread Loaf', 'Bakery & Desserts', 'CAT_BAK', 'Loaf', 90, 5, 20, 20, 15, 15, 50, 4500, 'SUP_005', 'Active', nowIso],
      ['ITM_005', 'Kesar Pista Kulfi Sticks', 'Bakery & Desserts', 'CAT_BAK', 'Pcs', 45, 18, 50, 60, 40, 50, 150, 6750, 'SUP_014', 'Active', nowIso],
      ['ITM_006', 'Hardwood Charcoal Briquettes 10kg', 'Kitchen Fuel', 'CAT_FUE', 'Bags', 850, 18, 10, 15, 5, 15, 35, 29750, 'SUP_008', 'Active', nowIso],
      ['ITM_007', 'Heavy Duty Floor Cleaner 5L', 'Housekeeping & Hygiene', 'CAT_HK', 'Can', 550, 18, 10, 10, 5, 10, 25, 13750, 'SUP_024', 'Active', nowIso],
      ['ITM_008', 'A4 Executive Copier Paper 75GSM', 'Stationery & Office', 'CAT_STA', 'Ream', 280, 12, 10, 10, 5, 15, 30, 8400, 'SUP_006', 'Active', nowIso],
      ['ITM_009', 'LED Warm Spotlight 12W GU10', 'Engineering & Electrical', 'CAT_ENG', 'Pcs', 195, 18, 15, 15, 10, 15, 40, 7800, 'SUP_009', 'Active', nowIso],
      ['ITM_010', 'Packaged Drinking Water 500ml (Crate 24)', 'Beverages', 'CAT_BEV', 'Crate', 240, 18, 20, 25, 15, 20, 60, 14400, 'SUP_003', 'Active', nowIso],
      ['ITM_011', 'Basmati Rice Classic 25kg Bag', 'Grocery & Staples', 'CAT_GRO', 'Bags', 2400, 5, 5, 8, 4, 6, 18, 43200, 'SUP_019', 'Active', nowIso],
      ['ITM_012', 'Signature Chai Cardboard Cups 150ml (Pack 100)', 'Packaging & Disposables', 'CAT_PAC', 'Pack', 180, 18, 30, 40, 30, 20, 90, 16200, 'SUP_023', 'Active', nowIso]
    ];

    prodSheet.getRange(2, 1, demoItems.length, demoItems[0].length).setValues(demoItems);
  }
}

/**
 * 3. Supplier_Master Workbook (all 28 real suppliers)
 */
function initSupplierMaster(ss) {
  let supSheet = ss.getSheetByName('Supplier_Master');
  if (!supSheet) {
    supSheet = ss.getSheets()[0];
    supSheet.setName('Supplier_Master');
  }

  if (supSheet.getLastRow() === 0) {
    supSheet.appendRow(['Supplier Code', 'Supplier Name', 'Category', 'Contact Person', 'Phone', 'Email', 'Status']);
    formatHeaderRow(supSheet, 7);

    const suppliersData = [
      ['SUP_001', 'A.T. Overseas', 'Imported Gourmet & Dry Foods', 'Mr. Amit Thapar', '+91 98101 23456', 'sales@atoverseas.com', 'Active'],
      ['SUP_002', 'Agarwal Enterprises', 'Provisions & Kirana', 'Mr. R. K. Agarwal', '+91 98112 34567', 'agarwal.ent@gmail.com', 'Active'],
      ['SUP_003', 'Ajay Cold Drinks', 'Beverages & Soft Drinks', 'Mr. Ajay Verma', '+91 98123 45678', 'ajaycolddrinks@yahoo.com', 'Active'],
      ['SUP_004', 'Bankey Behari Dairy & Paneer Bhandar', 'Fresh Dairy, Paneer & Mawa', 'Mr. Bankey Lal', '+91 98134 56789', 'bbdairy@gmail.com', 'Active'],
      ['SUP_005', "Chef's Bakeology Ifo Dayal Singh", 'Breads, Buns & Pastries', 'Chef Dayal Singh', '+91 98145 67890', 'bakeology@outlook.com', 'Active'],
      ['SUP_006', 'Deepak Stationery Mart', 'Office Stationery & Billing Paper', 'Mr. Deepak Jain', '+91 98156 78901', 'deepakstationery@gmail.com', 'Active'],
      ['SUP_007', 'Design Xpress', 'Menu Cards, Collateral & Printing', 'Mr. Nitin Saxena', '+91 98167 89012', 'print@designxpress.in', 'Active'],
      ['SUP_008', 'Essel Charcoals', 'Tandoor Charcoal & Fuel', 'Mr. S. L. Singhania', '+91 98178 90123', 'esselcharcoals@gmail.com', 'Active'],
      ['SUP_009', 'Friends Electric Works', 'Lighting, Bulbs & Electricals', 'Mr. Joginder Pal', '+91 98189 01234', 'friendselectric@gmail.com', 'Active'],
      ['SUP_010', 'Goodwill Traders', 'General Hospitality Hardware', 'Mr. Manpreet Singh', '+91 98190 12345', 'goodwilltraders@gmail.com', 'Active'],
      ['SUP_011', 'Hari Shankar Singh', 'Vegetables & Fresh Greens', 'Mr. Hari Shankar', '+91 98201 23456', 'harishankar.fresh@gmail.com', 'Active'],
      ['SUP_012', 'Hot Cakes Private Limited', 'Confectionery & Special Cakes', 'Ms. Ananya Roy', '+91 98212 34567', 'orders@hotcakes.co.in', 'Active'],
      ['SUP_013', 'Jai Guru Ji Traders', 'Spices, Masala & Seasonings', 'Mr. Gurpreet Sethi', '+91 98223 45678', 'jaigurujitraders@gmail.com', 'Active'],
      ['SUP_014', 'K B Kulfi', 'Traditional Kulfi & Ice Creams', 'Mr. Kailash B.', '+91 98234 56789', 'kbkulfi@gmail.com', 'Active'],
      ['SUP_015', 'Kanshi Ram Enterprises', 'Commercial Kitchen Crockery', 'Mr. Kanshi Ram', '+91 98245 67890', 'krrkitchenware@gmail.com', 'Active'],
      ['SUP_016', 'L R Wholesale Services Pvt. Ltd.', 'Bulk Institutional Supplies', 'Mr. Lalit Rao', '+91 98256 78901', 'info@lrwholesale.com', 'Active'],
      ['SUP_017', 'M S Manufactures and Distributors Pvt Ltd.', 'Linens & Table Runners', 'Mr. M. S. Sodhi', '+91 98267 89012', 'msdistributors@gmail.com', 'Active'],
      ['SUP_018', 'MR International', 'Premium Cutlery & Holloware', 'Mr. Manish Rawat', '+91 98278 90123', 'mrinternational@gmail.com', 'Active'],
      ['SUP_019', 'Navpallav Agro Products Pvt. Ltd', 'Grains, Basmati Rice & Pulses', 'Mr. Pallav Gupta', '+91 98289 01234', 'agro@navpallav.com', 'Active'],
      ['SUP_020', 'New Harry Store', 'Imported Condiments & Syrups', 'Mr. Harinder Suri', '+91 98290 12345', 'harrystore@gmail.com', 'Active'],
      ['SUP_021', 'Pindi Kirana Store', 'Premium Tea, Sugar & Dry Grocery', 'Mr. Satish Pindi', '+91 98301 23456', 'pindikirana@gmail.com', 'Active'],
      ['SUP_022', 'Rajesh Kumar (29 Jaguar)', 'Poultry & Fresh Meats', 'Mr. Rajesh Kumar', '+91 98312 34567', 'rajesh29jaguar@gmail.com', 'Active'],
      ['SUP_023', 'Sms Commercial', 'Takeaway Cups, Boxes & Disposables', 'Mr. Sanjay Mishra', '+91 98323 45678', 'smscommercial@gmail.com', 'Active'],
      ['SUP_024', 'Sms Housekeeping', 'Cleaning Chemicals & Detergents', 'Mr. S. M. Sharma', '+91 98334 56789', 'smshk@gmail.com', 'Active'],
      ['SUP_025', 'Sms Marketing Solutions', 'Branded Signage & Packaging', 'Mr. Shivam Malhotra', '+91 98345 67890', 'smsmarketing@gmail.com', 'Active'],
      ['SUP_026', 'SRD Traders', 'Cooking Oils & Ghee', 'Mr. Suresh Dani', '+91 98356 78901', 'srdtraders@gmail.com', 'Active'],
      ['SUP_027', 'T K Traders', 'Cleaning Tools, Brooms & Mops', 'Mr. Tarun Kalra', '+91 98367 89012', 'tktraders@gmail.com', 'Active'],
      ['SUP_028', 'Tulip Enterprises', 'Tissue Paper, Napkins & Foil', 'Ms. Meenakshi Tulip', '+91 98378 90123', 'tulipenterprises@gmail.com', 'Active']
    ];

    supSheet.getRange(2, 1, suppliersData.length, suppliersData[0].length).setValues(suppliersData);
  }
}

/**
 * 4. Supplier_Transactions Workbook (Supplier-wise Purchasing & Stock In)
 */
function initSupplierTransactions(ss) {
  let sheet = ss.getSheetByName('Supplier_Transactions');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('Supplier_Transactions');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Transaction ID',
      'Timestamp',
      'Supplier Code',
      'Supplier Name',
      'Item Code',
      'Item Description',
      'Category',
      'Quantity',
      'UOM',
      'Rate',
      'Tax %',
      'Total Amount',
      'Receiving Store Code',
      'Receiving Store Name',
      'PO / Invoice Ref',
      'Received By',
      'Notes'
    ]);
    formatHeaderRow(sheet, 17);

    const now = new Date();
    const demoTxns = [
      ['TXN_SUP_001', new Date(now.getTime() - 86400000 * 3).toISOString(), 'SUP_004', 'Bankey Behari Dairy & Paneer Bhandar', 'ITM_003', 'Fresh Malai Paneer Block', 'Dairy & Fresh', 30, 'Kg', 380, 0, 11400, 'S_001', '21 GUN SOLUTE GGN SEC 29', 'INV-BBD-991', 'Head Chef', 'Morning delivery received in cold crate'],
      ['TXN_SUP_002', new Date(now.getTime() - 86400000 * 2).toISOString(), 'SUP_021', 'Pindi Kirana Store', 'ITM_001', 'Special Assam Orthodox Tea Leaves', 'Tea & Beverages', 40, 'Kg', 650, 5, 27300, 'S_000', 'Central Depot Warehouse', 'PO-2026-441', 'Store Keeper', 'Monthly bulk tea consignment'],
      ['TXN_SUP_003', new Date(now.getTime() - 86400000 * 1).toISOString(), 'SUP_023', 'Sms Commercial', 'ITM_012', 'Signature Chai Cardboard Cups 150ml (Pack 100)', 'Packaging & Disposables', 50, 'Pack', 180, 18, 10620, 'S_002', 'PAHLE CHAI GGN Sec 27', 'INV-SMS-802', 'Outlet Mgr', 'Restock for takeaway chai counter'],
      ['TXN_SUP_004', new Date(now.getTime() - 3600000 * 4).toISOString(), 'SUP_008', 'Essel Charcoals', 'ITM_006', 'Hardwood Charcoal Briquettes 10kg', 'Kitchen Fuel', 20, 'Bags', 850, 18, 20060, 'S_001', '21 GUN SOLUTE GGN SEC 29', 'INV-EC-109', 'Tandoor Chef', 'High calorific value coal batch']
    ];

    sheet.getRange(2, 1, demoTxns.length, demoTxns[0].length).setValues(demoTxns);
  }
}

/**
 * 5. Issuance_Transactions Workbook (Store to Selling Point & Department Issues)
 */
function initIssuanceTransactions(ss) {
  let sheet = ss.getSheetByName('Issuance_Transactions');
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheet.setName('Issuance_Transactions');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Issuance ID',
      'Timestamp',
      'Type',
      'Item Code',
      'Item Description',
      'Quantity',
      'UOM',
      'From Store Code',
      'From Store Name',
      'To Selling Point Code',
      'To Selling Point Name',
      'Unit Rate',
      'Total Value',
      'Requisition Ref',
      'Issued By',
      'Status',
      'Notes'
    ]);
    formatHeaderRow(sheet, 17);

    const now = new Date();
    const demoIssues = [
      ['ISS_001', new Date(now.getTime() - 86400000 * 2).toISOString(), 'DISBURSEMENT', 'ITM_003', 'Fresh Malai Paneer Block', 15, 'Kg', 'S_001', '21 GUN SOLUTE GGN SEC 29', 'SP_003', '21 Gun Salute - Kitchen Store', 380, 5700, 'REQ-8812', 'Store Keeper', 'Approved', 'Daily dinner prep issuance'],
      ['ISS_002', new Date(now.getTime() - 86400000 * 1).toISOString(), 'DISBURSEMENT', 'ITM_001', 'Special Assam Orthodox Tea Leaves', 10, 'Kg', 'S_000', 'Central Depot Warehouse', 'SP_002', 'PAHLE CHAI GGN Sec 27', 650, 6500, 'REQ-6812', 'Logistics Mgr', 'Shipped', 'Replenishment for Pahle Chai Sec 27'],
      ['ISS_003', new Date(now.getTime() - 3600000 * 6).toISOString(), 'DISBURSEMENT', 'ITM_012', 'Signature Chai Cardboard Cups 150ml', 15, 'Pack', 'S_002', 'PAHLE CHAI GGN Sec 27', 'SP_002', 'PAHLE CHAI GGN Sec 27', 180, 2700, 'REQ-6815', 'Supervisor', 'Approved', 'Counter dispensers reload']
    ];

    sheet.getRange(2, 1, demoIssues.length, demoIssues[0].length).setValues(demoIssues);
  }
}

/**
 * 6. Users_and_Settings Workbook (Users & Settings sheets)
 */
function initUsersAndSettings(ss) {
  // Sheet 1: Users
  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.getSheets()[0];
    usersSheet.setName('Users');
  }
  if (usersSheet.getLastRow() === 0) {
    usersSheet.appendRow(['User ID', 'Name', 'Role', 'Email', 'Assigned Store / Selling Point', 'Status']);
    formatHeaderRow(usersSheet, 6);
    usersSheet.appendRow(['USR_001', 'Alsha Khan', 'Logistics & Inventory Manager', 'alsha.khan@dnphotels.com', 'ALL', 'Active']);
    usersSheet.appendRow(['USR_002', 'Store Incharge - GGN Sec 29', 'Store Keeper', 'store29@dnphotels.com', 'S_001', 'Active']);
    usersSheet.appendRow(['USR_003', 'Pahle Chai Supervisor - GGN Sec 27', 'Outlet Supervisor', 'pahlechai27@dnphotels.com', 'SP_002', 'Active']);
  }

  // Sheet 2: Settings
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
  }
  if (settingsSheet.getLastRow() === 0) {
    settingsSheet.appendRow(['Key', 'Value', 'Description']);
    formatHeaderRow(settingsSheet, 3);
    settingsSheet.appendRow(['HOTEL_NAME', '21 Gun Salute & Pahle Chai (DNP Hotels)', 'Organization Brand Name']);
    settingsSheet.appendRow(['CURRENCY_SYMBOL', '₹', 'Display Currency Symbol']);
    settingsSheet.appendRow(['DRIVE_FOLDER_ID', DEFAULT_DRIVE_FOLDER_ID, 'Designated Database Drive Folder']);
  }
}

/**
 * Validates and provisions all 6 workbooks inside the user's Drive folder
 */
function initAllWorkbooks() {
  const list = [
    WORKBOOKS.LOCATION_MASTER,
    WORKBOOKS.PRODUCT_MASTER,
    WORKBOOKS.SUPPLIER_MASTER,
    WORKBOOKS.SUPPLIER_TXNS,
    WORKBOOKS.ISSUANCE_TXNS,
    WORKBOOKS.USERS_SETTINGS
  ];

  const results = {};
  list.forEach(name => {
    const wb = getWorkbook(name);
    results[name] = {
      id: wb.getId(),
      name: wb.getName(),
      url: wb.getUrl(),
      sheets: wb.getSheets().map(s => s.getName())
    };
  });
  return results;
}

/**
 * Returns metadata of all 6 workbooks
 */
function getWorkbooksInfo() {
  const folder = getDriveFolder();
  const folderId = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID;

  const names = Object.values(WORKBOOKS);
  const workbooks = names.map(name => {
    try {
      const wb = getWorkbook(name);
      return {
        key: name,
        id: wb.getId(),
        name: wb.getName(),
        url: wb.getUrl(),
        sheets: wb.getSheets().map(s => s.getName())
      };
    } catch (e) {
      return {
        key: name,
        id: null,
        name: name,
        url: '#',
        sheets: []
      };
    }
  });

  return {
    folder: {
      id: folderId,
      name: folder ? folder.getName() : 'DNP Database Drive Folder',
      url: folder ? folder.getUrl() : 'https://drive.google.com/drive/folders/' + folderId
    },
    workbooks: workbooks
  };
}

/**
 * Settings helpers
 */
function getSettings() {
  try {
    const ss = getWorkbook(WORKBOOKS.USERS_SETTINGS);
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) return { CURRENCY_SYMBOL: '₹', HOTEL_NAME: '21 Gun Salute & Pahle Chai' };
    const data = sheet.getDataRange().getValues();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) {
        settings[String(data[i][0])] = String(data[i][1]);
      }
    }
    return settings;
  } catch (e) {
    return { CURRENCY_SYMBOL: '₹', HOTEL_NAME: '21 Gun Salute & Pahle Chai' };
  }
}

function saveSettings(settingsObj) {
  const ss = getWorkbook(WORKBOOKS.USERS_SETTINGS);
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) throw new Error('Settings sheet missing in Users_and_Settings');

  const data = sheet.getDataRange().getValues();
  for (const key in settingsObj) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === key) {
        sheet.getRange(i + 1, 2).setValue(String(settingsObj[key]));
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, String(settingsObj[key]), 'Custom setting']);
    }
  }
  return { success: true };
}

function setDriveFolderId(folderUrlOrId) {
  if (!folderUrlOrId || typeof folderUrlOrId !== 'string') {
    return { success: false, error: 'Please provide a valid Google Drive folder URL or ID.' };
  }
  let folderId = folderUrlOrId.trim();
  const match = folderId.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    folderId = match[1];
  }
  try {
    const folder = DriveApp.getFolderById(folderId);
    PropertiesService.getScriptProperties().setProperty('DRIVE_FOLDER_ID', folderId);
    // Clear cached workbook IDs to rebind to new folder
    const props = PropertiesService.getScriptProperties();
    Object.values(WORKBOOKS).forEach(wb => props.deleteProperty('WB_ID_' + wb));

    return {
      success: true,
      id: folderId,
      name: folder.getName(),
      url: folder.getUrl(),
      message: 'Connected to Drive folder: ' + folder.getName()
    };
  } catch (err) {
    return {
      success: false,
      error: 'Could not access the specified Google Drive folder. Details: ' + err.message
    };
  }
}
