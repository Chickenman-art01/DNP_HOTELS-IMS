/**
 * DNP HOTELS - IMS Database & Spreadsheet Management
 */

const SHEET_NAMES = {
  ITEMS: 'Items',
  TRANSACTIONS: 'Transactions',
  SUPPLIERS: 'Suppliers',
  DEPARTMENTS: 'Departments',
  LOCATIONS: 'Locations',
  SETTINGS: 'Settings'
};

const DEFAULT_DRIVE_FOLDER_ID = '1lkSx36mqaqnF8gfqNswdSPb0zqY4lvOx';

/**
 * Gets the connected Google Drive folder for database storage
 */
function getDriveFolder() {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID;
  try {
    return DriveApp.getFolderById(folderId);
  } catch (err) {
    console.warn('Could not access Google Drive folder by ID: ' + folderId, err);
    return null;
  }
}

/**
 * Scans the Google Drive folder for all spreadsheets
 */
function getFolderSheets() {
  const folder = getDriveFolder();
  if (!folder) return [];
  const sheets = [];
  try {
    const files = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      const f = files.next();
      sheets.push({
        id: f.getId(),
        name: f.getName(),
        url: f.getUrl(),
        lastUpdated: f.getLastUpdated().toISOString()
      });
    }
  } catch (err) {
    console.warn('Error fetching sheets in Drive folder', err);
  }
  return sheets;
}

/**
 * Updates the designated Google Drive folder ID or URL
 */
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
      error: 'Could not access the specified Google Drive folder. Please verify permissions. Details: ' + err.message
    };
  }
}

/**
 * Gets or initializes the connected Google Spreadsheet.
 * If not connected or invalid, checks the Drive folder or creates a new spreadsheet inside it.
 */
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let sheetId = props.getProperty('SPREADSHEET_ID');
  let ss = null;

  if (sheetId) {
    try {
      ss = SpreadsheetApp.openById(sheetId);
    } catch (err) {
      console.warn('Could not open spreadsheet by saved ID: ' + sheetId, err);
      ss = null;
    }
  }

  // If container-bound, try active spreadsheet
  if (!ss) {
    try {
      const active = SpreadsheetApp.getActiveSpreadsheet();
      if (active) {
        ss = active;
        sheetId = ss.getId();
        props.setProperty('SPREADSHEET_ID', sheetId);
      }
    } catch (e) {
      // Standalone script
    }
  }

  // Check designated Google Drive folder for existing sheets
  const folder = getDriveFolder();

  if (!ss && folder) {
    try {
      const folderSheets = getFolderSheets();
      if (folderSheets.length > 0) {
        const match = folderSheets.find(s => s.name.toUpperCase().includes('IMS') || s.name.toUpperCase().includes('DATABASE')) || folderSheets[0];
        ss = SpreadsheetApp.openById(match.id);
        sheetId = ss.getId();
        props.setProperty('SPREADSHEET_ID', sheetId);
      }
    } catch (e) {
      console.warn('Could not auto-detect sheet from folder', e);
    }
  }

  // If still not available, create a new spreadsheet directly inside the Drive folder
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('DNP HOTELS - IMS Database');
      sheetId = ss.getId();
      props.setProperty('SPREADSHEET_ID', sheetId);

      if (folder) {
        try {
          const file = DriveApp.getFileById(sheetId);
          file.moveTo(folder);
        } catch (e) {
          console.warn('Could not move new database into Drive folder', e);
        }
      }

      initDatabase(ss);
      seedSampleDataIfEmpty(ss);
    } catch (err) {
      console.error('Failed to create new spreadsheet', err);
      throw new Error('Unable to access or create Google Sheet database: ' + err.message);
    }
  } else {
    // Ensure spreadsheet is housed in the configured Drive folder
    if (folder && sheetId) {
      try {
        const file = DriveApp.getFileById(sheetId);
        const parents = file.getParents();
        let inFolder = false;
        while (parents.hasNext()) {
          if (parents.next().getId() === folder.getId()) {
            inFolder = true;
            break;
          }
        }
        if (!inFolder) {
          file.moveTo(folder);
        }
      } catch (e) {}
    }

    initDatabase(ss);
  }

  return ss;
}

/**
 * Validates and ensures all necessary sheets and columns exist.
 */
function initDatabase(ss) {
  // 1. Items Sheet
  let itemsSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
  if (!itemsSheet) {
    itemsSheet = ss.insertSheet(SHEET_NAMES.ITEMS);
    itemsSheet.appendRow([
      'Item ID',
      'SKU / Barcode',
      'Item Name',
      'Category',
      'Unit',
      'Min Stock',
      'Central Stock',
      'Deneb Stock',
      'Pollux Stock',
      'Total Stock',
      'Unit Cost',
      'Total Value',
      'Primary Supplier',
      'Storage Location',
      'Status',
      'Last Updated'
    ]);
    formatHeaderRow(itemsSheet, 16);
  }

  // 2. Transactions Sheet
  let txnSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (!txnSheet) {
    txnSheet = ss.insertSheet(SHEET_NAMES.TRANSACTIONS);
    txnSheet.appendRow([
      'Transaction ID',
      'Timestamp',
      'Type',
      'Item ID',
      'SKU',
      'Item Name',
      'Category',
      'Quantity',
      'Unit',
      'Source Location',
      'Destination / Dept',
      'Unit Price',
      'Total Cost',
      'Reference / PO',
      'Performed By',
      'Notes'
    ]);
    formatHeaderRow(txnSheet, 16);
  }

  // 3. Suppliers Sheet
  let supSheet = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (!supSheet) {
    supSheet = ss.insertSheet(SHEET_NAMES.SUPPLIERS);
    supSheet.appendRow([
      'Supplier ID',
      'Supplier Name',
      'Contact Person',
      'Email',
      'Phone',
      'Category Supplied',
      'Address',
      'Status',
      'Notes'
    ]);
    formatHeaderRow(supSheet, 9);
  }

  // 4. Departments Sheet
  let deptSheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (!deptSheet) {
    deptSheet = ss.insertSheet(SHEET_NAMES.DEPARTMENTS);
    deptSheet.appendRow([
      'Department ID',
      'Department Name',
      'Property Scope',
      'Head of Department',
      'Notes'
    ]);
    formatHeaderRow(deptSheet, 5);
  }

  // 5. Locations Sheet
  let locSheet = ss.getSheetByName(SHEET_NAMES.LOCATIONS);
  if (!locSheet) {
    locSheet = ss.insertSheet(SHEET_NAMES.LOCATIONS);
    locSheet.appendRow([
      'Location ID',
      'Location Name',
      'Code',
      'Type',
      'Description'
    ]);
    formatHeaderRow(locSheet, 5);
  }

  // 6. Settings Sheet
  let setSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!setSheet) {
    setSheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
    setSheet.appendRow(['Key', 'Value', 'Description']);
    formatHeaderRow(setSheet, 3);
    setSheet.appendRow(['CURRENCY_SYMBOL', '₹', 'Display currency symbol']);
    setSheet.appendRow(['HOTEL_NAME', 'DNP HOTELS', 'Organization title']);
    setSheet.appendRow(['LOW_STOCK_THRESHOLD_DEFAULT', '10', 'Default minimum alert threshold']);
  }

  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(sheet1);
    } catch (e) {}
  }
}

function formatHeaderRow(sheet, colCount) {
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

function seedSampleDataIfEmpty(ss) {
  const itemsSheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
  if (itemsSheet && itemsSheet.getLastRow() <= 1) {
    const demoItems = [
      ['ITM-1001', 'DP-HK-001', 'King Size Bed Sheet 400TC (White)', 'Housekeeping & Linens', 'Pcs', 20, 50, 25, 20, 95, 1200, 114000, 'Apex Hospitality Linens', 'Aisle 1 - Shelf A', 'Active', new Date().toISOString()],
      ['ITM-1002', 'DP-HK-002', 'Bath Towel 700GSM Luxury Cotton', 'Housekeeping & Linens', 'Pcs', 30, 80, 40, 35, 155, 650, 100750, 'Apex Hospitality Linens', 'Aisle 1 - Shelf B', 'Active', new Date().toISOString()],
      ['ITM-1003', 'DP-AM-001', 'Organic Shampoo 30ml Eco-Bottle', 'Guest Amenities & Toiletries', 'Bottles', 100, 400, 150, 120, 670, 25, 16750, 'Luxe Spa & Botanicals', 'Aisle 2 - Shelf A', 'Active', new Date().toISOString()],
      ['ITM-1004', 'DP-AM-002', 'Organic Body Wash 30ml Eco-Bottle', 'Guest Amenities & Toiletries', 'Bottles', 100, 350, 130, 110, 590, 25, 14750, 'Luxe Spa & Botanicals', 'Aisle 2 - Shelf A', 'Active', new Date().toISOString()],
      ['ITM-1005', 'DP-AM-003', 'Dental Kit with Bamboo Toothbrush', 'Guest Amenities & Toiletries', 'Packs', 80, 250, 90, 80, 420, 35, 14700, 'Luxe Spa & Botanicals', 'Aisle 2 - Shelf B', 'Active', new Date().toISOString()],
      ['ITM-1006', 'DP-FB-001', 'Arabica Dark Roast Coffee Beans 1kg', 'Food & Beverage', 'Bags', 15, 40, 12, 10, 62, 1100, 68200, 'Estate Harvest Roasters', 'Pantry 1 - Cool Rack', 'Active', new Date().toISOString()],
      ['ITM-1007', 'DP-FB-002', 'Twinings Earl Grey Tea (100 Bags)', 'Food & Beverage', 'Boxes', 10, 30, 8, 7, 45, 680, 30600, 'Estate Harvest Roasters', 'Pantry 1 - Shelf C', 'Active', new Date().toISOString()],
      ['ITM-1008', 'DP-FB-003', 'Organic Extra Virgin Olive Oil 5L', 'Food & Beverage', 'Tins', 8, 20, 6, 5, 31, 3200, 99200, 'Gourmet Pantry Supplies', 'Kitchen Dry Store', 'Active', new Date().toISOString()],
      ['ITM-1009', 'DP-FO-001', 'RFID Keycards (DNP Branded)', 'Front Office & Stationery', 'Cards', 200, 600, 200, 150, 950, 45, 42750, 'TechSmart RFID Systems', 'Front Desk Safe', 'Active', new Date().toISOString()],
      ['ITM-1010', 'DP-FO-002', 'A4 Executive Hotel Letterhead (500s)', 'Front Office & Stationery', 'Reams', 5, 15, 5, 4, 24, 450, 10800, 'PrintCraft Media', 'Admin Supply Room', 'Active', new Date().toISOString()],
      ['ITM-1011', 'DP-MN-001', 'LED Bulb 9W Warm White E27', 'Maintenance & Engineering', 'Pcs', 25, 60, 15, 12, 87, 140, 12180, 'ElectroTech Solutions', 'Maintenance Store Bay 4', 'Active', new Date().toISOString()],
      ['ITM-1012', 'DP-MN-002', 'HVAC Air Filter 20x20 High Efficiency', 'Maintenance & Engineering', 'Pcs', 12, 30, 8, 6, 44, 450, 19800, 'ElectroTech Solutions', 'HVAC Plant Room', 'Active', new Date().toISOString()],
      ['ITM-1013', 'DP-CL-001', 'Commercial Multi-Surface Cleaner 5L', 'Cleaning & Chemicals', 'Cans', 10, 25, 8, 7, 40, 850, 34000, 'HygieneFirst Chemicals', 'Chemical Bunker 2', 'Active', new Date().toISOString()],
      ['ITM-1014', 'DP-CL-002', 'High Efficiency Laundry Detergent 25kg', 'Cleaning & Chemicals', 'Drums', 4, 10, 3, 2, 15, 3800, 57000, 'HygieneFirst Chemicals', 'Laundry Bay Store', 'Active', new Date().toISOString()],
      ['ITM-1015', 'DP-BV-001', 'Cabernet Sauvignon Reserve 750ml', 'Bar & Beverages', 'Bottles', 12, 48, 18, 14, 80, 1800, 144000, 'Vintage Wine Merchants', 'Cellar Bin 12', 'Active', new Date().toISOString()]
    ];
    itemsSheet.getRange(2, 1, demoItems.length, demoItems[0].length).setValues(demoItems);
  }

  // Seed Suppliers
  const supSheet = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (supSheet && supSheet.getLastRow() <= 1) {
    const demoSuppliers = [
      ['SUP-101', 'Apex Hospitality Linens', 'Rahul Sharma', 'orders@apexlinens.com', '+91 98765 43210', 'Housekeeping & Linens', 'Plot 45, Textile Park, Mumbai', 'Active', 'Contracted vendor with 15-day Net payment terms'],
      ['SUP-102', 'Luxe Spa & Botanicals', 'Priya Menon', 'supply@luxebotanicals.com', '+91 98234 56789', 'Guest Amenities & Toiletries', 'Indl Area Phase 2, Bengaluru', 'Active', 'Eco-friendly certified amenities supplier'],
      ['SUP-103', 'Estate Harvest Roasters', 'Vikram Sen', 'b2b@estateharvest.com', '+91 97112 34567', 'Food & Beverage', 'Estate Road, Coorg', 'Active', 'Supplies freshly roasted single-origin beans weekly'],
      ['SUP-104', 'Gourmet Pantry Supplies', 'Farhan Merchant', 'sales@gourmetpantry.in', '+91 98450 11223', 'Food & Beverage', 'APMC Market Yard, Pune', 'Active', 'Imported cooking oils, spices and culinary ingredients'],
      ['SUP-105', 'TechSmart RFID Systems', 'Ananya Roy', 'support@techsmartrfid.com', '+91 99001 22334', 'Front Office & Stationery', 'Okhla Phase III, New Delhi', 'Active', 'VingCard compatible RFID cards & readers'],
      ['SUP-106', 'ElectroTech Solutions', 'Suresh Nair', 'info@electrotech.co.in', '+91 94455 66778', 'Maintenance & Engineering', 'Electronics Complex, Kochi', 'Active', 'LED, HVAC parts, sensors, and electrical spares'],
      ['SUP-107', 'HygieneFirst Chemicals', 'Meera Joshi', 'hygiene@firstchem.in', '+91 91678 99001', 'Cleaning & Chemicals', 'GIDC Estate, Ankleshwar', 'Active', 'ISO certified hotel housekeeping chemicals'],
      ['SUP-108', 'Vintage Wine Merchants', 'David Fernandez', 'cellar@vintagemerchants.com', '+91 98200 44556', 'Bar & Beverages', 'Colaba, Mumbai', 'Active', 'Licensed alcoholic and non-alcoholic beverages distributor']
    ];
    supSheet.getRange(2, 1, demoSuppliers.length, demoSuppliers[0].length).setValues(demoSuppliers);
  }

  // Seed Departments
  const deptSheet = ss.getSheetByName(SHEET_NAMES.DEPARTMENTS);
  if (deptSheet && deptSheet.getLastRow() <= 1) {
    const demoDepts = [
      ['DEP-01', 'Housekeeping', 'Both Properties', 'Ms. Sunita Patil', 'Daily guest room preparation, linen distribution, and laundry'],
      ['DEP-02', 'F&B Production (Kitchen)', 'Both Properties', 'Chef Antoine Girard', 'Main restaurants, banqueting, and in-room dining culinary operations'],
      ['DEP-03', 'F&B Service (Restaurant & Bar)', 'Both Properties', 'Mr. Amit Kapoor', 'Dining room, lounge, poolside bar service and beverage management'],
      ['DEP-04', 'Front Office & Concierge', 'Both Properties', 'Ms. Neha Gupta', 'Reception, check-in, key handling, and guest stationery'],
      ['DEP-05', 'Engineering & Maintenance', 'Both Properties', 'Mr. K. R. Nambiar', 'Property upkeep, HVAC, plumbing, lighting and mechanical systems'],
      ['DEP-06', 'Spa & Wellness Center', 'Deneb Hotel', 'Ms. Clara D’Souza', 'Ayurvedic treatments, aromatherapy, and wellness retail'],
      ['DEP-07', 'Banqueting & Events', 'Pollux Hotel', 'Mr. Rohit Mehra', 'Conference halls, wedding venues, and audio-visual stores']
    ];
    deptSheet.getRange(2, 1, demoDepts.length, demoDepts[0].length).setValues(demoDepts);
  }

  // Seed Locations
  const locSheet = ss.getSheetByName(SHEET_NAMES.LOCATIONS);
  if (locSheet && locSheet.getLastRow() <= 1) {
    const demoLocs = [
      ['LOC-CENTRAL', 'Central Warehouse', 'CENTRAL', 'Main Central Depot', 'Primary receipt hub for bulk procurement and inter-hotel distribution'],
      ['LOC-DENEB', 'Deneb Hotel Store', 'DENEB', 'Property Main Store', 'On-site hotel inventory store serving Deneb Hotel departments'],
      ['LOC-POLLUX', 'Pollux Hotel Store', 'POLLUX', 'Property Main Store', 'On-site hotel inventory store serving Pollux Hotel departments']
    ];
    locSheet.getRange(2, 1, demoLocs.length, demoLocs[0].length).setValues(demoLocs);
  }

  // Seed Initial Sample Transactions
  const txnSheet = ss.getSheetByName(SHEET_NAMES.TRANSACTIONS);
  if (txnSheet && txnSheet.getLastRow() <= 1) {
    const now = new Date();
    const demoTxns = [
      ['TXN-INIT-001', new Date(now.getTime() - 86400000 * 3).toISOString(), 'STOCK_IN', 'ITM-1001', 'DP-HK-001', 'King Size Bed Sheet 400TC (White)', 'Housekeeping & Linens', 50, 'Pcs', 'Apex Hospitality Linens', 'Central Warehouse', 1200, 60000, 'PO-2026-901', 'Inventory Mgr', 'Initial seasonal bulk replenishment'],
      ['TXN-INIT-002', new Date(now.getTime() - 86400000 * 2).toISOString(), 'TRANSFER', 'ITM-1001', 'DP-HK-001', 'King Size Bed Sheet 400TC (White)', 'Housekeeping & Linens', 25, 'Pcs', 'Central Warehouse', 'Deneb Hotel Store', 1200, 30000, 'TRF-0881', 'Logistics Coord', 'Weekly transfer to Deneb property linen room'],
      ['TXN-INIT-003', new Date(now.getTime() - 86400000 * 2).toISOString(), 'TRANSFER', 'ITM-1001', 'DP-HK-001', 'King Size Bed Sheet 400TC (White)', 'Housekeeping & Linens', 20, 'Pcs', 'Central Warehouse', 'Pollux Hotel Store', 1200, 24000, 'TRF-0882', 'Logistics Coord', 'Weekly transfer to Pollux property linen room'],
      ['TXN-INIT-004', new Date(now.getTime() - 86400000 * 1).toISOString(), 'STOCK_OUT', 'ITM-1001', 'DP-HK-001', 'King Size Bed Sheet 400TC (White)', 'Housekeeping & Linens', 5, 'Pcs', 'Deneb Hotel Store', 'Housekeeping', 1200, 6000, 'REQ-3312', 'Sunita Patil', 'Issued for Deneb 3rd floor executive suites turnover'],
      ['TXN-INIT-005', new Date(now.getTime() - 3600000 * 5).toISOString(), 'STOCK_IN', 'ITM-1003', 'DP-AM-001', 'Organic Shampoo 30ml Eco-Bottle', 'Guest Amenities & Toiletries', 300, 'Bottles', 'Luxe Spa & Botanicals', 'Central Warehouse', 25, 7500, 'PO-2026-908', 'Inventory Mgr', 'Delivery received in good condition, batch #AM88']
    ];
    txnSheet.getRange(2, 1, demoTxns.length, demoTxns[0].length).setValues(demoTxns);
  }
}

function connectCustomSpreadsheet(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') {
    return { success: false, error: 'Please provide a valid Google Sheet URL or ID.' };
  }

  let sheetId = urlOrId.trim();
  const match = sheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    sheetId = match[1];
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    initDatabase(ss);
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', sheetId);

    return {
      success: true,
      id: sheetId,
      name: ss.getName(),
      url: ss.getUrl(),
      message: 'Successfully linked Google Sheet: ' + ss.getName()
    };
  } catch (err) {
    return {
      success: false,
      error: 'Could not access the specified Google Sheet. Please verify permissions or the ID. Details: ' + err.message
    };
  }
}

function resetOrSeedDemoData() {
  const ss = getSpreadsheet();
  seedSampleDataIfEmpty(ss);
  return { success: true, message: 'Database sample items and tables refreshed.' };
}

function getDatabaseInfo() {
  const ss = getSpreadsheet();
  const folder = getDriveFolder();
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('DRIVE_FOLDER_ID') || DEFAULT_DRIVE_FOLDER_ID;

  return {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl(),
    sheetNames: ss.getSheets().map(s => s.getName()),
    folder: {
      id: folderId,
      name: folder ? folder.getName() : 'DNP Database Drive Folder',
      url: folder ? folder.getUrl() : 'https://drive.google.com/drive/folders/' + folderId
    },
    folderSheets: getFolderSheets()
  };
}

function getSettingsInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return { CURRENCY_SYMBOL: '₹', HOTEL_NAME: 'DNP HOTELS' };
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      settings[String(data[i][0])] = String(data[i][1]);
    }
  }
  return settings;
}

function getSettings() {
  const ss = getSpreadsheet();
  return getSettingsInternal(ss);
}

function saveSettings(settingsObj) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) throw new Error('Settings sheet missing');

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
      sheet.appendRow([key, String(settingsObj[key]), 'User customized setting']);
    }
  }
  return { success: true };
}
