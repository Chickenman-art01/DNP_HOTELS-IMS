/**
 * DNP HOTELS - INVENTORY MANAGEMENT SYSTEM (IMS)
 * Google Apps Script Server-Side Backend
 * 
 * Multi-property support: Deneb Hotel, Pollux Hotel, Central Warehouse
 * Handles:
 *  - Automatic Google Sheet database creation & schema validation
 *  - Linking user-supplied Google Sheet via URL or ID
 *  - Real-time atomic stock operations (Stock In, Stock Out, Transfers, Audit Adjustments)
 *  - Dashboard KPIs, valuation, and low-stock alerts
 *  - Barcode and SKU search & lookup
 *  - Supplier & Department management
 */

// ==========================================
// WEB APP ENTRY POINTS
// ==========================================

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('DNP HOTELS | Inventory Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// DATABASE & SPREADSHEET MANAGEMENT
// ==========================================

const SHEET_NAMES = {
  ITEMS: 'Items',
  TRANSACTIONS: 'Transactions',
  SUPPLIERS: 'Suppliers',
  DEPARTMENTS: 'Departments',
  LOCATIONS: 'Locations',
  SETTINGS: 'Settings'
};

/**
 * Gets or initializes the connected Google Spreadsheet.
 * If not connected or invalid, creates a new spreadsheet and stores its ID in ScriptProperties.
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

  // If still not available, create a new spreadsheet automatically
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('DNP HOTELS - IMS Database');
      sheetId = ss.getId();
      props.setProperty('SPREADSHEET_ID', sheetId);
      initDatabase(ss);
      seedSampleDataIfEmpty(ss);
    } catch (err) {
      console.error('Failed to create new spreadsheet', err);
      throw new Error('Unable to access or create Google Sheet database: ' + err.message);
    }
  } else {
    // Ensure all required sheets exist
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

  // Remove default "Sheet1" if other sheets exist
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(sheet1);
    } catch (e) {}
  }
}

/**
 * Format headers with professional styling
 */
function formatHeaderRow(sheet, colCount) {
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground('#0F172A')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * Seeds demo inventory, suppliers, departments, and initial stock
 */
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
      ['ITM-1009', 'DP-FO-001', 'RFID Keycards (Deneb-Pollux Branded)', 'Front Office & Stationery', 'Cards', 200, 600, 200, 150, 950, 45, 42750, 'TechSmart RFID Systems', 'Front Desk Safe', 'Active', new Date().toISOString()],
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

/**
 * Connects user's custom Google Sheet by URL or ID.
 * Validates, initializes schemas if missing, and saves the ID in ScriptProperties.
 */
function connectCustomSpreadsheet(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') {
    return { success: false, error: 'Please provide a valid Google Sheet URL or ID.' };
  }

  let sheetId = urlOrId.trim();
  // Extract ID if URL is provided
  const match = sheetId.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    sheetId = match[1];
  }

  try {
    const ss = SpreadsheetApp.openById(sheetId);
    // Initialize structure if needed without deleting existing data
    initDatabase(ss);
    
    // Save to ScriptProperties
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

/**
 * Re-seeds demo data if requested by user
 */
function resetOrSeedDemoData() {
  const ss = getSpreadsheet();
  seedSampleDataIfEmpty(ss);
  return { success: true, message: 'Database sample items and tables refreshed.' };
}

/**
 * Returns database metadata and direct Google Sheet URL
 */
function getDatabaseInfo() {
  const ss = getSpreadsheet();
  return {
    id: ss.getId(),
    name: ss.getName(),
    url: ss.getUrl(),
    sheetNames: ss.getSheets().map(s => s.getName())
  };
}

// ==========================================
// CORE API ENDPOINTS
// ==========================================

/**
 * Returns full initial bundle for the client app:
 * - Dashboard KPIs & stats
 * - Items catalog
 * - Suppliers list
 * - Departments list
 * - Locations list
 * - System settings & Sheet info
 */
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

  // Compute Dashboard Metrics
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

/**
 * Computes dashboard statistics from current items & transactions
 */
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

  // Today's transaction count
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

// ==========================================
// ITEMS CATALOG OPERATIONS
// ==========================================

function getItemsInternal(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.ITEMS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[2]) continue; // Skip empty rows

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

/**
 * Creates or updates an inventory item
 */
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
      // Auto-generate SKU
      const catPrefix = (itemData.category || 'GEN').substring(0, 2).toUpperCase();
      sku = 'DP-' + catPrefix + '-' + Math.floor(1000 + Math.random() * 9000);
    }

    if (targetRow > 1) {
      // Update existing item
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
      // Create new item
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

/**
 * Deletes or archives an item
 */
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

// ==========================================
// STOCK OPERATIONS & TRANSACTIONS
// ==========================================

/**
 * Records an atomic stock movement:
 * - STOCK_IN: Receiving goods from supplier into Central, Deneb, or Pollux
 * - STOCK_OUT: Issuing stock to hotel department
 * - TRANSFER: Transferring between Central, Deneb, and Pollux
 * - ADJUSTMENT: Physical audit discrepancy correction
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

    // Property stock column helper
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

    // Process movement by type
    if (type === 'STOCK_IN') {
      // Receiving inventory into destination location
      updateStockByLoc(dest, qty);
      if (txn.unitPrice && Number(txn.unitPrice) > 0) {
        unitCost = Number(txn.unitPrice); // Update latest unit cost
      }
    } else if (type === 'STOCK_OUT') {
      // Issuing inventory to hotel department
      const currentLocStock = getStockByLoc(source);
      if (currentLocStock === null) {
        throw new Error('Invalid source location for Stock Out: ' + source);
      }
      if (currentLocStock < qty) {
        throw new Error('Insufficient stock in ' + source + '. Available: ' + currentLocStock + ', Requested: ' + qty);
      }
      updateStockByLoc(source, -qty);
    } else if (type === 'TRANSFER') {
      // Inter-property transfer
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
      // Direct physical count reconciliation
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

    // Recalculate totals
    const totalStock = centralStock + denebStock + polluxStock;
    const totalValue = totalStock * unitCost;
    const nowIso = new Date().toISOString();

    // Update item row
    itemsSheet.getRange(targetRow, 7, 1, 6).setValues([[
      centralStock,
      denebStock,
      polluxStock,
      totalStock,
      unitCost,
      totalValue
    ]]);
    itemsSheet.getRange(targetRow, 16).setValue(nowIso);

    // Generate Transaction ID
    const dateCode = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd');
    const txnId = 'TXN-' + dateCode + '-' + Math.floor(1000 + Math.random() * 9000);
    const lineTotalCost = qty * unitCost;

    // Record in Transactions Sheet
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
  // Read reverse order (newest first)
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

// ==========================================
// SUPPLIERS OPERATIONS
// ==========================================

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

// ==========================================
// DEPARTMENTS & LOCATIONS
// ==========================================

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

// ==========================================
// SETTINGS
// ==========================================

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
