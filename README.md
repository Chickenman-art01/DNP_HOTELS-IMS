# DENEB & POLLUX HOTELS - INVENTORY MANAGEMENT SYSTEM (IMS)

A full-stack, enterprise-grade Inventory Management System Web App tailored for **Deneb Hotel**, **Pollux Hotel**, and the **Central Distribution Warehouse**, built with **Google Apps Script**, **TypeScript**, and **Google Sheets**.

---

## 🌟 Key Features

1. **Multi-Property Stock Tracking**:
   - Distinct stock tracking across **Central Warehouse (Depot)**, **Deneb Hotel Store**, and **Pollux Hotel Store**.
   - Combined chain-wide inventory valuation and unified SKU catalog.

2. **Atomic Stock Movements**:
   - **Stock In (Receiving / Purchasing)**: Receive items from suppliers with invoice/PO references, automatically updating receiving store balances and latest unit costs.
   - **Stock Out (Department Issuance)**: Disburse items to hotel departments (Housekeeping, Kitchen, Restaurant & Bar, Front Desk, Maintenance, Spa) with real-time stock availability validation.
   - **Inter-Property Transfers**: Transfer stock between Central Warehouse, Deneb Hotel, and Pollux Hotel with waybill tracking.
   - **Physical Audit Adjustments**: Reconcile system records with physical cycle counts, logging audit reasons (cycle count, damage, expiry, shrinkage).

3. **Executive Dashboard**:
   - Real-time KPIs: Chain valuation, total physical units, active catalog SKUs, low stock warnings, and daily transaction volume.
   - Interactive **Chart.js** valuation breakdown by category.
   - One-click **Quick Restock** from priority low-stock alerts.

4. **Barcode & QR Scanner**:
   - **JsBarcode generation**: Printable barcode labels with SKU, item name, and price.
   - **Live Camera Barcode Scanner**: Scan barcodes or QR codes in real-time using mobile or laptop cameras to lookup items or execute transactions.

5. **Flexible Google Sheets Backend**:
   - **Zero-Config Setup**: Automatically provisions a structured Google Sheet (`Items`, `Transactions`, `Suppliers`, `Departments`, `Locations`, `Settings`) with realistic demo data on first launch.
   - **Custom Sheet Integration**: Connect your own Google Sheet anytime by pasting its URL or ID in the **Database & Settings** tab.
   - **Direct Google Sheet Link**: Open and inspect raw tables directly in Google Sheets anytime.

6. **Vendor & Department Directories**:
   - Comprehensive supplier management with contact details, phone, email, and supplied categories.
   - Department directory mapped to hotel property scopes.

7. **Data Portability**:
   - One-click CSV export for both the full inventory catalog and the complete transaction audit log.

---

## 🛠️ TypeScript & Clasp Developer Setup

This project uses `@types/google-apps-script` with `tsconfig.json` so you get full VS Code IntelliSense / autocomplete for all Google Apps Script classes (`SpreadsheetApp`, `PropertiesService`, `HtmlService`, `LockService`, etc.).

### Project Structure
```text
├── Code.ts            # Server-side TypeScript code (transpiled to GAS)
├── Index.html         # Main Web App UI layout & modal templates
├── Styles.html        # Custom styles & barcode label print styling
├── JavaScript.html    # Client-side reactive logic, Chart.js & scanner
├── appsscript.json    # Apps Script manifest with web app configuration
├── .clasp.json        # Clasp project bindings
├── .claspignore       # Prevents node_modules & local configs from pushing
├── tsconfig.json      # TypeScript compiler options & type definitions
└── package.json       # Node package configuration with dev dependencies
```

### Development Commands
- **Push changes to Apps Script**:
  ```bash
  clasp push
  ```
- **Deploy a new version**:
  ```bash
  clasp deploy --description "Release description"
  ```
- **Check deployment URLs**:
  ```bash
  clasp deployments
  ```
