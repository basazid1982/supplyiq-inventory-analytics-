/*************************************************************************
 * SUPPLYIQ — Supply Chain Inventory Analytics Platform
 * Backend: Google Apps Script | Database: Google Sheets
 * Architecture: Sections 1-9. Section 9 adds the analytics/BI layer
 * (inventory turnover, ABC/Pareto classification, reorder-point
 * suggestions, and demand forecasting) on top of the core CRUD engine.
 *************************************************************************/

/* =========================================================
 * SECTION 1 — GLOBAL CONFIGURATION
 * ========================================================= */

const SHEET_NAMES = {
  USERS: 'Users',
  PRODUCTS: 'Products',
  CATEGORIES: 'Categories',
  SUPPLIERS: 'Suppliers',
  PURCHASES: 'Purchases',
  PURCHASE_ITEMS: 'PurchaseItems',
  SALES: 'Sales',
  SALE_ITEMS: 'SaleItems',
  CUSTOMERS: 'Customers',
  INVENTORY_LOG: 'InventoryLog',
  SETTINGS: 'Settings',
  AUDIT_LOG: 'AuditLog'
};

const SHEET_SCHEMAS = {
  Users: ['UserID', 'FullName', 'Email', 'Password', 'Role', 'Phone', 'Status', 'CreatedAt', 'LastLogin'],
  Categories: ['CategoryID', 'CategoryName', 'Description', 'CreatedAt'],
  Products: ['ProductID', 'SKU', 'Barcode', 'ProductName', 'Category', 'Brand', 'Unit', 'CostPrice', 'SellingPrice',
    'CurrentStock', 'MinimumStock', 'MaximumStock', 'SupplierID', 'Location', 'Description', 'ImageFileID', 'Status', 'CreatedAt', 'UpdatedAt'],
  Suppliers: ['SupplierID', 'CompanyName', 'ContactPerson', 'Phone', 'Email', 'Address', 'City', 'Country', 'TaxNumber', 'Status', 'CreatedAt'],
  Customers: ['CustomerID', 'CustomerName', 'Phone', 'Email', 'Address', 'City', 'Country', 'Status', 'CreatedAt'],
  Purchases: ['PurchaseID', 'InvoiceNumber', 'SupplierID', 'PurchaseDate', 'Subtotal', 'Tax', 'Discount', 'GrandTotal', 'PaymentStatus', 'Remarks', 'CreatedBy', 'CreatedAt'],
  PurchaseItems: ['PurchaseItemID', 'PurchaseID', 'ProductID', 'Quantity', 'UnitCost', 'Total'],
  Sales: ['SaleID', 'InvoiceNumber', 'CustomerID', 'SaleDate', 'Subtotal', 'Tax', 'Discount', 'GrandTotal', 'PaymentMethod', 'PaymentStatus', 'CreatedBy', 'CreatedAt'],
  SaleItems: ['SaleItemID', 'SaleID', 'ProductID', 'Quantity', 'UnitPrice', 'Discount', 'Total'],
  InventoryLog: ['LogID', 'ProductID', 'TransactionType', 'ReferenceID', 'QuantityIn', 'QuantityOut', 'Balance', 'CreatedBy', 'CreatedAt'],
  Settings: ['Key', 'Value'],
  AuditLog: ['LogID', 'UserEmail', 'Action', 'Details', 'Timestamp']
};

const ID_PREFIXES = {
  Users: 'USR', Categories: 'CAT', Products: 'PRD', Suppliers: 'SUP',
  Customers: 'CUS', Purchases: 'PUR', PurchaseItems: 'PIT',
  Sales: 'SAL', SaleItems: 'SIT', InventoryLog: 'LOG', AuditLog: 'ALG'
};

const DEFAULT_SETTINGS = {
  BusinessName: 'My Business',
  BusinessLogo: '',
  Phone: '',
  Email: '',
  Address: '',
  Currency: 'CAD',
  Timezone: 'America/Winnipeg',
  LowStockLimit: '10',
  LeadTimeDays: '7'
};

/* =========================================================
 * SECTION 2 — UTILITY FUNCTIONS
 * ========================================================= */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}

// Response builder — every backend function returns this shape.
function ok_(message, data) {
  return { success: true, message: message || 'OK', data: (data === undefined ? {} : data) };
}
function fail_(message) {
  return { success: false, message: message || 'Something went wrong.', data: {} };
}

// Wrap any function so it never throws raw errors to the client.
function guard_(fn) {
  try {
    return fn();
  } catch (e) {
    logError_(e);
    return fail_(e && e.message ? e.message : String(e));
  }
}

function logError_(e) {
  try {
    console.error(e && e.stack ? e.stack : e);
  } catch (err) { /* no-op */ }
}

// Sequential, collision-free ID generator using sheet row counts + prefix + timestamp fragment.
function generateId_(sheetName) {
  const prefix = ID_PREFIXES[sheetName] || 'ID';
  const sh = sheet_(sheetName);
  const lastRow = sh.getLastRow();
  const seq = Math.max(0, lastRow - 1) + 1;
  return prefix + Utilities.formatString('%04d', seq) + '-' + Utilities.getUuid().substring(0, 4).toUpperCase();
}

function nowIso_() {
  return new Date().toISOString();
}

function formatDate_(d) {
  if (!d) return '';
  const date = (d instanceof Date) ? d : new Date(d);
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'America/Winnipeg', 'yyyy-MM-dd HH:mm:ss');
}

function toNumber_(v, fallback) {
  const n = parseFloat(v);
  return isNaN(n) ? (fallback || 0) : n;
}

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || ''));
}

function hashPassword_(plain) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(plain));
  return digest.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

// Convert a sheet's data range into an array of row objects keyed by header.
function readTable_(sheetName) {
  const sh = sheet_(sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  return values.map(function (row, idx) {
    const obj = { _row: idx + 2 };
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    return obj;
  });
}

function appendRow_(sheetName, rowObj) {
  const sh = sheet_(sheetName);
  const headers = SHEET_SCHEMAS[sheetName];
  const row = headers.map(function (h) { return (rowObj[h] === undefined || rowObj[h] === null) ? '' : rowObj[h]; });
  sh.appendRow(row);
  return row;
}

function updateRowById_(sheetName, idField, idValue, patch) {
  const sh = sheet_(sheetName);
  const headers = SHEET_SCHEMAS[sheetName];
  const idCol = headers.indexOf(idField) + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idValue)) {
      const rowNum = i + 2;
      headers.forEach(function (h, colIdx) {
        if (patch.hasOwnProperty(h)) {
          sh.getRange(rowNum, colIdx + 1).setValue(patch[h]);
        }
      });
      return true;
    }
  }
  return false;
}

function deleteRowById_(sheetName, idField, idValue) {
  const sh = sheet_(sheetName);
  const headers = SHEET_SCHEMAS[sheetName];
  const idCol = headers.indexOf(idField) + 1;
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;
  const ids = sh.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idValue)) {
      sh.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function findRowById_(sheetName, idField, idValue) {
  const rows = readTable_(sheetName);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][idField]) === String(idValue)) return rows[i];
  }
  return null;
}

function auditLog_(action, details) {
  try {
    const email = getSessionEmail_();
    appendRow_(SHEET_NAMES.AUDIT_LOG, {
      LogID: generateId_(SHEET_NAMES.AUDIT_LOG),
      UserEmail: email || 'system',
      Action: action,
      Details: JSON.stringify(details || {}),
      Timestamp: nowIso_()
    });
  } catch (e) { logError_(e); }
}

/* =========================================================
 * SECTION 3 — AUTHENTICATION
 * ========================================================= */

// Simple session using PropertiesService keyed by a random token returned to the client.
function getSessionEmail_() {
  const cache = CacheService.getUserCache();
  return cache.get('ims_session_email');
}

function login(email, password) {
  return guard_(function () {
    if (!email || !password) return fail_('Email and password are required.');
    const users = readTable_(SHEET_NAMES.USERS);
    const user = users.find(function (u) { return String(u.Email).toLowerCase() === String(email).toLowerCase(); });
    if (!user) return fail_('Invalid email or password.');
    if (String(user.Status).toLowerCase() !== 'active') return fail_('This account is inactive. Contact an administrator.');
    const hashed = hashPassword_(password);
    if (String(user.Password) !== hashed) return fail_('Invalid email or password.');

    const token = Utilities.getUuid();
    const cache = CacheService.getUserCache();
    cache.put('ims_session_email', user.Email, 21600); // 6 hours
    cache.put('ims_session_token_' + token, user.Email, 21600);

    updateRowById_(SHEET_NAMES.USERS, 'UserID', user.UserID, { LastLogin: nowIso_() });
    auditLog_('Login', { email: user.Email });

    return ok_('Login successful.', {
      token: token,
      user: { userId: user.UserID, fullName: user.FullName, email: user.Email, role: user.Role }
    });
  });
}

function logout(token) {
  return guard_(function () {
    const cache = CacheService.getUserCache();
    if (token) cache.remove('ims_session_token_' + token);
    cache.remove('ims_session_email');
    auditLog_('Logout', {});
    return ok_('Logged out.');
  });
}

function checkSession(token) {
  return guard_(function () {
    if (!token) return fail_('No session.');
    const cache = CacheService.getUserCache();
    const email = cache.get('ims_session_token_' + token);
    if (!email) return fail_('Session expired.');
    cache.put('ims_session_email', email, 21600);
    const user = readTable_(SHEET_NAMES.USERS).find(function (u) { return u.Email === email; });
    if (!user) return fail_('User not found.');
    return ok_('Session valid.', { userId: user.UserID, fullName: user.FullName, email: user.Email, role: user.Role });
  });
}

function getCurrentUser() {
  return guard_(function () {
    const email = getSessionEmail_();
    if (!email) return fail_('Not logged in.');
    const user = readTable_(SHEET_NAMES.USERS).find(function (u) { return u.Email === email; });
    if (!user) return fail_('User not found.');
    return ok_('OK', { userId: user.UserID, fullName: user.FullName, email: user.Email, role: user.Role });
  });
}

/* =========================================================
 * SECTION 4 — DATABASE (SHEET SETUP)
 * ========================================================= */

function setupSheets() {
  return guard_(function () {
    const spreadsheet = ss_();
    Object.keys(SHEET_SCHEMAS).forEach(function (name) {
      let sh = spreadsheet.getSheetByName(name);
      if (!sh) {
        sh = spreadsheet.insertSheet(name);
      }
      const headers = SHEET_SCHEMAS[name];
      const existingHeaderRange = sh.getRange(1, 1, 1, headers.length);
      const existingHeaders = existingHeaderRange.getValues()[0];
      const needsHeaders = headers.some(function (h, i) { return existingHeaders[i] !== h; });
      if (needsHeaders) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        sh.setFrozenRows(1);
        sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#2563EB').setFontColor('#FFFFFF');
      }
    });
    // Remove the default 'Sheet1' if it's empty and unused.
    const def = spreadsheet.getSheetByName('Sheet1');
    if (def && def.getLastRow() === 0 && spreadsheet.getSheets().length > 1) {
      spreadsheet.deleteSheet(def);
    }

    // Seed default settings row if empty.
    const settingsSheet = sheet_(SHEET_NAMES.SETTINGS);
    if (settingsSheet.getLastRow() < 2) {
      Object.keys(DEFAULT_SETTINGS).forEach(function (key) {
        settingsSheet.appendRow([key, DEFAULT_SETTINGS[key]]);
      });
    }

    // Seed default admin + categories if Users/Categories empty.
    const usersSheet = sheet_(SHEET_NAMES.USERS);
    if (usersSheet.getLastRow() < 2) {
      appendRow_(SHEET_NAMES.USERS, {
        UserID: 'USR0001',
        FullName: 'Admin',
        Email: 'admin@example.com',
        Password: hashPassword_('admin123'),
        Role: 'Admin',
        Phone: '',
        Status: 'Active',
        CreatedAt: nowIso_(),
        LastLogin: ''
      });
    }
    const catSheet = sheet_(SHEET_NAMES.CATEGORIES);
    if (catSheet.getLastRow() < 2) {
      ['Electrical', 'Mechanical', 'Safety'].forEach(function (name, i) {
        appendRow_(SHEET_NAMES.CATEGORIES, {
          CategoryID: 'CAT' + Utilities.formatString('%04d', i + 1),
          CategoryName: name,
          Description: name + ' products',
          CreatedAt: nowIso_()
        });
      });
    }

    return ok_('Sheets are set up.');
  });
}

/* =========================================================
 * SAMPLE DATA SEEDING
 * ========================================================= */

function seedSampleData() {
  return guard_(function () {
    setupSheets();

    const suppliersSheet = sheet_(SHEET_NAMES.SUPPLIERS);
    if (suppliersSheet.getLastRow() < 2) {
      for (let i = 1; i <= 10; i++) {
        appendRow_(SHEET_NAMES.SUPPLIERS, {
          SupplierID: 'SUP' + Utilities.formatString('%04d', i),
          CompanyName: 'Supplier Co ' + i,
          ContactPerson: 'Contact ' + i,
          Phone: '204-555-' + (1000 + i),
          Email: 'supplier' + i + '@example.com',
          Address: i + ' Main St',
          City: 'Winnipeg',
          Country: 'Canada',
          TaxNumber: 'TAX' + (10000 + i),
          Status: 'Active',
          CreatedAt: nowIso_()
        });
      }
    }

    const customersSheet = sheet_(SHEET_NAMES.CUSTOMERS);
    if (customersSheet.getLastRow() < 2) {
      for (let i = 1; i <= 15; i++) {
        appendRow_(SHEET_NAMES.CUSTOMERS, {
          CustomerID: 'CUS' + Utilities.formatString('%04d', i),
          CustomerName: 'Customer ' + i,
          Phone: '204-555-' + (2000 + i),
          Email: 'customer' + i + '@example.com',
          Address: i + ' Elm St',
          City: 'Winnipeg',
          Country: 'Canada',
          Status: 'Active',
          CreatedAt: nowIso_()
        });
      }
    }

    const categories = readTable_(SHEET_NAMES.CATEGORIES).map(function (c) { return c.CategoryName; });
    const suppliers = readTable_(SHEET_NAMES.SUPPLIERS).map(function (s) { return s.SupplierID; });

    const productsSheet = sheet_(SHEET_NAMES.PRODUCTS);
    if (productsSheet.getLastRow() < 2) {
      for (let i = 1; i <= 20; i++) {
        const cost = 10 + i * 2;
        appendRow_(SHEET_NAMES.PRODUCTS, {
          ProductID: 'PRD' + Utilities.formatString('%04d', i),
          SKU: 'SKU-' + (1000 + i),
          Barcode: '89' + (100000000 + i),
          ProductName: 'Product ' + i,
          Category: categories[i % categories.length],
          Brand: 'Brand ' + ((i % 5) + 1),
          Unit: 'pcs',
          CostPrice: cost,
          SellingPrice: Math.round(cost * 1.4 * 100) / 100,
          CurrentStock: 50,
          MinimumStock: 10,
          MaximumStock: 200,
          SupplierID: suppliers[i % suppliers.length],
          Location: 'Aisle ' + ((i % 4) + 1),
          Description: 'Sample product ' + i,
          ImageFileID: '',
          Status: 'Active',
          CreatedAt: nowIso_(),
          UpdatedAt: nowIso_()
        });
      }
    }

    // Sample purchases (each increases stock + logs inventory).
    const purchasesSheet = sheet_(SHEET_NAMES.PURCHASES);
    if (purchasesSheet.getLastRow() < 2) {
      const products = readTable_(SHEET_NAMES.PRODUCTS);
      for (let i = 1; i <= 10; i++) {
        const supplierId = suppliers[i % suppliers.length];
        const product = products[i % products.length];
        const qty = 10;
        const unitCost = toNumber_(product.CostPrice);
        const subtotal = qty * unitCost;
        const tax = Math.round(subtotal * 0.05 * 100) / 100;
        const grand = Math.round((subtotal + tax) * 100) / 100;
        const purchaseId = 'PUR' + Utilities.formatString('%04d', i);
        appendRow_(SHEET_NAMES.PURCHASES, {
          PurchaseID: purchaseId,
          InvoiceNumber: 'PINV-' + (1000 + i),
          SupplierID: supplierId,
          PurchaseDate: nowIso_(),
          Subtotal: subtotal, Tax: tax, Discount: 0, GrandTotal: grand,
          PaymentStatus: 'Paid', Remarks: 'Sample purchase',
          CreatedBy: 'admin@example.com', CreatedAt: nowIso_()
        });
        appendRow_(SHEET_NAMES.PURCHASE_ITEMS, {
          PurchaseItemID: 'PIT' + Utilities.formatString('%04d', i),
          PurchaseID: purchaseId, ProductID: product.ProductID,
          Quantity: qty, UnitCost: unitCost, Total: subtotal
        });
      }
    }

    // Sample sales (each decreases stock + logs inventory).
    const salesSheet = sheet_(SHEET_NAMES.SALES);
    if (salesSheet.getLastRow() < 2) {
      const products = readTable_(SHEET_NAMES.PRODUCTS);
      const customers = readTable_(SHEET_NAMES.CUSTOMERS).map(function (c) { return c.CustomerID; });
      for (let i = 1; i <= 10; i++) {
        const customerId = customers[i % customers.length];
        const product = products[(i + 3) % products.length];
        const qty = 3;
        const unitPrice = toNumber_(product.SellingPrice);
        const subtotal = qty * unitPrice;
        const tax = Math.round(subtotal * 0.05 * 100) / 100;
        const grand = Math.round((subtotal + tax) * 100) / 100;
        const saleId = 'SAL' + Utilities.formatString('%04d', i);
        appendRow_(SHEET_NAMES.SALES, {
          SaleID: saleId, InvoiceNumber: 'SINV-' + (1000 + i), CustomerID: customerId,
          SaleDate: nowIso_(), Subtotal: subtotal, Tax: tax, Discount: 0, GrandTotal: grand,
          PaymentMethod: 'Cash', PaymentStatus: 'Paid',
          CreatedBy: 'admin@example.com', CreatedAt: nowIso_()
        });
        appendRow_(SHEET_NAMES.SALE_ITEMS, {
          SaleItemID: 'SIT' + Utilities.formatString('%04d', i),
          SaleID: saleId, ProductID: product.ProductID, Quantity: qty,
          UnitPrice: unitPrice, Discount: 0, Total: subtotal
        });

        // Reflect the sample sale in stock + inventory log.
        const current = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', product.ProductID);
        const newStock = Math.max(0, toNumber_(current.CurrentStock) - qty);
        updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', product.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
        appendRow_(SHEET_NAMES.INVENTORY_LOG, {
          LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: product.ProductID,
          TransactionType: 'Sale', ReferenceID: saleId, QuantityIn: 0, QuantityOut: qty,
          Balance: newStock, CreatedBy: 'admin@example.com', CreatedAt: nowIso_()
        });
      }
      // Also log the purchases' stock-in movements now that products exist.
      const purchases = readTable_(SHEET_NAMES.PURCHASES);
      const items = readTable_(SHEET_NAMES.PURCHASE_ITEMS);
      items.forEach(function (item) {
        const prod = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', item.ProductID);
        if (!prod) return;
        const newStock = toNumber_(prod.CurrentStock) + toNumber_(item.Quantity);
        updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', item.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
        appendRow_(SHEET_NAMES.INVENTORY_LOG, {
          LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: item.ProductID,
          TransactionType: 'Purchase', ReferenceID: item.PurchaseID, QuantityIn: item.Quantity, QuantityOut: 0,
          Balance: newStock, CreatedBy: 'admin@example.com', CreatedAt: nowIso_()
        });
      });
    }

    return ok_('Sample data seeded.');
  });
}

/* =========================================================
 * SECTION 5 — BUSINESS LOGIC
 * ========================================================= */

/* ---------- CATEGORIES ---------- */

function getCategories() {
  return guard_(function () { return ok_('OK', readTable_(SHEET_NAMES.CATEGORIES)); });
}

function saveCategory(payload) {
  return guard_(function () {
    if (!payload || !payload.CategoryName) return fail_('Category name is required.');
    const dup = readTable_(SHEET_NAMES.CATEGORIES).find(function (c) {
      return String(c.CategoryName).toLowerCase() === String(payload.CategoryName).toLowerCase();
    });
    if (dup) return fail_('A category with this name already exists.');
    const id = generateId_(SHEET_NAMES.CATEGORIES);
    appendRow_(SHEET_NAMES.CATEGORIES, {
      CategoryID: id, CategoryName: payload.CategoryName,
      Description: payload.Description || '', CreatedAt: nowIso_()
    });
    auditLog_('Category Added', { id: id });
    return ok_('Category created.', { CategoryID: id });
  });
}

function updateCategory(payload) {
  return guard_(function () {
    if (!payload || !payload.CategoryID) return fail_('CategoryID is required.');
    const dup = readTable_(SHEET_NAMES.CATEGORIES).find(function (c) {
      return String(c.CategoryName).toLowerCase() === String(payload.CategoryName).toLowerCase() && c.CategoryID !== payload.CategoryID;
    });
    if (dup) return fail_('A category with this name already exists.');
    const okUpdated = updateRowById_(SHEET_NAMES.CATEGORIES, 'CategoryID', payload.CategoryID, {
      CategoryName: payload.CategoryName, Description: payload.Description || ''
    });
    if (!okUpdated) return fail_('Category not found.');
    auditLog_('Category Updated', { id: payload.CategoryID });
    return ok_('Category updated.');
  });
}

function deleteCategory(categoryId) {
  return guard_(function () {
    const linked = readTable_(SHEET_NAMES.PRODUCTS).some(function (p) { return p.Category === categoryId || p.Category === (findRowById_(SHEET_NAMES.CATEGORIES, 'CategoryID', categoryId) || {}).CategoryName; });
    if (linked) return fail_('Cannot delete: products are linked to this category.');
    const okDeleted = deleteRowById_(SHEET_NAMES.CATEGORIES, 'CategoryID', categoryId);
    if (!okDeleted) return fail_('Category not found.');
    auditLog_('Category Deleted', { id: categoryId });
    return ok_('Category deleted.');
  });
}

/* ---------- SUPPLIERS ---------- */

function getSuppliers() {
  return guard_(function () { return ok_('OK', readTable_(SHEET_NAMES.SUPPLIERS)); });
}

function saveSupplier(payload) {
  return guard_(function () {
    if (!payload || !payload.CompanyName) return fail_('Company name is required.');
    if (payload.Email && !isValidEmail_(payload.Email)) return fail_('Invalid email address.');
    if (payload.Email) {
      const dup = readTable_(SHEET_NAMES.SUPPLIERS).find(function (s) { return String(s.Email).toLowerCase() === String(payload.Email).toLowerCase(); });
      if (dup) return fail_('A supplier with this email already exists.');
    }
    const id = generateId_(SHEET_NAMES.SUPPLIERS);
    appendRow_(SHEET_NAMES.SUPPLIERS, {
      SupplierID: id, CompanyName: payload.CompanyName, ContactPerson: payload.ContactPerson || '',
      Phone: payload.Phone || '', Email: payload.Email || '', Address: payload.Address || '',
      City: payload.City || '', Country: payload.Country || '', TaxNumber: payload.TaxNumber || '',
      Status: payload.Status || 'Active', CreatedAt: nowIso_()
    });
    auditLog_('Supplier Added', { id: id });
    return ok_('Supplier created.', { SupplierID: id });
  });
}

function updateSupplier(payload) {
  return guard_(function () {
    if (!payload || !payload.SupplierID) return fail_('SupplierID is required.');
    if (payload.Email && !isValidEmail_(payload.Email)) return fail_('Invalid email address.');
    const okUpdated = updateRowById_(SHEET_NAMES.SUPPLIERS, 'SupplierID', payload.SupplierID, payload);
    if (!okUpdated) return fail_('Supplier not found.');
    auditLog_('Supplier Updated', { id: payload.SupplierID });
    return ok_('Supplier updated.');
  });
}

function deleteSupplier(supplierId) {
  return guard_(function () {
    const linked = readTable_(SHEET_NAMES.PURCHASES).some(function (p) { return p.SupplierID === supplierId; });
    if (linked) return fail_('Cannot delete: supplier has linked purchases.');
    const okDeleted = deleteRowById_(SHEET_NAMES.SUPPLIERS, 'SupplierID', supplierId);
    if (!okDeleted) return fail_('Supplier not found.');
    auditLog_('Supplier Deleted', { id: supplierId });
    return ok_('Supplier deleted.');
  });
}

/* ---------- CUSTOMERS ---------- */

function getCustomers() {
  return guard_(function () { return ok_('OK', readTable_(SHEET_NAMES.CUSTOMERS)); });
}

function saveCustomer(payload) {
  return guard_(function () {
    if (!payload || !payload.CustomerName) return fail_('Customer name is required.');
    if (payload.Email && !isValidEmail_(payload.Email)) return fail_('Invalid email address.');
    if (payload.Email) {
      const dup = readTable_(SHEET_NAMES.CUSTOMERS).find(function (c) { return String(c.Email).toLowerCase() === String(payload.Email).toLowerCase(); });
      if (dup) return fail_('A customer with this email already exists.');
    }
    const id = generateId_(SHEET_NAMES.CUSTOMERS);
    appendRow_(SHEET_NAMES.CUSTOMERS, {
      CustomerID: id, CustomerName: payload.CustomerName, Phone: payload.Phone || '',
      Email: payload.Email || '', Address: payload.Address || '', City: payload.City || '',
      Country: payload.Country || '', Status: payload.Status || 'Active', CreatedAt: nowIso_()
    });
    auditLog_('Customer Added', { id: id });
    return ok_('Customer created.', { CustomerID: id });
  });
}

function updateCustomer(payload) {
  return guard_(function () {
    if (!payload || !payload.CustomerID) return fail_('CustomerID is required.');
    if (payload.Email && !isValidEmail_(payload.Email)) return fail_('Invalid email address.');
    const okUpdated = updateRowById_(SHEET_NAMES.CUSTOMERS, 'CustomerID', payload.CustomerID, payload);
    if (!okUpdated) return fail_('Customer not found.');
    auditLog_('Customer Updated', { id: payload.CustomerID });
    return ok_('Customer updated.');
  });
}

function deleteCustomer(customerId) {
  return guard_(function () {
    const okDeleted = deleteRowById_(SHEET_NAMES.CUSTOMERS, 'CustomerID', customerId);
    if (!okDeleted) return fail_('Customer not found.');
    auditLog_('Customer Deleted', { id: customerId });
    return ok_('Customer deleted.');
  });
}

/* ---------- PRODUCTS ---------- */

function getProducts() {
  return guard_(function () { return ok_('OK', readTable_(SHEET_NAMES.PRODUCTS)); });
}

function getProductById(productId) {
  return guard_(function () {
    const p = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', productId);
    if (!p) return fail_('Product not found.');
    return ok_('OK', p);
  });
}

function validateProduct_(payload, ignoreId) {
  if (!payload.ProductName) return 'Product name is required.';
  if (!payload.Category) return 'Category is required.';
  const cost = toNumber_(payload.CostPrice);
  const selling = toNumber_(payload.SellingPrice);
  if (cost < 0) return 'Cost price cannot be negative.';
  if (selling < 0) return 'Selling price cannot be negative.';
  const minStock = toNumber_(payload.MinimumStock);
  const maxStock = toNumber_(payload.MaximumStock);
  if (maxStock > 0 && minStock > maxStock) return 'Minimum stock cannot exceed maximum stock.';

  const products = readTable_(SHEET_NAMES.PRODUCTS);
  const dupName = products.find(function (p) {
    return p.ProductID !== ignoreId && String(p.ProductName).toLowerCase() === String(payload.ProductName).toLowerCase()
      && String(p.Category).toLowerCase() === String(payload.Category).toLowerCase();
  });
  if (dupName) return 'A product with this name already exists in this category.';

  if (payload.SKU) {
    const dupSku = products.find(function (p) { return p.ProductID !== ignoreId && String(p.SKU).toLowerCase() === String(payload.SKU).toLowerCase(); });
    if (dupSku) return 'This SKU is already in use.';
  }
  if (payload.Barcode) {
    const dupBarcode = products.find(function (p) { return p.ProductID !== ignoreId && String(p.Barcode) === String(payload.Barcode); });
    if (dupBarcode) return 'This barcode is already in use.';
  }
  return null;
}

function generateSku_(productName) {
  const clean = String(productName || 'ITEM').toUpperCase().replace(/[^A-Z0-9]/g, '').substring(0, 5) || 'ITEM';
  return clean + '-' + Math.floor(1000 + Math.random() * 9000);
}

function generateBarcode_() {
  return '2' + Math.floor(100000000000 + Math.random() * 899999999999);
}

function saveProduct(payload) {
  return guard_(function () {
    if (!payload) return fail_('Product data is required.');
    if (!payload.SKU) payload.SKU = generateSku_(payload.ProductName);
    if (!payload.Barcode) payload.Barcode = generateBarcode_();
    const err = validateProduct_(payload, null);
    if (err) return fail_(err);

    const id = generateId_(SHEET_NAMES.PRODUCTS);
    appendRow_(SHEET_NAMES.PRODUCTS, {
      ProductID: id, SKU: payload.SKU, Barcode: payload.Barcode, ProductName: payload.ProductName,
      Category: payload.Category, Brand: payload.Brand || '', Unit: payload.Unit || 'pcs',
      CostPrice: toNumber_(payload.CostPrice), SellingPrice: toNumber_(payload.SellingPrice),
      CurrentStock: toNumber_(payload.CurrentStock), MinimumStock: toNumber_(payload.MinimumStock),
      MaximumStock: toNumber_(payload.MaximumStock), SupplierID: payload.SupplierID || '',
      Location: payload.Location || '', Description: payload.Description || '',
      ImageFileID: payload.ImageFileID || '', Status: payload.Status || 'Active',
      CreatedAt: nowIso_(), UpdatedAt: nowIso_()
    });
    if (toNumber_(payload.CurrentStock) > 0) {
      appendRow_(SHEET_NAMES.INVENTORY_LOG, {
        LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: id, TransactionType: 'Initial Stock',
        ReferenceID: id, QuantityIn: toNumber_(payload.CurrentStock), QuantityOut: 0,
        Balance: toNumber_(payload.CurrentStock), CreatedBy: getSessionEmail_() || 'system', CreatedAt: nowIso_()
      });
    }
    auditLog_('Product Added', { id: id });
    return ok_('Product created.', { ProductID: id });
  });
}

function updateProduct(payload) {
  return guard_(function () {
    if (!payload || !payload.ProductID) return fail_('ProductID is required.');
    const err = validateProduct_(payload, payload.ProductID);
    if (err) return fail_(err);
    payload.UpdatedAt = nowIso_();
    const clean = {};
    SHEET_SCHEMAS.Products.forEach(function (h) { if (payload[h] !== undefined) clean[h] = payload[h]; });
    delete clean.CurrentStock; // stock only changes via purchases/sales/adjustments
    const okUpdated = updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', payload.ProductID, clean);
    if (!okUpdated) return fail_('Product not found.');
    auditLog_('Product Updated', { id: payload.ProductID });
    return ok_('Product updated.');
  });
}

function deleteProduct(productId) {
  return guard_(function () {
    const linkedPurchase = readTable_(SHEET_NAMES.PURCHASE_ITEMS).some(function (i) { return i.ProductID === productId; });
    const linkedSale = readTable_(SHEET_NAMES.SALE_ITEMS).some(function (i) { return i.ProductID === productId; });
    if (linkedPurchase || linkedSale) return fail_('Cannot delete: product has transaction history. Consider marking it Inactive instead.');
    const okDeleted = deleteRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', productId);
    if (!okDeleted) return fail_('Product not found.');
    auditLog_('Product Deleted', { id: productId });
    return ok_('Product deleted.');
  });
}

function getInventoryLogRows() {
  return readTable_(SHEET_NAMES.INVENTORY_LOG);
}

function lowStock() {
  return guard_(function () {
    const products = readTable_(SHEET_NAMES.PRODUCTS);
    const low = products.filter(function (p) { return toNumber_(p.CurrentStock) <= toNumber_(p.MinimumStock) && toNumber_(p.CurrentStock) > 0; });
    const out = products.filter(function (p) { return toNumber_(p.CurrentStock) <= 0; });
    return ok_('OK', { lowStock: low, outOfStock: out });
  });
}

/* ---------- PURCHASES ---------- */

function nextInvoiceNumber_(prefix, sheetName, field) {
  const rows = readTable_(sheetName);
  return prefix + '-' + (1000 + rows.length + 1);
}

function createPurchase(payload) {
  return guard_(function () {
    if (!payload || !payload.SupplierID) return fail_('Supplier is required.');
    if (!payload.items || !payload.items.length) return fail_('At least one product line is required.');

    let subtotal = 0;
    payload.items.forEach(function (it) {
      const qty = toNumber_(it.Quantity);
      const cost = toNumber_(it.UnitCost);
      if (qty <= 0) throw new Error('Quantity must be greater than zero for every line item.');
      if (cost < 0) throw new Error('Unit cost cannot be negative.');
      subtotal += qty * cost;
    });
    const tax = toNumber_(payload.Tax);
    const discount = toNumber_(payload.Discount);
    const grandTotal = Math.max(0, subtotal + tax - discount);

    const purchaseId = generateId_(SHEET_NAMES.PURCHASES);
    const invoiceNumber = payload.InvoiceNumber || nextInvoiceNumber_('PINV', SHEET_NAMES.PURCHASES);
    const who = getSessionEmail_() || 'system';

    appendRow_(SHEET_NAMES.PURCHASES, {
      PurchaseID: purchaseId, InvoiceNumber: invoiceNumber, SupplierID: payload.SupplierID,
      PurchaseDate: payload.PurchaseDate || nowIso_(), Subtotal: subtotal, Tax: tax, Discount: discount,
      GrandTotal: grandTotal, PaymentStatus: payload.PaymentStatus || 'Pending', Remarks: payload.Remarks || '',
      CreatedBy: who, CreatedAt: nowIso_()
    });

    payload.items.forEach(function (it) {
      const qty = toNumber_(it.Quantity);
      const cost = toNumber_(it.UnitCost);
      const lineTotal = qty * cost;
      const pitId = generateId_(SHEET_NAMES.PURCHASE_ITEMS);
      appendRow_(SHEET_NAMES.PURCHASE_ITEMS, {
        PurchaseItemID: pitId, PurchaseID: purchaseId, ProductID: it.ProductID, Quantity: qty, UnitCost: cost, Total: lineTotal
      });
      const product = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID);
      if (!product) throw new Error('Product not found: ' + it.ProductID);
      const newStock = toNumber_(product.CurrentStock) + qty;
      updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
      appendRow_(SHEET_NAMES.INVENTORY_LOG, {
        LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: it.ProductID, TransactionType: 'Purchase',
        ReferenceID: purchaseId, QuantityIn: qty, QuantityOut: 0, Balance: newStock, CreatedBy: who, CreatedAt: nowIso_()
      });
    });

    auditLog_('Purchase Created', { id: purchaseId });
    return ok_('Purchase recorded.', { PurchaseID: purchaseId, InvoiceNumber: invoiceNumber, GrandTotal: grandTotal });
  });
}

function getPurchases() {
  return guard_(function () {
    const purchases = readTable_(SHEET_NAMES.PURCHASES);
    const items = readTable_(SHEET_NAMES.PURCHASE_ITEMS);
    purchases.forEach(function (p) { p.items = items.filter(function (i) { return i.PurchaseID === p.PurchaseID; }); });
    return ok_('OK', purchases);
  });
}

function deletePurchase(purchaseId) {
  return guard_(function () {
    const items = readTable_(SHEET_NAMES.PURCHASE_ITEMS).filter(function (i) { return i.PurchaseID === purchaseId; });
    const who = getSessionEmail_() || 'system';
    items.forEach(function (it) {
      const product = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID);
      if (product) {
        const newStock = Math.max(0, toNumber_(product.CurrentStock) - toNumber_(it.Quantity));
        updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
        appendRow_(SHEET_NAMES.INVENTORY_LOG, {
          LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: it.ProductID, TransactionType: 'Purchase Reversal',
          ReferenceID: purchaseId, QuantityIn: 0, QuantityOut: it.Quantity, Balance: newStock, CreatedBy: who, CreatedAt: nowIso_()
        });
      }
      deleteRowById_(SHEET_NAMES.PURCHASE_ITEMS, 'PurchaseItemID', it.PurchaseItemID);
    });
    const okDeleted = deleteRowById_(SHEET_NAMES.PURCHASES, 'PurchaseID', purchaseId);
    if (!okDeleted) return fail_('Purchase not found.');
    auditLog_('Purchase Deleted', { id: purchaseId });
    return ok_('Purchase deleted and stock adjusted.');
  });
}

/* ---------- SALES ---------- */

function createSale(payload) {
  return guard_(function () {
    if (!payload || !payload.items || !payload.items.length) return fail_('At least one product line is required.');

    // Validate stock availability up front.
    for (let i = 0; i < payload.items.length; i++) {
      const it = payload.items[i];
      const qty = toNumber_(it.Quantity);
      if (qty <= 0) return fail_('Quantity must be greater than zero for every line item.');
      const product = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID);
      if (!product) return fail_('Product not found: ' + it.ProductID);
      if (toNumber_(product.CurrentStock) < qty) return fail_('Insufficient stock for ' + product.ProductName + '. Available: ' + product.CurrentStock);
    }

    let subtotal = 0;
    payload.items.forEach(function (it) {
      const qty = toNumber_(it.Quantity);
      const price = toNumber_(it.UnitPrice);
      const lineDiscount = toNumber_(it.Discount);
      subtotal += (qty * price) - lineDiscount;
    });
    const tax = toNumber_(payload.Tax);
    const discount = toNumber_(payload.Discount);
    const grandTotal = Math.max(0, subtotal + tax - discount);

    const saleId = generateId_(SHEET_NAMES.SALES);
    const invoiceNumber = payload.InvoiceNumber || nextInvoiceNumber_('SINV', SHEET_NAMES.SALES);
    const who = getSessionEmail_() || 'system';

    appendRow_(SHEET_NAMES.SALES, {
      SaleID: saleId, InvoiceNumber: invoiceNumber, CustomerID: payload.CustomerID || '',
      SaleDate: payload.SaleDate || nowIso_(), Subtotal: subtotal, Tax: tax, Discount: discount,
      GrandTotal: grandTotal, PaymentMethod: payload.PaymentMethod || 'Cash',
      PaymentStatus: payload.PaymentStatus || 'Paid', CreatedBy: who, CreatedAt: nowIso_()
    });

    payload.items.forEach(function (it) {
      const qty = toNumber_(it.Quantity);
      const price = toNumber_(it.UnitPrice);
      const lineDiscount = toNumber_(it.Discount);
      const lineTotal = (qty * price) - lineDiscount;
      const sitId = generateId_(SHEET_NAMES.SALE_ITEMS);
      appendRow_(SHEET_NAMES.SALE_ITEMS, {
        SaleItemID: sitId, SaleID: saleId, ProductID: it.ProductID, Quantity: qty,
        UnitPrice: price, Discount: lineDiscount, Total: lineTotal
      });
      const product = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID);
      const newStock = Math.max(0, toNumber_(product.CurrentStock) - qty);
      updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
      appendRow_(SHEET_NAMES.INVENTORY_LOG, {
        LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: it.ProductID, TransactionType: 'Sale',
        ReferenceID: saleId, QuantityIn: 0, QuantityOut: qty, Balance: newStock, CreatedBy: who, CreatedAt: nowIso_()
      });
    });

    auditLog_('Sale Created', { id: saleId });
    return ok_('Sale recorded.', { SaleID: saleId, InvoiceNumber: invoiceNumber, GrandTotal: grandTotal });
  });
}

function getSales() {
  return guard_(function () {
    const sales = readTable_(SHEET_NAMES.SALES);
    const items = readTable_(SHEET_NAMES.SALE_ITEMS);
    sales.forEach(function (s) { s.items = items.filter(function (i) { return i.SaleID === s.SaleID; }); });
    return ok_('OK', sales);
  });
}

function deleteSale(saleId) {
  return guard_(function () {
    const items = readTable_(SHEET_NAMES.SALE_ITEMS).filter(function (i) { return i.SaleID === saleId; });
    const who = getSessionEmail_() || 'system';
    items.forEach(function (it) {
      const product = findRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID);
      if (product) {
        const newStock = toNumber_(product.CurrentStock) + toNumber_(it.Quantity);
        updateRowById_(SHEET_NAMES.PRODUCTS, 'ProductID', it.ProductID, { CurrentStock: newStock, UpdatedAt: nowIso_() });
        appendRow_(SHEET_NAMES.INVENTORY_LOG, {
          LogID: generateId_(SHEET_NAMES.INVENTORY_LOG), ProductID: it.ProductID, TransactionType: 'Sale Reversal',
          ReferenceID: saleId, QuantityIn: it.Quantity, QuantityOut: 0, Balance: newStock, CreatedBy: who, CreatedAt: nowIso_()
        });
      }
      deleteRowById_(SHEET_NAMES.SALE_ITEMS, 'SaleItemID', it.SaleItemID);
    });
    const okDeleted = deleteRowById_(SHEET_NAMES.SALES, 'SaleID', saleId);
    if (!okDeleted) return fail_('Sale not found.');
    auditLog_('Sale Deleted', { id: saleId });
    return ok_('Sale deleted and stock restored.');
  });
}

/* ---------- SEARCH ---------- */

function searchProducts(query) {
  return guard_(function () {
    const q = String(query || '').toLowerCase();
    const results = readTable_(SHEET_NAMES.PRODUCTS).filter(function (p) {
      return String(p.ProductName).toLowerCase().indexOf(q) > -1 ||
        String(p.SKU).toLowerCase().indexOf(q) > -1 ||
        String(p.Barcode).toLowerCase().indexOf(q) > -1 ||
        String(p.Category).toLowerCase().indexOf(q) > -1;
    });
    return ok_('OK', results);
  });
}

function searchSales(query) {
  return guard_(function () {
    const q = String(query || '').toLowerCase();
    const customers = readTable_(SHEET_NAMES.CUSTOMERS);
    const results = readTable_(SHEET_NAMES.SALES).filter(function (s) {
      const cust = customers.find(function (c) { return c.CustomerID === s.CustomerID; });
      const custName = cust ? cust.CustomerName : '';
      return String(s.InvoiceNumber).toLowerCase().indexOf(q) > -1 || String(custName).toLowerCase().indexOf(q) > -1;
    });
    return ok_('OK', results);
  });
}

function searchPurchases(query) {
  return guard_(function () {
    const q = String(query || '').toLowerCase();
    const suppliers = readTable_(SHEET_NAMES.SUPPLIERS);
    const results = readTable_(SHEET_NAMES.PURCHASES).filter(function (p) {
      const sup = suppliers.find(function (s) { return s.SupplierID === p.SupplierID; });
      const supName = sup ? sup.CompanyName : '';
      return String(p.InvoiceNumber).toLowerCase().indexOf(q) > -1 || String(supName).toLowerCase().indexOf(q) > -1;
    });
    return ok_('OK', results);
  });
}

/* =========================================================
 * SECTION 6 — DASHBOARD / ANALYTICS
 * ========================================================= */

function getDashboard() {
  return guard_(function () {
    const products = readTable_(SHEET_NAMES.PRODUCTS);
    const suppliers = readTable_(SHEET_NAMES.SUPPLIERS);
    const customers = readTable_(SHEET_NAMES.CUSTOMERS);
    const sales = readTable_(SHEET_NAMES.SALES);
    const purchases = readTable_(SHEET_NAMES.PURCHASES);

    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Winnipeg', 'yyyy-MM-dd');
    const isToday = function (isoStr) { return String(isoStr).indexOf(todayStr) === 0; };

    const revenue = sales.reduce(function (sum, s) { return sum + toNumber_(s.GrandTotal); }, 0);
    const expenses = purchases.reduce(function (sum, p) { return sum + toNumber_(p.GrandTotal); }, 0);
    const inventoryValue = products.reduce(function (sum, p) { return sum + (toNumber_(p.CostPrice) * toNumber_(p.CurrentStock)); }, 0);
    const todaySales = sales.filter(function (s) { return isToday(s.SaleDate) || isToday(s.CreatedAt); })
      .reduce(function (sum, s) { return sum + toNumber_(s.GrandTotal); }, 0);
    const todayPurchases = purchases.filter(function (p) { return isToday(p.PurchaseDate) || isToday(p.CreatedAt); })
      .reduce(function (sum, p) { return sum + toNumber_(p.GrandTotal); }, 0);

    const lowStockItems = products.filter(function (p) { return toNumber_(p.CurrentStock) > 0 && toNumber_(p.CurrentStock) <= toNumber_(p.MinimumStock); });
    const outOfStockItems = products.filter(function (p) { return toNumber_(p.CurrentStock) <= 0; });

    // Recent transactions (last 8 sales + purchases merged, newest first).
    const recentSales = sales.slice(-8).map(function (s) { return { type: 'Sale', ref: s.InvoiceNumber, amount: s.GrandTotal, date: s.CreatedAt }; });
    const recentPurchases = purchases.slice(-8).map(function (p) { return { type: 'Purchase', ref: p.InvoiceNumber, amount: p.GrandTotal, date: p.CreatedAt }; });
    const recentTransactions = recentSales.concat(recentPurchases)
      .sort(function (a, b) { return new Date(b.date) - new Date(a.date); }).slice(0, 10);

    // Top selling products by quantity.
    const saleItems = readTable_(SHEET_NAMES.SALE_ITEMS);
    const qtyByProduct = {};
    saleItems.forEach(function (i) { qtyByProduct[i.ProductID] = (qtyByProduct[i.ProductID] || 0) + toNumber_(i.Quantity); });
    const topProducts = Object.keys(qtyByProduct).map(function (pid) {
      const p = products.find(function (pr) { return pr.ProductID === pid; });
      return { productId: pid, name: p ? p.ProductName : pid, quantitySold: qtyByProduct[pid] };
    }).sort(function (a, b) { return b.quantitySold - a.quantitySold; }).slice(0, 5);

    // Monthly revenue for the trailing 6 months (by CreatedAt month).
    const monthlyRevenue = monthlyBuckets_(sales, 'GrandTotal', 6);
    const monthlyPurchasesData = monthlyBuckets_(purchases, 'GrandTotal', 6);

    return ok_('OK', {
      totalProducts: products.length,
      totalSuppliers: suppliers.length,
      totalCustomers: customers.length,
      totalSales: sales.length,
      totalPurchases: purchases.length,
      revenue: revenue,
      expenses: expenses,
      profit: revenue - expenses,
      inventoryValue: inventoryValue,
      todaySales: todaySales,
      todayPurchases: todayPurchases,
      lowStockCount: lowStockItems.length,
      outOfStockCount: outOfStockItems.length,
      lowStockItems: lowStockItems,
      outOfStockItems: outOfStockItems,
      recentTransactions: recentTransactions,
      topProducts: topProducts,
      monthlyRevenue: monthlyRevenue,
      monthlyPurchases: monthlyPurchasesData
    });
  });
}

function monthlyBuckets_(rows, amountField, monthsBack) {
  const buckets = {};
  const labels = [];
  const now = new Date();
  for (let i = monthsBack - 1; i >= 0; i--) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const key = Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Winnipeg', 'yyyy-MM');
    buckets[key] = 0;
    labels.push(key);
  }
  rows.forEach(function (r) {
    const d = new Date(r.CreatedAt);
    if (isNaN(d.getTime())) return;
    const key = Utilities.formatDate(d, Session.getScriptTimeZone() || 'America/Winnipeg', 'yyyy-MM');
    if (buckets.hasOwnProperty(key)) buckets[key] += toNumber_(r[amountField]);
  });
  return labels.map(function (l) { return { month: l, total: Math.round(buckets[l] * 100) / 100 }; });
}

/* =========================================================
 * SECTION 7 — BACKUP / RESTORE / CSV
 * ========================================================= */

function exportCSV(sheetName) {
  return guard_(function () {
    if (!SHEET_SCHEMAS[sheetName]) return fail_('Unknown sheet: ' + sheetName);
    const sh = sheet_(sheetName);
    const data = sh.getDataRange().getValues();
    const csv = data.map(function (row) {
      return row.map(function (cell) {
        const val = String(cell === null || cell === undefined ? '' : cell);
        return /[",\n]/.test(val) ? '"' + val.replace(/"/g, '""') + '"' : val;
      }).join(',');
    }).join('\n');
    return ok_('CSV generated.', { csv: csv, filename: sheetName + '.csv' });
  });
}

function backupDatabase() {
  return guard_(function () {
    const backup = {};
    Object.keys(SHEET_SCHEMAS).forEach(function (name) {
      backup[name] = readTable_(name).map(function (row) {
        const clean = {};
        SHEET_SCHEMAS[name].forEach(function (h) { clean[h] = row[h]; });
        return clean;
      });
    });
    auditLog_('Backup Created', {});
    return ok_('Backup generated.', { backup: backup, generatedAt: nowIso_() });
  });
}

function restoreDatabase(backupJsonString) {
  return guard_(function () {
    const backup = typeof backupJsonString === 'string' ? JSON.parse(backupJsonString) : backupJsonString;
    if (!backup) return fail_('Invalid backup payload.');
    Object.keys(SHEET_SCHEMAS).forEach(function (name) {
      if (!backup[name]) return;
      const sh = sheet_(name);
      const lastRow = sh.getLastRow();
      if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
      const headers = SHEET_SCHEMAS[name];
      const rows = backup[name].map(function (obj) { return headers.map(function (h) { return obj[h] === undefined ? '' : obj[h]; }); });
      if (rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    });
    auditLog_('Backup Restored', {});
    return ok_('Database restored.');
  });
}

/* ---------- SETTINGS ---------- */

function getSettings() {
  return guard_(function () {
    const rows = readTable_(SHEET_NAMES.SETTINGS);
    const obj = {};
    rows.forEach(function (r) { obj[r.Key] = r.Value; });
    return ok_('OK', obj);
  });
}

function saveSettings(payload) {
  return guard_(function () {
    Object.keys(payload || {}).forEach(function (key) {
      const updated = updateRowById_(SHEET_NAMES.SETTINGS, 'Key', key, { Value: payload[key] });
      if (!updated) appendRow_(SHEET_NAMES.SETTINGS, { Key: key, Value: payload[key] });
    });
    auditLog_('Settings Updated', {});
    return ok_('Settings saved.');
  });
}

/* =========================================================
 * SECTION 9 — ANALYTICS / BUSINESS INTELLIGENCE
 * ========================================================= *
 * Everything here is derived from Products, Sales, SaleItems and
 * Purchases — no new sheets required. Methodology notes are returned
 * alongside the numbers so the "how" is transparent, not a black box.
 * ========================================================= */

function getInventoryAnalytics() {
  return guard_(function () {
    const products = readTable_(SHEET_NAMES.PRODUCTS);
    const saleItems = readTable_(SHEET_NAMES.SALE_ITEMS);
    const sales = readTable_(SHEET_NAMES.SALES);
    const settings = readTable_(SHEET_NAMES.SETTINGS).reduce(function (o, r) { o[r.Key] = r.Value; return o; }, {});
    const leadTimeDays = toNumber_(settings.LeadTimeDays, 7);

    // Join sale line items with their sale date for velocity windows.
    const saleDateBySaleId = {};
    sales.forEach(function (s) { saleDateBySaleId[s.SaleID] = new Date(s.CreatedAt || s.SaleDate); });

    const WINDOW_DAYS = 90;
    const windowStart = new Date(); windowStart.setDate(windowStart.getDate() - WINDOW_DAYS);

    // Per-product aggregates: all-time revenue/COGS (for ABC + turnover) and last-90-day quantity (for velocity).
    const perProduct = {};
    products.forEach(function (p) {
      perProduct[p.ProductID] = {
        productId: p.ProductID, name: p.ProductName, category: p.Category, sku: p.SKU,
        currentStock: toNumber_(p.CurrentStock), minimumStock: toNumber_(p.MinimumStock),
        costPrice: toNumber_(p.CostPrice), sellingPrice: toNumber_(p.SellingPrice),
        totalRevenue: 0, totalCOGS: 0, qtySoldAllTime: 0, qtySoldWindow: 0
      };
    });
    saleItems.forEach(function (item) {
      const agg = perProduct[item.ProductID];
      if (!agg) return;
      const qty = toNumber_(item.Quantity);
      const revenue = toNumber_(item.Total);
      agg.totalRevenue += revenue;
      agg.totalCOGS += qty * agg.costPrice;
      agg.qtySoldAllTime += qty;
      const saleDate = saleDateBySaleId[item.SaleID];
      if (saleDate && saleDate >= windowStart) agg.qtySoldWindow += qty;
    });

    const productList = Object.keys(perProduct).map(function (k) { return perProduct[k]; });

    /* ---------- ABC / PARETO CLASSIFICATION ----------
     * Rank products by revenue contribution. Classic 80/15/5 Pareto bands:
     * A = items making up the first 80% of cumulative revenue,
     * B = next 15% (80-95%), C = remaining 5% (95-100%).
     * This tells you which SKUs deserve tight control vs. which barely matter. */
    const totalRevenueAll = productList.reduce(function (s, p) { return s + p.totalRevenue; }, 0);
    const ranked = productList.slice().sort(function (a, b) { return b.totalRevenue - a.totalRevenue; });
    let cumulative = 0;
    const abcRows = ranked.map(function (p) {
      cumulative += p.totalRevenue;
      const cumulativePct = totalRevenueAll > 0 ? (cumulative / totalRevenueAll) * 100 : 0;
      let cls = 'C';
      if (cumulativePct <= 80) cls = 'A'; else if (cumulativePct <= 95) cls = 'B';
      return {
        productId: p.productId, name: p.name, category: p.category,
        revenue: Math.round(p.totalRevenue * 100) / 100,
        revenueSharePct: totalRevenueAll > 0 ? Math.round((p.totalRevenue / totalRevenueAll) * 1000) / 10 : 0,
        cumulativePct: Math.round(cumulativePct * 10) / 10,
        class: cls
      };
    });
    const abcSummary = { A: 0, B: 0, C: 0 };
    abcRows.forEach(function (r) { abcSummary[r.class]++; });

    /* ---------- REORDER POINT / DAYS-OF-STOCK ----------
     * Daily velocity = quantity sold in the trailing 90 days / 90.
     * Days of stock remaining = current stock / daily velocity.
     * Reorder point = (daily velocity * supplier lead time) + safety stock,
     * where safety stock is the product's configured Minimum Stock.
     * A product "needs reorder" once current stock drops to/below that point. */
    const reorderRows = productList.map(function (p) {
      const dailyVelocity = p.qtySoldWindow / WINDOW_DAYS;
      const daysOfStock = dailyVelocity > 0 ? Math.round((p.currentStock / dailyVelocity) * 10) / 10 : null;
      const reorderPoint = Math.ceil((dailyVelocity * leadTimeDays) + p.minimumStock);
      return {
        productId: p.productId, name: p.name, sku: p.sku, currentStock: p.currentStock,
        dailyVelocity: Math.round(dailyVelocity * 100) / 100,
        daysOfStock: daysOfStock, reorderPoint: reorderPoint,
        needsReorder: p.currentStock <= reorderPoint && dailyVelocity > 0
      };
    }).sort(function (a, b) {
      // Items needing reorder first, then soonest-to-stock-out.
      if (a.needsReorder !== b.needsReorder) return a.needsReorder ? -1 : 1;
      const ad = a.daysOfStock === null ? Infinity : a.daysOfStock;
      const bd = b.daysOfStock === null ? Infinity : b.daysOfStock;
      return ad - bd;
    });
    const reorderCount = reorderRows.filter(function (r) { return r.needsReorder; }).length;

    /* ---------- INVENTORY TURNOVER RATIO ----------
     * Turnover = trailing-12-month COGS / current inventory value (at cost).
     * Note: this uses CURRENT inventory value as the denominator (a
     * snapshot proxy for average inventory) since the system doesn't
     * store historical stock-value snapshots — a known simplification
     * worth stating out loud rather than presenting as exact. */
    const oneYearAgo = new Date(); oneYearAgo.setFullYear(oneYearAgo.getFullYear() - 1);
    let trailingCOGS = 0;
    saleItems.forEach(function (item) {
      const saleDate = saleDateBySaleId[item.SaleID];
      const agg = perProduct[item.ProductID];
      if (!agg || !saleDate || saleDate < oneYearAgo) return;
      trailingCOGS += toNumber_(item.Quantity) * agg.costPrice;
    });
    const currentInventoryValue = productList.reduce(function (s, p) { return s + (p.costPrice * p.currentStock); }, 0);
    const turnoverRatio = currentInventoryValue > 0 ? Math.round((trailingCOGS / currentInventoryValue) * 100) / 100 : 0;

    /* ---------- DEMAND FORECAST (simple linear trend) ----------
     * Trailing 6 months of revenue, plus a naive least-squares linear
     * regression projecting next month. Good enough to show forecasting
     * literacy without pretending to be a production forecasting engine. */
    const monthly = monthlyBuckets_(sales, 'GrandTotal', 6);
    const n = monthly.length;
    const xs = monthly.map(function (_, i) { return i; });
    const ys = monthly.map(function (m) { return m.total; });
    const xMean = xs.reduce(function (a, b) { return a + b; }, 0) / n;
    const yMean = ys.reduce(function (a, b) { return a + b; }, 0) / n;
    let num = 0, den = 0;
    for (let i = 0; i < n; i++) { num += (xs[i] - xMean) * (ys[i] - yMean); den += (xs[i] - xMean) * (xs[i] - xMean); }
    const slope = den !== 0 ? num / den : 0;
    const intercept = yMean - slope * xMean;
    const forecastNextMonth = Math.max(0, Math.round((intercept + slope * n) * 100) / 100);

    return ok_('OK', {
      methodology: {
        abc: 'Products ranked by trailing revenue; A = first 80% of cumulative revenue, B = next 15%, C = last 5%.',
        reorder: 'Daily velocity = qty sold in trailing 90 days / 90. Reorder point = (daily velocity x lead time) + minimum stock (safety stock).',
        turnover: 'Trailing 12-month COGS / current inventory value at cost (a snapshot proxy for average inventory).',
        forecast: 'Least-squares linear regression over the trailing 6 months of revenue, projected one month forward.'
      },
      leadTimeDays: leadTimeDays,
      turnoverRatio: turnoverRatio,
      trailingCOGS: Math.round(trailingCOGS * 100) / 100,
      currentInventoryValue: Math.round(currentInventoryValue * 100) / 100,
      abcSummary: abcSummary,
      abcRows: abcRows,
      reorderRows: reorderRows,
      reorderCount: reorderCount,
      monthlyRevenue: monthly,
      forecastNextMonth: forecastNextMonth,
      forecastSlope: Math.round(slope * 100) / 100
    });
  });
}

/* =========================================================
 * SECTION 8 — WEB APP ENTRY POINT
 * ========================================================= */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SupplyIQ — Inventory Analytics Platform')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
