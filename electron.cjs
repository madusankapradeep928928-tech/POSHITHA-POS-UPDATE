const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const bcrypt = require('bcryptjs');

const { autoUpdater } = require('electron-updater');
let mainWindow;
let db = null;

// SQLite3 Async Helpers
function dbAll(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.all(sql, params, (err, rows) => err ? reject(err) : resolve(rows));
  });
}

function dbGet(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.get(sql, params, (err, row) => err ? reject(err) : resolve(row));
  });
}

function dbRun(sql, params = []) {
  return new Promise((resolve, reject) => {
    db.run(sql, params, function(err) {
      err ? reject(err) : resolve(this);
    });
  });
}

// Safely add a column only when it does not already exist.
async function ensureColumn(tableName, columnName, columnDefinition) {
  const cols = await dbAll(`PRAGMA table_info(${tableName})`);
  if (!cols.some(c => c.name === columnName)) {
    await dbRun(`ALTER TABLE ${tableName} ADD COLUMN ${columnName} ${columnDefinition}`);
  }
}

// Keep sale identity stable across the UI, localStorage and SQLite.
function normalizeInvoiceNo(value) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? Math.trunc(n) : null;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 850,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true,
      spellcheck: false
    }
  });

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    mainWindow.focus();
  });

  const isDev = !app.isPackaged;

  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    const startUrl = path.join(__dirname, 'dist/index.html');
    mainWindow.loadFile(startUrl);
  }
}

// ==========================================
// DATABASE INITIALIZATION
// ==========================================

app.whenReady().then(() => {
  const sqlite3 = require('sqlite3').verbose();

  const dbPath = path.join(
    app.getPath('userData'),
    'pos_data.db'
  );

  global.db = new sqlite3.Database(dbPath);
  db = global.db;
  // Check for updates
  autoUpdater.checkForUpdatesAndNotify();

  // Production-safe SQLite settings for a desktop POS.
  db.configure('busyTimeout', 5000);
  db.run('PRAGMA journal_mode = WAL');
  db.run('PRAGMA synchronous = NORMAL');
  db.run('PRAGMA foreign_keys = ON');

  db.serialize(() => {

    // ======================================
    // 1. USERS TABLE
    // ======================================

    db.run(`
      CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        shop_name TEXT
      );
    `, () => {

      db.get(
        "SELECT * FROM users WHERE username = 'admin'",
        (err, row) => {

          if (!row) {

            const hashPassword =
              bcrypt.hashSync('admin123', 10);

            db.run(
              "INSERT INTO users (username, password, shop_name) VALUES (?, ?, ?)",
              [
                'admin',
                hashPassword,
                'Default Shop'
              ],
              (err) => {

                if (!err) {
                  console.log(
                    "Default admin created: admin / admin123"
                  );
                }

              }
            );
          }

        }
      );
    });

    // ======================================
    // 2. SHOPS TABLE
    // ======================================

    db.run(`
      CREATE TABLE IF NOT EXISTS shops (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        shop_name TEXT NOT NULL,
        owner_name TEXT,
        phone TEXT,
        address TEXT,
        created_at TEXT
      );
    `, () => {

      db.get(
        "SELECT * FROM shops WHERE id = 1",
        (err, row) => {

          if (!row) {

            const currentDate =
              new Date().toISOString();

            db.run(
              `INSERT INTO shops
              (shop_name, owner_name, phone, address, created_at)
              VALUES (?, ?, ?, ?, ?)`,
              [
                'Default Shop',
                'admin',
                '0712345678',
                'Sri Lanka',
                currentDate
              ]
            );
          }

        }
      );
    });

    // ======================================
    // 3. PRODUCTS TABLE
    // ======================================
    //
    // barcode is NOT UNIQUE.
    //
    // Same barcode can have multiple lots.
    //
    // Example:
    //
    // LOT-001 -> Buy 280 -> Sell 330
    // LOT-002 -> Buy 300 -> Sell 350
    //
    // Both can have the same barcode.
    // ======================================

    db.run(`
      CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        barcode TEXT,
        name TEXT,
        lot_number TEXT,
        grn_rate REAL DEFAULT 0,
        grn_date TEXT,
        buying_price REAL DEFAULT 0,
        price REAL,
        stock REAL DEFAULT 0,
        min_stock_alert REAL DEFAULT 5,
        unit TEXT DEFAULT 'Pcs',
        offer_type TEXT DEFAULT 'none',
        offer_value REAL DEFAULT 0,
        expiry_date TEXT,
        supplier TEXT
      );
    `);

    // ======================================
    // 4. SALES TABLE
    // ======================================

    db.run(`
      CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        total REAL,
        payment_method TEXT,
        paid_amount REAL,
        change_amount REAL,
        customer_name TEXT,
        customer_phone TEXT,
        date TEXT,
        invoice_no INTEGER,
        subtotal REAL DEFAULT 0,
        discount REAL DEFAULT 0,
        item_offer_discount REAL DEFAULT 0,
        bill_discount_percent REAL DEFAULT 0,
        bill_discount_lkr REAL DEFAULT 0,
        bill_discount_type TEXT DEFAULT 'none',
        bill_discount_amount REAL DEFAULT 0,
        profit REAL DEFAULT 0,
        customer_id INTEGER,
        credit_amount REAL DEFAULT 0,
        outstanding_balance REAL DEFAULT 0
      );
    `);

    // ======================================
    // 5. SALE ITEMS TABLE
    // ======================================

    db.run(`
      CREATE TABLE IF NOT EXISTS sale_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sale_id INTEGER,
        product_name TEXT,
        buying_price REAL DEFAULT 0,
        price REAL,
        qty REAL,
        offer_type TEXT DEFAULT 'none',
        offer_value REAL DEFAULT 0,
        item_discount REAL DEFAULT 0,
        final_price REAL DEFAULT 0,
        unit_price REAL DEFAULT 0
      );
    `);

    // ======================================
    // SAFE / IDEMPOTENT MIGRATIONS
    // ======================================
    // Never blindly ALTER TABLE: older databases already contain many of
    // these columns. Repeated ALTER TABLE calls can otherwise race startup
    // and make the database schema unreliable.

    (async () => {
      try {
        await ensureColumn('products', 'min_stock_alert', 'REAL DEFAULT 5');

        await ensureColumn('sales', 'payment_method', 'TEXT');
        await ensureColumn('sales', 'paid_amount', 'REAL');
        await ensureColumn('sales', 'change_amount', 'REAL');
        await ensureColumn('sales', 'customer_name', 'TEXT');
        await ensureColumn('sales', 'customer_phone', 'TEXT');
        await ensureColumn('sales', 'invoice_no', 'INTEGER');
        await ensureColumn('sales', 'subtotal', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'discount', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'item_offer_discount', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'bill_discount_percent', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'bill_discount_lkr', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'bill_discount_type', "TEXT DEFAULT 'none'");
        await ensureColumn('sales', 'bill_discount_amount', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'profit', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'customer_id', 'INTEGER');
        await ensureColumn('sales', 'credit_amount', 'REAL DEFAULT 0');
        await ensureColumn('sales', 'outstanding_balance', 'REAL DEFAULT 0');

        await ensureColumn('sale_items', 'buying_price', 'REAL DEFAULT 0');
        await ensureColumn('sale_items', 'offer_type', "TEXT DEFAULT 'none'");
        await ensureColumn('sale_items', 'offer_value', 'REAL DEFAULT 0');
        await ensureColumn('sale_items', 'product_id', 'INTEGER');
        await ensureColumn('sale_items', 'barcode', 'TEXT');
        await ensureColumn('sale_items', 'lot_number', 'TEXT');
        await ensureColumn('sale_items', 'item_discount', 'REAL DEFAULT 0');
        await ensureColumn('sale_items', 'final_price', 'REAL DEFAULT 0');
        await ensureColumn('sale_items', 'unit_price', 'REAL DEFAULT 0');

        // Old records pre-date invoice_no. Their SQLite id is their original
        // bill identity, so backfill once and keep it stable forever.
        await dbRun(`
          UPDATE sales
          SET invoice_no = id
          WHERE invoice_no IS NULL
             OR trim(CAST(invoice_no AS TEXT)) = ''
        `);

        // Fast and deterministic lookups used by billing, reports and the
        // customer dashboard. invoice_no is the canonical bill identity.
        await dbRun(`
          CREATE UNIQUE INDEX IF NOT EXISTS idx_sales_invoice_no
          ON sales(invoice_no)
        `);

        await dbRun(`
          CREATE INDEX IF NOT EXISTS idx_sales_customer_phone
          ON sales(customer_phone)
        `);

        await dbRun(`
          CREATE INDEX IF NOT EXISTS idx_sales_customer_id
          ON sales(customer_id)
        `);

        await dbRun(`
          CREATE INDEX IF NOT EXISTS idx_sale_items_sale_id
          ON sale_items(sale_id)
        `);

        await dbRun(`
          CREATE INDEX IF NOT EXISTS idx_sale_items_product_id
          ON sale_items(product_id)
        `);

        // Repair legacy rows that were saved before the full pricing fields existed.
        // final_price is the FINAL LINE TOTAL after item offers; unit_price is one-unit price.
        await dbRun(`
          UPDATE sale_items
          SET unit_price = COALESCE(NULLIF(unit_price, 0), price, 0)
          WHERE unit_price IS NULL OR unit_price = 0
        `);

        await dbRun(`
          UPDATE sale_items
          SET final_price = COALESCE(price, 0) * COALESCE(qty, 0)
          WHERE (final_price IS NULL OR final_price = 0)
            AND COALESCE(item_discount, 0) = 0
            AND COALESCE(offer_value, 0) = 0
        `);

        await dbRun(`
          UPDATE sale_items
          SET lot_number = COALESCE((
            SELECT p.lot_number FROM products p WHERE p.id = sale_items.product_id
          ), lot_number, '')
          WHERE (lot_number IS NULL OR trim(lot_number) = '')
            AND product_id IS NOT NULL
        `);

        await dbRun(`
          UPDATE sales
          SET subtotal = COALESCE((
            SELECT SUM(COALESCE(si.price, 0) * COALESCE(si.qty, 0))
            FROM sale_items si WHERE si.sale_id = sales.id
          ), total, 0)
          WHERE subtotal IS NULL OR subtotal = 0
        `);

        await dbRun(`
          UPDATE sales
          SET profit = COALESCE((
            SELECT SUM(
              COALESCE(NULLIF(si.final_price, 0), COALESCE(si.price, 0) * COALESCE(si.qty, 0))
              - (COALESCE(si.buying_price, 0) * COALESCE(si.qty, 0))
            )
            FROM sale_items si WHERE si.sale_id = sales.id
          ), 0) - COALESCE(bill_discount_amount, 0)
          WHERE profit IS NULL OR profit = 0
        `);

        // SQLite is fully suitable for the normal POS workload here. WAL and
        // a busy timeout reduce transient locking during backup/report reads.
        await dbRun(`PRAGMA journal_mode = WAL`);
        await dbRun(`PRAGMA synchronous = NORMAL`);
        await dbRun(`PRAGMA busy_timeout = 5000`);

      } catch (migrationErr) {
        console.error('Safe schema migration error:', migrationErr);
      }
    })();

    // ======================================
    // LOT / BATCH MIGRATION
    // ======================================

    (async () => {

      try {

        const cols =
          await dbAll(`PRAGMA table_info(products)`);

        const names =
          cols.map(c => c.name);

        if (!names.includes('lot_number')) {
          await dbRun(`
            ALTER TABLE products
            ADD COLUMN lot_number TEXT
          `);
        }

        if (!names.includes('grn_rate')) {
          await dbRun(`
            ALTER TABLE products
            ADD COLUMN grn_rate REAL DEFAULT 0
          `);
        }

        if (!names.includes('grn_date')) {
          await dbRun(`
            ALTER TABLE products
            ADD COLUMN grn_date TEXT
          `);
        }

        if (!names.includes('expiry_date')) {
          await dbRun(`
            ALTER TABLE products
            ADD COLUMN expiry_date TEXT
          `);
        }

        if (!names.includes('supplier')) {
          await dbRun(`
            ALTER TABLE products
            ADD COLUMN supplier TEXT
          `);
        }

        // ==================================
        // Detect old UNIQUE barcode schema
        // ==================================

        const createRow = await dbGet(`
          SELECT sql
          FROM sqlite_master
          WHERE type='table'
          AND name='products'
        `);

        const createSql =
          String(createRow?.sql || '').toUpperCase();

        if (createSql.includes('BARCODE TEXT UNIQUE')) {

          console.log(
            'Old products table detected. Migrating...'
          );

          await dbRun(`PRAGMA foreign_keys = OFF`);

          await dbRun('BEGIN TRANSACTION');

          // New table WITHOUT UNIQUE barcode
          await dbRun(`
            CREATE TABLE products_new (
              id INTEGER PRIMARY KEY AUTOINCREMENT,
              barcode TEXT,
              name TEXT,
              lot_number TEXT,
              grn_rate REAL DEFAULT 0,
              grn_date TEXT,
              buying_price REAL DEFAULT 0,
              price REAL,
              stock REAL DEFAULT 0,
              min_stock_alert REAL DEFAULT 5,
              unit TEXT DEFAULT 'Pcs',
              offer_type TEXT DEFAULT 'none',
              offer_value REAL DEFAULT 0,
              expiry_date TEXT,
              supplier TEXT
            )
          `);

          // Move existing data
          await dbRun(`
            INSERT INTO products_new
            (
              id,
              barcode,
              name,
              lot_number,
              grn_rate,
              grn_date,
              buying_price,
              price,
              stock,
              min_stock_alert,
              unit,
              offer_type,
              offer_value,
              expiry_date,
              supplier
            )
            SELECT
              id,
              barcode,
              name,
              NULL,
              buying_price,
              NULL,
              buying_price,
              price,
              stock,
              min_stock_alert,
              unit,
              offer_type,
              offer_value,
              NULL,
              NULL
            FROM products
          `);

          await dbRun(`
            DROP TABLE products
          `);

          await dbRun(`
            ALTER TABLE products_new
            RENAME TO products
          `);

          await dbRun('COMMIT');

          await dbRun(`PRAGMA foreign_keys = ON`);

          console.log(
            'Products table migrated successfully.'
          );
        }

      } catch (migrationErr) {

        try {
          await dbRun('ROLLBACK');
        } catch (_) {}

        console.error(
          'Lot migration error:',
          migrationErr
        );
      }

    })();

  });

  createWindow();
});

// ==========================================
// APP LIFECYCLE
// ==========================================

app.on('will-quit', () => {

  if (global.db) {

    try {
      global.db.close();
    } catch (err) {
      console.error(
        'Database close error:',
        err
      );
    }

  }

});

app.on('window-all-closed', () => {

  if (process.platform !== 'darwin') {
    app.quit();
  }

});

app.on('activate', () => {

  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }

});

// ==========================================
// AUTHENTICATION & SHOP IPC HANDLERS
// ==========================================

ipcMain.handle(
  'login-user',
  async (event, { username, password }) => {

    try {

      const user = await dbGet(
        "SELECT * FROM users WHERE username = ?",
        [username]
      );

      if (
        user &&
        bcrypt.compareSync(
          password,
          user.password
        )
      ) {

        return {
          success: true,
          user: {
            id: user.id,
            username: user.username,
            shop_name: user.shop_name
          }
        };

      }

      return {
        success: false,
        error:
          'වැරදි Username එකක් හෝ Password එකක්!'
      };

    } catch (err) {

      return {
        success: false,
        error: err.message
      };

    }

  }
);

ipcMain.handle(
  'register-user',
  async (
    event,
    { username, password, shopName }
  ) => {

    try {

      const hashPassword =
        bcrypt.hashSync(password, 10);

      const res = await dbRun(
        `INSERT INTO users
        (username, password, shop_name)
        VALUES (?, ?, ?)`,
        [
          username,
          hashPassword,
          shopName || 'My Shop'
        ]
      );

      await dbRun(
        `INSERT INTO shops
        (shop_name, owner_name, created_at)
        VALUES (?, ?, ?)`,
        [
          shopName || 'My Shop',
          username,
          new Date().toISOString()
        ]
      );

      return {
        success: true,
        id: res.lastID
      };

    } catch (err) {

      return {
        success: false,
        error:
          'Username එක දැනටමත් භාවිතා කර ඇත / Already exists!'
      };

    }

  }
);

ipcMain.handle(
  'get-shops',
  async () => {

    try {

      return await dbAll(
        'SELECT * FROM shops ORDER BY id DESC'
      );

    } catch (err) {

      return [];

    }

  }
);

ipcMain.handle(
  'update-shop',
  async (event, shop) => {

    try {

      await dbRun(
        `UPDATE shops
         SET shop_name = ?,
             owner_name = ?,
             phone = ?,
             address = ?
         WHERE id = ?`,
        [
          shop.shop_name,
          shop.owner_name,
          shop.phone,
          shop.address,
          shop.id
        ]
      );

      return {
        success: true
      };

    } catch (err) {

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// DATABASE BACKUP & RESTORE
// ==========================================

ipcMain.handle(
  'db-backup',
  async () => {

    try {

      const dbPath =
        path.join(
          app.getPath('userData'),
          'pos_data.db'
        );

      const { filePath } =
        await dialog.showSaveDialog({
          title: 'Backup Database',
          defaultPath:
            `pos_backup_${new Date()
              .toISOString()
              .slice(0, 10)}.db`,
          filters: [
            {
              name: 'Database Files',
              extensions: ['db']
            }
          ]
        });

      if (filePath) {

        fs.copyFileSync(
          dbPath,
          filePath
        );

        return {
          success: true,
          message:
            'ඩේටාබේස් බැකප් කිරීම සාර්ථකයි!'
        };

      }

      return {
        success: false,
        message:
          'බැකප් කිරීම අවලංගු කරන ලදී.'
      };

    } catch (error) {

      console.error(
        'Backup error:',
        error
      );

      return {
        success: false,
        error: error.message
      };

    }

  }
);

ipcMain.handle(
  'db-restore',
  async () => {

    try {

      const dbPath =
        path.join(
          app.getPath('userData'),
          'pos_data.db'
        );

      const { filePaths } =
        await dialog.showOpenDialog({
          title: 'Restore Database',
          filters: [
            {
              name: 'Database Files',
              extensions: ['db']
            }
          ],
          properties: ['openFile']
        });

      if (
        filePaths &&
        filePaths.length > 0
      ) {

        const selectedFile =
          filePaths[0];

        if (global.db) {
          global.db.close();
        }

        fs.copyFileSync(
          selectedFile,
          dbPath
        );

        return {
          success: true,
          message:
            'ඩේටාබේස් එක සාර්ථකව ප්‍රතිස්ථාපනය විය! කරුණාකර ඇප් එක Restart කරන්න.'
        };

      }

      return {
        success: false,
        message:
          'Restore කිරීම අවලංගු කරන ලදී.'
      };

    } catch (error) {

      console.error(
        'Restore error:',
        error
      );

      return {
        success: false,
        error: error.message
      };

    }

  }
);

// ==========================================
// PRODUCTS & SALES IPC HANDLERS
// ==========================================

ipcMain.handle(
  'get-products',
  async () => {

    try {

      return await dbAll(
        'SELECT * FROM products ORDER BY id DESC'
      );

    } catch (err) {

      return [];

    }

  }
);

// ==========================================
// ADD PRODUCT / LOT
// ==========================================

ipcMain.handle(
  'add-product',
  async (event, product) => {

    try {

      const sql = `
        INSERT INTO products
        (
          barcode,
          name,
          lot_number,
          grn_rate,
          grn_date,
          buying_price,
          price,
          stock,
          min_stock_alert,
          unit,
          offer_type,
          offer_value,
          expiry_date,
          supplier
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;

      const params = [

        product.barcode || '',

        product.name || '',

        product.lot_number || '',

        parseFloat(product.grn_rate) || 0,

        product.grn_date ||
          new Date().toISOString(),

        parseFloat(
          product.buying_price
        ) || 0,

        parseFloat(
          product.price
        ) || 0,

        parseFloat(
          product.stock
        ) || 0,

        parseFloat(
          product.min_stock_alert
        ) || 5,

        product.unit || 'Pcs',

        product.offer_type || 'none',

        parseFloat(
          product.offer_value
        ) || 0,

        product.expiry_date || '',

        product.supplier || ''

      ];

      const res =
        await dbRun(sql, params);

      return {
        success: true,
        id: res.lastID
      };

    } catch (err) {

      console.error(
        'Add product error:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// BULK ADD PRODUCTS
// ==========================================

ipcMain.handle(
  'bulk-add-products',
  async (
    event,
    productList
  ) => {

    try {

      const sql = `
        INSERT INTO products
        (
          barcode,
          name,
          lot_number,
          grn_rate,
          grn_date,
          buying_price,
          price,
          stock,
          min_stock_alert,
          unit,
          offer_type,
          offer_value,
          expiry_date,
          supplier
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;

      for (
        const p of
        (
          Array.isArray(productList)
            ? productList
            : []
        )
      ) {

        await dbRun(
          sql,
          [

            p.barcode || '',

            p.name || '',

            p.lot_number || '',

            parseFloat(
              p.grn_rate
            ) || 0,

            p.grn_date ||
              new Date().toISOString(),

            parseFloat(
              p.buying_price
            ) || 0,

            parseFloat(
              p.price
            ) || 0,

            parseFloat(
              p.stock
            ) || 0,

            parseFloat(
              p.min_stock_alert
            ) || 5,

            p.unit || 'Pcs',

            p.offer_type || 'none',

            parseFloat(
              p.offer_value
            ) || 0,

            p.expiry_date || '',

            p.supplier || ''

          ]
        );

      }

      return {
        success: true
      };

    } catch (err) {

      console.error(
        'Bulk add products error:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// UPDATE PRODUCT / LOT
// ==========================================

ipcMain.handle(
  'update-product',
  async (
    event,
    product
  ) => {

    try {

      const sql = `
        UPDATE products
        SET
          barcode = ?,
          name = ?,
          lot_number = ?,
          grn_rate = ?,
          grn_date = ?,
          buying_price = ?,
          price = ?,
          stock = ?,
          min_stock_alert = ?,
          unit = ?,
          offer_type = ?,
          offer_value = ?,
          expiry_date = ?,
          supplier = ?
        WHERE id = ?
      `;

      const params = [

        product.barcode || '',

        product.name || '',

        product.lot_number || '',

        parseFloat(
          product.grn_rate
        ) || 0,

        product.grn_date || '',

        parseFloat(
          product.buying_price
        ) || 0,

        parseFloat(
          product.price
        ) || 0,

        parseFloat(
          product.stock
        ) || 0,

        parseFloat(
          product.min_stock_alert
        ) || 5,

        product.unit || 'Pcs',

        product.offer_type || 'none',

        parseFloat(
          product.offer_value
        ) || 0,

        product.expiry_date || '',

        product.supplier || '',

        product.id

      ];

      await dbRun(
        sql,
        params
      );

      return {
        success: true
      };

    } catch (err) {

      console.error(
        'Update product error:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// DELETE PRODUCT / LOT
// ==========================================

ipcMain.handle(
  'delete-product',
  async (
    event,
    id
  ) => {

    try {

      await dbRun(
        'DELETE FROM products WHERE id = ?',
        [id]
      );

      return {
        success: true
      };

    } catch (err) {

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// CLEAR ALL PRODUCTS
// ==========================================

ipcMain.handle(
  'clear-all-products',
  async () => {

    try {

      await dbRun(
        'DELETE FROM products'
      );

      return {
        success: true
      };

    } catch (err) {

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// SAVE SALE
// ==========================================
//
// IMPORTANT:
// Exact product ID is used for stock deduction.
//
// This prevents:
// LOT-001 and LOT-002 having same barcode
// from both losing stock.
// ==========================================

ipcMain.handle(
  'save-sale',
  async (
    event,
    saleData
  ) => {

    try {

      const formattedDate =
        saleData.date || new Date().toISOString();

      const requestedInvoiceNo = normalizeInvoiceNo(
        saleData.invoiceNo ?? saleData.invoice_no
      );

      // The renderer already creates the invoice number used by the receipt.
      // Reusing that number makes this IPC operation idempotent: if the same
      // bill is sent twice, the second call returns the existing bill instead
      // of creating a second customer-dashboard entry or deducting stock twice.
      if (requestedInvoiceNo != null) {
        const existingSale = await dbGet(
          `SELECT id, invoice_no FROM sales WHERE invoice_no = ? LIMIT 1`,
          [requestedInvoiceNo]
        );

        if (existingSale) {
          return {
            success: true,
            saleId: existingSale.id,
            invoiceNo: existingSale.invoice_no,
            alreadySaved: true
          };
        }
      }

      await dbRun('BEGIN IMMEDIATE TRANSACTION');

      let invoiceNo = requestedInvoiceNo;

      if (invoiceNo == null) {
        const maxRow = await dbGet(
          `SELECT COALESCE(MAX(invoice_no), 0) AS max_invoice FROM sales`
        );
        invoiceNo = Number(maxRow?.max_invoice || 0) + 1;
      }

      // A second check is required after BEGIN IMMEDIATE so a stale renderer
      // cannot create the same invoice twice.
      const duplicateSale = await dbGet(
        `SELECT id, invoice_no FROM sales WHERE invoice_no = ? LIMIT 1`,
        [invoiceNo]
      );

      if (duplicateSale) {
        await dbRun('ROLLBACK');
        return {
          success: true,
          saleId: duplicateSale.id,
          invoiceNo: duplicateSale.invoice_no,
          alreadySaved: true
        };
      }

      const insertSaleSql = `
        INSERT INTO sales
        (
          total,
          payment_method,
          paid_amount,
          change_amount,
          customer_name,
          customer_phone,
          date,
          invoice_no,
          subtotal,
          discount,
          item_offer_discount,
          bill_discount_percent,
          bill_discount_lkr,
          bill_discount_type,
          bill_discount_amount,
          profit,
          customer_id,
          credit_amount,
          outstanding_balance
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;

      const saleParams = [
        Number(saleData.total) || 0,
        saleData.paymentMethod || 'Cash',
        Number(saleData.paidAmount) || Number(saleData.total) || 0,
        Number(saleData.changeAmount) || 0,
        saleData.customerName || 'General Customer',
        saleData.customerPhone || '-',
        formattedDate,
        invoiceNo,
        Number(saleData.subtotal) || 0,
        Number(saleData.discount) || 0,
        Number(saleData.itemOfferDiscount) || 0,
        Number(saleData.billDiscountPercent) || 0,
        Number(saleData.billDiscountLkr) || 0,
        saleData.billDiscountType || 'none',
        Number(saleData.billDiscountAmount) || 0,
        Number(saleData.profit) || 0,
        saleData.customer_id ?? saleData.customerId ?? null,
        Number(saleData.creditAmount) || 0,
        Number(saleData.outstandingBalance) || 0
      ];

      const saleRes =
        await dbRun(
          insertSaleSql,
          saleParams
        );

      const saleId =
        saleRes.lastID;

      const insertItemSql = `
        INSERT INTO sale_items
        (
          sale_id,
          product_name,
          buying_price,
          price,
          qty,
          offer_type,
          offer_value,
          product_id,
          barcode,
          lot_number,
          item_discount,
          final_price,
          unit_price
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `;

      if (
        saleData.items &&
        Array.isArray(
          saleData.items
        )
      ) {

        for (
          const item of
          saleData.items
        ) {

          const itemQty =
            parseFloat(
              item.qty
            ) || 1;

          const buyPrice =
            parseFloat(
              item.buying_price ??
              item.buyingPrice
            ) || 0;

          const sellPrice =
            parseFloat(
              item.price
            ) || 0;

          const itemName =
            item.name ||
            item.product_name ||
            '';

          const itemBarcode =
            item.barcode ||
            '';

          const itemId =
            item.id ??
            item.product_id ??
            null;

          const lotNumber =
            item.lot_number ||
            item.lotNumber ||
            '';

          const offerType =
            item.offer_type ||
            item.offerType ||
            'none';

          const offerValue =
            parseFloat(
              item.offer_value ??
              item.offerValue
            ) || 0;

          // finalPrice is the final LINE TOTAL; unitPrice is the normal one-unit selling price.
          const itemDiscount =
            parseFloat(
              item.item_discount ??
              item.itemDiscount ??
              item.discount
            ) || 0;

          const unitPrice =
            parseFloat(
              item.unit_price ??
              item.unitPrice ??
              item.price
            ) || sellPrice;

          const finalPriceRaw =
            parseFloat(
              item.finalPrice ??
              item.final_price
            );

          const finalPrice =
            Number.isFinite(finalPriceRaw)
              ? finalPriceRaw
              : Math.max(0, (unitPrice * itemQty) - itemDiscount);

          if (itemQty <= 0) {
            continue;
          }

          // ==================================
          // EXACT PRODUCT / LOT
          // ==================================

          let targetId =
            itemId;

          if (
            targetId != null
          ) {

            const product =
              await dbGet(
                `
                SELECT
                  id,
                  stock,
                  barcode,
                  lot_number
                FROM products
                WHERE id = ?
                `,
                [targetId]
              );

            if (!product) {

              throw new Error(
                `Product/Lot not found: ${targetId}`
              );

            }

            if (
              Number(product.stock) <
              itemQty
            ) {

              throw new Error(
                `Insufficient stock for ${itemName || 'product'}${lotNumber ? ` (Lot ${lotNumber})` : ''}`
              );

            }

          }

          // ==================================
          // LEGACY BARCODE FALLBACK
          // ==================================

          else if (
            itemBarcode
          ) {

            const product =
              await dbGet(
                `
                SELECT
                  id,
                  stock
                FROM products
                WHERE
                  trim(barcode) =
                  trim(?)
                  AND stock >= ?
                ORDER BY
                  CASE
                    WHEN grn_date IS NULL
                    OR grn_date = ''
                    THEN 1
                    ELSE 0
                  END,
                  grn_date ASC,
                  id ASC
                LIMIT 1
                `,
                [
                  itemBarcode,
                  itemQty
                ]
              );

            if (product) {

              targetId =
                product.id;

            }

          }

          await dbRun(
            insertItemSql,
            [
              saleId,
              itemName,
              buyPrice,
              sellPrice,
              itemQty,
              offerType,
              offerValue,
              targetId,
              itemBarcode,
              lotNumber,
              itemDiscount,
              finalPrice,
              unitPrice
            ]
          );

          // Exact lot stock deduction. Same barcode rows are never all
          // touched; only the selected product_id loses stock.
          if (
            targetId != null
          ) {

            const res =
              await dbRun(
                `
                UPDATE products
                SET stock =
                  stock - ?
                WHERE
                  id = ?
                  AND stock >= ?
                `,
                [
                  itemQty,
                  targetId,
                  itemQty
                ]
              );

            if (
              res.changes === 0
            ) {

              throw new Error(
                `Stock changed before sale was completed for ${itemName || 'product'}`
              );

            }

          }

          // Old records without a product id can still use name fallback.
          else if (
            itemName
          ) {

            const product =
              await dbGet(
                `
                SELECT
                  id,
                  stock
                FROM products
                WHERE
                  trim(LOWER(name)) =
                  trim(LOWER(?))
                  AND stock >= ?
                ORDER BY id ASC
                LIMIT 1
                `,
                [
                  itemName,
                  itemQty
                ]
              );

            if (product) {

              const res =
                await dbRun(
                  `
                  UPDATE products
                  SET stock =
                    stock - ?
                  WHERE
                    id = ?
                    AND stock >= ?
                  `,
                  [
                    itemQty,
                    product.id,
                    itemQty
                  ]
                );

              if (
                res.changes === 0
              ) {

                throw new Error(
                  `Stock update failed for ${itemName}`
                );

              }

            }

          }

        }

      }

      await dbRun(
        'COMMIT'
      );

      return {
        success: true,
        saleId,
        invoiceNo,
        alreadySaved: false
      };

    } catch (err) {

      try {
        await dbRun(
          'ROLLBACK'
        );
      } catch (_) {}

      console.error(
        'Error saving sale:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// DELETE SALE
// ==========================================
//
// When a sale is deleted,
// stock is returned to the EXACT lot.
// ==========================================

ipcMain.handle(
  'delete-sale',
  async (
    event,
    saleId
  ) => {

    try {

      await dbRun(
        'BEGIN TRANSACTION'
      );

      // The UI uses the displayed invoice number. Resolve it to the real
      // SQLite sales.id before touching sale_items or stock. Legacy callers
      // that still pass the DB id continue to work.
      const saleRow = await dbGet(
        `SELECT id, COALESCE(invoice_no, id) AS invoice_no
         FROM sales
         WHERE invoice_no = ? OR id = ?
         ORDER BY CASE WHEN invoice_no = ? THEN 0 ELSE 1 END
         LIMIT 1`,
        [saleId, saleId, saleId]
      );

      if (!saleRow) {
        await dbRun('ROLLBACK');
        return { success: false, error: `Sale/Invoice ${saleId} not found` };
      }

      const dbSaleId = saleRow.id;

      const items =
        await dbAll(
          `
          SELECT
            product_id,
            qty,
            product_name,
            barcode
          FROM sale_items
          WHERE sale_id = ?
          `,
          [dbSaleId]
        );

      for (
        const item of items
      ) {

        const qty =
          parseFloat(
            item.qty
          ) || 0;

        if (qty <= 0) {
          continue;
        }

        // Exact lot
        if (
          item.product_id != null
        ) {

          await dbRun(
            `
            UPDATE products
            SET stock =
              stock + ?
            WHERE id = ?
            `,
            [
              qty,
              item.product_id
            ]
          );

        }

        // Legacy fallback
        else if (
          item.barcode
        ) {

          const product =
            await dbGet(
              `
              SELECT id
              FROM products
              WHERE
                trim(barcode) =
                trim(?)
              ORDER BY id ASC
              LIMIT 1
              `,
              [item.barcode]
            );

          if (product) {

            await dbRun(
              `
              UPDATE products
              SET stock =
                stock + ?
              WHERE id = ?
              `,
              [
                qty,
                product.id
              ]
            );

          }

        }

      }

      await dbRun(
        `
        DELETE FROM sale_items
        WHERE sale_id = ?
        `,
        [dbSaleId]
      );

      await dbRun(
        `
        DELETE FROM sales
        WHERE id = ?
        `,
        [dbSaleId]
      );

      await dbRun(
        'COMMIT'
      );

      return {
        success: true
      };

    } catch (err) {

      try {
        await dbRun(
          'ROLLBACK'
        );
      } catch (_) {}

      console.error(
        'Error deleting sale:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// CLEAR ALL SALES
// ==========================================

ipcMain.handle(
  'clear-all-sales',
  async () => {

    try {

      await dbRun(
        'BEGIN IMMEDIATE TRANSACTION'
      );

      await dbRun(
        'DELETE FROM sale_items'
      );

      await dbRun(
        'DELETE FROM sales'
      );

      await dbRun(
        'COMMIT'
      );

      return {
        success: true
      };

    } catch (err) {

      try {
        await dbRun('ROLLBACK');
      } catch (_) {}

      console.error(
        'Error clearing sales:',
        err
      );

      return {
        success: false,
        error: err.message
      };

    }

  }
);

// ==========================================
// NEXT INVOICE NUMBER
// ==========================================

ipcMain.handle(
  'get-next-invoice-no',
  async () => {
    try {
      const row = await dbGet(`SELECT COALESCE(MAX(invoice_no), 0) AS max_invoice FROM sales`);
      return Number(row?.max_invoice || 0) + 1;
    } catch (err) {
      console.error('Next invoice number error:', err);
      return 1;
    }
  }
);

// ==========================================
// GET SALES
// ==========================================

ipcMain.handle(
  'get-sales',
  async () => {

    try {

      return await dbAll(
        `
        SELECT
          s.*,
          COALESCE(s.invoice_no, s.id) AS invoiceNo
        FROM sales s
        ORDER BY COALESCE(s.invoice_no, s.id) DESC, s.id DESC
        `
      );

    } catch (err) {

      return [];

    }

  }
);

// ==========================================
// REPORT DATA
// ==========================================

ipcMain.handle(
  'get-reports-data',
  async () => {

    try {

      const salesReportSql = `
        SELECT

          s.id AS sale_id,

          COALESCE(s.invoice_no, s.id) AS invoice_no,

          s.date,

          s.customer_name,

          s.payment_method,

          s.total AS sale_total,

          si.product_name,

          COALESCE(
            si.buying_price,
            0
          ) AS buying_price,

          si.price,

          si.qty,

          si.offer_type,

          si.offer_value,

          si.item_discount,

          si.final_price,

          si.unit_price,

          si.lot_number,

          si.barcode,

          si.product_id,

          (
            COALESCE(
              NULLIF(si.final_price, 0),
              COALESCE(si.price, 0) * COALESCE(si.qty, 0)
            )
            -
            (COALESCE(si.buying_price, 0) * COALESCE(si.qty, 0))
          ) AS profit

        FROM sales s

        LEFT JOIN sale_items si
          ON s.id = si.sale_id

        ORDER BY
          s.id DESC
      `;

      const salesReport =
        await dbAll(
          salesReportSql
        );

      const lowStock =
        await dbAll(
          `
          SELECT *
          FROM products
          WHERE stock <=
            COALESCE(
              min_stock_alert,
              5
            )
          ORDER BY stock ASC
          `
        );

      return {
        sales:
          salesReport,
        lowStock
      };

    } catch (err) {

      console.error(
        'Report error:',
        err
      );

      return {
        sales: [],
        lowStock: []
      };

    }

  }
);