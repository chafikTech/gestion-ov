const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');
const { app } = require('electron');

let db = null;
let dbPath = null;

/**
 * Initialize the SQLite database
 * @returns {Promise<Database>} Database instance
 */
async function initDatabase() {
  try {
    // Get user data path
    const userDataPath = app.getPath('userData');
    dbPath = path.join(userDataPath, 'gestion_ouvriers.db');

    // Ensure directory exists
    if (!fs.existsSync(userDataPath)) {
      fs.mkdirSync(userDataPath, { recursive: true });
    }

    // Create or open database
    db = new Database(dbPath, { verbose: console.log });

    // Enable foreign keys
    db.pragma('foreign_keys = ON');

    // Read and execute schema
    const schemaPath = path.join(__dirname, 'schema.sql');
    const schema = fs.readFileSync(schemaPath, 'utf8');

    // Execute schema - execute all at once to handle triggers properly
    db.exec(schema);

    // Lightweight migration for new columns
    const workersColumns = db.prepare('PRAGMA table_info(workers)').all().map(col => col.name);
    if (!workersColumns.includes('cin_validite')) {
      db.exec('ALTER TABLE workers ADD COLUMN cin_validite TEXT');
    }

    // Lightweight migration for bordereau scope columns
    const bordereauColumns = db.prepare('PRAGMA table_info(bordereau_monthly_totals)').all().map(col => col.name);
    if (!bordereauColumns.includes('commune_id')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN commune_id TEXT');
    }
    if (!bordereauColumns.includes('commune_name')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN commune_name TEXT');
    }
    if (!bordereauColumns.includes('exercise_year')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN exercise_year TEXT');
    }
    if (!bordereauColumns.includes('chap')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN chap TEXT');
    }
    if (!bordereauColumns.includes('art')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN art TEXT');
    }
    if (!bordereauColumns.includes('prog')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN prog TEXT');
    }
    if (!bordereauColumns.includes('proj')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN proj TEXT');
    }
    if (!bordereauColumns.includes('ligne')) {
      db.exec('ALTER TABLE bordereau_monthly_totals ADD COLUMN ligne TEXT');
    }

    // Backfill newly introduced scope fields for legacy rows when possible.
    db.exec(`
      UPDATE bordereau_monthly_totals
      SET commune_id = commune_name
      WHERE IFNULL(commune_id, '') = '' AND IFNULL(commune_name, '') <> ''
    `);
    db.exec(`
      UPDATE bordereau_monthly_totals
      SET exercise_year = CAST(year AS TEXT)
      WHERE IFNULL(exercise_year, '') = ''
    `);

    // Create scoped bordereau index only after scope columns are guaranteed to exist.
    db.exec('DROP INDEX IF EXISTS idx_bordereau_scope_period');
    db.exec(`
      CREATE INDEX IF NOT EXISTS idx_bordereau_scope_period
      ON bordereau_monthly_totals(commune_id, exercise_year, chap, art, prog, proj, ligne, year, month)
    `);

    console.log('Database initialized successfully at:', dbPath);
    return db;
  } catch (error) {
    console.error('Failed to initialize database:', error);
    throw error;
  }
}

/**
 * Get database instance
 * @returns {Database} Database instance
 */
function getDatabase() {
  if (!db) {
    throw new Error('Database not initialized. Call initDatabase() first.');
  }
  return db;
}

/**
 * Get database file path
 * @returns {string} Database path
 */
function getDatabasePath() {
  if (!dbPath) {
    const userDataPath = app.getPath('userData');
    dbPath = path.join(userDataPath, 'gestion_ouvriers.db');
  }
  return dbPath;
}

/**
 * Close database connection
 */
function closeDatabase() {
  if (db) {
    db.close();
    db = null;
  }
}

module.exports = {
  initDatabase,
  getDatabase,
  getDatabasePath,
  closeDatabase
};
