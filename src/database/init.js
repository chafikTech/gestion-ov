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
