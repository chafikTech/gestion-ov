const fs = require('fs');
const path = require('path');
const { getDatabasePath, closeDatabase } = require('../database/init');

function ensureDbExtension(filePath) {
  if (filePath.toLowerCase().endsWith('.db')) {
    return filePath;
  }
  return `${filePath}.db`;
}

async function exportDatabase(targetPath) {
  const dbPath = getDatabasePath();
  if (!fs.existsSync(dbPath)) {
    throw new Error('Database file not found');
  }

  const finalPath = ensureDbExtension(targetPath);
  const dir = path.dirname(finalPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }

  fs.copyFileSync(dbPath, finalPath);
  return { success: true, filePath: finalPath };
}

async function importDatabase(sourcePath) {
  if (!fs.existsSync(sourcePath)) {
    throw new Error('Import file not found');
  }

  const dbPath = getDatabasePath();
  closeDatabase();
  fs.copyFileSync(sourcePath, dbPath);
  return { success: true, filePath: dbPath };
}

module.exports = {
  exportDatabase,
  importDatabase
};
