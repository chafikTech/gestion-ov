const { getDatabase } = require('../database/init');
const { getDailySalary, WORKER_TYPES } = require('./constants');

/**
 * Get all workers
 * @returns {Array} Array of worker objects
 */
function getAllWorkers() {
  const db = getDatabase();
  const stmt = db.prepare('SELECT * FROM workers ORDER BY nom_prenom ASC');
  return stmt.all();
}

/**
 * Get worker by ID
 * @param {number} id - Worker ID
 * @returns {Object|null} Worker object or null
 */
function getWorkerById(id) {
  const db = getDatabase();
  const stmt = db.prepare('SELECT * FROM workers WHERE id = ?');
  return stmt.get(id);
}

/**
 * Search workers by name or CIN
 * @param {string} query - Search query
 * @returns {Array} Array of matching worker objects
 */
function searchWorkers(query) {
  const db = getDatabase();
  const searchPattern = `%${query}%`;
  const stmt = db.prepare(`
    SELECT * FROM workers 
    WHERE nom_prenom LIKE ? OR cin LIKE ?
    ORDER BY nom_prenom ASC
  `);
  return stmt.all(searchPattern, searchPattern);
}

/**
 * Validate worker data
 * @param {Object} workerData - Worker data to validate
 * @throws {Error} If validation fails
 */
function validateWorkerData(workerData) {
  const { nom_prenom, cin, cin_validite, date_naissance, type } = workerData;

  if (!nom_prenom || typeof nom_prenom !== 'string' || nom_prenom.trim().length === 0) {
    throw new Error('Le nom et prénom sont obligatoires');
  }

  if (!cin || typeof cin !== 'string' || cin.trim().length === 0) {
    throw new Error('Le CIN est obligatoire');
  }

  if (!date_naissance || !/^\d{4}-\d{2}-\d{2}$/.test(date_naissance)) {
    throw new Error('La date de naissance doit être au format YYYY-MM-DD');
  }

  // Validate date is valid
  const dateObj = new Date(date_naissance);
  if (isNaN(dateObj.getTime())) {
    throw new Error('La date de naissance n\'est pas valide');
  }

  // Date should not be in the future
  if (dateObj > new Date()) {
    throw new Error('La date de naissance ne peut pas être dans le futur');
  }

  if (!type || !WORKER_TYPES[type]) {
    throw new Error(`Le type doit être 'OS' ou 'ONS'`);
  }

  if (cin_validite && !/^\d{4}-\d{2}-\d{2}$/.test(cin_validite)) {
    throw new Error('La validité CIN doit être au format YYYY-MM-DD');
  }

  if (cin_validite) {
    const cinDate = new Date(cin_validite);
    if (isNaN(cinDate.getTime())) {
      throw new Error('La date de validité CIN n\'est pas valide');
    }
  }
}

/**
 * Create a new worker
 * @param {Object} workerData - Worker data
 * @param {string} workerData.nom_prenom - Full name
 * @param {string} workerData.cin - National ID
 * @param {string} workerData.date_naissance - Date of birth (YYYY-MM-DD)
 * @param {string} workerData.type - Worker type (OS or ONS)
 * @returns {Object} Created worker with ID
 */
function createWorker(workerData) {
  // Validate input
  validateWorkerData(workerData);

  const { nom_prenom, cin, cin_validite, date_naissance, type } = workerData;

  // Calculate salary based on type
  const salaire_journalier = getDailySalary(type);

  const db = getDatabase();

  // Check if CIN already exists
  const existingWorker = db.prepare('SELECT id FROM workers WHERE cin = ?').get(cin);
  if (existingWorker) {
    throw new Error(`Un ouvrier avec le CIN ${cin} existe déjà`);
  }

  // Insert worker
  const stmt = db.prepare(`
    INSERT INTO workers (nom_prenom, cin, cin_validite, date_naissance, type, salaire_journalier)
    VALUES (?, ?, ?, ?, ?, ?)
  `);

  try {
    const result = stmt.run(
      nom_prenom.trim(),
      cin.trim(),
      cin_validite ? cin_validite : null,
      date_naissance,
      type,
      salaire_journalier
    );

    // Return created worker
    return getWorkerById(result.lastInsertRowid);
  } catch (error) {
    if (error.message.includes('UNIQUE constraint failed')) {
      throw new Error(`Un ouvrier avec le CIN ${cin} existe déjà`);
    }
    throw error;
  }
}

/**
 * Update an existing worker
 * @param {number} id - Worker ID
 * @param {Object} workerData - Updated worker data
 * @returns {Object} Updated worker
 */
function updateWorker(id, workerData) {
  // Validate input
  validateWorkerData(workerData);

  const { nom_prenom, cin, cin_validite, date_naissance, type } = workerData;

  // Calculate salary based on type
  const salaire_journalier = getDailySalary(type);

  const db = getDatabase();

  // Check if worker exists
  const existingWorker = getWorkerById(id);
  if (!existingWorker) {
    throw new Error(`L'ouvrier avec l'ID ${id} n'existe pas`);
  }

  // Check if CIN is taken by another worker
  const cinCheck = db.prepare('SELECT id FROM workers WHERE cin = ? AND id != ?').get(cin, id);
  if (cinCheck) {
    throw new Error(`Un autre ouvrier avec le CIN ${cin} existe déjà`);
  }

  // Update worker
  const stmt = db.prepare(`
    UPDATE workers 
    SET nom_prenom = ?, cin = ?, cin_validite = ?, date_naissance = ?, type = ?, salaire_journalier = ?
    WHERE id = ?
  `);

  try {
    stmt.run(
      nom_prenom.trim(),
      cin.trim(),
      cin_validite ? cin_validite : null,
      date_naissance,
      type,
      salaire_journalier,
      id
    );

    // Return updated worker
    return getWorkerById(id);
  } catch (error) {
    if (error.message.includes('UNIQUE constraint failed')) {
      throw new Error(`Un ouvrier avec le CIN ${cin} existe déjà`);
    }
    throw error;
  }
}

/**
 * Delete a worker
 * @param {number} id - Worker ID
 * @returns {boolean} True if deleted
 */
function deleteWorker(id) {
  const db = getDatabase();

  // Check if worker exists
  const existingWorker = getWorkerById(id);
  if (!existingWorker) {
    throw new Error(`L'ouvrier avec l'ID ${id} n'existe pas`);
  }

  // Delete worker (cascade will delete related presence records)
  const stmt = db.prepare('DELETE FROM workers WHERE id = ?');
  const result = stmt.run(id);

  return result.changes > 0;
}

/**
 * Get worker count by type
 * @returns {Object} Count by type
 */
function getWorkerCountByType() {
  const db = getDatabase();
  const stmt = db.prepare(`
    SELECT type, COUNT(*) as count 
    FROM workers 
    GROUP BY type
  `);
  const results = stmt.all();
  
  return {
    OS: results.find(r => r.type === 'OS')?.count || 0,
    ONS: results.find(r => r.type === 'ONS')?.count || 0,
    total: results.reduce((sum, r) => sum + r.count, 0)
  };
}

module.exports = {
  getAllWorkers,
  getWorkerById,
  searchWorkers,
  createWorker,
  updateWorker,
  deleteWorker,
  getWorkerCountByType
};
