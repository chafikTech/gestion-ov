const { getDatabase } = require('../database/init');
const { getWorkerById } = require('./workers');
const { getDaysInMonth, getMonthsInQuarter } = require('./constants');

function validateMonth(month) {
  if (month < 1 || month > 12) {
    throw new Error('Le mois doit être entre 1 et 12');
  }
}

function validateQuarter(quarter) {
  if (quarter < 1 || quarter > 4) {
    throw new Error('Le trimestre doit être entre 1 et 4');
  }
}

function normalizeAttachmentNumber(value) {
  if (value === null || value === undefined || value === '') {
    return null;
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return null;
  }

  const asInt = Math.trunc(numeric);
  if (asInt < 0) {
    return null;
  }

  return asInt;
}

function upsertAttachmentNumberInTransaction(db, workerId, year, month, attachmentNumber) {
  const normalized = normalizeAttachmentNumber(attachmentNumber);

  if (normalized === null) {
    const deleteStmt = db.prepare(`
      DELETE FROM presence_attachment_orders
      WHERE worker_id = ? AND year = ? AND month = ?
    `);
    deleteStmt.run(workerId, year, month);
    return;
  }

  const upsertStmt = db.prepare(`
    INSERT INTO presence_attachment_orders (worker_id, year, month, attachment_number)
    VALUES (?, ?, ?, ?)
    ON CONFLICT(worker_id, year, month)
    DO UPDATE SET attachment_number = excluded.attachment_number
  `);
  upsertStmt.run(workerId, year, month, normalized);
}

function normalizeAttachmentEntries(entries) {
  if (Array.isArray(entries)) {
    return entries;
  }

  if (!entries || typeof entries !== 'object') {
    return [];
  }

  return Object.entries(entries).map(([workerId, attachmentNumber]) => ({
    workerId: Number(workerId),
    attachmentNumber
  }));
}

function applyMonthlyAttachmentEntriesInTransaction(db, year, month, entries) {
  const normalizedEntries = normalizeAttachmentEntries(entries);
  for (const entry of normalizedEntries) {
    const workerId = Number(entry?.workerId);
    if (!Number.isInteger(workerId) || workerId <= 0) {
      continue;
    }

    const worker = getWorkerById(workerId);
    if (!worker) {
      continue;
    }

    upsertAttachmentNumberInTransaction(db, workerId, year, month, entry?.attachmentNumber);
  }
}

function getAttachmentOrderMapForMonth(year, month) {
  const db = getDatabase();
  validateMonth(month);

  const stmt = db.prepare(`
    SELECT worker_id, attachment_number
    FROM presence_attachment_orders
    WHERE year = ? AND month = ? AND attachment_number IS NOT NULL
  `);

  const rows = stmt.all(year, month);
  const orderMap = new Map();
  rows.forEach((row) => {
    orderMap.set(row.worker_id, row.attachment_number);
  });
  return orderMap;
}

function getAttachmentOrderMapForQuarter(year, quarter) {
  const db = getDatabase();
  validateQuarter(quarter);

  const months = getMonthsInQuarter(quarter);
  const placeholders = months.map(() => '?').join(',');

  const stmt = db.prepare(`
    SELECT worker_id, MIN(attachment_number) AS attachment_number
    FROM presence_attachment_orders
    WHERE year = ?
      AND month IN (${placeholders})
      AND attachment_number IS NOT NULL
    GROUP BY worker_id
  `);

  const rows = stmt.all(year, ...months);
  const orderMap = new Map();
  rows.forEach((row) => {
    orderMap.set(row.worker_id, row.attachment_number);
  });
  return orderMap;
}

/**
 * Get all workers sorted by monthly attachment number (nulls at the end)
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Array} Workers sorted for the selected month
 */
function getWorkersWithAttachmentOrderForMonth(year, month) {
  const db = getDatabase();
  validateMonth(month);

  const stmt = db.prepare(`
    SELECT w.*, pao.attachment_number
    FROM workers w
    LEFT JOIN presence_attachment_orders pao
      ON pao.worker_id = w.id AND pao.year = ? AND pao.month = ?
    ORDER BY
      CASE WHEN pao.attachment_number IS NULL THEN 1 ELSE 0 END,
      pao.attachment_number ASC,
      w.nom_prenom COLLATE NOCASE ASC
  `);

  const workers = stmt.all(year, month);
  return workers.map((worker) => ({
    ...worker,
    attachmentNumber: worker.attachment_number ?? null
  }));
}

/**
 * Get attachment number for a worker in a specific month
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {number|null} Attachment number
 */
function getAttachmentNumber(workerId, year, month) {
  const db = getDatabase();

  // Verify worker exists
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateMonth(month);

  const stmt = db.prepare(`
    SELECT attachment_number
    FROM presence_attachment_orders
    WHERE worker_id = ? AND year = ? AND month = ?
  `);

  const row = stmt.get(workerId, year, month);
  return row?.attachment_number ?? null;
}

/**
 * Save attachment number for a worker in a specific month
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @param {number|string|null} attachmentNumber - Attachment number (null clears value)
 * @returns {Object} Save result
 */
function saveAttachmentNumber(workerId, year, month, attachmentNumber) {
  const db = getDatabase();

  // Verify worker exists
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateMonth(month);

  upsertAttachmentNumberInTransaction(db, workerId, year, month, attachmentNumber);

  return {
    success: true,
    workerId,
    year,
    month,
    attachmentNumber: normalizeAttachmentNumber(attachmentNumber)
  };
}

/**
 * Save monthly attachment numbers for multiple workers
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @param {Array<{workerId:number,attachmentNumber:number|string|null}>|Object} entries
 * @returns {{success:boolean, updated:number}}
 */
function saveMonthlyAttachmentNumbers(year, month, entries) {
  const db = getDatabase();
  validateMonth(month);

  const normalizedEntries = normalizeAttachmentEntries(entries);
  let updated = 0;

  const transaction = db.transaction(() => {
    for (const entry of normalizedEntries) {
      const workerId = Number(entry?.workerId);
      if (!Number.isInteger(workerId) || workerId <= 0) {
        continue;
      }

      const worker = getWorkerById(workerId);
      if (!worker) {
        continue;
      }

      upsertAttachmentNumberInTransaction(db, workerId, year, month, entry?.attachmentNumber);
      updated += 1;
    }
  });

  transaction();

  return { success: true, updated };
}

/**
 * Get monthly attachment numbers for all workers (sorted)
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Array<{workerId:number,nom_prenom:string,cin:string,type:string,attachmentNumber:number|null}>}
 */
function getMonthlyAttachmentNumbers(year, month) {
  const workers = getWorkersWithAttachmentOrderForMonth(year, month);
  return workers.map((worker) => ({
    workerId: worker.id,
    nom_prenom: worker.nom_prenom,
    cin: worker.cin,
    type: worker.type,
    attachmentNumber: worker.attachmentNumber
  }));
}

/**
 * Get presence data for a worker in a specific month
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Array} Array of days where worker was present
 */
function getPresence(workerId, year, month) {
  const db = getDatabase();

  // Verify worker exists
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateMonth(month);

  const stmt = db.prepare(`
    SELECT day
    FROM presence
    WHERE worker_id = ? AND year = ? AND month = ? AND is_present = 1
    ORDER BY day ASC
  `);

  const results = stmt.all(workerId, year, month);
  return results.map(r => r.day);
}

/**
 * Save presence data for a worker in a specific month
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @param {Array<number>} days - Array of days where worker was present
 * @param {number|string|null|undefined} attachmentNumber - Optional current worker attachment number
 * @param {Array<{workerId:number,attachmentNumber:number|string|null}>|Object|undefined} monthlyAttachmentEntries - Optional monthly values for all workers
 * @returns {Object} Save result
 */
function savePresence(workerId, year, month, days, attachmentNumber = undefined, monthlyAttachmentEntries = undefined) {
  const db = getDatabase();

  // Verify worker exists
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateMonth(month);

  const safeDays = Array.isArray(days) ? days : [];

  // Validate days
  const maxDays = getDaysInMonth(year, month);
  const invalidDays = safeDays.filter(day => day < 1 || day > maxDays);
  if (invalidDays.length > 0) {
    throw new Error(`Jours invalides pour ce mois: ${invalidDays.join(', ')}`);
  }

  // Use transaction for atomicity
  const transaction = db.transaction(() => {
    // Delete all existing presence for this worker/year/month
    const deleteStmt = db.prepare(`
      DELETE FROM presence
      WHERE worker_id = ? AND year = ? AND month = ?
    `);
    deleteStmt.run(workerId, year, month);

    // Insert new presence records
    if (safeDays.length > 0) {
      const insertStmt = db.prepare(`
        INSERT INTO presence (worker_id, year, month, day, is_present)
        VALUES (?, ?, ?, ?, 1)
      `);

      for (const day of safeDays) {
        insertStmt.run(workerId, year, month, day);
      }
    }

    if (attachmentNumber !== undefined) {
      upsertAttachmentNumberInTransaction(db, workerId, year, month, attachmentNumber);
    }

    if (monthlyAttachmentEntries !== undefined) {
      applyMonthlyAttachmentEntriesInTransaction(db, year, month, monthlyAttachmentEntries);
    }
  });

  transaction();

  return {
    success: true,
    workerId,
    year,
    month,
    daysWorked: safeDays.length,
    attachmentNumber: attachmentNumber === undefined
      ? getAttachmentNumber(workerId, year, month)
      : normalizeAttachmentNumber(attachmentNumber)
  };
}

/**
 * Calculate monthly statistics for a worker
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Object} Statistics including days worked and total salary
 */
function calculateMonthlyStats(workerId, year, month) {
  // Get worker
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateMonth(month);

  // Get presence days
  const presenceDays = getPresence(workerId, year, month);
  const daysWorked = presenceDays.length;

  // Calculate total salary
  const totalSalary = daysWorked * worker.salaire_journalier;

  return {
    workerId,
    workerName: worker.nom_prenom,
    workerType: worker.type,
    year,
    month,
    dailySalary: worker.salaire_journalier,
    daysWorked,
    presenceDays,
    totalSalary,
    totalDaysInMonth: getDaysInMonth(year, month),
    attachmentNumber: getAttachmentNumber(workerId, year, month)
  };
}

/**
 * Get quarterly presence for a worker
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} quarter - Quarter (1-4)
 * @returns {Object} Quarterly statistics
 */
function getQuarterlyPresence(workerId, year, quarter) {
  const worker = getWorkerById(workerId);
  if (!worker) {
    throw new Error(`L'ouvrier avec l'ID ${workerId} n'existe pas`);
  }

  validateQuarter(quarter);

  const months = getMonthsInQuarter(quarter);
  const monthlyStats = [];
  let totalDaysWorked = 0;
  let totalSalary = 0;

  for (const month of months) {
    const stats = calculateMonthlyStats(workerId, year, month);
    monthlyStats.push(stats);
    totalDaysWorked += stats.daysWorked;
    totalSalary += stats.totalSalary;
  }

  const orderMap = getAttachmentOrderMapForQuarter(year, quarter);

  return {
    workerId,
    workerName: worker.nom_prenom,
    workerType: worker.type,
    cin: worker.cin,
    dateNaissance: worker.date_naissance,
    year,
    quarter,
    months,
    monthlyStats,
    totalDaysWorked,
    totalSalary,
    dailySalary: worker.salaire_journalier,
    attachmentNumber: orderMap.get(workerId) ?? null
  };
}

/**
 * Get all workers with presence in a specific month
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {Array} Array of workers with their presence stats
 */
function getWorkersWithPresence(year, month) {
  const db = getDatabase();
  validateMonth(month);

  const stmt = db.prepare(`
    SELECT DISTINCT w.*, pao.attachment_number
    FROM workers w
    INNER JOIN presence p ON w.id = p.worker_id
    LEFT JOIN presence_attachment_orders pao
      ON pao.worker_id = w.id AND pao.year = ? AND pao.month = ?
    WHERE p.year = ? AND p.month = ? AND p.is_present = 1
    ORDER BY
      CASE WHEN pao.attachment_number IS NULL THEN 1 ELSE 0 END,
      pao.attachment_number ASC,
      w.nom_prenom COLLATE NOCASE ASC
  `);

  const workers = stmt.all(year, month, year, month);

  // Calculate stats for each worker
  return workers.map(worker => {
    const stats = calculateMonthlyStats(worker.id, year, month);
    return {
      ...worker,
      ...stats,
      attachmentNumber: worker.attachment_number ?? stats.attachmentNumber ?? null
    };
  });
}

/**
 * Get all workers with presence in a specific quarter
 * @param {number} year - Year
 * @param {number} quarter - Quarter (1-4)
 * @returns {Array} Array of workers with their quarterly stats
 */
function getWorkersWithQuarterlyPresence(year, quarter) {
  const db = getDatabase();
  validateQuarter(quarter);

  const months = getMonthsInQuarter(quarter);
  const monthPlaceholders = months.map(() => '?').join(',');

  const stmt = db.prepare(`
    SELECT DISTINCT w.*, pao.attachment_number
    FROM workers w
    INNER JOIN presence p ON w.id = p.worker_id
    LEFT JOIN (
      SELECT worker_id, MIN(attachment_number) AS attachment_number
      FROM presence_attachment_orders
      WHERE year = ? AND month IN (${monthPlaceholders}) AND attachment_number IS NOT NULL
      GROUP BY worker_id
    ) pao ON pao.worker_id = w.id
    WHERE p.year = ? AND p.month IN (${monthPlaceholders}) AND p.is_present = 1
    ORDER BY
      CASE WHEN pao.attachment_number IS NULL THEN 1 ELSE 0 END,
      pao.attachment_number ASC,
      w.nom_prenom COLLATE NOCASE ASC
  `);

  const workers = stmt.all(year, ...months, year, ...months);

  // Calculate quarterly stats for each worker
  return workers.map(worker => {
    const stats = getQuarterlyPresence(worker.id, year, quarter);
    return {
      ...stats,
      attachmentNumber: worker.attachment_number ?? stats.attachmentNumber ?? null
    };
  });
}

/**
 * Delete all presence records for a specific month
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month (1-12)
 * @returns {boolean} True if deleted
 */
function deletePresence(workerId, year, month) {
  const db = getDatabase();
  validateMonth(month);

  const stmt = db.prepare(`
    DELETE FROM presence
    WHERE worker_id = ? AND year = ? AND month = ?
  `);

  const result = stmt.run(workerId, year, month);
  return result.changes > 0;
}

module.exports = {
  getPresence,
  savePresence,
  calculateMonthlyStats,
  getQuarterlyPresence,
  getWorkersWithPresence,
  getWorkersWithQuarterlyPresence,
  deletePresence,
  getAttachmentNumber,
  saveAttachmentNumber,
  getMonthlyAttachmentNumbers,
  saveMonthlyAttachmentNumbers,
  getWorkersWithAttachmentOrderForMonth,
  getAttachmentOrderMapForMonth,
  getAttachmentOrderMapForQuarter
};
