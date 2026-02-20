const fs = require('node:fs');
const path = require('node:path');
const { app } = require('electron');
const { getAllWorkers } = require('./workers');
const {
  calculateMonthlyStats,
  getWorkersWithQuarterlyPresence,
  getAttachmentOrderMapForMonth,
  getAttachmentOrderMapForQuarter
} = require('./presence');
const { getMonthsInQuarter, RCAR_AGE_LIMIT, RCAR_RATE, calculateAgeAt, getMonthHalfRanges, calculatePresenceStatsForRange } = require('./constants');
const { getAppSettings } = require('./settingsStore');
const pythonBridge = require('./pythonBridge');

// Role des journées (special-case: exact scanned template via src/python/generate_role.py)
const roleJournees = require('../documents/monthly/roleJournees');

const PYTHON_GENERIC_SCRIPT = path.join('src', 'python', 'generate_document.py');

function getDefaultDocumentsOutputDir() {
  return path.join(app.getPath('documents'), 'Gestion_Ouvriers', 'Documents');
}

async function generateWithPython({ documentType, year, month, quarter, report, options, outputDir }) {
  const finalOutputDir = (outputDir || '').toString().trim() || getDefaultDocumentsOutputDir();
  if (!fs.existsSync(finalOutputDir)) {
    fs.mkdirSync(finalOutputDir, { recursive: true });
  }

  return await pythonBridge.runPythonJson(PYTHON_GENERIC_SCRIPT, {
    documentType,
    year,
    month,
    quarter,
    report,
    options,
    outputDir: finalOutputDir,
    // DOCX-only output (forced)
    format: 'docx'
  });
}

function buildGeneratorOptions(userOptions = {}) {
  const appSettings = getAppSettings();
  return {
    ...appSettings,
    ...userOptions,
    // Force DOCX even if the UI passes something else
    format: 'docx'
  };
}

function sortWorkersByAttachmentOrder(workers, attachmentOrderMap) {
  return [...workers].sort((a, b) => {
    const aOrder = attachmentOrderMap.get(a.id);
    const bOrder = attachmentOrderMap.get(b.id);
    const aHas = Number.isInteger(aOrder);
    const bHas = Number.isInteger(bOrder);

    if (aHas && bHas && aOrder !== bOrder) {
      return aOrder - bOrder;
    }
    if (aHas && !bHas) {
      return -1;
    }
    if (!aHas && bHas) {
      return 1;
    }
    return String(a.nom_prenom || '').localeCompare(String(b.nom_prenom || ''), 'fr', { sensitivity: 'base' });
  });
}

function buildMonthlyReport(year, month) {
  const workers = sortWorkersByAttachmentOrder(
    getAllWorkers(),
    getAttachmentOrderMapForMonth(year, month)
  );
  if (workers.length === 0) {
    throw new Error('Aucun ouvrier enregistré');
  }

  const rows = workers.map(worker => {
    const stats = calculateMonthlyStats(worker.id, year, month);
    return {
      workerId: worker.id,
      nom_prenom: worker.nom_prenom,
      cin: worker.cin,
      cin_validite: worker.cin_validite,
      type: worker.type,
      date_naissance: worker.date_naissance,
      salaire_journalier: worker.salaire_journalier,
      days_worked: stats.daysWorked,
      amount: stats.totalSalary,
      presenceDays: stats.presenceDays
    };
  });

  const totals = rows.reduce(
    (acc, row) => {
      acc.total_days += row.days_worked;
      acc.total_amount += row.amount;
      return acc;
    },
    { total_days: 0, total_amount: 0 }
  );

  return {
    period: { year, month },
    rows,
    totals
  };
}

function buildQuarterlyReport(year, quarter, rcarAgeLimit = RCAR_AGE_LIMIT) {
  const months = getMonthsInQuarter(quarter);
  if (!months || months.length === 0) {
    throw new Error(`Trimestre invalide: ${quarter}`);
  }

  const workers = sortWorkersByAttachmentOrder(
    getAllWorkers(),
    getAttachmentOrderMapForQuarter(year, quarter)
  );
  if (workers.length === 0) {
    throw new Error('Aucun ouvrier enregistré');
  }

  const rows = workers.map(worker => {
    let totalDays = 0;
    let totalAmount = 0;
    const monthlyStats = months.map(month => {
      const stats = calculateMonthlyStats(worker.id, year, month);
      totalDays += stats.daysWorked;
      totalAmount += stats.totalSalary;
      return {
        month,
        days_worked: stats.daysWorked,
        amount: stats.totalSalary
      };
    });

    return {
      workerId: worker.id,
      nom_prenom: worker.nom_prenom,
      cin: worker.cin,
      cin_validite: worker.cin_validite,
      type: worker.type,
      date_naissance: worker.date_naissance,
      salaire_journalier: worker.salaire_journalier,
      monthlyStats,
      total_days: totalDays,
      total_amount: totalAmount
    };
  });

  const totals = rows.reduce(
    (acc, row) => {
      acc.total_days += row.total_days;
      acc.total_amount += row.total_amount;
      return acc;
    },
    { total_days: 0, total_amount: 0 }
  );

  const lastMonth = months[months.length - 1];
  const quarterEndDate = new Date(year, lastMonth, 0);
  const effectiveAgeLimit = Number.isFinite(Number(rcarAgeLimit)) ? Number(rcarAgeLimit) : RCAR_AGE_LIMIT;
  const eligibleRows = rows.filter((row) => {
    const age = calculateAgeAt(row.date_naissance, quarterEndDate);
    return age === null || age <= effectiveAgeLimit;
  });

  return {
    period: { year, quarter, months, quarterEndDate },
    rows,
    eligibleRows,
    totals
  };
}

/**
 * Generate a monthly document
 * @param {string} documentType - Type of document
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} month - Month
 * @returns {Promise<Object>} Generation result
 */
async function generateMonthlyDocument(documentType, workerId, year, month) {
  const report = buildMonthlyReport(year, month);
  const generatorOptions = buildGeneratorOptions();

  if (documentType === 'depense-regie-salaire') {
    return await roleJournees.generate({ report, year, month, options: generatorOptions });
  }

  const py = await generateWithPython({ documentType, year, month, report, options: generatorOptions });
  return {
    success: true,
    docxFileName: py.docxFileName,
    docxFilePath: py.docxFilePath,
    files: [py.docxFileName],
    documentType
  };
}

/**
 * Generate a quarterly document
 * @param {string} documentType - Type of document
 * @param {number} workerId - Worker ID
 * @param {number} year - Year
 * @param {number} quarter - Quarter (1-4)
 * @returns {Promise<Object>} Generation result
 */
async function generateQuarterlyDocument(documentType, workerId, year, quarter) {
  const appSettings = getAppSettings();
  const report = buildQuarterlyReport(year, quarter, appSettings.rcarAgeLimit);

  const generatorOptions = buildGeneratorOptions();
  const reportForDoc = documentType === 'rcar-salariale'
    ? { ...report, rows: report.eligibleRows || [] }
    : report;

  const py = await generateWithPython({ documentType, year, quarter, report: reportForDoc, options: generatorOptions });
  return {
    success: true,
    docxFileName: py.docxFileName,
    docxFilePath: py.docxFilePath,
    files: [py.docxFileName],
    documentType
  };
}


/**
 * Get all workers with presence for a specific month
 * @param {number} year - Year
 * @param {number} month - Month
 * @returns {Array} Workers with presence data
 */
function getAllWorkersForMonth(year, month) {
  const { getWorkersWithPresence } = require('./presence');
  return getWorkersWithPresence(year, month);
}

/**
 * Get all workers with presence for a specific quarter
 * @param {number} year - Year
 * @param {number} quarter - Quarter
 * @returns {Array} Workers with quarterly presence data
 */
function getAllWorkersForQuarter(year, quarter) {
  return getWorkersWithQuarterlyPresence(year, quarter);
}

/**
 * Generate monthly documents for all workers with presence
 * @param {string} documentType - Type of document
 * @param {number} year - Year
 * @param {number} month - Month
 * @param {string} outputDir - Output directory (optional)
 * @returns {Promise<Object>} Generation results
 */
async function generateMonthlyDocumentBatch(documentType, year, month, outputDir = null, options = {}) {
  const report = buildMonthlyReport(year, month);
  
  const results = {
    total: report.rows.length,
    success: 0,
    failed: 0,
    files: [],
    errors: []
  };

  const generatorOptions = buildGeneratorOptions(options);

  try {
    if (documentType === 'depense-regie-salaire') {
      // Exact Role des Journées (Python: generate_role.py) via JS wrapper
      const originalDir = roleJournees.outputDir;
      if (outputDir) roleJournees.outputDir = outputDir;
      try {
        const result = await roleJournees.generate({ report, year, month, options: generatorOptions });
        if (result && result.docxFileName) results.files.push(result.docxFileName);
      } finally {
        roleJournees.outputDir = originalDir;
      }

      results.success = 1;
      return results;
    }

    const py = await generateWithPython({ documentType, year, month, report, options: generatorOptions, outputDir });
    if (py.docxFileName) results.files.push(py.docxFileName);
    results.success = 1;
  } catch (error) {
    results.failed = 1;
    results.errors.push({ error: error.message });
  }

  return results;
}

/**
 * Generate quarterly documents for all workers with presence
 * @param {string} documentType - Type of document
 * @param {number} year - Year
 * @param {number} quarter - Quarter (1-4)
 * @param {string} outputDir - Output directory (optional)
 * @returns {Promise<Object>} Generation results
 */
async function generateQuarterlyDocumentBatch(documentType, year, quarter, outputDir = null, options = {}) {
  const appSettings = getAppSettings();
  const report = buildQuarterlyReport(year, quarter, appSettings.rcarAgeLimit);
  
  const reportForDoc = documentType === 'rcar-salariale'
    ? { ...report, rows: report.eligibleRows || [] }
    : report;

  const results = {
    total: (reportForDoc.rows || []).length,
    success: 0,
    failed: 0,
    files: [],
    errors: []
  };

  const generatorOptions = buildGeneratorOptions(options);

  try {
    const py = await generateWithPython({ documentType, year, quarter, report: reportForDoc, options: generatorOptions, outputDir });
    if (py.docxFileName) results.files.push(py.docxFileName);
    results.success = 1;
  } catch (error) {
    results.failed = 1;
    results.errors.push({ error: error.message });
  }

  return results;
}

/**
 * Generate combined monthly document for all workers
 * @param {string} documentType - Type of document
 * @param {number} year - Year
 * @param {number} month - Month
 * @param {string} outputDir - Output directory (optional)
 * @returns {Promise<Object>} Generation result
 */
async function generateCombinedMonthlyDocument(documentType, year, month, outputDir = null, options = {}) {
  const report = buildMonthlyReport(year, month);
  const generatorOptions = buildGeneratorOptions(options);

  const computeCombinedNetTotal = () => {
    const ranges = getMonthHalfRanges(year, month);
    const rangeEndDate = new Date(year, month - 1, ranges.combined.endDay);
    const effectiveAgeLimit = Number.isFinite(Number(generatorOptions.rcarAgeLimit))
      ? Number(generatorOptions.rcarAgeLimit)
      : RCAR_AGE_LIMIT;
    const totalCents = report.rows.reduce((sumCents, worker) => {
      const stats = calculatePresenceStatsForRange(
        worker.presenceDays,
        worker.salaire_journalier,
        ranges.combined.startDay,
        ranges.combined.endDay
      );
      const gross = Number(stats.totalSalary || 0);
      if (gross <= 0) return sumCents;

      const grossCents = Math.round(gross * 100);
      const age = calculateAgeAt(worker.date_naissance, rangeEndDate);
      const applyDeduction = age === null || age <= effectiveAgeLimit;
      const deductionCents = applyDeduction ? Math.round(grossCents * RCAR_RATE) : 0;
      const netCents = grossCents - deductionCents;
      return sumCents + netCents;
    }, 0);
    return totalCents / 100;
  };

  if (documentType === 'depense-regie-salaire-combined' || documentType === 'role-journees') {
    const originalDir = roleJournees.outputDir;
    if (outputDir) roleJournees.outputDir = outputDir;
    try {
      return await roleJournees.generate({ report, year, month, options: generatorOptions });
    } finally {
      roleJournees.outputDir = originalDir;
    }
  }

  if (documentType !== 'recu-combined' && documentType !== 'certificat-paiement-combined') {
    throw new Error(`Type de document combiné inconnu: ${documentType}`);
  }

  const py = await generateWithPython({ documentType, year, month, report, options: generatorOptions, outputDir });

  const combinedTotal = documentType === 'recu-combined' ? computeCombinedNetTotal() : null;

  return {
    success: true,
    docxFileName: py.docxFileName,
    docxFilePath: py.docxFilePath,
    documentType,
    totalWorkers: report.rows.length,
    totalNetSalary: combinedTotal == null ? undefined : combinedTotal.toFixed(2),
    totalAmount: combinedTotal == null ? undefined : combinedTotal.toFixed(2),
    files: [py.docxFileName],
    results: [{ success: true, docxFileName: py.docxFileName, docxFilePath: py.docxFilePath }]
  };
}

/**
 * Generate combined quarterly document for all workers
 * @param {string} documentType - Type of document
 * @param {number} year - Year
 * @param {number} quarter - Quarter (1-4)
 * @param {string} outputDir - Output directory (optional)
 * @returns {Promise<Object>} Generation result
 */
async function generateCombinedQuarterlyDocument(documentType, year, quarter, outputDir = null, options = {}) {
  const appSettings = getAppSettings();
  const report = buildQuarterlyReport(year, quarter, appSettings.rcarAgeLimit);

  switch(documentType) {
    case 'rcar-combined':
      return await generateQuarterlyCombinedPair(report, year, quarter, outputDir, options);
    default:
      throw new Error(`Type de document combiné inconnu: ${documentType}`);
  }
}

async function generateQuarterlyCombinedPair(report, year, quarter, outputDir, options = {}) {
  const generatorOptions = buildGeneratorOptions(options);
  const results = [];

  const docs = [
    { id: 'rcar-salariale', report: { ...report, rows: report.eligibleRows || [] } },
    { id: 'rcar-patronale', report }
  ];

  for (const item of docs) {
    const py = await generateWithPython({
      documentType: item.id,
      year,
      quarter,
      report: item.report,
      options: generatorOptions,
      outputDir
    });

    results.push({ success: true, docxFileName: py.docxFileName, docxFilePath: py.docxFilePath });
  }

  return {
    success: true,
    files: results.map(r => r.docxFileName),
    results
  };
}

module.exports = {
  generateMonthlyDocument,
  generateQuarterlyDocument,
  getAllWorkersForMonth,
  getAllWorkersForQuarter,
  generateMonthlyDocumentBatch,
  generateQuarterlyDocumentBatch,
  generateCombinedMonthlyDocument,
  generateCombinedQuarterlyDocument
};
