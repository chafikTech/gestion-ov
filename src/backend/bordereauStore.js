const { getDatabase } = require('../database/init');

function roundMoney(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return 0;
  }
  return Math.round(numeric * 100) / 100;
}

function normalizeMoney(value, fallback = 0, { allowNegative = false } = {}) {
  if (value === null || value === undefined || value === '') {
    return roundMoney(fallback);
  }

  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return roundMoney(fallback);
  }

  if (!allowNegative && numeric < 0) {
    return 0;
  }

  return roundMoney(numeric);
}

function validateYearMonth(year, month) {
  const y = Number(year);
  const m = Number(month);

  if (!Number.isInteger(y) || y <= 0) {
    throw new Error('Année invalide pour le bordereau');
  }

  if (!Number.isInteger(m) || m < 1 || m > 12) {
    throw new Error('Mois invalide pour le bordereau');
  }
}

function getPreviousTotalGeneral(year, month) {
  validateYearMonth(year, month);
  const db = getDatabase();

  const row = db.prepare(`
    SELECT total_general
    FROM bordereau_monthly_totals
    WHERE (year < ?) OR (year = ? AND month < ?)
    ORDER BY year DESC, month DESC
    LIMIT 1
  `).get(year, year, month);

  return normalizeMoney(row?.total_general, 0, { allowNegative: true });
}

function getMonthlyTotals(year, month) {
  validateYearMonth(year, month);
  const db = getDatabase();

  const row = db.prepare(`
    SELECT year, month, present_amount, admitted_amount, report_previous, rejected_amount, total_general, file_path
    FROM bordereau_monthly_totals
    WHERE year = ? AND month = ?
    LIMIT 1
  `).get(year, month);

  if (!row) {
    return null;
  }

  return {
    year: row.year,
    month: row.month,
    presentAmount: normalizeMoney(row.present_amount, 0, { allowNegative: true }),
    admittedAmount: normalizeMoney(row.admitted_amount, 0, { allowNegative: true }),
    reportPrevious: normalizeMoney(row.report_previous, 0, { allowNegative: true }),
    rejectedAmount: normalizeMoney(row.rejected_amount, 0, { allowNegative: true }),
    totalGeneral: normalizeMoney(row.total_general, 0, { allowNegative: true }),
    filePath: row.file_path || null
  };
}

function upsertMonthlyTotals({
  year,
  month,
  presentAmount,
  admittedAmount,
  reportPrevious,
  rejectedAmount,
  totalGeneral,
  filePath
}) {
  validateYearMonth(year, month);
  const db = getDatabase();

  const normalized = {
    year: Number(year),
    month: Number(month),
    presentAmount: normalizeMoney(presentAmount, 0),
    admittedAmount: normalizeMoney(admittedAmount, 0),
    reportPrevious: normalizeMoney(reportPrevious, 0),
    rejectedAmount: normalizeMoney(rejectedAmount, 0),
    totalGeneral: normalizeMoney(totalGeneral, 0, { allowNegative: true }),
    filePath: filePath ? String(filePath).trim() : null
  };

  db.prepare(`
    INSERT INTO bordereau_monthly_totals (
      year, month, present_amount, admitted_amount, report_previous, rejected_amount, total_general, file_path
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(year, month)
    DO UPDATE SET
      present_amount = excluded.present_amount,
      admitted_amount = excluded.admitted_amount,
      report_previous = excluded.report_previous,
      rejected_amount = excluded.rejected_amount,
      total_general = excluded.total_general,
      file_path = excluded.file_path
  `).run(
    normalized.year,
    normalized.month,
    normalized.presentAmount,
    normalized.admittedAmount,
    normalized.reportPrevious,
    normalized.rejectedAmount,
    normalized.totalGeneral,
    normalized.filePath
  );

  return getMonthlyTotals(normalized.year, normalized.month);
}

module.exports = {
  getPreviousTotalGeneral,
  getMonthlyTotals,
  upsertMonthlyTotals
};
