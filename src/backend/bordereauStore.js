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

function resolveDatabase(dbOverride) {
  return dbOverride || getDatabase();
}

function normalizeScopeValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  const normalized = String(value).trim();
  return normalized.length > 0 ? normalized : null;
}

function normalizeScope(scope = {}) {
  const normalizedCommuneName = normalizeScopeValue(scope.communeName ?? scope.commune_name ?? scope.commune);
  const normalizedCommuneId = normalizeScopeValue(scope.communeId ?? scope.commune_id ?? normalizedCommuneName);
  const normalizedExerciseYear = normalizeScopeValue(scope.exerciseYear ?? scope.exercise_year ?? scope.exercice);

  return {
    communeId: normalizedCommuneId,
    communeName: normalizedCommuneName,
    exerciseYear: normalizedExerciseYear,
    chap: normalizeScopeValue(scope.chap),
    art: normalizeScopeValue(scope.art),
    prog: normalizeScopeValue(scope.prog),
    proj: normalizeScopeValue(scope.proj),
    ligne: normalizeScopeValue(scope.ligne)
  };
}

function resolveScopeAndDb(scopeOrDb, dbOverride = null) {
  if (scopeOrDb && typeof scopeOrDb.prepare === 'function' && !dbOverride) {
    return {
      scope: normalizeScope({}),
      db: resolveDatabase(scopeOrDb)
    };
  }

  return {
    scope: normalizeScope(scopeOrDb || {}),
    db: resolveDatabase(dbOverride)
  };
}

function getPreviousPeriod(year, month) {
  validateYearMonth(year, month);
  const y = Number(year);
  const m = Number(month);
  if (m === 1) {
    return { year: y - 1, month: 12 };
  }
  return { year: y, month: m - 1 };
}

function mapMonthlyTotalsRow(row) {
  if (!row) {
    return null;
  }

  return {
    year: row.year,
    month: row.month,
    communeId: normalizeScopeValue(row.commune_id),
    communeName: normalizeScopeValue(row.commune_name),
    exerciseYear: normalizeScopeValue(row.exercise_year),
    chap: normalizeScopeValue(row.chap),
    art: normalizeScopeValue(row.art),
    prog: normalizeScopeValue(row.prog),
    proj: normalizeScopeValue(row.proj),
    ligne: normalizeScopeValue(row.ligne),
    presentAmount: normalizeMoney(row.present_amount, 0, { allowNegative: true }),
    admittedAmount: normalizeMoney(row.admitted_amount, 0, { allowNegative: true }),
    reportPrevious: normalizeMoney(row.report_previous, 0, { allowNegative: true }),
    rejectedAmount: normalizeMoney(row.rejected_amount, 0, { allowNegative: true }),
    totalGeneral: normalizeMoney(row.total_general, 0, { allowNegative: true }),
    filePath: row.file_path || null
  };
}

function buildPreviousLookupQuery(previousPeriod, scope) {
  const seriesColumns = [
    { field: 'communeId', column: 'commune_id' },
    { field: 'exerciseYear', column: 'exercise_year' },
    { field: 'chap', column: 'chap' },
    { field: 'art', column: 'art' },
    { field: 'prog', column: 'prog' },
    { field: 'proj', column: 'proj' },
    { field: 'ligne', column: 'ligne' }
  ];

  const conditions = ['year = ?', 'month = ?'];
  const params = [previousPeriod.year, previousPeriod.month];

  for (const key of seriesColumns) {
    const value = normalizeScopeValue(scope?.[key.field]);
    if (!value) continue;
    conditions.push(
      `(IFNULL(TRIM(CAST(${key.column} AS TEXT)), '') = '' OR LOWER(TRIM(CAST(${key.column} AS TEXT))) = LOWER(TRIM(?)))`
    );
    params.push(value);
  }

  return {
    sql: `
      SELECT year, month, commune_id, commune_name, exercise_year, chap, art, prog, proj, ligne, present_amount, admitted_amount, report_previous, rejected_amount, total_general, file_path
      FROM bordereau_monthly_totals
      WHERE ${conditions.join('\n        AND ')}
      LIMIT 1
    `,
    params
  };
}

function getPreviousMonthBordereau(year, month, scopeOrDb = null, dbOverride = null, options = {}) {
  validateYearMonth(year, month);
  const { scope, db } = resolveScopeAndDb(scopeOrDb, dbOverride);
  const previous = getPreviousPeriod(year, month);

  if (previous.year <= 0) {
    return null;
  }

  const query = buildPreviousLookupQuery(previous, scope);
  if (typeof options?.onDebug === 'function') {
    options.onDebug({
      currentPeriod: { year: Number(year), month: Number(month) },
      previousPeriod: previous,
      keys: scope,
      sql: query.sql,
      params: query.params
    });
  }
  const row = db.prepare(query.sql).get(...query.params);

  return mapMonthlyTotalsRow(row);
}

function computeCarryOver(previousBordereau) {
  if (!previousBordereau || typeof previousBordereau !== 'object') {
    return 0;
  }

  return normalizeMoney(
    previousBordereau.totalGeneral ?? previousBordereau.total_general,
    0,
    { allowNegative: true }
  );
}

function computeCumulativeTotalGeneral(presentTotal, previousBordereau) {
  const presentAmount = normalizeMoney(presentTotal, 0, { allowNegative: true });
  const previousTotalGeneral = computeCarryOver(previousBordereau);
  return normalizeMoney(previousTotalGeneral + presentAmount, 0, { allowNegative: true });
}

function getPreviousTotalGeneral(year, month, scopeOrDb = null, dbOverride = null) {
  const previousMonthRecord = getPreviousMonthBordereau(year, month, scopeOrDb, dbOverride);
  return computeCarryOver(previousMonthRecord);
}

function getMonthlyTotals(year, month, dbOverride = null) {
  validateYearMonth(year, month);
  const db = resolveDatabase(dbOverride);

  const row = db.prepare(`
    SELECT year, month, commune_id, commune_name, exercise_year, chap, art, prog, proj, ligne, present_amount, admitted_amount, report_previous, rejected_amount, total_general, file_path
    FROM bordereau_monthly_totals
    WHERE year = ? AND month = ?
    LIMIT 1
  `).get(year, month);
  return mapMonthlyTotalsRow(row);
}

function upsertMonthlyTotals({
  year,
  month,
  scope,
  presentAmount,
  admittedAmount,
  reportPrevious,
  rejectedAmount,
  totalGeneral,
  filePath
}, dbOverride = null) {
  validateYearMonth(year, month);
  const db = resolveDatabase(dbOverride);
  const normalizedScope = normalizeScope(scope || {});

  const normalized = {
    year: Number(year),
    month: Number(month),
    communeId: normalizedScope.communeId,
    communeName: normalizedScope.communeName,
    exerciseYear: normalizedScope.exerciseYear,
    chap: normalizedScope.chap,
    art: normalizedScope.art,
    prog: normalizedScope.prog,
    proj: normalizedScope.proj,
    ligne: normalizedScope.ligne,
    presentAmount: normalizeMoney(presentAmount, 0),
    admittedAmount: normalizeMoney(admittedAmount, 0),
    reportPrevious: normalizeMoney(reportPrevious, 0),
    rejectedAmount: normalizeMoney(rejectedAmount, 0),
    totalGeneral: normalizeMoney(totalGeneral, 0, { allowNegative: true }),
    filePath: filePath ? String(filePath).trim() : null
  };

  db.prepare(`
    INSERT INTO bordereau_monthly_totals (
      year, month, commune_id, commune_name, exercise_year, chap, art, prog, proj, ligne, present_amount, admitted_amount, report_previous, rejected_amount, total_general, file_path
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(year, month)
    DO UPDATE SET
      commune_id = excluded.commune_id,
      commune_name = excluded.commune_name,
      exercise_year = excluded.exercise_year,
      chap = excluded.chap,
      art = excluded.art,
      prog = excluded.prog,
      proj = excluded.proj,
      ligne = excluded.ligne,
      present_amount = excluded.present_amount,
      admitted_amount = excluded.admitted_amount,
      report_previous = excluded.report_previous,
      rejected_amount = excluded.rejected_amount,
      total_general = excluded.total_general,
      file_path = excluded.file_path
  `).run(
    normalized.year,
    normalized.month,
    normalized.communeId,
    normalized.communeName,
    normalized.exerciseYear,
    normalized.chap,
    normalized.art,
    normalized.prog,
    normalized.proj,
    normalized.ligne,
    normalized.presentAmount,
    normalized.admittedAmount,
    normalized.reportPrevious,
    normalized.rejectedAmount,
    normalized.totalGeneral,
    normalized.filePath
  );

  return getMonthlyTotals(normalized.year, normalized.month, db);
}

module.exports = {
  getPreviousPeriod,
  getPreviousMonthBordereau,
  computeCarryOver,
  computeCumulativeTotalGeneral,
  getPreviousTotalGeneral,
  getMonthlyTotals,
  upsertMonthlyTotals
};
