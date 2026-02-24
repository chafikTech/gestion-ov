const test = require('node:test');
const assert = require('node:assert/strict');

const {
  getPreviousPeriod,
  getPreviousMonthBordereau,
  computeCarryOver,
  computeCumulativeTotalGeneral,
  getPreviousTotalGeneral
} = require('../../src/backend/bordereauStore');

const DEFAULT_SCOPE = {
  communeId: 'OULED NACEUR',
  communeName: 'OULED NACEUR',
  chap: '10',
  art: '20',
  prog: '20',
  proj: '10',
  ligne: '14'
};

function normalizeScopeValue(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function buildScopeKey(scope = {}) {
  return [
    normalizeScopeValue(scope.communeId),
    normalizeScopeValue(scope.communeName),
    normalizeScopeValue(scope.exerciseYear),
    normalizeScopeValue(scope.chap),
    normalizeScopeValue(scope.art),
    normalizeScopeValue(scope.prog),
    normalizeScopeValue(scope.proj),
    normalizeScopeValue(scope.ligne)
  ].join('|');
}

function buildStoreKey(year, month, scope = {}) {
  return `${year}-${month}-${buildScopeKey(scope)}`;
}

function createInMemoryDb() {
  const store = new Map();
  return {
    __store: store,
    prepare(sql) {
      if (sql.includes('FROM bordereau_monthly_totals')) {
        return {
          get(year, month, ...scopeValues) {
            const matchingRows = Array.from(store.values()).filter((row) => row.year === Number(year) && row.month === Number(month));
            if (matchingRows.length === 0) return undefined;
            if (scopeValues.length === 0) return matchingRows[0];

            const scopeKeys = [];
            if (sql.includes('CAST(commune_id AS TEXT)')) scopeKeys.push('commune_id');
            if (sql.includes('CAST(exercise_year AS TEXT)')) scopeKeys.push('exercise_year');
            if (sql.includes('CAST(chap AS TEXT)')) scopeKeys.push('chap');
            if (sql.includes('CAST(art AS TEXT)')) scopeKeys.push('art');
            if (sql.includes('CAST(prog AS TEXT)')) scopeKeys.push('prog');
            if (sql.includes('CAST(proj AS TEXT)')) scopeKeys.push('proj');
            if (sql.includes('CAST(ligne AS TEXT)')) scopeKeys.push('ligne');

            return matchingRows.find((row) => {
              return scopeKeys.every((columnName, idx) => {
                const incoming = normalizeScopeValue(scopeValues[idx]);
                if (!incoming) return true;
                const rowValue = normalizeScopeValue(row[columnName]);
                if (!rowValue) return true;
                return rowValue.toLowerCase() === incoming.toLowerCase();
              });
            });
          }
        };
      }
      throw new Error(`Unsupported SQL in test mock: ${sql}`);
    }
  };
}

function insertMonthlyTotal(db, { year, month, totalGeneral, scope = DEFAULT_SCOPE }) {
  db.__store.set(buildStoreKey(year, month, scope), {
    year,
    month,
    commune_id: scope.communeId,
    commune_name: scope.communeName,
    exercise_year: scope.exerciseYear,
    chap: scope.chap,
    art: scope.art,
    prog: scope.prog,
    proj: scope.proj,
    ligne: scope.ligne,
    present_amount: 0,
    admitted_amount: 0,
    report_previous: 0,
    rejected_amount: 0,
    total_general: totalGeneral,
    file_path: null
  });
}

test('first month with no previous bordereau returns carry-over 0.00', () => {
  const db = createInMemoryDb();
  const previous = getPreviousMonthBordereau(2026, 3, DEFAULT_SCOPE, db);
  assert.equal(previous, null);
  assert.equal(computeCarryOver(previous), 0);
  assert.equal(getPreviousTotalGeneral(2026, 3, DEFAULT_SCOPE, db), 0);
});

test('normal month carry-over equals previous month total général', () => {
  const db = createInMemoryDb();
  insertMonthlyTotal(db, { year: 2026, month: 3, totalGeneral: 1000.0, scope: DEFAULT_SCOPE });

  const previous = getPreviousMonthBordereau(2026, 4, DEFAULT_SCOPE, db);
  assert.ok(previous);
  assert.equal(previous.year, 2026);
  assert.equal(previous.month, 3);
  assert.equal(previous.totalGeneral, 1000.0);
  assert.equal(computeCarryOver(previous), 1000.0);
  assert.equal(getPreviousTotalGeneral(2026, 4, DEFAULT_SCOPE, db), 1000.0);
});

test('January reads carry-over from previous year December', () => {
  const db = createInMemoryDb();
  insertMonthlyTotal(db, { year: 2025, month: 12, totalGeneral: 1234.56, scope: DEFAULT_SCOPE });

  const previous = getPreviousMonthBordereau(2026, 1, DEFAULT_SCOPE, db);
  assert.ok(previous);
  assert.equal(previous.year, 2025);
  assert.equal(previous.month, 12);
  assert.equal(previous.totalGeneral, 1234.56);
  assert.equal(computeCarryOver(previous), 1234.56);
  assert.equal(getPreviousTotalGeneral(2026, 1, DEFAULT_SCOPE, db), 1234.56);
});

test('getPreviousPeriod returns strict previous calendar month', () => {
  assert.deepEqual(getPreviousPeriod(2026, 6), { year: 2026, month: 5 });
  assert.deepEqual(getPreviousPeriod(2026, 1), { year: 2025, month: 12 });
});

test('cumulative total général across three consecutive months (June -> July -> August)', () => {
  const db = createInMemoryDb();

  // June: no previous month record in scope
  const junePrevious = getPreviousMonthBordereau(2026, 6, DEFAULT_SCOPE, db);
  const juneTotalGeneral = computeCumulativeTotalGeneral(1000, junePrevious);
  assert.equal(juneTotalGeneral, 1000);
  insertMonthlyTotal(db, { year: 2026, month: 6, totalGeneral: juneTotalGeneral, scope: DEFAULT_SCOPE });

  // July: previous total général (June) + present (2000) = 3000
  const julyPrevious = getPreviousMonthBordereau(2026, 7, DEFAULT_SCOPE, db);
  const julyTotalGeneral = computeCumulativeTotalGeneral(2000, julyPrevious);
  assert.equal(julyTotalGeneral, 3000);
  insertMonthlyTotal(db, { year: 2026, month: 7, totalGeneral: julyTotalGeneral, scope: DEFAULT_SCOPE });

  // August: previous total général (July) + present (500) = 3500
  const augustPrevious = getPreviousMonthBordereau(2026, 8, DEFAULT_SCOPE, db);
  const augustTotalGeneral = computeCumulativeTotalGeneral(500, augustPrevious);
  assert.equal(augustTotalGeneral, 3500);
});

test('cumulative Total Général for January, February, March 2026', () => {
  const db = createInMemoryDb();

  // January 2026 presentTotal = 3000 -> Total Général = 3000
  const januaryPrevious = getPreviousMonthBordereau(2026, 1, DEFAULT_SCOPE, db);
  const januaryTotalGeneral = computeCumulativeTotalGeneral(3000, januaryPrevious);
  assert.equal(januaryTotalGeneral, 3000);
  insertMonthlyTotal(db, { year: 2026, month: 1, totalGeneral: januaryTotalGeneral, scope: DEFAULT_SCOPE });

  // February 2026 presentTotal = 2000 -> Report = 3000 -> Total Général = 5000
  const februaryPrevious = getPreviousMonthBordereau(2026, 2, DEFAULT_SCOPE, db);
  assert.ok(februaryPrevious);
  assert.equal(computeCarryOver(februaryPrevious), 3000);
  const februaryTotalGeneral = computeCumulativeTotalGeneral(2000, februaryPrevious);
  assert.equal(februaryTotalGeneral, 5000);
  insertMonthlyTotal(db, { year: 2026, month: 2, totalGeneral: februaryTotalGeneral, scope: DEFAULT_SCOPE });

  // March 2026 presentTotal = 1000 -> Report = 5000 -> Total Général = 6000
  const marchPrevious = getPreviousMonthBordereau(2026, 3, DEFAULT_SCOPE, db);
  assert.ok(marchPrevious);
  assert.equal(computeCarryOver(marchPrevious), 5000);
  const marchTotalGeneral = computeCumulativeTotalGeneral(1000, marchPrevious);
  assert.equal(marchTotalGeneral, 6000);
});
