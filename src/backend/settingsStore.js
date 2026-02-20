const { getDatabase } = require('../database/init');
const config = require('./config');
const { REFERENCE_VALUES, RCAR_AGE_LIMIT } = require('./constants');

const DB_KEYS = {
  rcarAdhesionNumber: 'rcar_adhesion_number',
  chap: 'ref_chap',
  art: 'ref_art',
  prog: 'ref_prog',
  proj: 'ref_proj',
  ligne: 'ref_ligne',
  rcarAgeLimit: 'rcar_age_limit',
  decisionNumber: 'decision_number',
  decisionDate: 'decision_date'
};

function getDefaultSettings() {
  return {
    rcarAdhesionNumber: (config?.rcar?.adhesionNumber || '').toString().trim(),
    chap: (REFERENCE_VALUES?.chapitre || '').toString().trim(),
    art: (REFERENCE_VALUES?.article || '').toString().trim(),
    prog: (REFERENCE_VALUES?.programme || '').toString().trim(),
    proj: (REFERENCE_VALUES?.projet || '').toString().trim(),
    ligne: (REFERENCE_VALUES?.ligne || '').toString().trim(),
    rcarAgeLimit: RCAR_AGE_LIMIT,
    decisionNumber: '02/2024',
    decisionDate: '02/02/2024'
  };
}

function normalizeOptionalString(value) {
  if (value === undefined) return undefined;
  if (value === null) return null;
  const str = String(value).trim();
  return str.length === 0 ? null : str;
}

function normalizeOptionalInt(value, { min, max } = {}) {
  if (value === undefined) return undefined;
  if (value === null) return null;
  if (value === '') return null;
  const parsed = Number.parseInt(String(value).trim(), 10);
  if (!Number.isFinite(parsed)) {
    throw new Error('Valeur numérique invalide');
  }
  if (typeof min === 'number' && parsed < min) {
    throw new Error(`La valeur doit être ≥ ${min}`);
  }
  if (typeof max === 'number' && parsed > max) {
    throw new Error(`La valeur doit être ≤ ${max}`);
  }
  return parsed;
}

function getAppSettings() {
  const db = getDatabase();
  const defaults = getDefaultSettings();
  const settings = { ...defaults };

  const rows = db.prepare('SELECT key, value FROM app_settings').all();
  for (const row of rows) {
    switch (row.key) {
      case DB_KEYS.rcarAdhesionNumber:
        settings.rcarAdhesionNumber = String(row.value || '').trim();
        break;
      case DB_KEYS.chap:
        settings.chap = String(row.value || '').trim();
        break;
      case DB_KEYS.art:
        settings.art = String(row.value || '').trim();
        break;
      case DB_KEYS.prog:
        settings.prog = String(row.value || '').trim();
        break;
      case DB_KEYS.proj:
        settings.proj = String(row.value || '').trim();
        break;
      case DB_KEYS.ligne:
        settings.ligne = String(row.value || '').trim();
        break;
      case DB_KEYS.rcarAgeLimit: {
        const parsed = Number.parseInt(String(row.value || '').trim(), 10);
        if (Number.isFinite(parsed)) {
          settings.rcarAgeLimit = parsed;
        }
        break;
      }
      case DB_KEYS.decisionNumber:
        settings.decisionNumber = String(row.value || '').trim();
        break;
      case DB_KEYS.decisionDate:
        settings.decisionDate = String(row.value || '').trim();
        break;
      default:
        break;
    }
  }

  return settings;
}

function saveAppSettings(input = {}) {
  if (!input || typeof input !== 'object') {
    throw new Error('Paramètres invalides');
  }

  const db = getDatabase();

  const normalized = {
    rcarAdhesionNumber: normalizeOptionalString(input.rcarAdhesionNumber),
    chap: normalizeOptionalString(input.chap),
    art: normalizeOptionalString(input.art),
    prog: normalizeOptionalString(input.prog),
    proj: normalizeOptionalString(input.proj),
    ligne: normalizeOptionalString(input.ligne),
    rcarAgeLimit: normalizeOptionalInt(input.rcarAgeLimit, { min: 0, max: 120 }),
    decisionNumber: normalizeOptionalString(input.decisionNumber),
    decisionDate: normalizeOptionalString(input.decisionDate)
  };

  const upserts = [];
  const deletions = [];

  for (const [field, dbKey] of Object.entries(DB_KEYS)) {
    const value = normalized[field];
    if (value === undefined) continue;
    if (value === null) {
      deletions.push(dbKey);
      continue;
    }
    upserts.push([dbKey, String(value)]);
  }

  const upsertStmt = db.prepare('INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)');
  const deleteStmt = db.prepare('DELETE FROM app_settings WHERE key = ?');

  const tx = db.transaction(() => {
    for (const [key, value] of upserts) {
      upsertStmt.run(key, value);
    }
    for (const key of deletions) {
      deleteStmt.run(key);
    }
  });

  tx();
  return getAppSettings();
}

module.exports = {
  getAppSettings,
  saveAppSettings
};
