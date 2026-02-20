-- Database schema for Gestion des Ouvriers Occasionnels
-- SQLite database structure

-- Workers table
CREATE TABLE IF NOT EXISTS workers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nom_prenom TEXT NOT NULL,
    cin TEXT NOT NULL UNIQUE,
    cin_validite TEXT, -- Format: YYYY-MM-DD
    date_naissance TEXT NOT NULL,  -- Format: YYYY-MM-DD
    type TEXT NOT NULL CHECK(type IN ('OS', 'ONS')),
    salaire_journalier REAL NOT NULL,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

-- Presence table (stores daily presence for each worker)
CREATE TABLE IF NOT EXISTS presence (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    worker_id INTEGER NOT NULL,
    year INTEGER NOT NULL,
    month INTEGER NOT NULL CHECK(month >= 1 AND month <= 12),
    day INTEGER NOT NULL CHECK(day >= 1 AND day <= 31),
    is_present INTEGER NOT NULL DEFAULT 0 CHECK(is_present IN (0, 1)),
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (worker_id) REFERENCES workers(id) ON DELETE CASCADE,
    UNIQUE(worker_id, year, month, day)
);

-- Monthly attachment order used for sorting workers in presence/documents
CREATE TABLE IF NOT EXISTS presence_attachment_orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    worker_id INTEGER NOT NULL,
    year INTEGER NOT NULL,
    month INTEGER NOT NULL CHECK(month >= 1 AND month <= 12),
    attachment_number INTEGER CHECK(attachment_number >= 0),
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (worker_id) REFERENCES workers(id) ON DELETE CASCADE,
    UNIQUE(worker_id, year, month)
);

-- Generated documents tracking (optional, for audit trail)
CREATE TABLE IF NOT EXISTS generated_documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    document_type TEXT NOT NULL,
    worker_id INTEGER,
    year INTEGER NOT NULL,
    month INTEGER,
    quarter INTEGER,
    file_path TEXT NOT NULL,
    generated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (worker_id) REFERENCES workers(id) ON DELETE CASCADE
);

-- Application settings (key/value)
CREATE TABLE IF NOT EXISTS app_settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- Indexes for better query performance
CREATE INDEX IF NOT EXISTS idx_workers_cin ON workers(cin);
CREATE INDEX IF NOT EXISTS idx_workers_type ON workers(type);
CREATE INDEX IF NOT EXISTS idx_presence_worker_date ON presence(worker_id, year, month);
CREATE INDEX IF NOT EXISTS idx_presence_year_month ON presence(year, month);
CREATE INDEX IF NOT EXISTS idx_presence_attach_worker_period ON presence_attachment_orders(worker_id, year, month);
CREATE INDEX IF NOT EXISTS idx_presence_attach_period ON presence_attachment_orders(year, month);
CREATE INDEX IF NOT EXISTS idx_documents_worker ON generated_documents(worker_id);
CREATE INDEX IF NOT EXISTS idx_documents_period ON generated_documents(year, month, quarter);

-- Trigger to update updated_at timestamp on workers
CREATE TRIGGER IF NOT EXISTS update_workers_timestamp 
AFTER UPDATE ON workers
BEGIN
    UPDATE workers SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
END;

-- Trigger to update updated_at timestamp on presence
CREATE TRIGGER IF NOT EXISTS update_presence_timestamp 
AFTER UPDATE ON presence
BEGIN
    UPDATE presence SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
END;

-- Trigger to update updated_at timestamp on attachment order records
CREATE TRIGGER IF NOT EXISTS update_presence_attachment_orders_timestamp 
AFTER UPDATE ON presence_attachment_orders
BEGIN
    UPDATE presence_attachment_orders SET updated_at = CURRENT_TIMESTAMP WHERE id = NEW.id;
END;
