'use strict';
require('dotenv').config();
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

const DB_PATH = path.join(__dirname, '..', 'auditease.db');
const VAULT_PATH = path.join(__dirname, 'storage', 'vault');

// Ensure vault directory exists
if (!fs.existsSync(VAULT_PATH)) {
  fs.mkdirSync(VAULT_PATH, { recursive: true });
}

const db = new Database(DB_PATH);

// Enable WAL mode for better concurrent read performance
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

// --- Schema ---
db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'company'
  );

  CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    version INTEGER NOT NULL DEFAULT 1,
    original_filename TEXT NOT NULL,
    stored_filename TEXT NOT NULL,
    category TEXT NOT NULL,
    subcategory TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Uploaded',
    original_uploader_id INTEGER NOT NULL,
    last_uploader_id INTEGER NOT NULL,
    approver_id INTEGER,
    upload_date TEXT NOT NULL,
    due_date TEXT,
    month TEXT NOT NULL,
    year TEXT NOT NULL,
    is_editable INTEGER NOT NULL DEFAULT 1,
    is_archived INTEGER NOT NULL DEFAULT 0,
    parent_document_id INTEGER,
    FOREIGN KEY (original_uploader_id) REFERENCES users(id),
    FOREIGN KEY (last_uploader_id) REFERENCES users(id),
    FOREIGN KEY (approver_id) REFERENCES users(id),
    FOREIGN KEY (parent_document_id) REFERENCES documents(id)
  );

  CREATE TABLE IF NOT EXISTS user_page_visits (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER NOT NULL,
    page_key TEXT NOT NULL,
    page_label TEXT NOT NULL,
    page_url TEXT NOT NULL,
    visit_count INTEGER NOT NULL DEFAULT 1,
    last_visited TEXT NOT NULL,
    UNIQUE(user_id, page_key),
    FOREIGN KEY (user_id) REFERENCES users(id)
  );

  CREATE TABLE IF NOT EXISTS audit_engagements (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    client_name TEXT NOT NULL,
    financial_year TEXT NOT NULL,
    period_start TEXT NOT NULL,
    period_end TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Active',
    created_by INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS ie_pl_groups (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    engagement_id INTEGER NOT NULL REFERENCES audit_engagements(id) ON DELETE CASCADE,
    ie_pl_type TEXT NOT NULL,
    group_code TEXT NOT NULL,
    group_name TEXT NOT NULL,
    subgroup_code TEXT NOT NULL,
    subgroup_name TEXT NOT NULL,
    display_order INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS trial_balance_ledgers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    engagement_id INTEGER NOT NULL REFERENCES audit_engagements(id) ON DELETE CASCADE,
    ledger_code TEXT NOT NULL,
    ledger_name TEXT NOT NULL,
    opening_balance REAL NOT NULL DEFAULT 0,
    debit_transactions REAL NOT NULL DEFAULT 0,
    credit_transactions REAL NOT NULL DEFAULT 0,
    closing_balance REAL NOT NULL DEFAULT 0,
    ie_pl_group_id INTEGER REFERENCES ie_pl_groups(id),
    is_mapped INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE(engagement_id, ledger_code, ledger_name)
  );

  CREATE TABLE IF NOT EXISTS audit_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    engagement_id INTEGER NOT NULL REFERENCES audit_engagements(id) ON DELETE CASCADE,
    entry_number TEXT NOT NULL,
    entry_type TEXT NOT NULL,
    description TEXT NOT NULL,
    narration TEXT,
    entry_date TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Draft',
    submitted_by INTEGER REFERENCES users(id),
    submitted_at TEXT,
    reviewed_by INTEGER REFERENCES users(id),
    reviewed_at TEXT,
    rejection_reason TEXT,
    created_by INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS audit_entry_lines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    audit_entry_id INTEGER NOT NULL REFERENCES audit_entries(id) ON DELETE CASCADE,
    ledger_id INTEGER NOT NULL REFERENCES trial_balance_ledgers(id),
    line_type TEXT NOT NULL,
    amount REAL NOT NULL,
    line_order INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS report_templates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    description TEXT,
    template_json TEXT NOT NULL,
    created_by INTEGER NOT NULL REFERENCES users(id),
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS generated_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    engagement_id INTEGER NOT NULL REFERENCES audit_engagements(id) ON DELETE CASCADE,
    template_id INTEGER REFERENCES report_templates(id),
    report_type TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'Draft',
    generated_by INTEGER NOT NULL REFERENCES users(id),
    approved_by INTEGER REFERENCES users(id),
    approved_at TEXT,
    stored_filename TEXT,
    report_data_json TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT (datetime('now'))
  );
`);

// Migration: fix trial_balance_ledgers uniqueness constraint
function migrateTrialBalanceUniqueness() {
  const tableInfo = db.prepare(`
    SELECT sql FROM sqlite_master WHERE type='table' AND name='trial_balance_ledgers'
  `).get();

  if (tableInfo && tableInfo.sql.includes('UNIQUE(engagement_id, ledger_code)') 
      && !tableInfo.sql.includes('UNIQUE(engagement_id, ledger_code, ledger_name)')) {
    
    console.log('[Migration] Fixing trial_balance_ledgers uniqueness constraint...');
    
    db.exec(`
      BEGIN TRANSACTION;

      CREATE TABLE trial_balance_ledgers_new (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        engagement_id INTEGER NOT NULL REFERENCES audit_engagements(id) ON DELETE CASCADE,
        ledger_code TEXT NOT NULL,
        ledger_name TEXT NOT NULL,
        opening_balance REAL NOT NULL DEFAULT 0,
        debit_transactions REAL NOT NULL DEFAULT 0,
        credit_transactions REAL NOT NULL DEFAULT 0,
        closing_balance REAL NOT NULL DEFAULT 0,
        ie_pl_group_id INTEGER REFERENCES ie_pl_groups(id),
        is_mapped INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        UNIQUE(engagement_id, ledger_code, ledger_name)
      );

      INSERT INTO trial_balance_ledgers_new 
        SELECT id, engagement_id, ledger_code, ledger_name, opening_balance,
               debit_transactions, credit_transactions, closing_balance,
               ie_pl_group_id, is_mapped, created_at
        FROM trial_balance_ledgers;

      DROP TABLE trial_balance_ledgers;
      ALTER TABLE trial_balance_ledgers_new RENAME TO trial_balance_ledgers;

      COMMIT;
    `);
    
    console.log('[Migration] trial_balance_ledgers uniqueness fixed.');
  }
}

// Migration: add role to users
function migrateUsersAddRole() {
  const tableInfo = db.pragma('table_info(users)');
  const hasRole = tableInfo.some(col => col.name === 'role');
  if (!hasRole) {
    console.log('[Migration] Adding role column to users table...');
    db.exec(`ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'company'`);
    console.log('[Migration] role column added.');
  }
}

migrateTrialBalanceUniqueness();
migrateUsersAddRole();

module.exports = db;
