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
    password_hash TEXT NOT NULL
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
`);

module.exports = db;
