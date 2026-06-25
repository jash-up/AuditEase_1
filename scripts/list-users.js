#!/usr/bin/env node
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const Database = require('better-sqlite3');
const path = require('path');

const DB_PATH = path.join(__dirname, '..', 'auditease.db');

(() => {
  try {
    const db = new Database(DB_PATH);
    const users = db.prepare('SELECT id, name, username, role FROM users ORDER BY id ASC').all();

    if (users.length === 0) {
      console.log('No users found in database.');
      db.close();
      return;
    }

    console.log('\nAuditEase User Directory:');
    console.log('='.repeat(65));
    console.log(String('ID').padEnd(6) + String('Username').padEnd(16) + String('Role').padEnd(12) + 'Full Name');
    console.log('-'.repeat(65));
    for (const u of users) {
      console.log(String(u.id).padEnd(6) + String(u.username).padEnd(16) + String(u.role).padEnd(12) + u.name);
    }
    console.log('='.repeat(65) + '\n');
    db.close();
  } catch (err) {
    console.error('❌ Failed to list users:', err.message);
    process.exit(1);
  }
})();
