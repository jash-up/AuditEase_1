#!/usr/bin/env node
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const Database = require('better-sqlite3');
const bcrypt = require('bcrypt');
const path = require('path');
const { VALID_ROLES } = require('../backend/constants');

const SALT_ROUNDS = 12;
const DB_PATH = path.join(__dirname, '..', 'auditease.db');

// Parse CLI arguments
const args = process.argv.slice(2);
function getArg(flag) {
  const idx = args.indexOf(flag);
  return idx !== -1 ? args[idx + 1] : null;
}

const name = getArg('--name');
const username = getArg('--username');
const password = getArg('--password');
const roleArg = getArg('--role');

if (!name || !username || !password) {
  console.error('Usage: node scripts/add-user.js --name "John Doe" --username johndoe --password secret123 [--role company|auditor]');
  process.exit(1);
}

const role = roleArg || 'company';
if (!VALID_ROLES.includes(role)) {
  console.error(`❌ Error: Invalid role "${role}". Valid roles are: ${VALID_ROLES.join(', ')}`);
  process.exit(1);
}

(async () => {
  try {
    const db = new Database(DB_PATH);
    db.pragma('foreign_keys = ON');

    // Check if username already exists
    const existing = db.prepare('SELECT id FROM users WHERE username = ?').get(username);
    if (existing) {
      console.error(`❌ Error: Username "${username}" already exists.`);
      process.exit(1);
    }

    const passwordHash = await bcrypt.hash(password, SALT_ROUNDS);
    const result = db.prepare(
      'INSERT INTO users (name, username, password_hash, role) VALUES (?, ?, ?, ?)'
    ).run(name, username, passwordHash, role);

    console.log(`✅ User created successfully!`);
    console.log(`   ID:       ${result.lastInsertRowid}`);
    console.log(`   Name:     ${name}`);
    console.log(`   Username: ${username}`);
    console.log(`   Role:     ${role}`);
    db.close();
  } catch (err) {
    console.error('❌ Failed to create user:', err.message);
    process.exit(1);
  }
})();
