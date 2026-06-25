#!/usr/bin/env node
'use strict';
require('dotenv').config({ path: require('path').join(__dirname, '..', '.env') });
const Database = require('better-sqlite3');
const path = require('path');
const { VALID_ROLES } = require('../backend/constants');

const DB_PATH = path.join(__dirname, '..', 'auditease.db');

const args = process.argv.slice(2);
function getArg(flag) {
  const idx = args.indexOf(flag);
  return idx !== -1 ? args[idx + 1] : null;
}

const username = getArg('--username');
const role = getArg('--role');

if (!username || !role) {
  console.error('Usage: node scripts/update-role.js --username johndoe --role auditor');
  process.exit(1);
}

if (!VALID_ROLES.includes(role)) {
  console.error(`❌ Error: Invalid role "${role}". Valid roles are: ${VALID_ROLES.join(', ')}`);
  process.exit(1);
}

(async () => {
  try {
    const db = new Database(DB_PATH);
    
    // Check if user exists
    const user = db.prepare('SELECT id, role FROM users WHERE username = ?').get(username);
    if (!user) {
      console.error(`❌ Error: User "${username}" not found.`);
      process.exit(1);
    }

    db.prepare('UPDATE users SET role = ? WHERE id = ?').run(role, user.id);

    console.log(`✅ Role updated successfully for user "${username}".`);
    console.log(`   Old role: ${user.role}`);
    console.log(`   New role: ${role}`);
    db.close();
  } catch (err) {
    console.error('❌ Failed to update role:', err.message);
    process.exit(1);
  }
})();
