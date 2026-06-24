'use strict';
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// The AES algorithm used in your app
const ALGO = 'aes-256-cbc';

const oldKeyHex = process.env.OLD_ENCRYPTION_KEY;
const newKeyHex = process.env.NEW_ENCRYPTION_KEY;

if (!oldKeyHex || oldKeyHex.length !== 64) {
  console.error("❌ Please provide the old key in OLD_ENCRYPTION_KEY (must be 64-character hex)");
  process.exit(1);
}
if (!newKeyHex || newKeyHex.length !== 64) {
  console.error("❌ Please provide the new key in NEW_ENCRYPTION_KEY (must be 64-character hex)");
  process.exit(1);
}

const OLD_KEY = Buffer.from(oldKeyHex, 'hex');
const NEW_KEY = Buffer.from(newKeyHex, 'hex');
const VAULT_DIR = path.join(__dirname, '..', 'storage', 'vault');

function decrypt(buffer, key) {
  const iv = buffer.slice(0, 16);
  const data = buffer.slice(16);
  const decipher = crypto.createDecipheriv(ALGO, key, iv);
  return Buffer.concat([decipher.update(data), decipher.final()]);
}

function encrypt(buffer, key) {
  const iv = crypto.randomBytes(16);
  const cipher = crypto.createCipheriv(ALGO, key, iv);
  const encrypted = Buffer.concat([cipher.update(buffer), cipher.final()]);
  return Buffer.concat([iv, encrypted]);
}

if (!fs.existsSync(VAULT_DIR)) {
  console.log("No vault directory found at:", VAULT_DIR);
  process.exit(0);
}

const files = fs.readdirSync(VAULT_DIR);
console.log(`Starting re-encryption of ${files.length} files in the vault...`);

let success = 0;
let failed = 0;

for (const file of files) {
  const filePath = path.join(VAULT_DIR, file);
  try {
    const encryptedData = fs.readFileSync(filePath);
    
    // 1. Decrypt with OLD key
    const rawData = decrypt(encryptedData, OLD_KEY);
    
    // 2. Encrypt with NEW key
    const newEncryptedData = encrypt(rawData, NEW_KEY);
    
    // 3. Save it back
    fs.writeFileSync(filePath, newEncryptedData);
    success++;
    console.log(`✅ Rekeyed: ${file}`);
  } catch (err) {
    failed++;
    console.error(`❌ Failed to rekey ${file}:`, err.message);
  }
}

console.log(`\n🎉 Done! Successfully rekeyed ${success} files. Failed: ${failed}.`);
console.log("You can now update your main .env ENCRYPTION_KEY to match the NEW_ENCRYPTION_KEY.");
