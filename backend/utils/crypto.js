'use strict';
const crypto = require('crypto');

const ALGO = 'aes-256-cbc';

const KEY_HEX = process.env.ENCRYPTION_KEY;
if (!KEY_HEX || KEY_HEX.length !== 64) {
  throw new Error(`ENCRYPTION_KEY must be a 64-character hex string. Got length: ${KEY_HEX?.length ?? 0}`);
}
const KEY = Buffer.from(KEY_HEX, 'hex');

function getKey() {
  return KEY;
}

/**
 * Encrypts a file buffer using AES-256-CBC.
 * Returns a buffer with the IV prepended (first 16 bytes = IV, rest = ciphertext).
 * @param {Buffer} buffer - Raw file data
 * @returns {Buffer}
 */
function encryptFile(buffer) {
  const KEY = getKey();
  const iv = crypto.randomBytes(16);
  const cipher = crypto.createCipheriv(ALGO, KEY, iv);
  const encrypted = Buffer.concat([cipher.update(buffer), cipher.final()]);
  return Buffer.concat([iv, encrypted]);
}

/**
 * Decrypts an encrypted file buffer.
 * Expects IV prepended as first 16 bytes.
 * @param {Buffer} buffer - Encrypted file data (IV + ciphertext)
 * @returns {Buffer}
 */
function decryptFile(buffer) {
  const KEY = getKey();
  const iv = buffer.slice(0, 16);
  const data = buffer.slice(16);
  const decipher = crypto.createDecipheriv(ALGO, KEY, iv);
  return Buffer.concat([decipher.update(data), decipher.final()]);
}

module.exports = { encryptFile, decryptFile };
