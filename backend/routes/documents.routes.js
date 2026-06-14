'use strict';
require('dotenv').config();
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const db = require('../db');
const { requireAuth } = require('../auth');
const { encryptFile, decryptFile } = require('../utils/crypto');

const router = express.Router();
const VAULT_PATH = path.join(__dirname, '..', 'storage', 'vault');

// Multer: store files in memory for encryption before writing to disk
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB limit
});

// Valid categories/subcategories
const CATEGORIES = {
  'Indirect Tax': ['GST Returns', 'VAT Filings', 'Customs Duty', 'Service Tax', 'Excise'],
  'Employees': ['Payroll', 'PF/ESI', 'TDS on Salary', 'Leave Records', 'Contracts'],
  'Income Tax': ['ITR Filing', 'Advance Tax', 'TDS Returns', 'Form 16', 'Tax Audit Report']
};

const VALID_STATUSES = ['Uploaded', 'Pending Approval', 'Action Required', 'Verified', 'Submitted', 'Overdue'];

// ── GET /api/documents/summary ──────────────────────────────────────────────
router.get('/summary', requireAuth, (req, res) => {
  try {
    // Get latest version per chain (group by root doc id)
    const rows = db.prepare(`
      WITH latest AS (
        SELECT COALESCE(parent_document_id, id) AS root_id, MAX(id) AS latest_id
        FROM documents
        WHERE is_archived = 0
        GROUP BY COALESCE(parent_document_id, id)
      )
      SELECT d.category, d.status, COUNT(*) as count
      FROM documents d
      INNER JOIN latest l ON d.id = l.latest_id
      GROUP BY d.category, d.status
    `).all();

    // Structure: { "Indirect Tax": { "Uploaded": 3, ... }, ... }
    const summary = {};
    for (const cat of Object.keys(CATEGORIES)) {
      summary[cat] = {};
      for (const s of VALID_STATUSES) {
        summary[cat][s] = 0;
      }
    }
    for (const row of rows) {
      if (summary[row.category]) {
        summary[row.category][row.status] = row.count;
      }
    }
    res.json(summary);
  } catch (err) {
    console.error('Summary error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── GET /api/documents ──────────────────────────────────────────────────────
router.get('/', requireAuth, (req, res) => {
  try {
    const {
      category, subcategory, status, archived, search,
      month, year, uploader_id, latest
    } = req.query;

    let query = `
      SELECT d.*,
             u1.name AS original_uploader_name,
             u2.name AS last_uploader_name,
             u3.name AS approver_name
      FROM documents d
      LEFT JOIN users u1 ON d.original_uploader_id = u1.id
      LEFT JOIN users u2 ON d.last_uploader_id = u2.id
      LEFT JOIN users u3 ON d.approver_id = u3.id
      WHERE 1=1
    `;
    const params = [];

    if (archived !== undefined) {
      query += ' AND d.is_archived = ?';
      params.push(archived === 'true' ? 1 : 0);
    }
    if (category) { query += ' AND d.category = ?'; params.push(category); }
    if (subcategory) { query += ' AND d.subcategory = ?'; params.push(subcategory); }
    if (status) { query += ' AND d.status = ?'; params.push(status); }
    if (month) { query += ' AND d.month = ?'; params.push(month); }
    if (year) { query += ' AND d.year = ?'; params.push(year); }
    if (uploader_id) { query += ' AND d.last_uploader_id = ?'; params.push(uploader_id); }
    if (search) { query += ' AND d.name LIKE ?'; params.push(`%${search}%`); }

    if (latest === 'true') {
      // Only return the latest version per chain
      query += `
        AND d.id IN (
          SELECT MAX(id) FROM documents
          GROUP BY COALESCE(parent_document_id, id)
        )
      `;
    }

    query += ' ORDER BY d.upload_date DESC, d.id DESC';

    const rows = db.prepare(query).all(...params);
    res.json(rows);
  } catch (err) {
    console.error('Documents list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── GET /api/documents/:id ──────────────────────────────────────────────────
router.get('/:id', requireAuth, (req, res) => {
  try {
    const doc = db.prepare(`
      SELECT d.*,
             u1.name AS original_uploader_name,
             u2.name AS last_uploader_name,
             u3.name AS approver_name
      FROM documents d
      LEFT JOIN users u1 ON d.original_uploader_id = u1.id
      LEFT JOIN users u2 ON d.last_uploader_id = u2.id
      LEFT JOIN users u3 ON d.approver_id = u3.id
      WHERE d.id = ?
    `).get(req.params.id);

    if (!doc) return res.status(404).json({ error: 'Document not found' });
    res.json(doc);
  } catch (err) {
    console.error('Document get error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── GET /api/documents/:id/versions ─────────────────────────────────────────
router.get('/:id/versions', requireAuth, (req, res) => {
  try {
    const doc = db.prepare('SELECT * FROM documents WHERE id = ?').get(req.params.id);
    if (!doc) return res.status(404).json({ error: 'Document not found' });

    // Find root document
    const rootId = doc.parent_document_id || doc.id;

    const versions = db.prepare(`
      SELECT d.*,
             u.name AS last_uploader_name
      FROM documents d
      LEFT JOIN users u ON d.last_uploader_id = u.id
      WHERE d.id = ? OR d.parent_document_id = ?
      ORDER BY d.version DESC
    `).all(rootId, rootId);

    res.json(versions);
  } catch (err) {
    console.error('Versions error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── POST /api/documents/upload ───────────────────────────────────────────────
router.post('/upload', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file received. Check that a file was selected and the form field is named "file".' });
    }

    const {
      name, category, subcategory, status, month, year,
      due_date, approver_id, is_editable
    } = req.body;

    // Validate required fields
    if (!name || !category || !subcategory || !status || !month || !year) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    // Validate category/subcategory
    if (!CATEGORIES[category]) return res.status(400).json({ error: 'Invalid category' });
    if (!CATEGORIES[category].includes(subcategory)) {
      return res.status(400).json({ error: 'Invalid subcategory for this category' });
    }
    if (!VALID_STATUSES.includes(status)) {
      return res.status(400).json({ error: 'Invalid status' });
    }

    const version = 1;
    const uuid = crypto.randomUUID();
    const safeSubcat = subcategory.replace(/[^a-zA-Z0-9_-]/g, '_');
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const mon = String(now.getMonth() + 1).padStart(2, '0');
    const yr = now.getFullYear();
    const storedFilename = `${safeSubcat}_${version}_${day}_${mon}_${yr}_${uuid}.enc`;

    // Encrypt and write file
    const encrypted = encryptFile(req.file.buffer);
    fs.writeFileSync(path.join(VAULT_PATH, storedFilename), encrypted);

    const uploadDate = now.toISOString();
    const stmt = db.prepare(`
      INSERT INTO documents
        (name, version, original_filename, stored_filename, category, subcategory,
         status, original_uploader_id, last_uploader_id, approver_id, upload_date,
         due_date, month, year, is_editable, is_archived, parent_document_id)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, NULL)
    `);
    const result = stmt.run(
      name,
      version,
      req.file.originalname,
      storedFilename,
      category,
      subcategory,
      status,
      req.user.id,
      req.user.id,
      approver_id || null,
      uploadDate,
      due_date || null,
      month,
      year,
      is_editable === 'false' || is_editable === '0' ? 0 : 1
    );

    const newDoc = db.prepare('SELECT * FROM documents WHERE id = ?').get(result.lastInsertRowid);
    res.status(201).json(newDoc);
  } catch (err) {
    console.error('[UPLOAD ERROR]', err); // log full stack
    return res.status(500).json({
      error: 'Upload failed',
      detail: err.message,       // send message to client
      stack: err.stack           // remove this in production
    });
  }
});

// ── POST /api/documents/:id/edit (new version) ──────────────────────────────
router.post('/:id/edit', requireAuth, upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file received. Check that a file was selected and the form field is named "file".' });
    }

    const parentDoc = db.prepare('SELECT * FROM documents WHERE id = ?').get(req.params.id);
    if (!parentDoc) return res.status(404).json({ error: 'Document not found' });
    if (parentDoc.is_archived) return res.status(400).json({ error: 'Cannot edit archived document' });
    if (!parentDoc.is_editable) return res.status(400).json({ error: 'Document is locked' });

    // Find root (for parent_document_id chain)
    const rootId = parentDoc.parent_document_id || parentDoc.id;

    // Get current max version in the chain
    const maxVersionRow = db.prepare(`
      SELECT MAX(version) as maxv FROM documents
      WHERE id = ? OR parent_document_id = ?
    `).get(rootId, rootId);
    const newVersion = (maxVersionRow.maxv || 1) + 1;

    const { status, due_date, approver_id, is_editable } = req.body;
    const effectiveStatus = status && VALID_STATUSES.includes(status) ? status : parentDoc.status;

    const uuid = crypto.randomUUID();
    const safeSubcat = parentDoc.subcategory.replace(/[^a-zA-Z0-9_-]/g, '_');
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const mon = String(now.getMonth() + 1).padStart(2, '0');
    const yr = now.getFullYear();
    const storedFilename = `${safeSubcat}_${newVersion}_${day}_${mon}_${yr}_${uuid}.enc`;

    // Encrypt and write file
    const encrypted = encryptFile(req.file.buffer);
    fs.writeFileSync(path.join(VAULT_PATH, storedFilename), encrypted);

    const uploadDate = now.toISOString();
    const stmt = db.prepare(`
      INSERT INTO documents
        (name, version, original_filename, stored_filename, category, subcategory,
         status, original_uploader_id, last_uploader_id, approver_id, upload_date,
         due_date, month, year, is_editable, is_archived, parent_document_id)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
    `);
    const result = stmt.run(
      parentDoc.name,
      newVersion,
      req.file.originalname,
      storedFilename,
      parentDoc.category,
      parentDoc.subcategory,
      effectiveStatus,
      parentDoc.original_uploader_id,
      req.user.id,
      approver_id !== undefined ? (approver_id || null) : parentDoc.approver_id,
      uploadDate,
      due_date !== undefined ? (due_date || null) : parentDoc.due_date,
      parentDoc.month,
      parentDoc.year,
      is_editable !== undefined ? (is_editable === 'false' || is_editable === '0' ? 0 : 1) : parentDoc.is_editable,
      rootId
    );

    const newDoc = db.prepare('SELECT * FROM documents WHERE id = ?').get(result.lastInsertRowid);
    res.status(201).json(newDoc);
  } catch (err) {
    console.error('[EDIT UPLOAD ERROR]', err); // log full stack
    return res.status(500).json({
      error: 'Upload failed',
      detail: err.message,       // send message to client
      stack: err.stack           // remove this in production
    });
  }
});

// ── PATCH /api/documents/:id/status ─────────────────────────────────────────
router.patch('/:id/status', requireAuth, (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;

    const validStatuses = ['Uploaded', 'Pending Approval', 'Action Required', 'Verified', 'Submitted', 'Overdue'];

    if (!status) {
      return res.status(400).json({ error: 'Status is required.' });
    }

    if (!validStatuses.includes(status)) {
      return res.status(400).json({ error: `Invalid status. Must be one of: ${validStatuses.join(', ')}` });
    }

    const doc = db.prepare('SELECT * FROM documents WHERE id = ?').get(id);
    if (!doc) {
      return res.status(404).json({ error: 'Document not found.' });
    }

    if (doc.is_archived) {
      return res.status(400).json({ error: 'Cannot change status of an archived document. Unarchive it first.' });
    }

    db.prepare('UPDATE documents SET status = ?, last_uploader_id = ? WHERE id = ?')
      .run(status, req.user.id, id);

    const updated = db.prepare('SELECT * FROM documents WHERE id = ?').get(id);
    return res.json({ message: 'Status updated successfully.', document: updated });

  } catch (err) {
    console.error('[STATUS UPDATE ERROR]', err);
    return res.status(500).json({ error: 'Failed to update status.', detail: err.message });
  }
});

// ── PATCH /api/documents/:id/archive ────────────────────────────────────────
router.patch('/:id/archive', requireAuth, (req, res) => {
  try {
    const doc = db.prepare('SELECT * FROM documents WHERE id = ?').get(req.params.id);
    if (!doc) return res.status(404).json({ error: 'Document not found' });

    const newArchived = doc.is_archived ? 0 : 1;

    // Archive/unarchive all versions in the chain
    const rootId = doc.parent_document_id || doc.id;
    db.prepare(`
      UPDATE documents SET is_archived = ?
      WHERE id = ? OR parent_document_id = ?
    `).run(newArchived, rootId, rootId);

    res.json({ archived: !!newArchived });
  } catch (err) {
    console.error('Archive error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── GET /api/documents/:id/download ─────────────────────────────────────────
router.get('/:id/download', requireAuth, (req, res) => {
  try {
    const doc = db.prepare('SELECT * FROM documents WHERE id = ?').get(req.params.id);
    if (!doc) return res.status(404).json({ error: 'Document not found' });

    const filePath = path.join(VAULT_PATH, doc.stored_filename);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'File not found on disk' });
    }

    const encrypted = fs.readFileSync(filePath);
    const decrypted = decryptFile(encrypted);

    res.setHeader('Content-Disposition', `attachment; filename="${doc.original_filename}"`);
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Length', decrypted.length);
    res.send(decrypted);
  } catch (err) {
    console.error('Download error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

module.exports = router;
