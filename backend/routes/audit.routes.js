'use strict';
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const XLSX = require('xlsx');
const { Document, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType, Packer } = require('docx');
const { db, seedDefaultGroupsForEngagement } = require('../db');
const { requireAuth } = require('../auth');
const { encryptFile } = require('../utils/crypto');

const router = express.Router();
const VAULT_PATH = path.join(__dirname, '..', 'storage', 'vault');

// Multer: store files in memory for processing
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 } // 100MB limit
});

// Apply auth to all routes
router.use(requireAuth);

// Valid IE/PL types
const VALID_IE_PL_TYPES = ['Income', 'Expenditure', 'Asset', 'Liability', 'Equity'];
const VALID_ENTRY_TYPES = ['one_to_one', 'one_to_many', 'many_to_many'];

// ─────────────────────────────────────────────────────────────────────────────
// ENGAGEMENTS
// ─────────────────────────────────────────────────────────────────────────────

// GET /api/audit/engagements
router.get('/engagements', (req, res) => {
  try {
    const engagements = db.prepare(`
      SELECT e.*,
             u.name AS created_by_name,
             (SELECT COUNT(*) FROM trial_balance_ledgers WHERE engagement_id = e.id) AS ledger_count,
             (SELECT COUNT(*) FROM trial_balance_ledgers WHERE engagement_id = e.id AND is_mapped = 1) AS mapped_count,
             (SELECT COUNT(*) FROM audit_entries WHERE engagement_id = e.id) AS entry_count,
             (SELECT COUNT(*) FROM audit_entries WHERE engagement_id = e.id AND status = 'Approved') AS approved_entry_count
      FROM audit_engagements e
      LEFT JOIN users u ON e.created_by = u.id
      ORDER BY e.created_at DESC
    `).all();
    res.json(engagements);
  } catch (err) {
    console.error('Engagements list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/engagements
router.post('/engagements', (req, res) => {
  try {
    const { client_name, financial_year, period_start, period_end } = req.body;
    if (!client_name || !financial_year || !period_start || !period_end) {
      return res.status(400).json({ error: 'client_name, financial_year, period_start, and period_end are required' });
    }

    const now = new Date().toISOString();
    const result = db.prepare(`
      INSERT INTO audit_engagements (client_name, financial_year, period_start, period_end, created_by, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `).run(client_name, financial_year, period_start, period_end, req.user.id, now, now);

    const engagement = db.prepare(`
      SELECT e.*, u.name AS created_by_name
      FROM audit_engagements e
      LEFT JOIN users u ON e.created_by = u.id
      WHERE e.id = ?
    `).get(result.lastInsertRowid);

    seedDefaultGroupsForEngagement(result.lastInsertRowid);

    res.status(201).json(engagement);
  } catch (err) {
    console.error('Engagement create error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/engagements/:id
router.get('/engagements/:id', (req, res) => {
  try {
    const engagement = db.prepare(`
      SELECT e.*,
             u.name AS created_by_name,
             (SELECT COUNT(*) FROM trial_balance_ledgers WHERE engagement_id = e.id) AS ledger_count,
             (SELECT COUNT(*) FROM trial_balance_ledgers WHERE engagement_id = e.id AND is_mapped = 1) AS mapped_count,
             (SELECT COUNT(*) FROM audit_entries WHERE engagement_id = e.id) AS entry_count,
             (SELECT COUNT(*) FROM audit_entries WHERE engagement_id = e.id AND status = 'Approved') AS approved_entry_count
      FROM audit_engagements e
      LEFT JOIN users u ON e.created_by = u.id
      WHERE e.id = ?
    `).get(req.params.id);

    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });
    res.json(engagement);
  } catch (err) {
    console.error('Engagement get error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/engagements/:id
router.patch('/engagements/:id', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const allowed = ['client_name', 'financial_year', 'period_start', 'period_end', 'status'];
    const updates = [];
    const values = [];
    for (const field of allowed) {
      if (req.body[field] !== undefined) {
        updates.push(`${field} = ?`);
        values.push(req.body[field]);
      }
    }

    if (updates.length === 0) {
      return res.status(400).json({ error: 'No valid fields to update' });
    }

    updates.push('updated_at = ?');
    values.push(new Date().toISOString());
    values.push(req.params.id);

    db.prepare(`UPDATE audit_engagements SET ${updates.join(', ')} WHERE id = ?`).run(...values);

    const updated = db.prepare(`
      SELECT e.*, u.name AS created_by_name
      FROM audit_engagements e
      LEFT JOIN users u ON e.created_by = u.id
      WHERE e.id = ?
    `).get(req.params.id);

    res.json(updated);
  } catch (err) {
    console.error('Engagement update error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// TRIAL BALANCE
// ─────────────────────────────────────────────────────────────────────────────

// Helper: convert a 0-based column index to a spreadsheet letter (0=A, 1=B, 25=Z, 26=AA...)
function columnIndexToLetter(index) {
  let letter = '';
  let num = index;
  while (num >= 0) {
    letter = String.fromCharCode((num % 26) + 65) + letter;
    num = Math.floor(num / 26) - 1;
  }
  return letter;
}

// POST /api/audit/:id/trial-balance/preview
router.post('/:id/trial-balance/preview', upload.single('file'), (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    if (!req.file) {
      return res.status(400).json({ error: 'No file received' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const allSheetNames = workbook.SheetNames;

    // Sheet selection: use requested sheet, or default to first
    const requestedSheet = req.body.sheet_name;
    const sheetName = (requestedSheet && allSheetNames.includes(requestedSheet))
      ? requestedSheet
      : allSheetNames[0];

    const sheet = workbook.Sheets[sheetName];

    // Always read as raw array-of-arrays — no header assumption at all here
    const rawRows = XLSX.utils.sheet_to_json(sheet, {
      header: 1,
      defval: '',
      raw: false,
      blankrows: true  // KEEP blank rows so row numbers stay accurate for header row selection
    });

    if (rawRows.length === 0) {
      return res.status(400).json({ error: 'Selected sheet is empty.' });
    }

    const maxCols = Math.max(...rawRows.map(r => r.length), 1);

    // Header row index: user explicitly picks it (0-based). Default guess: first row with 3+ non-empty cells.
    let headerRowIndex = req.body.header_row_index !== undefined
      ? parseInt(req.body.header_row_index, 10)
      : null;

    if (isNaN(headerRowIndex) || headerRowIndex === null) {
      headerRowIndex = rawRows.findIndex(row =>
        row.filter(cell => String(cell).trim() !== '').length >= 3
      );
      if (headerRowIndex === -1) headerRowIndex = 0;
    }

    const headerRow = rawRows[headerRowIndex] || [];

    const columns = [];
    for (let i = 0; i < maxCols; i++) {
      const letter = columnIndexToLetter(i);
      const headerText = headerRow[i] ? String(headerRow[i]).trim() : '';
      columns.push({
        index: i,
        letter,
        header: headerText,
        label: headerText ? `${letter}: ${headerText}` : `Column ${letter}`
      });
    }

    // Data rows = everything after the header row
    const dataRows = rawRows.slice(headerRowIndex + 1);

    res.json({
      all_sheet_names: allSheetNames,
      selected_sheet: sheetName,
      // First 15 RAW rows (unfiltered) so the user can visually pick the header row
      raw_preview_rows: rawRows.slice(0, 15),
      header_row_index: headerRowIndex,
      columns,
      total_data_rows: dataRows.length,
      // First 10 data rows AFTER the chosen header, for the mapping preview
      preview_rows: dataRows.slice(0, 10)
    });

  } catch (err) {
    console.error('[TB PREVIEW ERROR]', err);
    res.status(500).json({ error: 'Failed to parse file.', detail: err.message });
  }
});

// POST /api/audit/:id/trial-balance/import
router.post('/:id/trial-balance/import', upload.single('file'), (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    if (!req.file) {
      return res.status(400).json({ error: 'No file received' });
    }

    let columnMap;
    try {
      columnMap = JSON.parse(req.body.column_map);
    } catch (e) {
      return res.status(400).json({ error: 'Invalid column_map format.' });
    }

    const sheetName = req.body.sheet_name;
    const headerRowIndex = parseInt(req.body.header_row_index, 10);

    if (!sheetName || isNaN(headerRowIndex)) {
      return res.status(400).json({ error: 'sheet_name and header_row_index are required.' });
    }

    const requiredFields = [
      'opening_balance', 'debit_transactions', 'credit_transactions', 'closing_balance'
    ];
    for (const field of requiredFields) {
      if (columnMap[field] === undefined || columnMap[field] === null) {
        return res.status(400).json({ error: `Missing column mapping for: ${field}` });
      }
    }
    const usedIndices = Object.values(columnMap).filter(v => v !== null && v !== undefined);
    if (new Set(usedIndices).size !== usedIndices.length) {
      return res.status(400).json({ error: 'The same column cannot be mapped to more than one field.' });
    }

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    if (!workbook.SheetNames.includes(sheetName)) {
      return res.status(400).json({ error: `Sheet "${sheetName}" not found in file.` });
    }
    const sheet = workbook.Sheets[sheetName];
    const rawRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: true });
    const dataRows = rawRows.slice(headerRowIndex + 1);

    const parseNum = (v) => {
      if (v === '' || v === null || v === undefined) return 0;
      const cleaned = String(v).replace(/[,₹\s\u00A0]/g, ''); // also strips non-breaking spaces (\xa0) seen in real exports
      const num = parseFloat(cleaned);
      return isNaN(num) ? 0 : num;
    };

    // Optional: explicit list of row indices (relative to dataRows, 0-based) to SKIP
    // e.g. group/subtotal rows the user identified in the row-filter preview step
    let skipRowIndices = new Set();
    if (req.body.skip_row_indices) {
      try {
        skipRowIndices = new Set(JSON.parse(req.body.skip_row_indices));
      } catch (e) { /* ignore malformed, just import everything with a code */ }
    }

    const insertMany = db.transaction((rows) => {
      const stmt = db.prepare(`
        INSERT INTO trial_balance_ledgers 
          (engagement_id, ledger_code, ledger_name, opening_balance, 
           debit_transactions, credit_transactions, closing_balance)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(engagement_id, ledger_code, ledger_name) DO UPDATE SET
          opening_balance     = opening_balance + excluded.opening_balance,
          debit_transactions  = debit_transactions + excluded.debit_transactions,
          credit_transactions = credit_transactions + excluded.credit_transactions,
          closing_balance     = closing_balance + excluded.closing_balance,
          is_mapped = 0,
          sub_subgroup_id = NULL
      `);

      let imported = 0;
      let merged = 0; // count of rows that hit a true code+name collision and got summed
      let skippedBlank = 0;
      let skippedManual = 0;
      const errors = [];
      const seenInThisImport = new Set(); // tracks code+name pairs seen so far in this same import batch

      const dataStartIndex = headerRowIndex + 1;

      rows.forEach((row, i) => {
        if (skipRowIndices.has(i)) { skippedManual++; return; }

        // columnMap.ledger_code / columnMap.ledger_name may now be null if the user chose
        // not to map them. Only read from the row if a column was actually mapped.
        const code = (columnMap.ledger_code !== null && columnMap.ledger_code !== undefined)
          ? String(row[columnMap.ledger_code] ?? '').trim()
          : '';
        const name = (columnMap.ledger_name !== null && columnMap.ledger_name !== undefined)
          ? String(row[columnMap.ledger_name] ?? '').trim()
          : '';

        // The existing "skip rows with no code — likely a group/subtotal row" heuristic
        // only applies when ledger_code WAS mapped. If the user chose not to map ledger_code
        // at all, we can no longer use "blank code" to detect subtotal rows — every row
        // will have a blank code by definition, so skip that heuristic entirely in this case.
        if (columnMap.ledger_code !== null && columnMap.ledger_code !== undefined && !code) {
          skippedBlank++;
          return;
        }

        // Previously a missing ledger_name was always an error. Now it's only worth flagging
        // if ledger_name WAS mapped but happened to be blank on this specific row — not an
        // error if the user chose not to map ledger_name at all.
        if (columnMap.ledger_name !== null && columnMap.ledger_name !== undefined && !name) {
          errors.push(`Row with ledger code "${code || '(none)'}" has no ledger name — skipped`);
          return;
        }

        let finalCode = code;
        let finalName = name;

        // If BOTH ledger_code and ledger_name were left unmapped for this entire import,
        // every row would otherwise resolve to identical blank values and collide on the
        // UNIQUE constraint, incorrectly merging unrelated ledgers together. Guard against
        // this by giving each such row a distinct synthetic name based on its row position.
        const codeWasMapped = columnMap.ledger_code !== null && columnMap.ledger_code !== undefined;
        const nameWasMapped = columnMap.ledger_name !== null && columnMap.ledger_name !== undefined;

        if (!codeWasMapped && !nameWasMapped) {
          finalName = `Unnamed Ledger (Row ${i + dataStartIndex + 1})`;
        }

        const dedupeKey = `${finalCode}|||${finalName}`;
        if (seenInThisImport.has(dedupeKey)) {
          merged++;
        } else {
          seenInThisImport.add(dedupeKey);
          imported++;
        }

        stmt.run(
          req.params.id, finalCode, finalName,
          parseNum(row[columnMap.opening_balance]),
          parseNum(row[columnMap.debit_transactions]),
          parseNum(row[columnMap.credit_transactions]),
          parseNum(row[columnMap.closing_balance])
        );
      });

      return { imported, merged, skippedBlank, skippedManual, errors };
    });

    const result = insertMany(dataRows);

    const totals = db.prepare(`
      SELECT SUM(debit_transactions) AS total_debit, SUM(credit_transactions) AS total_credit, COUNT(*) AS total_count
      FROM trial_balance_ledgers WHERE engagement_id = ?
    `).get(req.params.id);

    const isBalanced = Math.abs((totals.total_debit || 0) - (totals.total_credit || 0)) < 0.01;

    db.prepare(`UPDATE audit_engagements SET status = 'Trial Balance Imported', updated_at = datetime('now') WHERE id = ?`)
      .run(req.params.id);

    res.json({ ...result, totals, isBalanced });

  } catch (err) {
    console.error('[TB IMPORT ERROR]', err);
    res.status(500).json({ error: 'Failed to import trial balance.', detail: err.message });
  }
});

// GET /api/audit/:id/trial-balance
router.get('/:id/trial-balance', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { search, mapped, group_id, subgroup_id, sub_subgroup_id } = req.query;

    let query = `
      SELECT tb.*,
             g.id AS group_id, g.name AS group_name,
             sg.id AS subgroup_id, sg.name AS subgroup_name,
             ssg.id AS sub_subgroup_id, ssg.name AS sub_subgroup_name
      FROM trial_balance_ledgers tb
      LEFT JOIN sub_subgroups ssg ON tb.sub_subgroup_id = ssg.id
      LEFT JOIN subgroups sg ON ssg.subgroup_id = sg.id
      LEFT JOIN groups g ON sg.group_id = g.id
      WHERE tb.engagement_id = ?
    `;
    const params = [req.params.id];

    if (search) {
      query += ' AND (tb.ledger_code LIKE ? OR tb.ledger_name LIKE ?)';
      params.push(`%${search}%`, `%${search}%`);
    }
    if (mapped === 'true') {
      query += ' AND tb.is_mapped = 1';
    } else if (mapped === 'false') {
      query += ' AND tb.is_mapped = 0';
    }

    if (sub_subgroup_id) {
      query += ' AND ssg.id = ?';
      params.push(sub_subgroup_id);
    } else if (subgroup_id) {
      query += ' AND sg.id = ?';
      params.push(subgroup_id);
    } else if (group_id) {
      query += ' AND g.id = ?';
      params.push(group_id);
    }

    query += ' ORDER BY tb.ledger_code ASC';

    const ledgers = db.prepare(query).all(...params);

    // Compute totals from all ledgers (ignoring filters for overall totals)
    const totalsRow = db.prepare(`
      SELECT
        COALESCE(SUM(opening_balance), 0) AS opening_balance,
        COALESCE(SUM(debit_transactions), 0) AS debit_transactions,
        COALESCE(SUM(credit_transactions), 0) AS credit_transactions,
        COALESCE(SUM(closing_balance), 0) AS closing_balance,
        COUNT(*) AS total_count,
        SUM(CASE WHEN is_mapped = 1 THEN 1 ELSE 0 END) AS mapped_count,
        SUM(CASE WHEN is_mapped = 0 THEN 1 ELSE 0 END) AS unmapped_count
      FROM trial_balance_ledgers
      WHERE engagement_id = ?
    `).get(req.params.id);

    const totals = {
      opening_balance: Math.round(totalsRow.opening_balance * 100) / 100,
      debit_transactions: Math.round(totalsRow.debit_transactions * 100) / 100,
      credit_transactions: Math.round(totalsRow.credit_transactions * 100) / 100,
      closing_balance: Math.round(totalsRow.closing_balance * 100) / 100,
      is_balanced: Math.abs(totalsRow.debit_transactions - totalsRow.credit_transactions) < 0.01,
      total_count: totalsRow.total_count,
      mapped_count: totalsRow.mapped_count,
      unmapped_count: totalsRow.unmapped_count
    };

    res.json({ ledgers, totals });
  } catch (err) {
    console.error('TB list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// GROUPS HIERARCHY
// ─────────────────────────────────────────────────────────────────────────────

// GET /api/audit/:id/groups-tree
router.get('/:id/groups-tree', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const treeData = db.prepare(`
      SELECT 
        g.id AS group_id, g.name AS group_name, g.display_order AS group_order,
        sg.id AS subgroup_id, sg.name AS subgroup_name, sg.display_order AS subgroup_order,
        ssg.id AS sub_subgroup_id, ssg.name AS sub_subgroup_name, ssg.display_order AS sub_subgroup_order,
        (SELECT COUNT(*) FROM trial_balance_ledgers WHERE sub_subgroup_id = ssg.id) AS ssg_ledger_count
      FROM groups g
      JOIN subgroups sg ON sg.group_id = g.id
      JOIN sub_subgroups ssg ON ssg.subgroup_id = sg.id
      WHERE g.engagement_id = ?
      ORDER BY g.display_order ASC, sg.display_order ASC, ssg.display_order ASC
    `).all(req.params.id);

    const tree = [];
    let currentGroup = null;
    let currentSubgroup = null;

    for (const row of treeData) {
      if (!currentGroup || currentGroup.id !== row.group_id) {
        currentGroup = {
          id: row.group_id,
          name: row.group_name,
          display_order: row.group_order,
          ledger_count: 0,
          subgroups: []
        };
        tree.push(currentGroup);
      }

      if (!currentSubgroup || currentSubgroup.id !== row.subgroup_id) {
        currentSubgroup = {
          id: row.subgroup_id,
          name: row.subgroup_name,
          display_order: row.subgroup_order,
          ledger_count: 0,
          sub_subgroups: []
        };
        currentGroup.subgroups.push(currentSubgroup);
      }

      const ssg = {
        id: row.sub_subgroup_id,
        name: row.sub_subgroup_name,
        display_order: row.sub_subgroup_order,
        ledger_count: row.ssg_ledger_count
      };
      
      currentSubgroup.sub_subgroups.push(ssg);
      currentSubgroup.ledger_count += ssg.ledger_count;
      currentGroup.ledger_count += ssg.ledger_count;
    }

    res.json(tree);
  } catch (err) {
    console.error('Groups tree error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// SUBGROUP & SUB-SUBGROUP CRUD
// ─────────────────────────────────────────────────────────────────────────────

// POST /api/audit/:id/subgroups
router.post('/:id/subgroups', (req, res) => {
  try {
    const engagement = db.prepare('SELECT id FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { group_id, name } = req.body;
    if (!name || typeof name !== 'string' || name.trim() === '') {
      return res.status(400).json({ error: 'Name is required' });
    }

    const group = db.prepare('SELECT id FROM groups WHERE id = ? AND engagement_id = ?').get(group_id, req.params.id);
    if (!group) return res.status(404).json({ error: 'Group not found in this engagement' });

    const maxOrderRes = db.prepare('SELECT MAX(display_order) as max_order FROM subgroups WHERE group_id = ?').get(group_id);
    const display_order = maxOrderRes.max_order !== null ? maxOrderRes.max_order + 1 : 0;

    const insert = db.prepare('INSERT INTO subgroups (group_id, name, display_order) VALUES (?, ?, ?)');
    const info = insert.run(group_id, name.trim(), display_order);

    res.status(201).json({ id: info.lastInsertRowid, group_id, name: name.trim(), display_order });
  } catch (err) {
    console.error('Create subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/subgroups/:sgid
router.patch('/:id/subgroups/:sgid', (req, res) => {
  try {
    const { name } = req.body;
    if (!name || typeof name !== 'string' || name.trim() === '') {
      return res.status(400).json({ error: 'Name is required' });
    }

    const check = db.prepare(`
      SELECT sg.id FROM subgroups sg
      JOIN groups g ON sg.group_id = g.id
      WHERE sg.id = ? AND g.engagement_id = ?
    `).get(req.params.sgid, req.params.id);

    if (!check) return res.status(404).json({ error: 'Subgroup not found in this engagement' });

    db.prepare('UPDATE subgroups SET name = ? WHERE id = ?').run(name.trim(), req.params.sgid);
    const updated = db.prepare('SELECT id, group_id, name, display_order FROM subgroups WHERE id = ?').get(req.params.sgid);

    res.json(updated);
  } catch (err) {
    console.error('Update subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// DELETE /api/audit/:id/subgroups/:sgid
router.delete('/:id/subgroups/:sgid', (req, res) => {
  try {
    const check = db.prepare(`
      SELECT sg.id FROM subgroups sg
      JOIN groups g ON sg.group_id = g.id
      WHERE sg.id = ? AND g.engagement_id = ?
    `).get(req.params.sgid, req.params.id);

    if (!check) return res.status(404).json({ error: 'Subgroup not found in this engagement' });

    let unmappedCount = 0;
    const tx = db.transaction(() => {
      const updateRes = db.prepare(`
        UPDATE trial_balance_ledgers
        SET sub_subgroup_id = NULL, is_mapped = 0
        WHERE sub_subgroup_id IN (
          SELECT id FROM sub_subgroups WHERE subgroup_id = ?
        )
      `).run(req.params.sgid);
      unmappedCount = updateRes.changes;

      db.prepare('DELETE FROM subgroups WHERE id = ?').run(req.params.sgid);
    });
    tx();

    res.json({ success: true, unmapped_count: unmappedCount });
  } catch (err) {
    console.error('Delete subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/sub-subgroups
router.post('/:id/sub-subgroups', (req, res) => {
  try {
    const { subgroup_id, name } = req.body;
    if (!name || typeof name !== 'string' || name.trim() === '') {
      return res.status(400).json({ error: 'Name is required' });
    }

    const check = db.prepare(`
      SELECT sg.id FROM subgroups sg
      JOIN groups g ON sg.group_id = g.id
      WHERE sg.id = ? AND g.engagement_id = ?
    `).get(subgroup_id, req.params.id);

    if (!check) return res.status(404).json({ error: 'Subgroup not found in this engagement' });

    const maxOrderRes = db.prepare('SELECT MAX(display_order) as max_order FROM sub_subgroups WHERE subgroup_id = ?').get(subgroup_id);
    const display_order = maxOrderRes.max_order !== null ? maxOrderRes.max_order + 1 : 0;

    const insert = db.prepare('INSERT INTO sub_subgroups (subgroup_id, name, display_order) VALUES (?, ?, ?)');
    const info = insert.run(subgroup_id, name.trim(), display_order);

    res.status(201).json({ id: info.lastInsertRowid, subgroup_id, name: name.trim(), display_order });
  } catch (err) {
    console.error('Create sub-subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/sub-subgroups/:ssgid
router.patch('/:id/sub-subgroups/:ssgid', (req, res) => {
  try {
    const { name } = req.body;
    if (!name || typeof name !== 'string' || name.trim() === '') {
      return res.status(400).json({ error: 'Name is required' });
    }

    const check = db.prepare(`
      SELECT ssg.id FROM sub_subgroups ssg
      JOIN subgroups sg ON ssg.subgroup_id = sg.id
      JOIN groups g ON sg.group_id = g.id
      WHERE ssg.id = ? AND g.engagement_id = ?
    `).get(req.params.ssgid, req.params.id);

    if (!check) return res.status(404).json({ error: 'Sub-subgroup not found in this engagement' });

    db.prepare('UPDATE sub_subgroups SET name = ? WHERE id = ?').run(name.trim(), req.params.ssgid);
    const updated = db.prepare('SELECT id, subgroup_id, name, display_order FROM sub_subgroups WHERE id = ?').get(req.params.ssgid);

    res.json(updated);
  } catch (err) {
    console.error('Update sub-subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// DELETE /api/audit/:id/sub-subgroups/:ssgid
router.delete('/:id/sub-subgroups/:ssgid', (req, res) => {
  try {
    const check = db.prepare(`
      SELECT ssg.id FROM sub_subgroups ssg
      JOIN subgroups sg ON ssg.subgroup_id = sg.id
      JOIN groups g ON sg.group_id = g.id
      WHERE ssg.id = ? AND g.engagement_id = ?
    `).get(req.params.ssgid, req.params.id);

    if (!check) return res.status(404).json({ error: 'Sub-subgroup not found in this engagement' });

    let unmappedCount = 0;
    const tx = db.transaction(() => {
      const updateRes = db.prepare(`
        UPDATE trial_balance_ledgers
        SET sub_subgroup_id = NULL, is_mapped = 0
        WHERE sub_subgroup_id = ?
      `).run(req.params.ssgid);
      unmappedCount = updateRes.changes;

      db.prepare('DELETE FROM sub_subgroups WHERE id = ?').run(req.params.ssgid);
    });
    tx();

    res.json({ success: true, unmapped_count: unmappedCount });
  } catch (err) {
    console.error('Delete sub-subgroup error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});


// ─────────────────────────────────────────────────────────────────────────────
// LEDGER MAPPING
// ─────────────────────────────────────────────────────────────────────────────

// PATCH /api/audit/:id/ledgers/:lid/map
router.patch('/:id/ledgers/:lid/map', (req, res) => {
  try {
    const ledger = db.prepare('SELECT * FROM trial_balance_ledgers WHERE id = ? AND engagement_id = ?').get(req.params.lid, req.params.id);
    if (!ledger) return res.status(404).json({ error: 'Ledger not found' });

    const { sub_subgroup_id } = req.body;

    if (sub_subgroup_id !== null && sub_subgroup_id !== undefined) {
      // Validate sub_subgroup exists and belongs to this engagement
      const groupCheck = db.prepare(`
        SELECT g.engagement_id
        FROM sub_subgroups ssg
        JOIN subgroups sg ON ssg.subgroup_id = sg.id
        JOIN groups g ON sg.group_id = g.id
        WHERE ssg.id = ?
      `).get(sub_subgroup_id);
      
      if (!groupCheck || groupCheck.engagement_id != req.params.id) {
        return res.status(404).json({ error: 'Sub-subgroup not found in this engagement' });
      }

      db.prepare('UPDATE trial_balance_ledgers SET sub_subgroup_id = ?, is_mapped = 1 WHERE id = ?').run(sub_subgroup_id, req.params.lid);
    } else {
      // Unmap
      db.prepare('UPDATE trial_balance_ledgers SET sub_subgroup_id = NULL, is_mapped = 0 WHERE id = ?').run(req.params.lid);
    }

    const updated = db.prepare(`
      SELECT tb.*,
             g.id AS group_id, g.name AS group_name,
             sg.id AS subgroup_id, sg.name AS subgroup_name,
             ssg.id AS sub_subgroup_id, ssg.name AS sub_subgroup_name
      FROM trial_balance_ledgers tb
      LEFT JOIN sub_subgroups ssg ON tb.sub_subgroup_id = ssg.id
      LEFT JOIN subgroups sg ON ssg.subgroup_id = sg.id
      LEFT JOIN groups g ON sg.group_id = g.id
      WHERE tb.id = ?
    `).get(req.params.lid);

    res.json(updated);
  } catch (err) {
    console.error('Ledger map error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/ledgers/bulk-map
router.post('/:id/ledgers/bulk-map', (req, res) => {
  try {
    const { ledger_ids, sub_subgroup_id } = req.body;

    if (!Array.isArray(ledger_ids) || ledger_ids.length === 0) {
      return res.status(400).json({ error: 'ledger_ids must be a non-empty array' });
    }

    if (sub_subgroup_id !== null && sub_subgroup_id !== undefined) {
      const groupCheck = db.prepare(`
        SELECT g.engagement_id
        FROM sub_subgroups ssg
        JOIN subgroups sg ON ssg.subgroup_id = sg.id
        JOIN groups g ON sg.group_id = g.id
        WHERE ssg.id = ?
      `).get(sub_subgroup_id);
      
      if (!groupCheck || groupCheck.engagement_id != req.params.id) {
        return res.status(404).json({ error: 'Sub-subgroup not found in this engagement' });
      }
    }

    const updateStmt = sub_subgroup_id !== null && sub_subgroup_id !== undefined
      ? db.prepare('UPDATE trial_balance_ledgers SET sub_subgroup_id = ?, is_mapped = 1 WHERE id = ? AND engagement_id = ?')
      : db.prepare('UPDATE trial_balance_ledgers SET sub_subgroup_id = NULL, is_mapped = 0 WHERE id = ? AND engagement_id = ?');

    let updated = 0;
    const bulkMapTx = db.transaction(() => {
      for (const lid of ledger_ids) {
        let result;
        if (sub_subgroup_id !== null && sub_subgroup_id !== undefined) {
          result = updateStmt.run(sub_subgroup_id, lid, req.params.id);
        } else {
          result = updateStmt.run(lid, req.params.id);
        }
        updated += result.changes;
      }
    });

    bulkMapTx();
    res.json({ updated });
  } catch (err) {
    console.error('Bulk map error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// AUDIT ENTRIES
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Helper: fetch a single entry with all its lines and computed fields.
 */
function fetchEntryWithLines(entryId) {
  const entry = db.prepare(`
    SELECT ae.*,
           u1.name AS created_by_name,
           u2.name AS submitted_by_name,
           u3.name AS reviewed_by_name
    FROM audit_entries ae
    LEFT JOIN users u1 ON ae.created_by = u1.id
    LEFT JOIN users u2 ON ae.submitted_by = u2.id
    LEFT JOIN users u3 ON ae.reviewed_by = u3.id
    WHERE ae.id = ?
  `).get(entryId);

  if (!entry) return null;

  const lines = db.prepare(`
    SELECT ael.*,
           tb.ledger_code, tb.ledger_name
    FROM audit_entry_lines ael
    LEFT JOIN trial_balance_ledgers tb ON ael.ledger_id = tb.id
    WHERE ael.audit_entry_id = ?
    ORDER BY ael.line_order ASC
  `).all(entryId);

  const total_debit = Math.round(lines.filter(l => l.line_type === 'debit').reduce((s, l) => s + l.amount, 0) * 100) / 100;
  const total_credit = Math.round(lines.filter(l => l.line_type === 'credit').reduce((s, l) => s + l.amount, 0) * 100) / 100;
  const is_balanced = Math.abs(total_debit - total_credit) < 0.01;

  return { ...entry, lines, total_debit, total_credit, is_balanced };
}

// GET /api/audit/:id/entries
router.get('/:id/entries', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { status, search } = req.query;

    let query = `
      SELECT ae.*,
             u1.name AS created_by_name,
             u2.name AS submitted_by_name,
             u3.name AS reviewed_by_name
      FROM audit_entries ae
      LEFT JOIN users u1 ON ae.created_by = u1.id
      LEFT JOIN users u2 ON ae.submitted_by = u2.id
      LEFT JOIN users u3 ON ae.reviewed_by = u3.id
      WHERE ae.engagement_id = ?
    `;
    const params = [req.params.id];

    if (status) {
      query += ' AND ae.status = ?';
      params.push(status);
    }
    if (search) {
      query += ' AND (ae.entry_number LIKE ? OR ae.description LIKE ?)';
      params.push(`%${search}%`, `%${search}%`);
    }

    query += ' ORDER BY ae.created_at DESC';

    const entries = db.prepare(query).all(...params);

    // Attach lines and computed totals for each entry
    const linesStmt = db.prepare(`
      SELECT ael.*, tb.ledger_code, tb.ledger_name
      FROM audit_entry_lines ael
      LEFT JOIN trial_balance_ledgers tb ON ael.ledger_id = tb.id
      WHERE ael.audit_entry_id = ?
      ORDER BY ael.line_order ASC
    `);

    const result = entries.map(entry => {
      const lines = linesStmt.all(entry.id);
      const total_debit = Math.round(lines.filter(l => l.line_type === 'debit').reduce((s, l) => s + l.amount, 0) * 100) / 100;
      const total_credit = Math.round(lines.filter(l => l.line_type === 'credit').reduce((s, l) => s + l.amount, 0) * 100) / 100;
      const is_balanced = Math.abs(total_debit - total_credit) < 0.01;
      return { ...entry, lines, total_debit, total_credit, is_balanced };
    });

    res.json(result);
  } catch (err) {
    console.error('Entries list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/entries
router.post('/:id/entries', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { entry_type, description, narration, entry_date, lines } = req.body;

    // Validate required fields
    if (!entry_type || !description || !entry_date) {
      return res.status(400).json({ error: 'entry_type, description, and entry_date are required' });
    }

    if (!VALID_ENTRY_TYPES.includes(entry_type)) {
      return res.status(400).json({ error: `Invalid entry_type. Must be one of: ${VALID_ENTRY_TYPES.join(', ')}` });
    }

    if (!Array.isArray(lines) || lines.length === 0) {
      return res.status(400).json({ error: 'lines array is required and must not be empty' });
    }

    // Validate lines
    const debitLines = lines.filter(l => l.line_type === 'debit');
    const creditLines = lines.filter(l => l.line_type === 'credit');

    if (debitLines.length === 0 || creditLines.length === 0) {
      return res.status(400).json({ error: 'Entry must have at least one debit and one credit line' });
    }

    // Validate entry type constraints
    if (entry_type === 'one_to_one') {
      if (debitLines.length !== 1 || creditLines.length !== 1) {
        return res.status(400).json({ error: 'one_to_one entry must have exactly 1 debit and 1 credit line' });
      }
    } else if (entry_type === 'one_to_many') {
      if (debitLines.length !== 1 || creditLines.length < 2) {
        return res.status(400).json({ error: 'one_to_many entry must have exactly 1 debit and 2+ credit lines' });
      }
    } else if (entry_type === 'many_to_many') {
      if (debitLines.length < 2 || creditLines.length < 2) {
        return res.status(400).json({ error: 'many_to_many entry must have 2+ debit and 2+ credit lines' });
      }
    }

    // Validate amounts
    for (const line of lines) {
      if (typeof line.amount !== 'number' || line.amount <= 0) {
        return res.status(400).json({ error: 'All line amounts must be positive numbers' });
      }
      if (!['debit', 'credit'].includes(line.line_type)) {
        return res.status(400).json({ error: 'line_type must be debit or credit' });
      }
    }

    // Validate balance
    const totalDebit = Math.round(debitLines.reduce((s, l) => s + l.amount, 0) * 100) / 100;
    const totalCredit = Math.round(creditLines.reduce((s, l) => s + l.amount, 0) * 100) / 100;
    if (Math.abs(totalDebit - totalCredit) > 0.01) {
      return res.status(400).json({ error: `Debits (${totalDebit}) must equal credits (${totalCredit})` });
    }

    // Validate ledger_ids belong to this engagement
    const ledgerCheckStmt = db.prepare('SELECT id FROM trial_balance_ledgers WHERE id = ? AND engagement_id = ?');
    for (const line of lines) {
      if (!line.ledger_id) {
        return res.status(400).json({ error: 'Each line must have a ledger_id' });
      }
      const ledger = ledgerCheckStmt.get(line.ledger_id, req.params.id);
      if (!ledger) {
        return res.status(400).json({ error: `Ledger ID ${line.ledger_id} not found in this engagement` });
      }
    }

    // Auto-generate entry_number
    const lastEntry = db.prepare(`
      SELECT entry_number FROM audit_entries
      WHERE engagement_id = ?
      ORDER BY id DESC LIMIT 1
    `).get(req.params.id);

    let nextNum = 1;
    if (lastEntry && lastEntry.entry_number) {
      const match = lastEntry.entry_number.match(/AE-(\d+)/);
      if (match) nextNum = parseInt(match[1], 10) + 1;
    }
    const entry_number = `AE-${String(nextNum).padStart(3, '0')}`;

    const now = new Date().toISOString();

    const insertEntryStmt = db.prepare(`
      INSERT INTO audit_entries (engagement_id, entry_number, entry_type, description, narration, entry_date, status, created_by, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, ?, 'Draft', ?, ?, ?)
    `);

    const insertLineStmt = db.prepare(`
      INSERT INTO audit_entry_lines (audit_entry_id, ledger_id, line_type, amount, line_order)
      VALUES (?, ?, ?, ?, ?)
    `);

    let entryId;
    const createTx = db.transaction(() => {
      const result = insertEntryStmt.run(req.params.id, entry_number, entry_type, description, narration || null, entry_date, req.user.id, now, now);
      entryId = result.lastInsertRowid;

      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        insertLineStmt.run(entryId, line.ledger_id, line.line_type, line.amount, i);
      }
    });

    createTx();

    const created = fetchEntryWithLines(entryId);
    res.status(201).json(created);
  } catch (err) {
    console.error('Entry create error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/:id/entries/:eid
router.get('/:id/entries/:eid', (req, res) => {
  try {
    const entry = fetchEntryWithLines(req.params.eid);
    if (!entry || entry.engagement_id !== parseInt(req.params.id, 10)) {
      return res.status(404).json({ error: 'Entry not found' });
    }
    res.json(entry);
  } catch (err) {
    console.error('Entry get error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/entries/:eid
router.patch('/:id/entries/:eid', (req, res) => {
  try {
    const existing = db.prepare('SELECT * FROM audit_entries WHERE id = ? AND engagement_id = ?').get(req.params.eid, req.params.id);
    if (!existing) return res.status(404).json({ error: 'Entry not found' });

    if (existing.status !== 'Draft' && existing.status !== 'Rejected') {
      return res.status(400).json({ error: 'Can only edit entries in Draft or Rejected status' });
    }

    const { entry_type, description, narration, entry_date, lines } = req.body;

    // Validate same as create
    if (!entry_type || !description || !entry_date) {
      return res.status(400).json({ error: 'entry_type, description, and entry_date are required' });
    }

    if (!VALID_ENTRY_TYPES.includes(entry_type)) {
      return res.status(400).json({ error: `Invalid entry_type. Must be one of: ${VALID_ENTRY_TYPES.join(', ')}` });
    }

    if (!Array.isArray(lines) || lines.length === 0) {
      return res.status(400).json({ error: 'lines array is required and must not be empty' });
    }

    const debitLines = lines.filter(l => l.line_type === 'debit');
    const creditLines = lines.filter(l => l.line_type === 'credit');

    if (debitLines.length === 0 || creditLines.length === 0) {
      return res.status(400).json({ error: 'Entry must have at least one debit and one credit line' });
    }

    if (entry_type === 'one_to_one') {
      if (debitLines.length !== 1 || creditLines.length !== 1) {
        return res.status(400).json({ error: 'one_to_one entry must have exactly 1 debit and 1 credit line' });
      }
    } else if (entry_type === 'one_to_many') {
      if (debitLines.length !== 1 || creditLines.length < 2) {
        return res.status(400).json({ error: 'one_to_many entry must have exactly 1 debit and 2+ credit lines' });
      }
    } else if (entry_type === 'many_to_many') {
      if (debitLines.length < 2 || creditLines.length < 2) {
        return res.status(400).json({ error: 'many_to_many entry must have 2+ debit and 2+ credit lines' });
      }
    }

    for (const line of lines) {
      if (typeof line.amount !== 'number' || line.amount <= 0) {
        return res.status(400).json({ error: 'All line amounts must be positive numbers' });
      }
      if (!['debit', 'credit'].includes(line.line_type)) {
        return res.status(400).json({ error: 'line_type must be debit or credit' });
      }
    }

    const totalDebit = Math.round(debitLines.reduce((s, l) => s + l.amount, 0) * 100) / 100;
    const totalCredit = Math.round(creditLines.reduce((s, l) => s + l.amount, 0) * 100) / 100;
    if (Math.abs(totalDebit - totalCredit) > 0.01) {
      return res.status(400).json({ error: `Debits (${totalDebit}) must equal credits (${totalCredit})` });
    }

    const ledgerCheckStmt = db.prepare('SELECT id FROM trial_balance_ledgers WHERE id = ? AND engagement_id = ?');
    for (const line of lines) {
      if (!line.ledger_id) {
        return res.status(400).json({ error: 'Each line must have a ledger_id' });
      }
      const ledger = ledgerCheckStmt.get(line.ledger_id, req.params.id);
      if (!ledger) {
        return res.status(400).json({ error: `Ledger ID ${line.ledger_id} not found in this engagement` });
      }
    }

    const now = new Date().toISOString();
    const newStatus = existing.status === 'Rejected' ? 'Draft' : existing.status;

    const updateTx = db.transaction(() => {
      db.prepare(`
        UPDATE audit_entries
        SET entry_type = ?, description = ?, narration = ?, entry_date = ?, status = ?, rejection_reason = NULL, updated_at = ?
        WHERE id = ?
      `).run(entry_type, description, narration || null, entry_date, newStatus, now, req.params.eid);

      // Delete old lines and re-insert
      db.prepare('DELETE FROM audit_entry_lines WHERE audit_entry_id = ?').run(req.params.eid);

      const insertLineStmt = db.prepare(`
        INSERT INTO audit_entry_lines (audit_entry_id, ledger_id, line_type, amount, line_order)
        VALUES (?, ?, ?, ?, ?)
      `);
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        insertLineStmt.run(req.params.eid, line.ledger_id, line.line_type, line.amount, i);
      }
    });

    updateTx();

    const updated = fetchEntryWithLines(req.params.eid);
    res.json(updated);
  } catch (err) {
    console.error('Entry update error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// DELETE /api/audit/:id/entries/:eid
router.delete('/:id/entries/:eid', (req, res) => {
  try {
    const entry = db.prepare('SELECT * FROM audit_entries WHERE id = ? AND engagement_id = ?').get(req.params.eid, req.params.id);
    if (!entry) return res.status(404).json({ error: 'Entry not found' });

    if (entry.status !== 'Draft') {
      return res.status(400).json({ error: 'Can only delete entries in Draft status' });
    }

    const deleteTx = db.transaction(() => {
      db.prepare('DELETE FROM audit_entry_lines WHERE audit_entry_id = ?').run(req.params.eid);
      db.prepare('DELETE FROM audit_entries WHERE id = ?').run(req.params.eid);
    });

    deleteTx();
    res.json({ success: true });
  } catch (err) {
    console.error('Entry delete error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/entries/:eid/submit
router.post('/:id/entries/:eid/submit', (req, res) => {
  try {
    const entry = db.prepare('SELECT * FROM audit_entries WHERE id = ? AND engagement_id = ?').get(req.params.eid, req.params.id);
    if (!entry) return res.status(404).json({ error: 'Entry not found' });

    if (entry.status !== 'Draft' && entry.status !== 'Rejected') {
      return res.status(400).json({ error: 'Can only submit entries in Draft or Rejected status' });
    }

    const now = new Date().toISOString();
    db.prepare(`
      UPDATE audit_entries SET status = 'Submitted', submitted_by = ?, submitted_at = ?, updated_at = ? WHERE id = ?
    `).run(req.user.id, now, now, req.params.eid);

    const updated = fetchEntryWithLines(req.params.eid);
    res.json(updated);
  } catch (err) {
    console.error('Entry submit error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/entries/:eid/approve
router.post('/:id/entries/:eid/approve', (req, res) => {
  try {
    const entry = db.prepare('SELECT * FROM audit_entries WHERE id = ? AND engagement_id = ?').get(req.params.eid, req.params.id);
    if (!entry) return res.status(404).json({ error: 'Entry not found' });

    if (entry.status !== 'Submitted') {
      return res.status(400).json({ error: 'Can only approve entries in Submitted status' });
    }

    const now = new Date().toISOString();
    db.prepare(`
      UPDATE audit_entries SET status = 'Approved', reviewed_by = ?, reviewed_at = ?, updated_at = ? WHERE id = ?
    `).run(req.user.id, now, now, req.params.eid);

    const updated = fetchEntryWithLines(req.params.eid);
    res.json(updated);
  } catch (err) {
    console.error('Entry approve error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/entries/:eid/reject
router.post('/:id/entries/:eid/reject', (req, res) => {
  try {
    const entry = db.prepare('SELECT * FROM audit_entries WHERE id = ? AND engagement_id = ?').get(req.params.eid, req.params.id);
    if (!entry) return res.status(404).json({ error: 'Entry not found' });

    if (entry.status !== 'Submitted') {
      return res.status(400).json({ error: 'Can only reject entries in Submitted status' });
    }

    const { rejection_reason } = req.body;
    if (!rejection_reason) {
      return res.status(400).json({ error: 'rejection_reason is required' });
    }

    const now = new Date().toISOString();
    db.prepare(`
      UPDATE audit_entries SET status = 'Rejected', reviewed_by = ?, reviewed_at = ?, rejection_reason = ?, updated_at = ? WHERE id = ?
    `).run(req.user.id, now, rejection_reason, now, req.params.eid);

    const updated = fetchEntryWithLines(req.params.eid);
    res.json(updated);
  } catch (err) {
    console.error('Entry reject error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// ADJUSTED TRIAL BALANCE
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Helper: compute adjusted trial balance for an engagement.
 */
function computeAdjustedTB(engagementId) {
  const ledgers = db.prepare(`
    SELECT tb.*,
           g.id AS group_id, g.name AS group_name,
           sg.id AS subgroup_id, sg.name AS subgroup_name,
           ssg.id AS sub_subgroup_id, ssg.name AS sub_subgroup_name
    FROM trial_balance_ledgers tb
    LEFT JOIN sub_subgroups ssg ON tb.sub_subgroup_id = ssg.id
    LEFT JOIN subgroups sg ON ssg.subgroup_id = sg.id
    LEFT JOIN groups g ON sg.group_id = g.id
    WHERE tb.engagement_id = ?
    ORDER BY tb.ledger_code ASC
  `).all(engagementId);

  // Get approved entry lines grouped by ledger
  const adjustments = db.prepare(`
    SELECT ael.ledger_id,
           SUM(CASE WHEN ael.line_type = 'debit' THEN ael.amount ELSE 0 END) AS total_adj_debit,
           SUM(CASE WHEN ael.line_type = 'credit' THEN ael.amount ELSE 0 END) AS total_adj_credit
    FROM audit_entry_lines ael
    INNER JOIN audit_entries ae ON ael.audit_entry_id = ae.id
    WHERE ae.engagement_id = ? AND ae.status = 'Approved'
    GROUP BY ael.ledger_id
  `).all(engagementId);

  const adjMap = {};
  for (const adj of adjustments) {
    adjMap[adj.ledger_id] = adj;
  }

  let totalOpening = 0, totalClosing = 0, totalAdjDebit = 0, totalAdjCredit = 0, totalAdjustedClosing = 0;

  const result = ledgers.map(ledger => {
    const adj = adjMap[ledger.id] || { total_adj_debit: 0, total_adj_credit: 0 };
    const net_adjustment = Math.round((adj.total_adj_debit - adj.total_adj_credit) * 100) / 100;
    const adjusted_closing = Math.round((ledger.closing_balance + net_adjustment) * 100) / 100;

    totalOpening += ledger.opening_balance;
    totalClosing += ledger.closing_balance;
    totalAdjDebit += adj.total_adj_debit;
    totalAdjCredit += adj.total_adj_credit;
    totalAdjustedClosing += adjusted_closing;

    return {
      ...ledger,
      adj_debit: Math.round(adj.total_adj_debit * 100) / 100,
      adj_credit: Math.round(adj.total_adj_credit * 100) / 100,
      net_adjustment,
      adjusted_closing
    };
  });

  const summary = {
    total_opening: Math.round(totalOpening * 100) / 100,
    total_closing: Math.round(totalClosing * 100) / 100,
    total_adj_debit: Math.round(totalAdjDebit * 100) / 100,
    total_adj_credit: Math.round(totalAdjCredit * 100) / 100,
    total_adjusted_closing: Math.round(totalAdjustedClosing * 100) / 100,
    is_balanced: Math.abs(totalAdjDebit - totalAdjCredit) < 0.01
  };

  return { ledgers: result, summary };
}

// GET /api/audit/:id/adjusted-tb
router.get('/:id/adjusted-tb', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const result = computeAdjustedTB(req.params.id);
    res.json(result);
  } catch (err) {
    console.error('Adjusted TB error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/adjusted-tb/approve
router.post('/:id/adjusted-tb/approve', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    db.prepare("UPDATE audit_engagements SET status = 'Adjusted TB Approved', updated_at = ? WHERE id = ?")
      .run(new Date().toISOString(), req.params.id);

    res.json({ success: true, status: 'Adjusted TB Approved' });
  } catch (err) {
    console.error('Adjusted TB approve error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// FINANCIALS
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Helper: build financial sections from adjusted TB data.
 */
function buildFinancialSections(engagementId) {
  const { ledgers } = computeAdjustedTB(engagementId);

  const sections = {};
  for (const type of VALID_IE_PL_TYPES) {
    sections[type] = { subgroups: {}, total: 0 };
  }

  for (const ledger of ledgers) {
    if (!ledger.group_name) continue; // unmapped ledgers

    const section = sections[ledger.group_name];
    if (!section) continue;

    const subgroupKey = ledger.subgroup_name;
    if (!section.subgroups[subgroupKey]) {
      section.subgroups[subgroupKey] = {
        subgroup_name: ledger.subgroup_name,
        subSubgroups: {},
        total: 0
      };
    }

    const subSubgroupKey = ledger.sub_subgroup_name;
    if (!section.subgroups[subgroupKey].subSubgroups[subSubgroupKey]) {
      section.subgroups[subgroupKey].subSubgroups[subSubgroupKey] = {
        sub_subgroup_name: ledger.sub_subgroup_name,
        ledgers: [],
        total: 0
      };
    }

    section.subgroups[subgroupKey].subSubgroups[subSubgroupKey].ledgers.push({
      id: ledger.id,
      ledger_code: ledger.ledger_code,
      ledger_name: ledger.ledger_name,
      opening_balance: ledger.opening_balance,
      closing_balance: ledger.closing_balance,
      adjusted_closing: ledger.adjusted_closing
    });

    section.subgroups[subgroupKey].subSubgroups[subSubgroupKey].total += ledger.adjusted_closing;
    section.subgroups[subgroupKey].total += ledger.adjusted_closing;
    section.total += ledger.adjusted_closing;
  }

  // Round totals
  for (const type of VALID_IE_PL_TYPES) {
    sections[type].total = Math.round(sections[type].total * 100) / 100;
    for (const sgk of Object.keys(sections[type].subgroups)) {
      sections[type].subgroups[sgk].total = Math.round(sections[type].subgroups[sgk].total * 100) / 100;
      for (const ssgk of Object.keys(sections[type].subgroups[sgk].subSubgroups)) {
        sections[type].subgroups[sgk].subSubgroups[ssgk].total = Math.round(sections[type].subgroups[sgk].subSubgroups[ssgk].total * 100) / 100;
      }
    }
  }

  return sections;
}

// GET /api/audit/:id/balance-sheet
router.get('/:id/balance-sheet', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const sections = buildFinancialSections(req.params.id);

    const total_assets = sections.Asset.total;
    const total_liabilities = sections.Liability.total;
    const total_equity = sections.Equity.total;
    const net_income = Math.round((sections.Income.total - sections.Expenditure.total) * 100) / 100;
    const liabilities_plus_equity = Math.round((total_liabilities + total_equity + net_income) * 100) / 100;
    const is_balanced = Math.abs(total_assets - liabilities_plus_equity) < 0.01;

    res.json({
      Asset: sections.Asset,
      Liability: sections.Liability,
      Equity: sections.Equity,
      Income: sections.Income,
      Expenditure: sections.Expenditure,
      summary: {
        total_assets,
        total_liabilities,
        total_equity,
        net_income,
        liabilities_plus_equity,
        is_balanced
      }
    });
  } catch (err) {
    console.error('Balance sheet error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/:id/pnl
router.get('/:id/pnl', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const sections = buildFinancialSections(req.params.id);

    const total_income = sections.Income.total;
    const total_expenditure = sections.Expenditure.total;
    const net_profit = Math.round((total_income - total_expenditure) * 100) / 100;

    res.json({
      Income: sections.Income,
      Expenditure: sections.Expenditure,
      summary: {
        total_income,
        total_expenditure,
        net_profit
      }
    });
  } catch (err) {
    console.error('P&L error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/financials/approve
router.post('/:id/financials/approve', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    db.prepare("UPDATE audit_engagements SET status = 'Financials Approved', updated_at = ? WHERE id = ?")
      .run(new Date().toISOString(), req.params.id);

    res.json({ success: true, status: 'Financials Approved' });
  } catch (err) {
    console.error('Financials approve error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// REPORTS
// ─────────────────────────────────────────────────────────────────────────────

// GET /api/audit/templates
router.get('/templates', (req, res) => {
  try {
    const templates = db.prepare(`
      SELECT rt.*, u.name AS created_by_name
      FROM report_templates rt
      LEFT JOIN users u ON rt.created_by = u.id
      ORDER BY rt.created_at DESC
    `).all();
    res.json(templates);
  } catch (err) {
    console.error('Templates list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/templates
router.post('/templates', (req, res) => {
  try {
    const { name, description, template_json } = req.body;

    if (!name || !template_json) {
      return res.status(400).json({ error: 'name and template_json are required' });
    }

    // Validate JSON
    let parsedJson;
    try {
      parsedJson = typeof template_json === 'string' ? JSON.parse(template_json) : template_json;
    } catch (parseErr) {
      return res.status(400).json({ error: 'template_json must be valid JSON' });
    }

    const jsonStr = typeof template_json === 'string' ? template_json : JSON.stringify(template_json);

    const result = db.prepare(`
      INSERT INTO report_templates (name, description, template_json, created_by)
      VALUES (?, ?, ?, ?)
    `).run(name, description || null, jsonStr, req.user.id);

    const template = db.prepare('SELECT * FROM report_templates WHERE id = ?').get(result.lastInsertRowid);
    res.status(201).json(template);
  } catch (err) {
    console.error('Template create error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/:id/report/preview
router.get('/:id/report/preview', (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const sections = buildFinancialSections(req.params.id);

    const total_income = sections.Income.total;
    const total_expenditure = sections.Expenditure.total;
    const net_profit = Math.round((total_income - total_expenditure) * 100) / 100;

    const total_assets = sections.Asset.total;
    const total_liabilities = sections.Liability.total;
    const total_equity = sections.Equity.total;
    const liabilities_plus_equity = Math.round((total_liabilities + total_equity + net_profit) * 100) / 100;
    const is_balanced = Math.abs(total_assets - liabilities_plus_equity) < 0.01;

    res.json({
      engagement: {
        id: engagement.id,
        client_name: engagement.client_name,
        financial_year: engagement.financial_year,
        period_start: engagement.period_start,
        period_end: engagement.period_end,
        status: engagement.status
      },
      balance_sheet: {
        Asset: sections.Asset,
        Liability: sections.Liability,
        Equity: sections.Equity,
        summary: { total_assets, total_liabilities, total_equity, liabilities_plus_equity, is_balanced }
      },
      profit_and_loss: {
        Income: sections.Income,
        Expenditure: sections.Expenditure,
        summary: { total_income, total_expenditure, net_profit }
      }
    });
  } catch (err) {
    console.error('Report preview error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/report/generate
router.post('/:id/report/generate', async (req, res) => {
  try {
    const engagement = db.prepare('SELECT * FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { template_id, report_type } = req.body;

    if (!report_type) {
      return res.status(400).json({ error: 'report_type is required' });
    }

    let template = null;
    if (template_id) {
      template = db.prepare('SELECT * FROM report_templates WHERE id = ?').get(template_id);
      if (!template) return res.status(404).json({ error: 'Template not found' });
    }

    // Build report data
    const sections = buildFinancialSections(req.params.id);
    const total_income = sections.Income.total;
    const total_expenditure = sections.Expenditure.total;
    const net_profit = Math.round((total_income - total_expenditure) * 100) / 100;
    const total_assets = sections.Asset.total;
    const total_liabilities = sections.Liability.total;
    const total_equity = sections.Equity.total;

    const reportData = {
      engagement: {
        client_name: engagement.client_name,
        financial_year: engagement.financial_year,
        period_start: engagement.period_start,
        period_end: engagement.period_end
      },
      balance_sheet: {
        Asset: sections.Asset,
        Liability: sections.Liability,
        Equity: sections.Equity,
        total_assets,
        total_liabilities,
        total_equity
      },
      profit_and_loss: {
        Income: sections.Income,
        Expenditure: sections.Expenditure,
        total_income,
        total_expenditure,
        net_profit
      },
      generated_at: new Date().toISOString()
    };

    // Generate DOCX
    const docChildren = [];

    // Title
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: `Audit Report - ${engagement.client_name}`, bold: true, size: 32 })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({
        children: [new TextRun({ text: `Financial Year: ${engagement.financial_year}`, size: 24 })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({
        children: [new TextRun({ text: `Period: ${engagement.period_start} to ${engagement.period_end}`, size: 20 })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ children: [new TextRun({ text: '' })] }) // spacer
    );

    // Helper to create financial section tables
    const addSectionTable = (title, section) => {
      docChildren.push(
        new Paragraph({
          children: [new TextRun({ text: title, bold: true, size: 26 })],
        }),
        new Paragraph({ children: [new TextRun({ text: '' })] })
      );

      const rows = [
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Account', bold: true })] })], width: { size: 60, type: WidthType.PERCENTAGE } }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: 'Amount', bold: true })] })], width: { size: 40, type: WidthType.PERCENTAGE } })
          ]
        })
      ];

      for (const sgk of Object.keys(section.subgroups)) {
        const subgroup = section.subgroups[sgk];
        rows.push(new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: subgroup.subgroup_name, bold: true })] })] }),
            new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: subgroup.total.toFixed(2) })] })] })
          ]
        }));

        for (const ssgk of Object.keys(subgroup.subSubgroups)) {
          const subSubgroup = subgroup.subSubgroups[ssgk];
          rows.push(new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `    ${subSubgroup.sub_subgroup_name}`, bold: true })] })] }),
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: subSubgroup.total.toFixed(2) })] })] })
            ]
          }));

          for (const ledger of subSubgroup.ledgers) {
            rows.push(new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `        ${ledger.ledger_code} - ${ledger.ledger_name}` })] })] }),
                new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: ledger.adjusted_closing.toFixed(2) })] })] })
              ]
            }));
          }
        }
      }

      // Total row
      rows.push(new TableRow({
        children: [
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: `Total ${title}`, bold: true })] })] }),
          new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: section.total.toFixed(2), bold: true })] })] })
        ]
      }));

      docChildren.push(new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }));
      docChildren.push(new Paragraph({ children: [new TextRun({ text: '' })] }));
    };

    // Add sections
    addSectionTable('Assets', sections.Asset);
    addSectionTable('Liabilities', sections.Liability);
    addSectionTable('Equity', sections.Equity);
    addSectionTable('Income', sections.Income);
    addSectionTable('Expenditure', sections.Expenditure);

    // Net Profit
    docChildren.push(
      new Paragraph({
        children: [new TextRun({ text: `Net Profit: ${net_profit.toFixed(2)}`, bold: true, size: 24 })]
      })
    );

    const doc = new Document({
      sections: [{ children: docChildren }]
    });

    const buffer = await Packer.toBuffer(doc);

    // Encrypt and save
    const uuid = crypto.randomUUID();
    const storedFilename = `audit_report_${engagement.id}_${uuid}.enc`;
    const encrypted = encryptFile(buffer);

    if (!fs.existsSync(VAULT_PATH)) {
      fs.mkdirSync(VAULT_PATH, { recursive: true });
    }
    fs.writeFileSync(path.join(VAULT_PATH, storedFilename), encrypted);

    // Create documents table record
    const now = new Date().toISOString();
    const docResult = db.prepare(`
      INSERT INTO documents (name, version, original_filename, stored_filename, category, subcategory, status, original_uploader_id, last_uploader_id, upload_date, month, year)
      VALUES (?, 1, ?, ?, 'Audit Reports', ?, 'Uploaded', ?, ?, ?, ?, ?)
    `).run(
      `Audit Report - ${engagement.client_name} - ${engagement.financial_year}`,
      `audit_report_${engagement.financial_year}.docx`,
      storedFilename,
      report_type,
      req.user.id,
      req.user.id,
      now,
      new Date().toLocaleString('default', { month: 'long' }),
      String(new Date().getFullYear())
    );

    // Create generated_reports record
    const reportResult = db.prepare(`
      INSERT INTO generated_reports (engagement_id, template_id, report_type, status, generated_by, stored_filename, report_data_json)
      VALUES (?, ?, ?, 'Draft', ?, ?, ?)
    `).run(
      req.params.id,
      template_id || null,
      report_type,
      req.user.id,
      storedFilename,
      JSON.stringify(reportData)
    );

    const generatedReport = db.prepare('SELECT * FROM generated_reports WHERE id = ?').get(reportResult.lastInsertRowid);

    res.status(201).json({
      report: generatedReport,
      document_id: docResult.lastInsertRowid,
      stored_filename: storedFilename
    });
  } catch (err) {
    console.error('Report generate error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// ENGAGEMENT AUDITORS
// ─────────────────────────────────────────────────────────────────────────────

// GET /api/audit/engagements/:id/auditors
router.get('/engagements/:id/auditors', (req, res) => {
  try {
    const engagement = db.prepare('SELECT id FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const auditors = db.prepare(`
      SELECT ea.id, ea.user_id, ea.assigned_at, u.name, u.username
      FROM engagement_auditors ea
      JOIN users u ON ea.user_id = u.id
      WHERE ea.engagement_id = ?
      ORDER BY u.name ASC
    `).all(req.params.id);

    res.json(auditors);
  } catch (err) {
    console.error('Get auditors error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/engagements/:id/auditors
// Body: { user_ids: [1, 2, 3] } — replaces all existing assignments
router.post('/engagements/:id/auditors', (req, res) => {
  try {
    const engagement = db.prepare('SELECT id FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { user_ids } = req.body;
    if (!Array.isArray(user_ids)) {
      return res.status(400).json({ error: 'user_ids must be an array' });
    }

    const now = new Date().toISOString();
    const tx = db.transaction(() => {
      db.prepare('DELETE FROM engagement_auditors WHERE engagement_id = ?').run(req.params.id);
      const insert = db.prepare(
        'INSERT OR IGNORE INTO engagement_auditors (engagement_id, user_id, assigned_at, assigned_by) VALUES (?, ?, ?, ?)'
      );
      for (const uid of user_ids) {
        insert.run(req.params.id, uid, now, req.user.id);
      }
    });
    tx();

    const updated = db.prepare(`
      SELECT ea.id, ea.user_id, ea.assigned_at, u.name, u.username
      FROM engagement_auditors ea
      JOIN users u ON ea.user_id = u.id
      WHERE ea.engagement_id = ?
      ORDER BY u.name ASC
    `).all(req.params.id);

    res.json(updated);
  } catch (err) {
    console.error('Set auditors error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// DELETE /api/audit/engagements/:id/auditors/:uid
router.delete('/engagements/:id/auditors/:uid', (req, res) => {
  try {
    db.prepare('DELETE FROM engagement_auditors WHERE engagement_id = ? AND user_id = ?')
      .run(req.params.id, req.params.uid);
    res.json({ success: true });
  } catch (err) {
    console.error('Remove auditor error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ─────────────────────────────────────────────────────────────────────────────
// AUDIT QUERIES
// ─────────────────────────────────────────────────────────────────────────────

const VALID_QUERY_TYPES = ['General', 'Document Request', 'Clarification', 'Observation'];

// GET /api/audit/:id/queries
router.get('/:id/queries', (req, res) => {
  try {
    const engagement = db.prepare('SELECT id FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const queries = db.prepare(`
      SELECT q.*,
             u.name AS raised_by_name,
             (SELECT COUNT(*) FROM audit_query_replies WHERE query_id = q.id) AS reply_count,
             (SELECT MAX(created_at) FROM audit_query_replies WHERE query_id = q.id) AS last_reply_at
      FROM audit_queries q
      JOIN users u ON q.raised_by = u.id
      WHERE q.engagement_id = ?
      ORDER BY q.updated_at DESC
    `).all(req.params.id);

    res.json(queries);
  } catch (err) {
    console.error('Queries list error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/queries
router.post('/:id/queries', (req, res) => {
  try {
    const engagement = db.prepare('SELECT id FROM audit_engagements WHERE id = ?').get(req.params.id);
    if (!engagement) return res.status(404).json({ error: 'Engagement not found' });

    const { query_type, subject, description } = req.body;

    if (!subject || !subject.trim()) {
      return res.status(400).json({ error: 'Subject is required' });
    }
    if (query_type && !VALID_QUERY_TYPES.includes(query_type)) {
      return res.status(400).json({ error: `Invalid query_type. Must be one of: ${VALID_QUERY_TYPES.join(', ')}` });
    }

    const now = new Date().toISOString();
    const result = db.prepare(`
      INSERT INTO audit_queries (engagement_id, raised_by, query_type, subject, description, status, created_at, updated_at)
      VALUES (?, ?, ?, ?, ?, 'Open', ?, ?)
    `).run(req.params.id, req.user.id, query_type || 'General', subject.trim(), description || null, now, now);

    const query = db.prepare(`
      SELECT q.*, u.name AS raised_by_name
      FROM audit_queries q
      JOIN users u ON q.raised_by = u.id
      WHERE q.id = ?
    `).get(result.lastInsertRowid);

    res.status(201).json(query);
  } catch (err) {
    console.error('Query create error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/:id/queries/:qid
router.get('/:id/queries/:qid', (req, res) => {
  try {
    const query = db.prepare(`
      SELECT q.*, u.name AS raised_by_name
      FROM audit_queries q
      JOIN users u ON q.raised_by = u.id
      WHERE q.id = ? AND q.engagement_id = ?
    `).get(req.params.qid, req.params.id);

    if (!query) return res.status(404).json({ error: 'Query not found' });

    const replies = db.prepare(`
      SELECT r.*, u.name AS sent_by_name, u.role AS sent_by_role
      FROM audit_query_replies r
      JOIN users u ON r.sent_by = u.id
      WHERE r.query_id = ?
      ORDER BY r.created_at ASC
    `).all(req.params.qid);

    for (const reply of replies) {
      if (reply.stored_filename) {
        reply.download_url = `/api/audit/${req.params.id}/queries/${req.params.qid}/replies/${reply.id}/download`;
      }
    }

    res.json({ ...query, replies });
  } catch (err) {
    console.error('Query get error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/audit/:id/queries/:qid/replies  (multipart)
router.post('/:id/queries/:qid/replies', upload.single('file'), (req, res) => {
  try {
    const query = db.prepare('SELECT * FROM audit_queries WHERE id = ? AND engagement_id = ?').get(req.params.qid, req.params.id);
    if (!query) return res.status(404).json({ error: 'Query not found' });

    const message = (req.body.message || '').trim() || null;
    let storedFilename = null;
    let originalFilename = null;

    if (!message && !req.file) {
      return res.status(400).json({ error: 'Reply must have a message or attachment' });
    }

    if (req.file) {
      const uuid = crypto.randomUUID();
      storedFilename = `query_reply_${req.params.qid}_${uuid}.enc`;
      originalFilename = req.file.originalname;
      const encrypted = encryptFile(req.file.buffer);
      if (!fs.existsSync(VAULT_PATH)) fs.mkdirSync(VAULT_PATH, { recursive: true });
      fs.writeFileSync(path.join(VAULT_PATH, storedFilename), encrypted);
    }

    const now = new Date().toISOString();
    const result = db.prepare(`
      INSERT INTO audit_query_replies (query_id, sent_by, message, stored_filename, original_filename, created_at)
      VALUES (?, ?, ?, ?, ?, ?)
    `).run(req.params.qid, req.user.id, message, storedFilename, originalFilename, now);

    db.prepare("UPDATE audit_queries SET updated_at = ? WHERE id = ?").run(now, req.params.qid);

    const reply = db.prepare(`
      SELECT r.*, u.name AS sent_by_name, u.role AS sent_by_role
      FROM audit_query_replies r
      JOIN users u ON r.sent_by = u.id
      WHERE r.id = ?
    `).get(result.lastInsertRowid);

    if (reply.stored_filename) {
      reply.download_url = `/api/audit/${req.params.id}/queries/${req.params.qid}/replies/${reply.id}/download`;
    }

    res.status(201).json(reply);
  } catch (err) {
    console.error('Reply create error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/queries/:qid/resolve
router.patch('/:id/queries/:qid/resolve', (req, res) => {
  try {
    if (req.user.role !== 'auditor') {
      return res.status(403).json({ error: 'Only auditors can resolve queries' });
    }
    const query = db.prepare('SELECT * FROM audit_queries WHERE id = ? AND engagement_id = ?').get(req.params.qid, req.params.id);
    if (!query) return res.status(404).json({ error: 'Query not found' });

    const now = new Date().toISOString();
    db.prepare("UPDATE audit_queries SET status = 'Resolved', updated_at = ? WHERE id = ?").run(now, req.params.qid);
    res.json({ success: true, status: 'Resolved' });
  } catch (err) {
    console.error('Query resolve error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/queries/:qid/reopen
router.patch('/:id/queries/:qid/reopen', (req, res) => {
  try {
    if (req.user.role !== 'auditor') {
      return res.status(403).json({ error: 'Only auditors can reopen queries' });
    }
    const query = db.prepare('SELECT * FROM audit_queries WHERE id = ? AND engagement_id = ?').get(req.params.qid, req.params.id);
    if (!query) return res.status(404).json({ error: 'Query not found' });

    const now = new Date().toISOString();
    db.prepare("UPDATE audit_queries SET status = 'Open', updated_at = ? WHERE id = ?").run(now, req.params.qid);
    res.json({ success: true, status: 'Open' });
  } catch (err) {
    console.error('Query reopen error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// PATCH /api/audit/:id/queries/:qid/close
router.patch('/:id/queries/:qid/close', (req, res) => {
  try {
    const query = db.prepare('SELECT * FROM audit_queries WHERE id = ? AND engagement_id = ?').get(req.params.qid, req.params.id);
    if (!query) return res.status(404).json({ error: 'Query not found' });

    const now = new Date().toISOString();
    db.prepare("UPDATE audit_queries SET status = 'Closed', updated_at = ? WHERE id = ?").run(now, req.params.qid);
    res.json({ success: true, status: 'Closed' });
  } catch (err) {
    console.error('Query close error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/audit/:id/queries/:qid/replies/:rid/download
router.get('/:id/queries/:qid/replies/:rid/download', (req, res) => {
  try {
    const reply = db.prepare('SELECT * FROM audit_query_replies WHERE id = ? AND query_id = ?').get(req.params.rid, req.params.qid);
    if (!reply || !reply.stored_filename) return res.status(404).json({ error: 'Attachment not found' });

    const filePath = path.join(VAULT_PATH, reply.stored_filename);
    if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found on disk' });

    const encrypted = fs.readFileSync(filePath);
    const decrypted = decryptFile(encrypted);

    res.setHeader('Content-Disposition', `attachment; filename="${reply.original_filename}"`);
    res.setHeader('Content-Type', 'application/octet-stream');
    res.setHeader('Content-Length', decrypted.length);
    res.send(decrypted);
  } catch (err) {
    console.error('Reply download error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

module.exports = router;
