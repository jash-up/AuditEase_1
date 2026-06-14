'use strict';
require('dotenv').config();
const express = require('express');
const path = require('path');
const db = require('./db');
const { requireAuth } = require('./auth');
const authRoutes = require('./routes/auth.routes');
const documentsRoutes = require('./routes/documents.routes');
const fs = require('fs');

const vaultDir = path.join(__dirname, 'storage', 'vault');
if (!fs.existsSync(vaultDir)) {
  fs.mkdirSync(vaultDir, { recursive: true });
  console.log('[AuditEase] Created storage/vault directory');
}

const app = express();
const PORT = process.env.PORT || 3000;

// ── Middleware ───────────────────────────────────────────────────────────────
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve frontend static files
app.use(express.static(path.join(__dirname, '..', 'frontend')));

// ── API Routes ───────────────────────────────────────────────────────────────
app.use('/api/auth', authRoutes);
app.use('/api/documents', documentsRoutes);

// ── GET /api/users ───────────────────────────────────────────────────────────
app.get('/api/users', requireAuth, (req, res) => {
  try {
    const users = db.prepare('SELECT id, name, username FROM users ORDER BY name ASC').all();
    res.json(users);
  } catch (err) {
    console.error('Users error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── Page Visit Tracking ──────────────────────────────────────────────────────
// POST /api/visits
app.post('/api/visits', requireAuth, (req, res) => {
  try {
    const { page_key, page_label, page_url } = req.body;
    if (!page_key || !page_label || !page_url) {
      return res.status(400).json({ error: 'page_key, page_label, page_url required' });
    }
    const now = new Date().toISOString();
    db.prepare(`
      INSERT INTO user_page_visits (user_id, page_key, page_label, page_url, visit_count, last_visited)
      VALUES (?, ?, ?, ?, 1, ?)
      ON CONFLICT(user_id, page_key) DO UPDATE SET
        visit_count = visit_count + 1,
        last_visited = excluded.last_visited,
        page_label = excluded.page_label,
        page_url = excluded.page_url
    `).run(req.user.id, page_key, page_label, page_url, now);
    res.json({ success: true });
  } catch (err) {
    console.error('Visit tracking error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// GET /api/visits/top
app.get('/api/visits/top', requireAuth, (req, res) => {
  try {
    const visits = db.prepare(`
      SELECT page_key, page_label, page_url, visit_count, last_visited
      FROM user_page_visits
      WHERE user_id = ?
      ORDER BY visit_count DESC, last_visited DESC
      LIMIT 5
    `).all(req.user.id);
    res.json(visits);
  } catch (err) {
    console.error('Top visits error:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

// ── Catch-all for SPA routing ─────────────────────────────────────────────────
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'frontend', 'index.html'));
});

// ── Start server ─────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n🚀 AuditEase server running at http://localhost:${PORT}`);
  console.log(`   Environment: ${process.env.NODE_ENV || 'development'}`);
  console.log(`   Database: auditease.db\n`);
});

module.exports = app;
