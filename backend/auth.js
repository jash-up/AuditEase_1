'use strict';
const jwt = require('jsonwebtoken');

/**
 * Express middleware that requires a valid JWT token.
 * Attaches decoded user data to req.user.
 */
function requireAuth(req, res, next) {
  try {
    const header = req.headers.authorization;
    if (!header?.startsWith('Bearer ')) return res.status(401).json({ error: 'Not authenticated.' });
    const token = header.slice(7);
    if (!process.env.JWT_SECRET) throw new Error('JWT_SECRET not set in environment.');
    req.user = jwt.verify(token, process.env.JWT_SECRET);
    next();
  } catch (err) {
    return res.status(401).json({ error: 'Invalid or expired session. Please sign in again.' });
  }
}

module.exports = { requireAuth };
