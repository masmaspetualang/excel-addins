/**
 * server/middleware/error.js
 * ───────────────────────────
 * Menangani semua error yang tidak tertangani di middleware lain.
 * - 404: file tidak ditemukan
 * - 500: server error internal
 *
 * Selalu kembalikan response yang aman (jangan expose stack trace di production).
 */
const logger = require('../utils/logger');

function notFoundHandler(req, res) {
  logger.warn(`404 — ${req.method} ${req.url}`);
  res.writeHead(404, { 'Content-Type': 'text/plain' });
  res.end(`404 Not Found: ${req.url}`);
}

function serverErrorHandler(err, req, res) {
  logger.error(`500 — ${req.method} ${req.url}`, err);

  res.writeHead(500, { 'Content-Type': 'text/plain' });
  res.end('500 Internal Server Error');
}

module.exports = { notFoundHandler, serverErrorHandler };
