/**
 * server/utils/logger.js
 * ───────────────────────
 * Logger ringan tanpa dependensi eksternal.
 * Output ke console (dev) dan file /logs/ (semua env).
 *
 * Cara pakai:
 *   const logger = require('./utils/logger');
 *   logger.info('Server started');
 *   logger.error('Something broke', error);
 */
const fs   = require('fs');
const path = require('path');

const LOG_DIR = path.join(__dirname, '../../logs');

// Buat folder logs jika belum ada
if (!fs.existsSync(LOG_DIR)) {
  fs.mkdirSync(LOG_DIR, { recursive: true });
}

function formatLine(level, msg) {
  return `[${new Date().toISOString()}] [${level}] ${msg}`;
}

function writeToFile(filename, line) {
  fs.appendFile(path.join(LOG_DIR, filename), line + '\n', () => {});
}

const logger = {
  info(msg) {
    const line = formatLine('INFO', msg);
    console.log('\x1b[32m%s\x1b[0m', line); // hijau
    writeToFile('combined.log', line);
  },

  warn(msg) {
    const line = formatLine('WARN', msg);
    console.warn('\x1b[33m%s\x1b[0m', line); // kuning
    writeToFile('combined.log', line);
  },

  error(msg, err) {
    const detail = err ? `\n  ${err.stack || err.message || err}` : '';
    const line   = formatLine('ERROR', msg + detail);
    console.error('\x1b[31m%s\x1b[0m', line); // merah
    writeToFile('error.log', line);
    writeToFile('combined.log', line);
  },

  // Untuk logging setiap request HTTP
  request(req, statusCode) {
    const line = formatLine('REQ', `${statusCode} ${req.method} ${req.url} — ${req.ip}`);
    console.log(line);
    writeToFile('combined.log', line);
  },
};

module.exports = logger;
