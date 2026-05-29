/**
 * server/app.js
 * ───────────────────
 * Konfigurasi utama Express app.
 * Menangani middleware, file statik, routing, dan error handling.
 */
const express = require('express');
const path = require('path');
const corsMiddleware = require('./middleware/cors');
const rateLimitMiddleware = require('./middleware/rate-limit');
const { notFoundHandler, serverErrorHandler } = require('./middleware/error');
const logger = require('./utils/logger');
const MIME_TYPES = require('./config/mime-types.json');

const app = express();

// Daftarkan mime types tambahan untuk support file Office (docx, xlsx, pptx)
if (express.static.mime && typeof express.static.mime.define === 'function') {
  express.static.mime.define(MIME_TYPES, true);
}

// Request logging middleware
app.use((req, res, next) => {
  res.on('finish', () => {
    logger.request(req, res.statusCode);
  });
  next();
});

// Security & rate limiting middleware
app.use(corsMiddleware);
app.use(rateLimitMiddleware);

// URL Redirect untuk default root
app.get('/', (req, res) => {
  res.redirect('/pages/participant/taskpane.html');
});

// Konfigurasi client dari .env (browser tidak bisa baca .env langsung)
const { toJavaScript } = require('./config/client-config');
app.get('/js/config/app.config.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.send(toJavaScript());
});
// Alias lama — hindari 404 jika HTML/cache masih memuat path lama
app.get('/js/config/supabase.config.js', (req, res) => {
  res.redirect(301, '/js/config/app.config.js');
});

// Backward compatibility redirect (jika manifest lama masih tersimpan di cache Office)
app.get('/taskpane.html', (req, res) => {
  res.redirect('/pages/participant/taskpane.html');
});
app.get('/login.html', (req, res) => {
  res.redirect('/pages/participant/login.html');
});
app.get('/dashboard.html', (req, res) => {
  res.redirect('/pages/admin/dashboard.html');
});
app.get('/admin-login.html', (req, res) => {
  res.redirect('/pages/admin/admin-login.html');
});
app.get('/commands.html', (req, res) => {
  res.redirect('/pages/admin/commands.html');
});

// Serve static files dari public directory
app.use(express.static(path.join(__dirname, '../public')));

// Not Found Handler
app.use(notFoundHandler);

// Centralized Error Handler (500)
app.use(serverErrorHandler);

module.exports = app;
