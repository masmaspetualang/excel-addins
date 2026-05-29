/**
 * server/app.js — Express app (development lokal)
 */
const express = require('express');
const path = require('path');
const corsMiddleware = require('./middleware/cors');
const rateLimitMiddleware = require('./middleware/rate-limit');
const { notFoundHandler, serverErrorHandler } = require('./middleware/error');
const logger = require('./utils/logger');
const MIME_TYPES = require('./config/mime-types.json');
const { FILES, URLS, LEGACY_REDIRECTS } = require('./config/routes');

const app = express();
const publicDir = path.join(__dirname, '../public');

if (express.static.mime && typeof express.static.mime.define === 'function') {
  express.static.mime.define(MIME_TYPES, true);
}

app.use((req, res, next) => {
  res.on('finish', () => logger.request(req, res.statusCode));
  next();
});

app.use(corsMiddleware);
app.use(rateLimitMiddleware);

function servePage(relativePath) {
  return (_req, res) => res.sendFile(path.join(publicDir, relativePath));
}

// Root → app utama
app.get('/', (_req, res) => res.redirect(302, URLS.app));

// URL profesional
app.get(URLS.app, servePage(FILES.app));
app.get(URLS.login, servePage(FILES.login));
app.get(URLS.admin, servePage(FILES.admin));
app.get(URLS.adminLogin, servePage(FILES.adminLogin));
app.get(URLS.adminCommands, servePage(FILES.adminCommands));

// Config client dari .env
const { toJavaScript } = require('./config/client-config');
app.get('/js/config/app.config.js', (_req, res) => {
  res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.send(toJavaScript());
});
app.get('/js/config/supabase.config.js', (_req, res) => {
  res.redirect(301, '/js/config/app.config.js');
});

// Redirect path lama → URL baru
for (const [from, to] of Object.entries(LEGACY_REDIRECTS)) {
  app.get(from, (req, res) => {
    const qs = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
    res.redirect(301, to + qs);
  });
}

app.use(express.static(publicDir));
app.use(notFoundHandler);
app.use(serverErrorHandler);

module.exports = app;
