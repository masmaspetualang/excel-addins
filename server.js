/**
 * Excel Practice Quiz — Local Dev Server Entrypoint
 * Bootstraps the Express app and serves it over HTTPS (required by Office.js).
 */
const https = require('https');
const fs = require('fs');
const path = require('path');
const env = require('./server/config/env');
const logger = require('./server/utils/logger');
const app = require('./server/app');

const SSL_KEY_PATH = process.env.SSL_KEY_PATH || path.join(__dirname, 'certs', 'server.key');
const SSL_CERT_PATH = process.env.SSL_CERT_PATH || path.join(__dirname, 'certs', 'server.crt');

let sslOptions;
try {
  sslOptions = {
    key: fs.readFileSync(SSL_KEY_PATH),
    cert: fs.readFileSync(SSL_CERT_PATH),
  };
} catch (err) {
  logger.error('Gagal memuat sertifikat SSL! Pastikan file key dan cert sudah ada di folder /certs.', err);
  process.exit(1);
}

const server = https.createServer(sslOptions, app);

server.listen(env.port, env.host, () => {
  const baseUrl = `https://${env.host}:${env.port}`;
  logger.info(`Excel Quiz Add-in Server running at ${baseUrl}`);
  logger.info(`App:     ${baseUrl}/app`);
  logger.info(`Admin:   ${baseUrl}/admin`);
  logger.info(`Manifest: ${baseUrl}/manifest.xml`);
});

server.on('error', (err) => {
  if (err && err.code === 'EADDRINUSE') {
    logger.error(`Port ${env.port} sedang dipakai. Ubah PORT di file .env lalu jalankan lagi.`);
    process.exit(1);
  }
  throw err;
});
