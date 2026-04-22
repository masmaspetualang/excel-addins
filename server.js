/**
 * Excel Practice Quiz — Local Dev Server
 * Serves add-in over HTTPS (required by Office.js)
 *
 * Usage:
 *   1. npm install
 *   2. Generate self-signed cert (see README)
 *   3. node server.js
 */

require("dotenv").config();

const https = require("https");
const fs = require("fs");
const path = require("path");

const PORT = Number.parseInt(process.env.PORT || "3000", 10);
const HOST = process.env.HOST || "localhost";
const SSL_KEY_PATH =
  process.env.SSL_KEY_PATH || path.join(__dirname, "certs", "server.key");
const SSL_CERT_PATH =
  process.env.SSL_CERT_PATH || path.join(__dirname, "certs", "server.crt");

// MIME types
const MIME = {
  ".html": "text/html",
  ".js": "application/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".xml": "application/xml",
  ".json": "application/json",
};

// Try to load SSL certs; fallback message if not present
let sslOptions;
try {
  sslOptions = {
    key: fs.readFileSync(SSL_KEY_PATH),
    cert: fs.readFileSync(SSL_CERT_PATH),
  };
} catch {
  console.error("⚠️  SSL certs not found.");
  console.error(`   SSL_KEY_PATH:  ${SSL_KEY_PATH}`);
  console.error(`   SSL_CERT_PATH: ${SSL_CERT_PATH}`);
  console.error("   Buat sertifikat self-signed (lihat README) atau set env var di .env");
  process.exit(1);
}

const server = https.createServer(sslOptions, (req, res) => {
  // CORS headers (needed for Office.js)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") { res.writeHead(204); res.end(); return; }

  let urlPath = req.url.split("?")[0];
  if (urlPath === "/" || urlPath === "") urlPath = "/taskpane.html";

  const filePath = path.join(__dirname, urlPath);
  const ext = path.extname(filePath).toLowerCase();
  const contentType = MIME[ext] || "application/octet-stream";

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain" });
      res.end(`404 Not Found: ${urlPath}`);
      return;
    }
    res.writeHead(200, { "Content-Type": contentType });
    res.end(data);
  });
});

server.listen(PORT, () => {
  const baseUrl = `https://${HOST}:${PORT}`;
  console.log(`\n✅ Excel Quiz Add-in Server running at ${baseUrl}`);
  console.log(`   Taskpane: ${baseUrl}/taskpane.html`);
  console.log(`   Manifest: ${baseUrl}/manifest.xml\n`);
});

server.on("error", (err) => {
  if (err && err.code === "EADDRINUSE") {
    console.error(`\n❌ Port ${PORT} sedang dipakai.`);
    console.error("   Ubah PORT di file .env lalu jalankan lagi: npm run dev\n");
    process.exit(1);
  }
  throw err;
});
