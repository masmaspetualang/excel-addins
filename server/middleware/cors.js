/**
 * server/middleware/cors.js
 * ──────────────────────────
 * Mengatur CORS headers yang diperlukan oleh Office.js.
 * Office Add-in mengakses file dari browser bawaan Office,
 * sehingga CORS wajib diizinkan.
 */
function corsMiddleware(req, res, next) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');

  // Tangani preflight request dari browser
  if (req.method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  next();
}

module.exports = corsMiddleware;
