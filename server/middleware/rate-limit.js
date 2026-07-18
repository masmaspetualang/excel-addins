/**
 * server/middleware/rate-limit.js
 * ────────────────────────────────
 * Rate limiter sederhana in-memory.
 * Membatasi setiap IP agar tidak bisa spam request
 * lebih dari MAX_REQUESTS dalam window waktu tertentu.
 *
 * Tidak butuh package eksternal (express-rate-limit).
 */
const WINDOW_MS   = 15 * 60 * 1000; // 15 menit
const MAX_REQUESTS = 300;            // max request per window

const store = {}; // { ip: { count, resetAt } }

// Bersihkan store tiap 15 menit agar tidak memory leak
setInterval(() => {
  const now = Date.now();
  for (const ip in store) {
    if (store[ip].resetAt < now) delete store[ip];
  }
}, WINDOW_MS);

function rateLimitMiddleware(req, res, next) {
  // Bypass rate limiting in local development
  if (process.env.NODE_ENV !== 'production' && !process.env.VERCEL) {
    return next();
  }

  const ip  = req.headers['x-forwarded-for'] || req.socket?.remoteAddress || 'unknown';
  const now = Date.now();

  if (!store[ip] || store[ip].resetAt < now) {
    store[ip] = { count: 1, resetAt: now + WINDOW_MS };
  } else {
    store[ip].count++;
  }

  if (store[ip].count > MAX_REQUESTS) {
    res.writeHead(429, { 'Content-Type': 'text/plain' });
    res.end('429 Too Many Requests — Coba lagi dalam 15 menit.');
    return;
  }

  next();
}

module.exports = rateLimitMiddleware;
