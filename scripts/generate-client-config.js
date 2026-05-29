/**
 * Menulis app.config.js dari environment variables.
 * Vercel: jika env belum diset saat build, tulis placeholder (runtime pakai /api/app-config).
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');

const supabaseUrl = process.env.SUPABASE_URL || '';
const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || '';
const isVercel = Boolean(process.env.VERCEL);

const outDir = path.join(__dirname, '../public/js/config');
const outFile = path.join(outDir, 'app.config.js');

if (!supabaseUrl || !supabaseAnonKey) {
  if (isVercel) {
    console.warn('\n⚠ SUPABASE_URL / SUPABASE_ANON_KEY belum diset di Vercel Environment Variables.');
    console.warn('  Build tetap lanjut. WAJIB isi env vars agar add-in berfungsi (config dilayani via /api/app-config).\n');
    fs.mkdirSync(outDir, { recursive: true });
    fs.writeFileSync(outFile, 'window.APP_CONFIG = {};\n', 'utf8');
    const manifestSrc = path.join(__dirname, '../manifest.xml');
    const manifestDest = path.join(__dirname, '../public/manifest.xml');
    if (fs.existsSync(manifestSrc)) fs.copyFileSync(manifestSrc, manifestDest);
    process.exit(0);
  }

  console.error('\n❌ Set SUPABASE_URL dan SUPABASE_ANON_KEY');
  console.error('   Lokal: isi di file .env');
  console.error('   Vercel: Settings → Environment Variables\n');
  process.exit(1);
}

const js = `window.APP_CONFIG = ${JSON.stringify({
  SUPABASE_URL: supabaseUrl,
  SUPABASE_ANON_KEY: supabaseAnonKey,
})};\n`;

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, js, 'utf8');
console.log('✓ Generated', outFile);

// Salin manifest.xml ke public/ agar bisa diakses via URL (sideload web catalog)
const manifestSrc = path.join(__dirname, '../manifest.xml');
const manifestDest = path.join(__dirname, '../public/manifest.xml');
if (fs.existsSync(manifestSrc)) {
  fs.copyFileSync(manifestSrc, manifestDest);
  console.log('✓ Copied manifest.xml → public/manifest.xml');
}
