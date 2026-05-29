/**
 * Menulis app.config.js dari environment variables.
 * Dipakai saat build Vercel (hanya butuh SUPABASE_URL + SUPABASE_ANON_KEY).
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');

const supabaseUrl = process.env.SUPABASE_URL || '';
const supabaseAnonKey = process.env.SUPABASE_ANON_KEY || '';

if (!supabaseUrl || !supabaseAnonKey) {
  console.error('\n❌ Set SUPABASE_URL dan SUPABASE_ANON_KEY');
  console.error('   Lokal: isi di file .env');
  console.error('   Vercel: Settings → Environment Variables\n');
  process.exit(1);
}

const outDir = path.join(__dirname, '../public/js/config');
const outFile = path.join(outDir, 'app.config.js');
const js = `window.APP_CONFIG = ${JSON.stringify({
  SUPABASE_URL: supabaseUrl,
  SUPABASE_ANON_KEY: supabaseAnonKey,
})};\n`;

fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(outFile, js, 'utf8');
console.log('✓ Generated', outFile);
