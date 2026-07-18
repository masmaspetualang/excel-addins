/**
 * server/config/env.js
 * Satu sumber kebenaran untuk environment variables.
 */
require('dotenv').config();

const isVercel = Boolean(process.env.VERCEL);
const REQUIRED = isVercel
  ? ['SUPABASE_URL', 'SUPABASE_ANON_KEY']
  : ['PORT', 'HOST', 'SUPABASE_URL', 'SUPABASE_ANON_KEY'];

const missing = REQUIRED.filter((key) => !process.env[key]);
if (missing.length > 0) {
  console.error('\n❌ KONFIGURASI ERROR: Variable berikut belum diisi di .env:');
  missing.forEach((k) => console.error(`   → ${k}`));
  console.error('\nSalin .env.example ke .env lalu isi nilainya.\n');
  process.exit(1);
}

module.exports = {
  port: parseInt(process.env.PORT || '3000', 10),
  host: process.env.HOST || 'localhost',
  supabaseUrl: process.env.SUPABASE_URL,
  supabaseAnonKey: process.env.SUPABASE_ANON_KEY,
  supabaseServiceKey: process.env.SUPABASE_SERVICE_KEY || '',
  nodeEnv: process.env.NODE_ENV || 'development',
  isDev: (process.env.NODE_ENV || 'development') === 'development',
};
