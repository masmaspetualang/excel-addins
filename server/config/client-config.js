/**
 * Konfigurasi client (browser) — satu sumber dari env, tanpa file config statis terpisah.
 */
const env = require('./env');

function getClientConfig() {
  return {
    SUPABASE_URL: env.supabaseUrl,
    SUPABASE_ANON_KEY: env.supabaseAnonKey,
  };
}

function toJavaScript() {
  return `window.APP_CONFIG = ${JSON.stringify(getClientConfig())};`;
}

module.exports = { getClientConfig, toJavaScript };
