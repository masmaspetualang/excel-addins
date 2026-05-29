/**
 * Klien Supabase dengan service role — hanya untuk script server (seed, promote-admin).
 */
const { createClient } = require('@supabase/supabase-js');
const env = require('../config/env');

function createSupabaseAdmin() {
  if (!env.supabaseServiceKey) {
    throw new Error('SUPABASE_SERVICE_KEY belum diset di .env (diperlukan untuk script admin).');
  }
  return createClient(env.supabaseUrl, env.supabaseServiceKey);
}

module.exports = { createSupabaseAdmin };
