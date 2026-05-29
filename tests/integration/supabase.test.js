require('dotenv').config();
const test = require('node:test');
const assert = require('node:assert');
const { createClient } = require('@supabase/supabase-js');

test('Integration Test: Supabase Database Connection', async (t) => {
  const url = process.env.SUPABASE_URL;
  const key = process.env.SUPABASE_SERVICE_KEY;

  assert.ok(url, 'SUPABASE_URL harus diset di .env');
  assert.ok(key, 'SUPABASE_SERVICE_KEY harus diset di .env');

  const supabase = createClient(url, key);
  
  // Ambil berkas template untuk menguji query
  const { data, error } = await supabase
    .from('berkas_template')
    .select('*')
    .limit(1);

  assert.strictEqual(error, null, `Query ke berkas_template gagal: ${error ? error.message : ''}`);
  assert.ok(Array.isArray(data), 'Data yang dikembalikan harus berupa array');
  console.log('✓ Supabase connection integration test passed. Data length:', data.length);
});
