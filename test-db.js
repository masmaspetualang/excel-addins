const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const env = require('./server/config/env');

async function test() {
  try {
    let supabase;
    let users = [];

    try {
      const { createSupabaseAdmin } = require('./server/lib/supabase-admin');
      supabase = createSupabaseAdmin();
      
      const { data, error: authErr } = await supabase.auth.admin.listUsers({ perPage: 1000 });
      if (authErr) {
        console.warn('[API] Warning: Failed to list auth users via admin client:', authErr.message);
      } else {
        users = data.users || [];
      }
    } catch (adminErr) {
      console.warn('[API] Warning: Cannot initialize admin client or fetch auth users. Falling back to anon client. Error:', adminErr.message);
      supabase = createClient(env.supabaseUrl, env.supabaseAnonKey);
    }

    // Fetch all participants from pengguna table
    const { data: participants, error: dbErr } = await supabase
      .from('pengguna')
      .select('*')
      .eq('peran', 'participant')
      .order('nama_lengkap', { ascending: true });
    
    if (dbErr) {
      console.error('Database query error:', dbErr);
      throw dbErr;
    }

    console.log('Participants count:', participants ? participants.length : 0);
    console.log('First participant:', participants ? participants[0] : null);
  } catch (err) {
    console.error('Outer Catch Error:', err);
  }
}

test();
