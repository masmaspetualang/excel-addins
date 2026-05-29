/**
 * Script to promote a user to Admin role
 * Usage: node scripts/promote-admin.js user@email.com
 */
require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

const supabaseUrl = process.env.SUPABASE_URL;
const supabaseServiceKey = process.env.SUPABASE_SERVICE_KEY; // Use Service Role Key

if (!supabaseUrl || !supabaseServiceKey) {
  console.error('Error: SUPABASE_URL and SUPABASE_SERVICE_KEY must be set in .env');
  process.exit(1);
}

const supabase = createClient(supabaseUrl, supabaseServiceKey);

async function promote(email) {
  console.log(`Promoting ${email} to admin...`);

  // 1. Get user by email
  const { data: { users }, error: userError } = await supabase.auth.admin.listUsers();
  if (userError) throw userError;

  const user = users.find(u => u.email === email);
  if (!user) {
    console.error(`User with email ${email} not found.`);
    return;
  }

  // 2. Update profile
  const { error: profileError } = await supabase
    .from('profiles')
    .update({ role: 'admin' })
    .eq('id', user.id);

  if (profileError) throw profileError;

  console.log(`Successfully promoted ${email} to admin! ✓`);
}

const email = process.argv[2];
if (!email) {
  console.log('Please provide an email: node scripts/promote-admin.js your@email.com');
} else {
  promote(email).catch(err => console.error(err));
}
