/**
 * Promote user ke admin
 * Usage: node scripts/promote-admin.js user@email.com
 */
const { createSupabaseAdmin } = require('../server/lib/supabase-admin');

async function promote(email) {
  const supabase = createSupabaseAdmin();
  console.log(`Promoting ${email} to admin...`);

  const { data: { users }, error: userError } = await supabase.auth.admin.listUsers();
  if (userError) throw userError;

  const user = users.find((u) => u.email === email);
  if (!user) {
    console.error(`User with email ${email} not found.`);
    return;
  }

  const { error: profileError } = await supabase
    .from('profiles')
    .update({ role: 'admin' })
    .eq('id', user.id);

  if (profileError) throw profileError;

  console.log(`Successfully promoted ${email} to admin! ✓`);
}

const email = process.argv[2];
if (!email) {
  console.log('Usage: node scripts/promote-admin.js your@email.com');
  process.exit(1);
}

promote(email).catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
