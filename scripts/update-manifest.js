/**
 * Ganti URL localhost di manifest.xml dengan domain production (Vercel).
 *
 * Usage:
 *   node scripts/update-manifest.js https://nama-project.vercel.app
 *   npm run manifest:set -- https://nama-project.vercel.app
 */
const fs = require('fs');
const path = require('path');

const manifestPath = path.join(__dirname, '../manifest.xml');
const rawUrl = process.argv[2];

if (!rawUrl) {
  console.error('\n❌ Masukkan URL deployment Vercel.\n');
  console.error('   Contoh:');
  console.error('   npm run manifest:set -- https://excel-quiz-pro.vercel.app\n');
  process.exit(1);
}

let baseUrl = rawUrl.trim().replace(/\/$/, '');
if (!/^https:\/\//i.test(baseUrl)) {
  console.error('❌ URL harus diawali https:// (contoh: https://nama-project.vercel.app)');
  process.exit(1);
}

const content = fs.readFileSync(manifestPath, 'utf8');

// Ganti localhost dev dan placeholder lama (jika pernah di-set sebelumnya)
const patterns = [
  /https:\/\/localhost:3000/g,
  /https:\/\/YOUR_VERCEL_URL/g,
  /https:\/\/YOUR_DOMAIN_HERE/g,
];

let updated = content;
let totalReplaced = 0;
for (const pattern of patterns) {
  const matches = updated.match(pattern);
  if (matches) totalReplaced += matches.length;
  updated = updated.replace(pattern, baseUrl);
}

if (totalReplaced === 0 && !content.includes('localhost:3000')) {
  // Sudah production — ganti semua URL https://... kecuali microsoft link
  const urlRegex = /DefaultValue="(https:\/\/(?!go\.microsoft\.com)[^"]+)"/g;
  updated = content.replace(urlRegex, (match, url) => {
    if (url.startsWith(baseUrl)) return match;
    return `DefaultValue="${baseUrl}${new URL(url).pathname}"`;
  });
}

fs.writeFileSync(manifestPath, updated, 'utf8');

console.log('\n✅ manifest.xml diperbarui');
console.log(`   Base URL: ${baseUrl}`);
console.log(`   Taskpane: ${baseUrl}/pages/participant/taskpane.html`);
console.log(`   Commands: ${baseUrl}/pages/admin/commands.html`);
console.log('\nLangkah berikutnya: sideload manifest.xml di Excel (Insert → My Add-ins → Upload).\n');
