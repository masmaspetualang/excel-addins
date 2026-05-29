/**
 * Ganti URL localhost di manifest.xml dengan domain production (Vercel).
 *
 * Usage:
 *   node scripts/update-manifest.js https://nama-project.vercel.app
 *   npm run manifest:set -- https://nama-project.vercel.app
 */
const fs = require('fs');
const path = require('path');
const { MANIFEST_PATH_MAP, URLS } = require('../server/config/routes');

const manifestPath = path.join(__dirname, '../manifest.xml');
const rawUrl = process.argv[2];

if (!rawUrl) {
  console.error('\n❌ Masukkan URL deployment Vercel.\n');
  console.error('   Contoh:');
  console.error('   npm run manifest:set -- https://quiz-addins-v1.vercel.app\n');
  process.exit(1);
}

let baseUrl = rawUrl.trim().replace(/\/$/, '');
if (!/^https:\/\//i.test(baseUrl)) {
  console.error('❌ URL harus diawali https:// (contoh: https://nama-project.vercel.app)');
  process.exit(1);
}

let updated = fs.readFileSync(manifestPath, 'utf8');

// Ganti base URL dev / placeholder
const hostPatterns = [
  /https:\/\/localhost:3000/g,
  /https:\/\/YOUR_VERCEL_URL/g,
  /https:\/\/YOUR_DOMAIN_HERE/g,
];
for (const pattern of hostPatterns) {
  updated = updated.replace(pattern, baseUrl);
}

// Normalisasi path lama → path profesional
for (const [oldPath, newPath] of Object.entries(MANIFEST_PATH_MAP)) {
  updated = updated.split(baseUrl + oldPath).join(baseUrl + newPath);
}

fs.writeFileSync(manifestPath, updated, 'utf8');

console.log('\n✅ manifest.xml diperbarui');
console.log(`   Base URL: ${baseUrl}`);
console.log(`   App:      ${baseUrl}${URLS.app}`);
console.log(`   Login:    ${baseUrl}${URLS.login}`);
console.log(`   Admin:    ${baseUrl}${URLS.admin}`);
console.log(`   Commands: ${baseUrl}${URLS.adminCommands}`);
console.log('\nLangkah berikutnya: sideload manifest.xml di Excel (Insert → My Add-ins → Upload).\n');
