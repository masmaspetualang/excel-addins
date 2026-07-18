const fs = require('fs');
const path = require('path');
const test = require('node:test');
const assert = require('node:assert');

// Mock browser globals untuk menghindari error "window is not defined"
global.window = {};

const checkersDir = path.join(__dirname, '../../public/features/participant-exam/checkers');
const filesToLoad = ['helpers.js', 'excel.js', 'word.js', 'ppt.js', 'index.js'];

filesToLoad.forEach(file => {
  const code = fs.readFileSync(path.join(checkersDir, file), 'utf8');
  const runCode = new Function('window', code);
  runCode(global.window);
});

const OfficeCheckers = global.window.OfficeCheckers;

test('Unit Test: OfficeCheckers._isRed', (t) => {
  // Test warna merah standar (hex 6 digit)
  assert.strictEqual(OfficeCheckers._isRed('#ff0000'), true, '#ff0000 harus bernilai true');
  assert.strictEqual(OfficeCheckers._isRed('#e61919'), true, '#e61919 (merah tua) harus bernilai true');
  assert.strictEqual(OfficeCheckers._isRed('ff0000'), true, 'ff0000 tanpa hashtag harus bernilai true');

  // Test warna bukan merah
  assert.strictEqual(OfficeCheckers._isRed('#00ff00'), false, '#00ff00 (hijau) harus bernilai false');
  assert.strictEqual(OfficeCheckers._isRed('#0000ff'), false, '#0000ff (biru) harus bernilai false');
  assert.strictEqual(OfficeCheckers._isRed('#ffff00'), false, '#ffff00 (kuning) harus bernilai false');

  // Test warna merah format ARGB (hex 8 digit)
  assert.strictEqual(OfficeCheckers._isRed('#ffff0000'), true, '#ffff0000 ARGB harus bernilai true');
  assert.strictEqual(OfficeCheckers._isRed('#ff00ff00'), false, '#ff00ff00 ARGB hijau harus bernilai false');

  // Test nama warna teks
  assert.strictEqual(OfficeCheckers._isRed('red'), true, 'Teks "red" harus bernilai true');
  assert.strictEqual(OfficeCheckers._isRed('blue'), false, 'Teks "blue" harus bernilai false');
  assert.strictEqual(OfficeCheckers._isRed(''), false, 'Teks kosong harus bernilai false');
});

test('Unit Test: OfficeCheckers.confirm', async (t) => {
  const confirmFunc = OfficeCheckers.confirm('task_1');
  const result = await confirmFunc();
  assert.strictEqual(result.score, 100);
  assert.match(result.detail, /dikonfirmasi/i);
});
