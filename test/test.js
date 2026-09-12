// test/test.js
const assert = require('node:assert/strict');
const T = require('../thue.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('✓', name); }
  catch (e) { console.error('✗', name); console.error(e.message); process.exitCode = 1; }
}

// --- normalizeText ---
test('normalizeText bỏ dấu, thường hoá, gộp khoảng trắng', () => {
  assert.equal(T.normalizeText('  Nguyễn   Văn  ĐỨC '), 'nguyen van duc');
  assert.equal(T.normalizeText('Tổng số thuế phải nộp'), 'tong so thue phai nop');
  assert.equal(T.normalizeText(null), '');
  assert.equal(T.normalizeText(123), '123');
});

// --- parseMoney ---
test('parseMoney đọc số và chuỗi có phân cách nghìn', () => {
  assert.equal(T.parseMoney(54000), 54000);
  assert.equal(T.parseMoney('1.234.000'), 1234000);
  assert.equal(T.parseMoney('27,000'), 27000);
  assert.equal(T.parseMoney(' 494 '), 494);
  assert.equal(T.parseMoney(''), null);
  assert.equal(T.parseMoney('-'), null);
  assert.equal(T.parseMoney(null), null);
  assert.equal(T.parseMoney('abc'), null);
});

// --- formatMoney ---
test('formatMoney phân cách bằng dấu chấm và hậu tố đ', () => {
  assert.equal(T.formatMoney(1234000), '1.234.000 đ');
  assert.equal(T.formatMoney(0), '0 đ');
  assert.equal(T.formatMoney(494), '494 đ');
});

// --- DEFAULT_RATES ---
test('DEFAULT_RATES đúng theo spec', () => {
  assert.deepEqual(T.DEFAULT_RATES, { nghiaTrang: 15000, thienTai: 10000, moiTruong: 15000, soThang: 6 });
});

console.log(`\n${passed} test đạt`);
