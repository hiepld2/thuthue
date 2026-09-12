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

// --- parseSheetRows ---
const HEADER = ['STT', 'MST', 'Tên NNT', 'Ngày sinh', 'Số CCCD', 'Diện tích', 'Đoạn đường', 'Giá đất',
  'Mã phi nông nghiệp', 'Thôn/Tổ', 'Tiểu mục', 'Số thuế kỳ trước chuyển sang', 'Số phát sinh trong kỳ',
  'Số truy thu năm trước', 'Thuế miễn giảm', 'Tổng số thuế phải nộp', 'Số tiền đã nộp', 'Chuyển kỳ sau', 'Ghi chú'];
const NUMROW = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19];
function row(stt, mst, ten, ns, cccd, dt, thon, phaiNop) {
  return [stt, mst, ten, ns, cccd, dt, 'Các trục đường có mặt cắt dưới 2,5m', 900000,
    '1019031264', thon, 1601, '', '', '', '', phaiNop, '', '', ''];
}

test('parseSheetRows tìm tiêu đề sau dòng trống và bỏ dòng số thứ tự cột', () => {
  const m = [
    ['SỔ BỘ THUẾ SỬ DỤNG ĐẤT PHI NÔNG NGHIỆP'],
    [],
    HEADER,
    NUMROW,
    row(177, '8220707543', 'LÊ VĂN AN', '04/06/1936', '145665829', 200, 'Thôn Thống Nhất', '54.000'),
    row(343, '8220708963', 'NGUYỄN VĂN AN', '18/04/1962', '011618238', 325, 'Thôn Thống Nhất', 234000),
    [],
  ];
  const r = T.parseSheetRows(m);
  assert.equal(r.ok, true);
  assert.equal(r.headerRowIndex, 2);
  assert.equal(r.rows.length, 2);
  assert.deepEqual(r.rows[0], {
    mst: '8220707543', ten: 'LÊ VĂN AN', ngaySinh: '04/06/1936', cccd: '145665829',
    dienTich: '200', thon: 'Thôn Thống Nhất', thuePhaiNop: 54000,
  });
  assert.equal(r.rows[1].thuePhaiNop, 234000);
});

test('parseSheetRows: ô thuế trống thành null, CCCD giữ số 0 đầu', () => {
  const m = [HEADER, row(1, '8220707021', 'NGUYỄN CHIẾN BINH', '25/04/1959', '030509004203', 60, 'Thôn A', '')];
  const r = T.parseSheetRows(m);
  assert.equal(r.ok, true);
  assert.equal(r.rows[0].thuePhaiNop, null);
  assert.equal(r.rows[0].cccd, '030509004203');
});

test('parseSheetRows: không có tiêu đề', () => {
  const r = T.parseSheetRows([['a', 'b'], [1, 2]]);
  assert.equal(r.ok, false);
  assert.equal(r.error, 'Không tìm thấy dòng tiêu đề (cần có cột MST và Tên NNT).');
});

test('parseSheetRows: thiếu cột bắt buộc', () => {
  const r = T.parseSheetRows([['STT', 'MST', 'Tên NNT', 'Ngày sinh'], [1, 'x', 'A', '1/1/1970']]);
  assert.equal(r.ok, false);
  assert.deepEqual(r.missingColumns, ['Số CCCD', 'Tổng số thuế phải nộp']);
  assert.equal(r.error, 'Thiếu cột: Số CCCD, Tổng số thuế phải nộp');
});

test('parseSheetRows: dòng không có tên bị bỏ qua, CCCD số được đổi sang chuỗi', () => {
  const m = [HEADER, row(1, 'm', '', '', '1', 1, 't', 1), row(2, 'm2', 'B', '', 145266887, 1, 't', 2)];
  const r = T.parseSheetRows(m);
  assert.equal(r.rows.length, 1);
  assert.equal(r.rows[0].cccd, '145266887');
});

test('parseSheetRows: không bắt nhầm cột chỉ chứa chữ "tên"', () => {
  const r = T.parseSheetRows([['STT', 'MST', 'Tên chủ hộ', 'Số CCCD', 'Tổng số thuế phải nộp'], [1, 'x', 'A', '1', 1]]);
  assert.equal(r.ok, false);
  assert.deepEqual(r.missingColumns, ['Tên NNT']);
  const ok = T.parseSheetRows([['STT', 'MST', 'Tên chủ hộ', 'Tên NNT', 'Số CCCD', 'Tổng số thuế phải nộp'], [1, 'x', 'CHỦ', 'NNT', '1', 1]]);
  assert.equal(ok.ok, true);
  assert.equal(ok.rows[0].ten, 'NNT');
});

// --- searchRows ---
const ROWS = [
  { mst: '1', ten: 'LÊ VĂN AN', ngaySinh: '', cccd: '145665829', dienTich: '', thon: '', thuePhaiNop: 54000 },
  { mst: '2', ten: 'NGUYỄN VĂN AN', ngaySinh: '', cccd: '011618238', dienTich: '', thon: '', thuePhaiNop: 234000 },
  { mst: '3', ten: 'Giang Lê Bằng', ngaySinh: '', cccd: '033167009147', dienTich: '', thon: '', thuePhaiNop: null },
];

test('searchRows tìm tên không dấu, không phân biệt hoa thường', () => {
  assert.deepEqual(T.searchRows(ROWS, 'van an').map(r => r.mst), ['1', '2']);
  assert.deepEqual(T.searchRows(ROWS, 'LE van').map(r => r.mst), ['1']);
  assert.deepEqual(T.searchRows(ROWS, 'giang le bang').map(r => r.mst), ['3']);
});

test('searchRows tìm theo CCCD một phần', () => {
  assert.deepEqual(T.searchRows(ROWS, '0331').map(r => r.mst), ['3']);
  assert.deepEqual(T.searchRows(ROWS, '033 167 009 147').map(r => r.mst), ['3']);
});

test('searchRows: chuỗi rỗng trả rỗng, giới hạn kết quả', () => {
  assert.deepEqual(T.searchRows(ROWS, ''), []);
  assert.deepEqual(T.searchRows(ROWS, '   '), []);
  assert.equal(T.searchRows(ROWS, 'an', 1).length, 1);
});

console.log(`\n${passed} test đạt`);
