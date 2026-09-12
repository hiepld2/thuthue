# Trang tính thuế — Kế hoạch triển khai

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Trang web tĩnh tra cứu thuế đất phi nông nghiệp từ Excel và tính tổng 4 khoản thu theo số nhân khẩu.

**Architecture:** `thue.js` chứa toàn bộ hàm thuần (chuẩn hoá chuỗi, đọc ma trận ô Excel thành bản ghi, tìm kiếm, tính tiền) và chạy được cả trong Node lẫn trình duyệt. `index.html` chứa giao diện, gọi SheetJS để đổi file Excel thành ma trận ô rồi gọi `thue.js`. Test chạy bằng `node test/test.js`, không cần cài gói.

**Tech Stack:** HTML/CSS/JS thuần, SheetJS 0.18.5 từ cdnjs, Node 22 (chỉ để chạy test), Python 3 + openpyxl (chỉ để sinh file Excel mẫu).

## Global Constraints

- Không có bước build, không `package.json`, không `node_modules`.
- SheetJS chỉ dùng trong `index.html`, ghim phiên bản: `https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js`.
- Toàn bộ chữ trên giao diện bằng tiếng Việt có dấu.
- Đơn giá mặc định: nghĩa trang 15000; phòng chống thiên tai 10000; bảo vệ môi trường 15000 đ/người/tháng; số tháng mặc định 6.
- Định dạng tiền: phân cách nghìn bằng dấu chấm, hậu tố ` đ` (ví dụ `1.234.000 đ`).
- Cột bắt buộc trong Excel: Tên NNT, Số CCCD, Tổng số thuế phải nộp.
- Giao diện: cỡ chữ cơ bản 18px, ô nhập và nút cao tối thiểu 48px, các bước có số thứ tự, bước sau chỉ hiện khi bước trước xong.
- Không đưa file Excel thật và ảnh lên repo (đã có `.gitignore`).
- Commit dùng `git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit ...` nếu git chưa cấu hình user.

## Cấu trúc file

- `thue.js` — hàm thuần: `normalizeText`, `parseMoney`, `formatMoney`, `parseSheetRows`, `searchRows`, `computeTotals`, `DEFAULT_RATES`. Cuối file xuất ra `module.exports` khi có, ngược lại gắn vào `window.Thue`.
- `test/test.js` — test bằng `node:assert`, chạy `node test/test.js`.
- `test/tao_mau.py` — script openpyxl sinh `test/mau.xlsx`.
- `test/mau.xlsx` — file Excel mẫu 10 hộ, tiêu đề giống bảng in.
- `index.html` — giao diện, nạp SheetJS + `thue.js`, lưu đơn giá vào `localStorage`.
- `README.md` — cách dùng và cách bật GitHub Pages.

---

### Task 1: Hàm chuẩn hoá chuỗi và tiền

**Files:**
- Create: `thue.js`
- Create: `test/test.js`

**Interfaces:**
- Produces:
  - `normalizeText(s: any) -> string` — bỏ dấu tiếng Việt (kể cả đ/Đ), thường hoá, gộp khoảng trắng, trim. `null`/`undefined` → `""`.
  - `parseMoney(v: any) -> number | null` — số → số; chuỗi `"1.234.000"` / `"27,000"` / `" 494 "` → số; chuỗi rỗng, `"-"`, không phải số → `null`.
  - `formatMoney(n: number) -> string` — `1234000` → `"1.234.000 đ"`, `0` → `"0 đ"`.
  - `DEFAULT_RATES = { nghiaTrang: 15000, thienTai: 10000, moiTruong: 15000, soThang: 6 }`.

- [ ] **Step 1: Viết test thất bại**

```js
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
```

- [ ] **Step 2: Chạy test, xác nhận thất bại**

Run: `node test/test.js`
Expected: lỗi `Cannot find module '../thue.js'`

- [ ] **Step 3: Viết `thue.js` tối thiểu**

```js
// thue.js — hàm thuần dùng chung cho trang và test
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Thue = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_RATES = { nghiaTrang: 15000, thienTai: 10000, moiTruong: 15000, soThang: 6 };

  function normalizeText(s) {
    if (s === null || s === undefined) return '';
    return String(s)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/đ/g, 'd').replace(/Đ/g, 'D')
      .toLowerCase()
      .replace(/\s+/g, ' ')
      .trim();
  }

  function parseMoney(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === 'number') return Number.isFinite(v) ? v : null;
    const s = String(v).trim().replace(/[.,\s]/g, '');
    if (!/^-?\d+$/.test(s)) return null;
    return parseInt(s, 10);
  }

  function formatMoney(n) {
    const s = String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return s + ' đ';
  }

  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney };
});
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

Run: `node test/test.js`
Expected: 4 dòng `✓`, cuối cùng `4 test đạt`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add thue.js test/test.js
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: hàm chuẩn hoá chuỗi và tiền"
```

---

### Task 2: Đọc ma trận ô Excel thành bản ghi

**Files:**
- Modify: `thue.js`
- Modify: `test/test.js`

**Interfaces:**
- Consumes: `normalizeText`, `parseMoney` (Task 1).
- Produces:
  - `parseSheetRows(matrix: any[][]) -> { ok: true, rows: Row[], headerRowIndex: number } | { ok: false, error: string, missingColumns?: string[] }`
  - `Row = { mst: string, ten: string, ngaySinh: string, cccd: string, dienTich: string, thon: string, thuePhaiNop: number | null }`
  - Lỗi: `"Không tìm thấy dòng tiêu đề (cần có cột MST và Tên NNT)."` hoặc `"Thiếu cột: Số CCCD, Tổng số thuế phải nộp"` (tên cột thiếu nối bằng `, `).

- [ ] **Step 1: Viết test thất bại**

Thêm vào `test/test.js` trước dòng `console.log(\`\n${passed} test đạt\`);`:

```js
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
```

- [ ] **Step 2: Chạy test, xác nhận thất bại**

Run: `node test/test.js`
Expected: các test `parseSheetRows` báo `✗` với `T.parseSheetRows is not a function`, exit code 1.

- [ ] **Step 3: Cài đặt `parseSheetRows`**

Thêm vào `thue.js` sau `formatMoney`, và thêm `parseSheetRows` vào object `return`:

```js
  // Cột: khoá nội bộ -> { label hiển thị, danh sách từ khoá (đã bỏ dấu), bắt buộc }
  const COLUMNS = [
    { key: 'mst', label: 'MST', keywords: ['mst', 'ma so thue'], required: false },
    { key: 'ten', label: 'Tên NNT', keywords: ['ten nnt', 'ten nguoi nop thue', 'ho ten', 'ho va ten'], required: true },
    { key: 'ngaySinh', label: 'Ngày sinh', keywords: ['ngay sinh'], required: false },
    { key: 'cccd', label: 'Số CCCD', keywords: ['cccd', 'cmnd', 'can cuoc'], required: true },
    { key: 'dienTich', label: 'Diện tích', keywords: ['dien tich'], required: false },
    { key: 'thon', label: 'Thôn', keywords: ['thon'], required: false },
    { key: 'thuePhaiNop', label: 'Tổng số thuế phải nộp', keywords: ['tong so thue phai nop', 'thue phai nop', 'phai nop'], required: true },
  ];

  function findHeaderRow(matrix) {
    for (let i = 0; i < matrix.length; i++) {
      const cells = (matrix[i] || []).map(normalizeText);
      const hasMst = cells.some(c => c.includes('mst') || c.includes('ma so thue'));
      const hasTen = cells.some(c => c.includes('ten'));
      if (hasMst && hasTen) return i;
    }
    return -1;
  }

  function isColumnNumberRow(r) {
    const cells = (r || []).filter(c => c !== null && c !== undefined && String(c).trim() !== '');
    if (cells.length === 0) return false;
    return cells.every(c => /^\d{1,2}$/.test(String(c).trim()));
  }

  function mapColumns(headerCells) {
    const norm = headerCells.map(normalizeText);
    const map = {};
    for (const col of COLUMNS) {
      let idx = -1;
      for (const kw of col.keywords) {
        // ưu tiên khớp chính xác, sau đó khớp chứa
        idx = norm.findIndex(c => c === kw);
        if (idx === -1) idx = norm.findIndex(c => c.includes(kw));
        if (idx !== -1) break;
      }
      map[col.key] = idx;
    }
    return map;
  }

  function cellStr(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) {
      const d = String(v.getDate()).padStart(2, '0');
      const m = String(v.getMonth() + 1).padStart(2, '0');
      return `${d}/${m}/${v.getFullYear()}`;
    }
    return String(v).trim();
  }

  function parseSheetRows(matrix) {
    const h = findHeaderRow(matrix);
    if (h === -1) return { ok: false, error: 'Không tìm thấy dòng tiêu đề (cần có cột MST và Tên NNT).' };
    const map = mapColumns(matrix[h] || []);
    const missing = COLUMNS.filter(c => c.required && map[c.key] === -1).map(c => c.label);
    if (missing.length) return { ok: false, error: 'Thiếu cột: ' + missing.join(', '), missingColumns: missing };

    let start = h + 1;
    if (isColumnNumberRow(matrix[start])) start++;

    const rows = [];
    for (let i = start; i < matrix.length; i++) {
      const r = matrix[i] || [];
      const get = k => (map[k] === -1 ? '' : cellStr(r[map[k]]));
      const ten = get('ten');
      if (!ten) continue;
      rows.push({
        mst: get('mst'),
        ten,
        ngaySinh: get('ngaySinh'),
        cccd: get('cccd'),
        dienTich: get('dienTich'),
        thon: get('thon'),
        thuePhaiNop: map.thuePhaiNop === -1 ? null : parseMoney(r[map.thuePhaiNop]),
      });
    }
    return { ok: true, rows, headerRowIndex: h };
  }
```

Sửa dòng return thành:

```js
  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney, parseSheetRows };
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

Run: `node test/test.js`
Expected: `9 test đạt`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add thue.js test/test.js
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: đọc ma trận Excel thành bản ghi hộ"
```

---

### Task 3: Tìm kiếm theo tên hoặc CCCD

**Files:**
- Modify: `thue.js`
- Modify: `test/test.js`

**Interfaces:**
- Consumes: `normalizeText`, `Row` (Task 2).
- Produces: `searchRows(rows: Row[], query: string, limit = 50) -> Row[]` — chuỗi rỗng trả `[]`; khớp khi tên chuẩn hoá chứa chuỗi tìm chuẩn hoá, hoặc CCCD chứa chuỗi tìm (bỏ khoảng trắng); tối đa `limit` kết quả, giữ thứ tự gốc.

- [ ] **Step 1: Viết test thất bại**

Thêm vào `test/test.js` trước dòng in tổng:

```js
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
```

- [ ] **Step 2: Chạy test, xác nhận thất bại**

Run: `node test/test.js`
Expected: 3 test `searchRows` báo `✗` với `T.searchRows is not a function`.

- [ ] **Step 3: Cài đặt `searchRows`**

Thêm vào `thue.js` sau `parseSheetRows`, thêm `searchRows` vào `return`:

```js
  function searchRows(rows, query, limit) {
    if (limit === undefined) limit = 50;
    const q = normalizeText(query);
    if (!q) return [];
    const qDigits = q.replace(/\s+/g, '');
    const out = [];
    for (const r of rows) {
      const byName = normalizeText(r.ten).includes(q);
      const byId = qDigits !== '' && String(r.cccd || '').replace(/\s+/g, '').includes(qDigits);
      if (byName || byId) {
        out.push(r);
        if (out.length >= limit) break;
      }
    }
    return out;
  }
```

```js
  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney, parseSheetRows, searchRows };
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

Run: `node test/test.js`
Expected: `12 test đạt`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add thue.js test/test.js
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: tìm kiếm theo tên hoặc CCCD"
```

---

### Task 4: Tính tổng tiền

**Files:**
- Modify: `thue.js`
- Modify: `test/test.js`

**Interfaces:**
- Consumes: `DEFAULT_RATES` (Task 1).
- Produces:
  - `computeTotals(input, rates) -> { lines: Line[], total: number }`
  - `input = { thueDat: number | null, nguoiNghiaTrang: number, nguoiThienTai: number, nguoiMoiTruong: number, soThang: number }`
  - `rates = { nghiaTrang, thienTai, moiTruong }` (bỏ qua `soThang` nếu có)
  - `Line = { key: 'thueDat'|'nghiaTrang'|'thienTai'|'moiTruong', label: string, detail: string, amount: number, note?: string }`
  - `thueDat === null` → dòng thuế đất `amount: 0`, `note: 'Không tìm thấy trong file'`.
  - Số người/tháng âm, `NaN`, không phải số → coi là 0.
  - `sumThueDat(rows: Row[]) -> { thueDat: number | null, soDongThieu: number }` — cộng `thuePhaiNop` của các dòng đã chọn; dòng `null` tính 0 và đếm vào `soDongThieu`; nếu `rows` rỗng hoặc mọi dòng đều `null` thì `thueDat: null`.

- [ ] **Step 1: Viết test thất bại**

Thêm vào `test/test.js` trước dòng in tổng:

```js
// --- computeTotals ---
test('computeTotals tính đủ 4 dòng và tổng', () => {
  const r = T.computeTotals(
    { thueDat: 54000, nguoiNghiaTrang: 4, nguoiThienTai: 3, nguoiMoiTruong: 4, soThang: 6 },
    T.DEFAULT_RATES
  );
  assert.equal(r.lines.length, 4);
  assert.deepEqual(r.lines.map(l => l.key), ['thueDat', 'nghiaTrang', 'thienTai', 'moiTruong']);
  assert.equal(r.lines[0].amount, 54000);
  assert.equal(r.lines[1].amount, 60000);   // 15000 × 4
  assert.equal(r.lines[2].amount, 30000);   // 10000 × 3
  assert.equal(r.lines[3].amount, 360000);  // 15000 × 4 × 6
  assert.equal(r.total, 504000);
  assert.equal(r.lines[1].detail, '15.000 đ × 4 người');
  assert.equal(r.lines[3].detail, '15.000 đ × 4 người × 6 tháng');
  assert.equal(r.lines[0].note, undefined);
});

test('computeTotals: thuế đất null thành 0 và có ghi chú', () => {
  const r = T.computeTotals(
    { thueDat: null, nguoiNghiaTrang: 0, nguoiThienTai: 0, nguoiMoiTruong: 0, soThang: 6 },
    T.DEFAULT_RATES
  );
  assert.equal(r.lines[0].amount, 0);
  assert.equal(r.lines[0].note, 'Không tìm thấy trong file');
  assert.equal(r.total, 0);
});

test('sumThueDat cộng nhiều dòng, đếm dòng thiếu', () => {
  const r1 = { thuePhaiNop: 54000 }, r2 = { thuePhaiNop: 27000 }, r3 = { thuePhaiNop: null };
  assert.deepEqual(T.sumThueDat([r1, r2]), { thueDat: 81000, soDongThieu: 0 });
  assert.deepEqual(T.sumThueDat([r1, r3]), { thueDat: 54000, soDongThieu: 1 });
  assert.deepEqual(T.sumThueDat([r3]), { thueDat: null, soDongThieu: 1 });
  assert.deepEqual(T.sumThueDat([]), { thueDat: null, soDongThieu: 0 });
});

test('computeTotals: giá trị âm, NaN, chuỗi coi là 0; đơn giá tuỳ chỉnh', () => {
  const r = T.computeTotals(
    { thueDat: 1000, nguoiNghiaTrang: -2, nguoiThienTai: NaN, nguoiMoiTruong: '3', soThang: 'x' },
    { nghiaTrang: 1, thienTai: 1, moiTruong: 20000 }
  );
  assert.equal(r.lines[1].amount, 0);
  assert.equal(r.lines[2].amount, 0);
  assert.equal(r.lines[3].amount, 0);       // 3 người × 0 tháng
  assert.equal(r.total, 1000);
});
```

- [ ] **Step 2: Chạy test, xác nhận thất bại**

Run: `node test/test.js`
Expected: 3 test `computeTotals` và 1 test `sumThueDat` báo `✗` với `... is not a function`.

- [ ] **Step 3: Cài đặt `computeTotals`**

Thêm vào `thue.js` sau `searchRows`, thêm `computeTotals` vào `return`:

```js
  function toCount(v) {
    const n = typeof v === 'number' ? v : parseInt(String(v), 10);
    if (!Number.isFinite(n) || n < 0) return 0;
    return Math.floor(n);
  }

  function sumThueDat(rows) {
    let total = 0, soDongThieu = 0, coSo = false;
    for (const r of rows || []) {
      if (r.thuePhaiNop === null || r.thuePhaiNop === undefined) soDongThieu++;
      else { total += r.thuePhaiNop; coSo = true; }
    }
    return { thueDat: coSo ? total : null, soDongThieu };
  }

  function computeTotals(input, rates) {
    const thueDat = input.thueDat === null || input.thueDat === undefined ? null : Number(input.thueDat);
    const nNT = toCount(input.nguoiNghiaTrang);
    const nTT = toCount(input.nguoiThienTai);
    const nMT = toCount(input.nguoiMoiTruong);
    const thang = toCount(input.soThang);
    const gNT = toCount(rates.nghiaTrang);
    const gTT = toCount(rates.thienTai);
    const gMT = toCount(rates.moiTruong);

    const thueDatLine = {
      key: 'thueDat',
      label: 'Thuế sử dụng đất phi nông nghiệp',
      detail: 'Theo sổ bộ thuế',
      amount: thueDat === null || !Number.isFinite(thueDat) ? 0 : thueDat,
    };
    if (thueDat === null) thueDatLine.note = 'Không tìm thấy trong file';

    const lines = [
      thueDatLine,
      { key: 'nghiaTrang', label: 'Nghĩa trang nhân dân', detail: `${formatMoney(gNT)} × ${nNT} người`, amount: gNT * nNT },
      { key: 'thienTai', label: 'Quỹ phòng chống thiên tai', detail: `${formatMoney(gTT)} × ${nTT} người`, amount: gTT * nTT },
      { key: 'moiTruong', label: 'Bảo vệ môi trường', detail: `${formatMoney(gMT)} × ${nMT} người × ${thang} tháng`, amount: gMT * nMT * thang },
    ];
    const total = lines.reduce((s, l) => s + l.amount, 0);
    return { lines, total };
  }
```

```js
  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney, parseSheetRows, searchRows, sumThueDat, computeTotals };
```

- [ ] **Step 4: Chạy test, xác nhận đạt**

Run: `node test/test.js`
Expected: `16 test đạt`, exit code 0.

- [ ] **Step 5: Commit**

```bash
git add thue.js test/test.js
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: tính tổng 4 khoản thu và cộng thuế đất nhiều dòng"
```

---

### Task 5: File Excel mẫu

**Files:**
- Create: `test/tao_mau.py`
- Create: `test/mau.xlsx` (sinh ra từ script)

**Interfaces:**
- Produces: `test/mau.xlsx` có sheet đầu tên `Sổ bộ`, dòng 1 tiêu đề lớn, dòng 2 trống, dòng 3 tiêu đề cột giống bảng in, dòng 4 số thứ tự cột 1..19, 11 dòng dữ liệu từ dòng 5. Cột "Số CCCD" định dạng text. Hộ LÊ VĂN AN có 2 dòng (cùng MST, tên, CCCD; hai thửa đất 54.000 và 27.000). Dòng cuối (LÊ VĂN BÌNH) có ô "Tổng số thuế phải nộp" trống.

- [ ] **Step 1: Viết script sinh file**

```python
# test/tao_mau.py — sinh test/mau.xlsx (dữ liệu mẫu lấy từ ảnh sổ bộ)
from pathlib import Path
from openpyxl import Workbook

HEADER = ['STT', 'MST', 'Tên NNT', 'Ngày sinh', 'Số CCCD', 'Diện tích', 'Đoạn đường', 'Giá đất',
          'Mã phi nông nghiệp', 'Thôn/Tổ', 'Tiểu mục', 'Số thuế kỳ trước chuyển sang',
          'Số phát sinh trong kỳ', 'Số truy thu năm trước', 'Thuế miễn giảm',
          'Tổng số thuế phải nộp', 'Số tiền đã nộp', 'Chuyển kỳ sau', 'Ghi chú']

DUONG = 'Các trục đường có mặt cắt dưới 2,5m (xã Đồng Tảo)'
THON = 'Thôn Thống Nhất (Xã Đồng Tảo cũ)'

# (STT, MST, Tên, Ngày sinh, CCCD, Diện tích, Giá đất, Phát sinh, Tổng phải nộp)
DATA = [
    (177, '8220707543', 'LÊ VĂN AN',          '04/06/1936', '145665829',    200, 900000,  54000,  54000),
    (178, '8220707543', 'LÊ VĂN AN',          '04/06/1936', '145665829',    100, 900000,  27000,  27000),
    (343, '8220708963', 'NGUYỄN VĂN AN',      '18/04/1962', '011618238',    325, 2400000, 234000, 468000),
    (25,  '8023404726', 'LÊ THỊ VÂN ANH',     '17/09/1984', '001184013031', 120, 2400000, 4282,   4282),
    (133, '8220707166', 'GIANG LÊ ANH',       '01/01/1956', '145344429',    200, 900000,  54000,  54000),
    (369, '8220709131', 'GIANG LÊ TUẤN ANH',  '23/10/1997', '033097008492', 120, 2400000, 86400,  172800),
    (365, '8220709075', 'LÊ VĂN BÀN',         '23/03/1949', '033049004943', 200, 2400000, 144000, 144000),
    (465, '8905401838', 'CHU THỊ BẰNG',       '26/11/1967', '033167009147', 75,  5500000, 123750, 123750),
    (438, '8608470444', 'NGUYỄN CHIẾN BINH',  '25/04/1959', '030509004203', 60,  1800000, 32400,  32400),
    (256, '8220708226', 'GIANG LÊ BÌNH',      '01/08/1956', '033056006676', 50,  900000,  13500,  13500),
    (122, '8220707053', 'LÊ VĂN BÌNH',        '01/01/1970', '0',            50,  900000,  None,   None),
]

wb = Workbook()
ws = wb.active
ws.title = 'Sổ bộ'
ws.append(['SỔ BỘ THUẾ SỬ DỤNG ĐẤT PHI NÔNG NGHIỆP NĂM 2026'])
ws.append([])
ws.append(HEADER)
ws.append(list(range(1, len(HEADER) + 1)))
for stt, mst, ten, ns, cccd, dt, gia, ps, tong in DATA:
    ws.append([stt, mst, ten, ns, cccd, dt, DUONG, gia, '1019031264', THON, 1601,
               None, ps, None, None, tong, None, None, None])
# Cột CCCD (E) và MST (B) để dạng text để giữ số 0 đầu
for r in range(5, 5 + len(DATA)):
    ws.cell(row=r, column=2).number_format = '@'
    ws.cell(row=r, column=5).number_format = '@'

out = Path(__file__).with_name('mau.xlsx')
wb.save(out)
print('Đã ghi', out)
```

- [ ] **Step 2: Chạy script**

Run: `python3 test/tao_mau.py`
Expected: in `Đã ghi .../test/mau.xlsx`, file tồn tại.

- [ ] **Step 3: Kiểm tra file mẫu đọc được bằng `parseSheetRows`**

Kiểm tra bằng openpyxl đổi thành ma trận JSON rồi chạy qua Node:

```bash
python3 -c "
import json, openpyxl
ws = openpyxl.load_workbook('test/mau.xlsx').active
print(json.dumps([[c for c in r] for r in ws.iter_rows(values_only=True)], ensure_ascii=False, default=str))
" > /tmp/mau.json && node -e "
const T = require('./thue.js');
const m = JSON.parse(require('fs').readFileSync('/tmp/mau.json', 'utf8'));
const r = T.parseSheetRows(m);
console.log(r.ok, r.rows.length, r.rows[0].ten, r.rows[1].thuePhaiNop, r.rows[10].thuePhaiNop);
"
```

Expected: `true 11 LÊ VĂN AN 27000 null`

- [ ] **Step 4: Commit**

```bash
git add test/tao_mau.py test/mau.xlsx
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "test: file Excel mẫu 11 dòng, có hộ nhiều thửa"
```

(`.gitignore` đã có ngoại lệ `!test/mau.xlsx`; nếu `git add` báo bị ignore, kiểm tra lại dòng đó.)

---

### Task 6: Giao diện `index.html`

**Files:**
- Create: `index.html`
- Create: `.claude/launch.json`

**Interfaces:**
- Consumes: `window.Thue.{parseSheetRows, searchRows, sumThueDat, computeTotals, formatMoney, normalizeText, DEFAULT_RATES}`; `XLSX.read`, `XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' })`.
- Produces: trang hoàn chỉnh 3 bước, chọn nhiều dòng cho một hộ, khối đơn giá, lưu `localStorage` khoá `thue-settings`.

- [ ] **Step 1: Viết `index.html`**

```html
<!doctype html>
<html lang="vi">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Tính tiền thuế cần nộp</title>
<style>
  :root {
    --nen: #f6f7f9; --the: #ffffff; --chu: #1a1a1a; --mo: #555;
    --vien: #c9ced6; --chinh: #0b5cad; --chinh-dam: #08447f; --nhan: #e8f1fb;
    --loi: #b3261e; --loi-nen: #fdecea; --ok: #1b6e3a; --ok-nen: #e6f4ea; --vang-nen: #fff6d6;
  }
  * { box-sizing: border-box; }
  html { font-size: 18px; }
  body { margin: 0; background: var(--nen); color: var(--chu); font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif; line-height: 1.5; }
  main { max-width: 760px; margin: 0 auto; padding: 20px 16px 60px; }
  h1 { font-size: 1.6rem; margin: 0 0 6px; }
  .mo { color: var(--mo); margin: 0 0 20px; }
  .buoc { background: var(--the); border: 1px solid var(--vien); border-radius: 12px; padding: 18px; margin-bottom: 18px; }
  .buoc h2 { font-size: 1.15rem; margin: 0 0 12px; display: flex; align-items: center; gap: 10px; }
  .so { display: inline-flex; width: 34px; height: 34px; border-radius: 50%; background: var(--chinh); color: #fff; align-items: center; justify-content: center; font-weight: 700; flex: none; }
  .huongdan { color: var(--mo); margin: 0 0 12px; }
  button, input { font: inherit; }
  .nut { min-height: 48px; padding: 0 22px; border-radius: 10px; border: 2px solid var(--chinh); background: var(--chinh); color: #fff; font-weight: 700; cursor: pointer; }
  .nut:hover { background: var(--chinh-dam); }
  .nut.phu { background: #fff; color: var(--chinh); }
  .nut:disabled { opacity: .5; cursor: default; }
  input[type=text], input[type=number] { min-height: 48px; width: 100%; padding: 0 14px; border: 2px solid var(--vien); border-radius: 10px; }
  input:focus { outline: 3px solid #9cc4ee; border-color: var(--chinh); }
  input[type=checkbox] { width: 26px; height: 26px; flex: none; cursor: pointer; }
  .trangthai { margin-top: 12px; padding: 10px 14px; border-radius: 10px; }
  .trangthai.ok { background: var(--ok-nen); color: var(--ok); }
  .trangthai.loi { background: var(--loi-nen); color: var(--loi); }
  .ketqua { list-style: none; margin: 12px 0 0; padding: 0; border: 1px solid var(--vien); border-radius: 10px; overflow: hidden; }
  .ketqua li { display: flex; gap: 12px; align-items: center; padding: 12px 14px; border-top: 1px solid var(--vien); cursor: pointer; }
  .ketqua li:first-child { border-top: 0; }
  .ketqua li:hover { background: var(--nhan); }
  .ketqua li.chon { background: var(--nhan); }
  .ketqua .noidung { flex: 1; min-width: 0; }
  .ketqua .ten { font-weight: 700; }
  .phu-chu { color: var(--mo); font-size: .9rem; }
  .ketqua .tien { font-weight: 700; white-space: nowrap; }
  .thongtin { background: var(--nhan); border-radius: 10px; padding: 12px 14px; margin-bottom: 14px; }
  .thongtin .ten { font-size: 1.2rem; font-weight: 700; }
  .thongtin ul { margin: 8px 0 0; padding-left: 22px; }
  .thongtin li { margin: 2px 0; }
  table { width: 100%; border-collapse: collapse; }
  th, td { padding: 12px 8px; border-bottom: 1px solid var(--vien); text-align: left; vertical-align: middle; }
  th { color: var(--mo); font-weight: 600; font-size: .9rem; }
  td.tien, th.tien { text-align: right; white-space: nowrap; font-weight: 700; }
  td.nhap input { width: 110px; text-align: center; }
  .chitiet { color: var(--mo); font-size: .9rem; }
  .ghichu { display: inline-block; margin-top: 4px; padding: 2px 8px; border-radius: 6px; background: var(--vang-nen); font-size: .85rem; }
  tr.tong td { border-bottom: 0; border-top: 3px solid var(--chu); font-size: 1.35rem; font-weight: 800; padding-top: 16px; }
  details { border: 1px dashed var(--vien); border-radius: 12px; padding: 12px 18px; background: var(--the); }
  summary { cursor: pointer; font-weight: 700; min-height: 32px; }
  .dongia { display: grid; grid-template-columns: 1fr 160px; gap: 10px 14px; align-items: center; margin-top: 12px; }
  .an { display: none !important; }
  .hang-nut { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 12px; }
  input[type=file] { display: none; }
</style>
</head>
<body>
<main>
  <h1>Tính tiền thuế cần nộp</h1>
  <p class="mo">Làm theo 3 bước từ trên xuống dưới.</p>

  <!-- Bước 1 -->
  <section class="buoc" id="buoc1">
    <h2><span class="so">1</span> Chọn file Excel sổ bộ thuế</h2>
    <p class="huongdan">Bấm nút bên dưới rồi chọn file Excel (.xlsx hoặc .xls) trên máy.</p>
    <label class="nut" for="fileExcel" role="button" tabindex="0">Chọn file Excel</label>
    <input type="file" id="fileExcel" accept=".xlsx,.xls">
    <div id="trangThaiFile" class="trangthai an"></div>
  </section>

  <!-- Bước 2 -->
  <section class="buoc an" id="buoc2">
    <h2><span class="so">2</span> Tìm hộ cần tính</h2>
    <p class="huongdan">Gõ tên hoặc số CCCD, rồi bấm vào đúng hộ. Hộ có nhiều thửa đất sẽ có nhiều dòng: các dòng cùng số CCCD được tích sẵn, bạn có thể bỏ tích dòng không cần.</p>
    <input type="text" id="oTim" placeholder="Ví dụ: Lê Văn An hoặc 145665829" autocomplete="off">
    <ul id="dsKetQua" class="ketqua an"></ul>
    <div id="thanhChon" class="hang-nut an"><button type="button" class="nut" id="nutTinh">Tính cho 0 dòng đã chọn</button></div>
    <div id="khongThay" class="trangthai loi an">
      Không tìm thấy trong file.
      <div class="hang-nut"><button type="button" class="nut phu" id="nutKhongThue">Tính với thuế đất 0 đ</button></div>
    </div>
  </section>

  <!-- Bước 3 -->
  <section class="buoc an" id="buoc3">
    <h2><span class="so">3</span> Nhập số người và xem tổng tiền</h2>
    <div class="thongtin" id="thongTinHo"></div>
    <table>
      <thead><tr><th>Khoản thu</th><th>Số người</th><th class="tien">Số tiền</th></tr></thead>
      <tbody>
        <tr>
          <td>Thuế sử dụng đất phi nông nghiệp<div class="chitiet" id="ctThueDat">Theo sổ bộ thuế</div><span id="ghiChuThueDat" class="ghichu an"></span></td>
          <td></td>
          <td class="tien" id="tienThueDat">0 đ</td>
        </tr>
        <tr>
          <td>Nghĩa trang nhân dân<div class="chitiet" id="ctNghiaTrang"></div></td>
          <td class="nhap"><input type="number" id="nguoiNghiaTrang" min="0" step="1" value="0" inputmode="numeric"></td>
          <td class="tien" id="tienNghiaTrang">0 đ</td>
        </tr>
        <tr>
          <td>Quỹ phòng chống thiên tai<div class="chitiet" id="ctThienTai"></div></td>
          <td class="nhap"><input type="number" id="nguoiThienTai" min="0" step="1" value="0" inputmode="numeric"></td>
          <td class="tien" id="tienThienTai">0 đ</td>
        </tr>
        <tr>
          <td>Bảo vệ môi trường<div class="chitiet" id="ctMoiTruong"></div>
            <div class="chitiet" style="margin-top:6px">Số tháng: <input type="number" id="soThang" min="0" step="1" value="6" inputmode="numeric" style="width:80px;min-height:40px;text-align:center"></div>
          </td>
          <td class="nhap"><input type="number" id="nguoiMoiTruong" min="0" step="1" value="0" inputmode="numeric"></td>
          <td class="tien" id="tienMoiTruong">0 đ</td>
        </tr>
        <tr class="tong"><td colspan="2">TỔNG CỘNG</td><td class="tien" id="tienTong">0 đ</td></tr>
      </tbody>
    </table>
    <div class="hang-nut"><button type="button" class="nut phu" id="nutHoKhac">Tính cho hộ khác</button></div>
  </section>

  <!-- Đơn giá -->
  <details id="khoiDonGia">
    <summary>Đơn giá (chỉ sửa khi có thay đổi)</summary>
    <div class="dongia">
      <label for="giaNghiaTrang">Nghĩa trang nhân dân (đ/người)</label><input type="number" id="giaNghiaTrang" min="0" step="1000" inputmode="numeric">
      <label for="giaThienTai">Quỹ phòng chống thiên tai (đ/người)</label><input type="number" id="giaThienTai" min="0" step="1000" inputmode="numeric">
      <label for="giaMoiTruong">Bảo vệ môi trường (đ/người/tháng)</label><input type="number" id="giaMoiTruong" min="0" step="1000" inputmode="numeric">
      <label for="thangMacDinh">Số tháng mặc định</label><input type="number" id="thangMacDinh" min="0" step="1" inputmode="numeric">
    </div>
    <div class="hang-nut"><button type="button" class="nut phu" id="nutMacDinh">Khôi phục mặc định</button></div>
  </details>
</main>

<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="thue.js"></script>
<script>
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  const KHOA_LUU = 'thue-settings';

  // ---- Đơn giá + localStorage ----
  function docCaiDat() {
    try {
      const raw = localStorage.getItem(KHOA_LUU);
      if (!raw) return Object.assign({}, Thue.DEFAULT_RATES);
      const o = JSON.parse(raw);
      const out = Object.assign({}, Thue.DEFAULT_RATES);
      for (const k of Object.keys(out)) {
        const n = Number(o[k]);
        if (Number.isFinite(n) && n >= 0) out[k] = n;
      }
      return out;
    } catch (e) { return Object.assign({}, Thue.DEFAULT_RATES); }
  }
  function luuCaiDat(cd) {
    try { localStorage.setItem(KHOA_LUU, JSON.stringify(cd)); } catch (e) { /* bỏ qua */ }
  }
  function layCaiDatTuForm() {
    return {
      nghiaTrang: Number($('giaNghiaTrang').value) || 0,
      thienTai: Number($('giaThienTai').value) || 0,
      moiTruong: Number($('giaMoiTruong').value) || 0,
      soThang: Number($('thangMacDinh').value) || 0,
    };
  }
  function doCaiDatVaoForm(cd) {
    $('giaNghiaTrang').value = cd.nghiaTrang;
    $('giaThienTai').value = cd.thienTai;
    $('giaMoiTruong').value = cd.moiTruong;
    $('thangMacDinh').value = cd.soThang;
  }
  let caiDat = docCaiDat();
  doCaiDatVaoForm(caiDat);
  $('soThang').value = caiDat.soThang;
  for (const id of ['giaNghiaTrang', 'giaThienTai', 'giaMoiTruong', 'thangMacDinh']) {
    $(id).addEventListener('input', () => { caiDat = layCaiDatTuForm(); luuCaiDat(caiDat); tinhLai(); });
  }
  $('nutMacDinh').addEventListener('click', () => {
    caiDat = Object.assign({}, Thue.DEFAULT_RATES);
    doCaiDatVaoForm(caiDat); luuCaiDat(caiDat);
    $('soThang').value = caiDat.soThang; tinhLai();
  });

  // ---- Bước 1: đọc file ----
  let cacHo = [];
  function baoFile(loai, chu) {
    const el = $('trangThaiFile');
    el.className = 'trangthai ' + loai;
    el.textContent = chu;
  }
  $('fileExcel').addEventListener('change', async ev => {
    const f = ev.target.files && ev.target.files[0];
    if (!f) return;
    an('buoc2'); an('buoc3'); cacHo = []; daChon.clear();
    try {
      const buf = await f.arrayBuffer();
      const wb = XLSX.read(buf, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' });
      const r = Thue.parseSheetRows(matrix);
      if (!r.ok) { baoFile('loi', r.error); return; }
      cacHo = r.rows;
      baoFile('ok', `Đã nạp ${cacHo.length} dòng từ file "${f.name}". Chuyển sang bước 2.`);
      hien('buoc2'); $('oTim').value = ''; timKiem(); $('oTim').focus();
    } catch (e) {
      baoFile('loi', 'Không đọc được file. Hãy chọn file .xlsx hoặc .xls.');
    }
    ev.target.value = '';
  });
  // Cho phép bấm Enter/Space trên nhãn "Chọn file Excel"
  document.querySelector('label[for=fileExcel]').addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('fileExcel').click(); }
  });

  // ---- Bước 2: tìm và chọn nhiều dòng ----
  const daChon = new Set(); // các bản ghi (object trong cacHo) đang được tích
  let ketQuaHienTai = [];

  function moTaDong(h) {
    return [h.ngaySinh && `Sinh ${h.ngaySinh}`, h.cccd && `CCCD ${h.cccd}`, h.thon, h.dienTich && `${h.dienTich} m²`]
      .filter(Boolean).join(' · ');
  }
  function cungHo(a, b) {
    const cccdA = String(a.cccd || '').trim();
    if (!cccdA || cccdA === '0') return false;
    return cccdA === String(b.cccd || '').trim() && Thue.normalizeText(a.ten) === Thue.normalizeText(b.ten);
  }
  function capNhatNutTinh() {
    const n = daChon.size;
    $('nutTinh').textContent = `Tính cho ${n} dòng đã chọn`;
    $('nutTinh').disabled = n === 0;
    if (ketQuaHienTai.length) hien('thanhChon'); else an('thanhChon');
  }
  function veKetQua() {
    const ds = $('dsKetQua');
    ds.innerHTML = '';
    for (const h of ketQuaHienTai) {
      const li = document.createElement('li');
      if (daChon.has(h)) li.classList.add('chon');
      const cb = document.createElement('input');
      cb.type = 'checkbox'; cb.checked = daChon.has(h); cb.tabIndex = -1;
      const nd = document.createElement('div');
      nd.className = 'noidung';
      nd.innerHTML = '<div class="ten"></div><div class="phu-chu"></div>';
      nd.querySelector('.ten').textContent = h.ten;
      nd.querySelector('.phu-chu').textContent = moTaDong(h);
      const tien = document.createElement('div');
      tien.className = 'tien';
      tien.textContent = h.thuePhaiNop === null ? 'Không có số thuế' : Thue.formatMoney(h.thuePhaiNop);
      li.append(cb, nd, tien);
      li.addEventListener('click', () => batTat(h));
      ds.appendChild(li);
    }
    hien('dsKetQua');
    capNhatNutTinh();
  }
  function batTat(h) {
    if (daChon.has(h)) {
      daChon.delete(h);
    } else {
      daChon.add(h);
      // tự tích các dòng khác của cùng hộ (cùng CCCD hợp lệ và cùng tên)
      for (const k of cacHo) if (k !== h && cungHo(h, k)) daChon.add(k);
    }
    veKetQua();
  }
  function timKiem() {
    const q = $('oTim').value;
    ketQuaHienTai = Thue.searchRows(cacHo, q, 50);
    if (!q.trim()) { ketQuaHienTai = []; an('dsKetQua'); an('khongThay'); an('thanhChon'); return; }
    if (ketQuaHienTai.length === 0) { an('dsKetQua'); an('thanhChon'); hien('khongThay'); return; }
    an('khongThay');
    veKetQua();
  }
  $('oTim').addEventListener('input', timKiem);
  $('nutTinh').addEventListener('click', () => {
    if (daChon.size === 0) return;
    const rows = cacHo.filter(h => daChon.has(h)); // giữ thứ tự trong file
    chonHo({ ten: rows[0].ten, rows });
  });
  $('nutKhongThue').addEventListener('click', () => {
    chonHo({ ten: $('oTim').value.trim() || 'Hộ chưa có trong sổ', rows: [] });
  });

  // ---- Bước 3: tính ----
  let hoDangChon = null; // { ten, rows }
  function chonHo(ho) {
    hoDangChon = ho;
    const tt = $('thongTinHo');
    tt.innerHTML = '<div class="ten"></div><div class="phu-chu" id="soDong"></div><ul id="dsDaChon"></ul>';
    tt.querySelector('.ten').textContent = ho.ten;
    const ul = tt.querySelector('#dsDaChon');
    if (ho.rows.length === 0) {
      tt.querySelector('#soDong').textContent = 'Không có dòng nào trong file Excel.';
    } else {
      tt.querySelector('#soDong').textContent = ho.rows.length === 1 ? '1 dòng trong sổ bộ:' : `${ho.rows.length} dòng trong sổ bộ (cộng dồn thuế đất):`;
      for (const h of ho.rows) {
        const li = document.createElement('li');
        li.textContent = `${moTaDong(h) || h.ten} — ${h.thuePhaiNop === null ? 'không có số thuế' : Thue.formatMoney(h.thuePhaiNop)}`;
        ul.appendChild(li);
      }
    }
    $('nguoiNghiaTrang').value = 0; $('nguoiThienTai').value = 0; $('nguoiMoiTruong').value = 0;
    $('soThang').value = caiDat.soThang;
    hien('buoc3'); tinhLai();
    $('buoc3').scrollIntoView({ behavior: 'smooth', block: 'start' });
    $('nguoiNghiaTrang').focus();
  }
  function tinhLai() {
    if (!hoDangChon) return;
    const tong = Thue.sumThueDat(hoDangChon.rows);
    const r = Thue.computeTotals({
      thueDat: tong.thueDat,
      nguoiNghiaTrang: $('nguoiNghiaTrang').value,
      nguoiThienTai: $('nguoiThienTai').value,
      nguoiMoiTruong: $('nguoiMoiTruong').value,
      soThang: $('soThang').value,
    }, caiDat);
    const L = Object.fromEntries(r.lines.map(l => [l.key, l]));
    $('tienThueDat').textContent = Thue.formatMoney(L.thueDat.amount);
    $('ctThueDat').textContent = hoDangChon.rows.length > 1 ? `Cộng ${hoDangChon.rows.length} dòng trong sổ bộ` : 'Theo sổ bộ thuế';
    const gc = $('ghiChuThueDat');
    let ghiChu = L.thueDat.note || '';
    if (!ghiChu && tong.soDongThieu > 0) ghiChu = `${tong.soDongThieu} dòng không có số thuế, tính 0 đ`;
    if (ghiChu) { gc.textContent = ghiChu; hien('ghiChuThueDat'); } else an('ghiChuThueDat');
    $('ctNghiaTrang').textContent = L.nghiaTrang.detail; $('tienNghiaTrang').textContent = Thue.formatMoney(L.nghiaTrang.amount);
    $('ctThienTai').textContent = L.thienTai.detail; $('tienThienTai').textContent = Thue.formatMoney(L.thienTai.amount);
    $('ctMoiTruong').textContent = L.moiTruong.detail; $('tienMoiTruong').textContent = Thue.formatMoney(L.moiTruong.amount);
    $('tienTong').textContent = Thue.formatMoney(r.total);
  }
  for (const id of ['nguoiNghiaTrang', 'nguoiThienTai', 'nguoiMoiTruong', 'soThang']) {
    $(id).addEventListener('input', tinhLai);
    $(id).addEventListener('focus', e => e.target.select());
  }
  $('nutHoKhac').addEventListener('click', () => {
    hoDangChon = null; an('buoc3'); daChon.clear();
    $('oTim').value = ''; timKiem();
    $('buoc2').scrollIntoView({ behavior: 'smooth', block: 'start' });
    $('oTim').focus();
  });

  function hien(id) { $(id).classList.remove('an'); }
  function an(id) { $(id).classList.add('an'); }
})();
</script>
</body>
</html>
```

- [ ] **Step 2: Kiểm tra bằng trình duyệt trong app**

Tạo `.claude/launch.json`:

```json
{
  "version": "0.0.1",
  "configurations": [
    { "name": "thue", "runtimeExecutable": "python3", "runtimeArgs": ["-m", "http.server", "8765"], "port": 8765 }
  ]
}
```

Dùng `preview_start` với `name: "thue"`, mở `http://localhost:8765/index.html`. Kiểm tra:

1. Trang hiện Bước 1, không hiện Bước 2 và 3. Không có lỗi trong console (`read_console_messages`).
2. Nạp file mẫu vào ô chọn file bằng `javascript_tool` (Chromium cho phép gán `input.files` qua `DataTransfer`):

```js
const buf = await fetch('test/mau.xlsx').then(r => r.arrayBuffer());
const dt = new DataTransfer();
dt.items.add(new File([buf], 'mau.xlsx'));
const inp = document.getElementById('fileExcel');
inp.files = dt.files;
inp.dispatchEvent(new Event('change', { bubbles: true }));
await new Promise(r => setTimeout(r, 500));
document.getElementById('trangThaiFile').textContent;
```

   Expected: chuỗi bắt đầu bằng `Đã nạp 11 dòng`, Bước 2 hiện ra.
3. Bấm ô `#oTim`, gõ `van an`. Expected: danh sách 3 dòng (LÊ VĂN AN ×2, NGUYỄN VĂN AN), mỗi dòng có ô tích, ngày sinh, CCCD, thôn, diện tích, số thuế; nút "Tính cho 0 dòng đã chọn" bị mờ.
4. Bấm dòng LÊ VĂN AN thứ nhất. Expected: cả 2 dòng LÊ VĂN AN được tích, NGUYỄN VĂN AN không; nút đổi thành "Tính cho 2 dòng đã chọn".
5. Bấm nút đó. Expected: Bước 3 hiện, khối thông tin liệt kê 2 dòng, thuế đất `81.000 đ`, chi tiết "Cộng 2 dòng trong sổ bộ", tổng `81.000 đ`.
6. Nhập số người 4 / 3 / 4, số tháng 6. Expected: `60.000 đ`, `30.000 đ`, `360.000 đ`, tổng `531.000 đ`.
7. Bấm "Tính cho hộ khác", tìm `le van binh`, bấm dòng, bấm "Tính cho 1 dòng đã chọn". Expected: thuế đất `0 đ` kèm ghi chú "Không tìm thấy trong file".
8. Tìm `zzz`. Expected: hộp "Không tìm thấy trong file" và nút "Tính với thuế đất 0 đ"; bấm nút → Bước 3 với thuế đất 0 đ.
9. Mở khối Đơn giá, đổi nghĩa trang thành 20000, tải lại trang. Expected: giá trị vẫn là 20000. Bấm "Khôi phục mặc định" → 15000.
10. `resize_window` preset `mobile`, chụp `screenshot`. Expected: không tràn ngang, chữ và nút đủ to. Trả về `desktop`.

Ghi kết quả từng mục vào báo cáo.

- [ ] **Step 3: Commit**

```bash
git add index.html .claude/launch.json
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: giao diện 3 bước tra cứu, chọn nhiều dòng và tính thuế"
```

---

### Task 7: README, cập nhật spec, kiểm tra tổng

**Files:**
- Create: `README.md`
- Modify: `docs/superpowers/specs/2026-09-12-trang-tinh-thue-design.md` (mục Kiến trúc và Kiểm thử)

- [ ] **Step 1: Viết README**

```markdown
# Tính tiền thuế cần nộp

Trang web tĩnh tra cứu thuế sử dụng đất phi nông nghiệp từ file Excel sổ bộ
thuế và tính tổng cùng ba khoản thu theo nhân khẩu: nghĩa trang nhân dân, quỹ
phòng chống thiên tai, bảo vệ môi trường.

## Cách dùng

1. Mở trang (địa chỉ GitHub Pages hoặc mở `index.html`).
2. Bước 1: bấm **Chọn file Excel**, chọn file sổ bộ thuế (.xlsx/.xls).
   File chỉ được đọc trên máy bạn, không gửi đi đâu.
3. Bước 2: gõ tên hoặc số CCCD, bấm vào hộ cần tính.
4. Bước 3: nhập số người từng khoản, xem tổng tiền.

Đơn giá có thể sửa ở khối **Đơn giá** cuối trang, được ghi nhớ trên trình duyệt.

## Yêu cầu file Excel

Sheet đầu tiên phải có dòng tiêu đề chứa các cột: `MST`, `Tên NNT`, `Số CCCD`,
`Tổng số thuế phải nộp`. Các cột khác (Ngày sinh, Diện tích, Thôn) là tuỳ chọn.
Xem ví dụ tại `test/mau.xlsx`.

## Đưa lên GitHub Pages

1. Tạo repo trên GitHub, đẩy nhánh `main`.
2. Vào **Settings → Pages**, mục *Build and deployment* chọn *Deploy from a
   branch*, nhánh `main`, thư mục `/ (root)`, bấm Save.
3. Sau khoảng một phút, truy cập `https://<tên-user>.github.io/<tên-repo>/`.

## Phát triển

- Chạy test: `node test/test.js`
- Sinh lại file mẫu: `python3 test/tao_mau.py`
- Không đưa file Excel thật lên repo (`.gitignore` đã chặn `*.xlsx` trừ `test/mau.xlsx`).
```

- [ ] **Step 2: Cập nhật spec cho khớp cấu trúc file**

Trong mục "Kiến trúc" của spec, thay câu `Một file \`index.html\` duy nhất chứa HTML, CSS, JavaScript.` bằng:

```
Hai file: `thue.js` chứa các hàm thuần (chạy được trong Node và trình duyệt)
và `index.html` chứa giao diện, nạp SheetJS và `thue.js`.
```

Trong mục "Kiểm thử", thay đoạn `File \`test/test.html\` chạy các hàm thuần bằng assert đơn giản trong trình duyệt:` bằng `File \`test/test.js\` chạy bằng \`node test/test.js\` với \`node:assert\`:`.

- [ ] **Step 3: Chạy toàn bộ kiểm tra**

```bash
node test/test.js && git status --short
```

Expected: `16 test đạt`, exit 0; `git status` chỉ còn README.md, spec, và không có file ảnh/Excel thật.

- [ ] **Step 4: Commit**

```bash
git add README.md docs/superpowers/specs/2026-09-12-trang-tinh-thue-design.md
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "docs: README và cập nhật spec theo cấu trúc file"
```

---

### Task 8: Bố cục mới theo yêu cầu người dùng, tối ưu di động

**Files:**
- Modify: `index.html` (viết lại phần HTML/CSS và các hàm giao diện; giữ nguyên các hàm đọc file, tìm kiếm, đơn giá/localStorage đã có)
- Modify: `README.md` (mục "Cách dùng")

**Interfaces:**
- Consumes: `window.Thue.{parseSheetRows, searchInfo, sumThueDat, computeTotals, formatMoney, normalizeText, DEFAULT_RATES}`. Row: `{ mst, ten, ngaySinh, cccd, dienTich, thon, stt, tieuMuc, thuePhaiNop, daNop }`.
- Produces: trang theo mục "Giao diện" của spec (đã cập nhật), hàm `cungHo(a, b)` đối xứng.

- [ ] **Step 1: Sắp lại bố cục trong `index.html`**

Thứ tự phần tử trong `<main>`:

1. `<h1>Tính tiền thuế cần nộp</h1>` và câu mô tả ngắn.
2. `<details id="khoiDonGia">` chuyển lên ngay dưới tiêu đề. `<summary>` gồm chữ "Đơn giá" và một `<span id="tomTatDonGia" class="phu-chu">` hiện: `Nghĩa trang {NT} đ/người · Thiên tai {TT} đ/người · Môi trường {MT} đ/người/tháng · {thang} tháng` (định dạng bằng `formatMoney` bỏ hậu tố " đ" hoặc dùng `toLocaleString('vi-VN')`). Cập nhật tóm tắt mỗi khi đơn giá đổi. Nội dung bên trong giữ như cũ (4 ô nhập + "Khôi phục mặc định").
3. Bước 1 (giữ nguyên).
4. Bước 2 (giữ nguyên danh sách kết quả, thông báo cắt, hộp không tìm thấy). Bỏ nút "Tính cho N dòng đã chọn"; thay bằng việc hiện Bước 3 ngay khi `daChon.size > 0`, ẩn khi về 0.
5. Bước 3 `<section id="buoc3">` "Nhập số người":
   - `<div id="thongTinHo" class="thongtin">` hiện tên hộ và `N dòng đã chọn` (hoặc "Không có dòng nào trong file Excel" khi tính với 0 đ).
   - Ba nhóm nhập, mỗi nhóm một hàng: nhãn trái, ô số phải (rộng 120px trên máy tính, chiếm hết hàng trên điện thoại):
     - "Nghĩa trang nhân dân — số người" `#nguoiNghiaTrang`
     - "Quỹ phòng chống thiên tai — số người" `#nguoiThienTai`
     - "Bảo vệ môi trường — số người" `#nguoiMoiTruong`, ngay dưới là "Số tháng (chỉ áp dụng cho bảo vệ môi trường)" `#soThang`, mặc định `caiDat.soThang`.
   - Nút `#nutTinhTien` class `nut nut-lon` chữ "Tính tiền", chiều ngang 100%.
6. `<section id="hoaDon" class="buoc an">` "Hoá đơn":
   - Tên hộ, danh sách `<ul id="hdDong">` các dòng sổ bộ đã chọn (mỗi dòng: `STT x · Tiểu mục y · Đã nộp z` nếu có, và số thuế hoặc "không có số thuế").
   - `<div class="hd-hang">` cho mỗi khoản: bên trái `.hd-ten` (tên khoản) và `.hd-ct` (chi tiết: `15.000 đ × 4 người` / `15.000 đ × 4 người × 6 tháng` / `Cộng N dòng trong sổ bộ`), bên phải `.hd-tien`. Ghi chú thiếu số thuế hiện bằng `.ghichu` dưới tên khoản.
   - `<div class="hd-hang hd-tong">` "TỔNG CỘNG" + số tiền, cỡ chữ 1.4rem.
   - Nút "Tính cho hộ khác" `#nutHoKhac`.
   - Thời điểm tính: `Tính lúc HH:MM dd/mm/yyyy` bằng `.phu-chu`.

CSS bổ sung (mobile-first):

```css
.nut-lon { width: 100%; font-size: 1.15rem; min-height: 56px; }
.nhom { display: flex; flex-wrap: wrap; gap: 8px 12px; align-items: center; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid var(--vien); }
.nhom label { flex: 1 1 220px; }
.nhom input { flex: 0 0 120px; text-align: center; }
.hd-hang { display: flex; justify-content: space-between; gap: 12px; padding: 12px 0; border-bottom: 1px solid var(--vien); }
.hd-hang .hd-tien { white-space: nowrap; font-weight: 700; text-align: right; }
.hd-tong { border-top: 3px solid var(--chu); border-bottom: 0; font-size: 1.4rem; font-weight: 800; }
@media (max-width: 480px) {
  main { padding: 12px 10px 40px; }
  .buoc { padding: 14px; border-radius: 10px; }
  .nhom input { flex: 1 1 100%; }
  .ketqua li { flex-wrap: wrap; }
  .ketqua .tien { width: 100%; text-align: right; }
  summary { font-size: 1rem; }
}
```

Bỏ `<table>` của Bước 3 cũ.

- [ ] **Step 2: Sửa logic JS**

- `cungHo(a, b)` đối xứng:

```js
function cccdHopLe(c) { const s = String(c || '').trim(); return s !== '' && s !== '0'; }
function cungHo(a, b) {
  if (Thue.normalizeText(a.ten) !== Thue.normalizeText(b.ten)) return false;
  const ca = String(a.cccd || '').trim(), cb = String(b.cccd || '').trim();
  if (cccdHopLe(ca) && cccdHopLe(cb) && ca === cb) return true;
  const ma = String(a.mst || '').trim(), mb = String(b.mst || '').trim();
  return ma !== '' && ma === mb;
}
```

- `batTat(h)` như cũ; sau `veKetQua()` gọi `capNhatBuoc3()`: nếu `daChon.size > 0` thì `hoDangChon = { ten: rows[0].ten, rows }` (rows theo thứ tự trong `cacHo`), cập nhật `#thongTinHo`, hiện Bước 3; nếu 0 và không ở chế độ "0 đ" thì ẩn Bước 3 và hoá đơn. Không reset các ô số người khi thay đổi tích (giữ số người đã gõ).
- `#nutKhongThue` → `hoDangChon = { ten: <chuỗi tìm> || 'Hộ chưa có trong sổ', rows: [] }`, hiện Bước 3, cuộn tới.
- `#nutTinhTien` → gọi `tinhVaVeHoaDon()`: dùng `Thue.sumThueDat(hoDangChon.rows)` và `Thue.computeTotals(...)` như hiện tại, vẽ hoá đơn, hiện `#hoaDon`, cuộn tới. Không có listener `input` tự tính nữa.
- `#nutHoKhac` → xoá `daChon`, `hoDangChon = null`, ẩn Bước 3 và hoá đơn, xoá ô tìm, focus ô tìm, cuộn về Bước 2.
- Đổi đơn giá → cập nhật `#tomTatDonGia`; nếu hoá đơn đang hiện thì vẽ lại hoá đơn với đơn giá mới (gọi `tinhVaVeHoaDon()`).
- Khi nạp file mới → xoá `daChon`, ẩn Bước 3 và hoá đơn.

- [ ] **Step 3: Kiểm tra trên trình duyệt trong app**

`preview_start` `name: "thue"`, `navigate` `http://localhost:8765/index.html`. Nạp file thật `Thue_dat_PNN_Thon_Thong_Nhat.xlsx` bằng `javascript_tool` (DataTransfer như các task trước). Kiểm tra và ghi lại:

1. Thứ tự khối: Đơn giá (thu gọn, có dòng tóm tắt) → Bước 1 → Bước 2; Bước 3 và Hoá đơn ẩn.
2. Tìm `pham van yem`, bấm dòng đầu → 4 dòng tích, Bước 3 hiện ngay bên dưới với "4 dòng đã chọn"; Hoá đơn vẫn ẩn.
3. Nhập 4 / 4 / 4, số tháng 6, bấm "Tính tiền" → Hoá đơn hiện: thuế đất `245.223 đ`, nghĩa trang `60.000 đ`, thiên tai `40.000 đ`, môi trường `360.000 đ`, tổng `705.223 đ`.
4. Đổi số tháng thành 3 mà không bấm → hoá đơn không đổi; bấm "Tính tiền" → môi trường `180.000 đ`, tổng `525.223 đ`.
5. Bỏ tích 1 dòng → Bước 3 "3 dòng đã chọn"; bỏ tích hết → Bước 3 và Hoá đơn ẩn.
6. Gộp hộ đối xứng: tìm `nguyen thi dien` (hộ có CCCD 0, nhiều dòng), bấm lần lượt từng dòng riêng lẻ (bỏ tích rồi bấm dòng khác) → luôn tích đủ các dòng cùng MST.
7. Mở Đơn giá, đổi nghĩa trang 20000 → tóm tắt đổi, hoá đơn (nếu đang hiện) vẽ lại; "Khôi phục mặc định" trả về.
8. `resize_window` width 360 height 740 → chụp ảnh; không tràn ngang (`document.documentElement.scrollWidth <= 360`); nút "Tính tiền" và "Chọn file Excel" rộng hết hàng; ô số chiếm hết hàng. Trả về `desktop`.
9. `read_console_messages` onlyErrors → không có.
10. Nạp `test/mau.xlsx` → "Đã nạp 11 dòng"; tìm `zzz` → nút "Tính với thuế đất 0 đ" → Bước 3 hiện với "Không có dòng nào trong file Excel"; Tính tiền → thuế đất 0 đ và ghi chú "Không tìm thấy trong file".

- [ ] **Step 4: README**

Sửa mục "Cách dùng" thành 5 bước: (0) kiểm tra đơn giá ở đầu trang, chỉ sửa khi có thay đổi; (1) chọn file; (2) tìm và tích hộ; (3) nhập số người, số tháng chỉ áp dụng cho bảo vệ môi trường, bấm "Tính tiền"; (4) xem hoá đơn ở cuối, bấm "Tính cho hộ khác" để tiếp.

- [ ] **Step 5: Commit**

```bash
git add index.html README.md
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: bố cục mới — đơn giá lên đầu, form số người, hoá đơn cuối trang, tối ưu di động"
```

---

### Task 9: Nhớ file Excel đã nạp trong trình duyệt; không cuộn khi focus lại dòng

**Files:**
- Modify: `index.html`
- Modify: `README.md` (mục "Cách dùng", bước 1)

**Interfaces:**
- Consumes: `Thue.parseSheetRows` (đã có), `localStorage`.
- Produces: khoá `localStorage` `thue-data` = JSON `{ v: 1, tenFile: string, luuLuc: number (ms), rows: Row[] }`.

- [ ] **Step 1: Sửa focus gây cuộn**

Trong `veKetQua()` đổi `liCanFocus.focus()` thành `liCanFocus.focus({ preventScroll: true })`.

- [ ] **Step 2: Lưu và nạp lại dữ liệu**

Thêm vào phần Bước 1 của HTML, ngay dưới `#trangThaiFile`, một khối `<div id="daLuu" class="hang-nut an"><button type="button" class="nut phu" id="nutChonKhac">Chọn file khác</button><button type="button" class="nut phu" id="nutXoaLuu">Xoá dữ liệu đã lưu</button></div>`.

JS:

```js
const KHOA_DATA = 'thue-data';
function luuData(tenFile, rows) {
  try {
    localStorage.setItem(KHOA_DATA, JSON.stringify({ v: 1, tenFile, luuLuc: Date.now(), rows }));
    return true;
  } catch (e) { return false; }
}
function docData() {
  try {
    const o = JSON.parse(localStorage.getItem(KHOA_DATA) || 'null');
    if (!o || o.v !== 1 || !Array.isArray(o.rows) || !o.rows.length) return null;
    return o;
  } catch (e) { return null; }
}
function xoaData() { try { localStorage.removeItem(KHOA_DATA); } catch (e) { /* bỏ qua */ } }
function dinhDangLuc(ms) {
  const d = new Date(ms), p = n => String(n).padStart(2, '0');
  return `${p(d.getHours())}:${p(d.getMinutes())} ${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()}`;
}
function apDungData(rows, tenFile, luuLuc, daLuuOk) {
  cacHo = rows; daChon.clear(); hoDangChon = null; an('buoc3'); an('hoaDon');
  const chu = `Đang dùng file "${tenFile}", nạp lúc ${dinhDangLuc(luuLuc)}, ${rows.length} dòng.`;
  baoFile('ok', daLuuOk ? chu : chu + ' Không lưu được vào trình duyệt, lần sau cần chọn file lại.');
  hien('daLuu'); hien('buoc2'); $('oTim').value = ''; timKiem();
}
```

- Trong handler `change` của `#fileExcel`, sau khi `parseSheetRows` thành công: `const ok = luuData(f.name, r.rows); apDungData(r.rows, f.name, Date.now(), ok);` (thay cho đoạn gán `cacHo` và `baoFile` cũ). Khi lỗi: giữ nguyên thông báo lỗi, không đụng dữ liệu đã lưu.
- Khi khởi động (cuối IIFE, sau khi kiểm tra `XLSX`): `const d = docData(); if (d) apDungData(d.rows, d.tenFile, d.luuLuc, true);`
- `#nutChonKhac` → `$('fileExcel').click()`.
- `#nutXoaLuu` → `xoaData(); cacHo = []; daChon.clear(); hoDangChon = null; an('daLuu'); an('buoc2'); an('buoc3'); an('hoaDon'); baoFile('ok', 'Đã xoá dữ liệu đã lưu. Hãy chọn file Excel.');`
- Câu hướng dẫn Bước 1 thêm: "Trang sẽ nhớ file này trên trình duyệt cho lần sau."

- [ ] **Step 3: Kiểm tra trên trình duyệt**

1. Xoá `localStorage` (`localStorage.clear()`), tải lại: Bước 1 như cũ, không có khối `#daLuu`.
2. Nạp file thật qua DataTransfer → trạng thái "Đang dùng file "that.xlsx", nạp lúc ..., 464 dòng."; `JSON.parse(localStorage.getItem('thue-data')).rows.length === 464`; kích cỡ `localStorage.getItem('thue-data').length` ghi lại (kỳ vọng dưới 500.000 ký tự).
3. Tải lại trang (`navigate` cùng URL) → Bước 1 tự hiện trạng thái đang dùng file, Bước 2 hiện, tìm `pham van yem` ra 4 dòng mà không cần chọn file.
4. Bấm "Xoá dữ liệu đã lưu" → `localStorage.getItem('thue-data') === null`, Bước 2 ẩn, thông báo "Đã xoá dữ liệu đã lưu. Hãy chọn file Excel."
5. Nạp `test/mau.xlsx` → 11 dòng, rồi nạp file thật → ghi đè, 464 dòng.
6. Cuộn: tìm `le`, `window.scrollTo(0,0)`, bấm dòng đầu → `window.scrollY` thay đổi không quá 5px, `document.activeElement` là `li`.
7. `read_console_messages` onlyErrors → không có. 360px không tràn ngang.

- [ ] **Step 4: README** — bước 1 ghi rõ trang nhớ file trên trình duyệt, cách chọn file khác và xoá dữ liệu.

- [ ] **Step 5: Commit**

```bash
git add index.html README.md
git -c user.name="phucvm3" -c user.email="phucvm3@fpt.com" commit -m "feat: nhớ file Excel đã nạp trong trình duyệt, không cuộn khi tích dòng"
```
