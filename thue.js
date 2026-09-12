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
      .replace(/[̀-ͯ]/g, '')
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

  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney, parseSheetRows, searchRows };
});
