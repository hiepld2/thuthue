// thue.js — hàm thuần dùng chung cho trang và test
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Thue = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const DEFAULT_RATES = { nghiaTrang: 15000, thienTai: 10000, moiTruong: 10000, soThang: 6 };

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
    const v = Number.isFinite(n) ? n : 0;
    const s = String(Math.round(v)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return s + ' đ';
  }

  // Cột: khoá nội bộ -> { label hiển thị, danh sách từ khoá (đã bỏ dấu), bắt buộc }
  const COLUMNS = [
    { key: 'stt', label: 'STT', keywords: ['stt', 'so thu tu'], required: false },
    { key: 'mst', label: 'MST', keywords: ['mst', 'ma so thue'], required: false },
    { key: 'ten', label: 'Tên NNT', keywords: ['ten nnt', 'ten nguoi nop thue', 'ho ten', 'ho va ten'], required: true },
    { key: 'ngaySinh', label: 'Ngày sinh', keywords: ['ngay sinh'], required: false },
    { key: 'cccd', label: 'Số CCCD', keywords: ['cccd', 'cmnd', 'can cuoc'], required: true },
    { key: 'dienTich', label: 'Diện tích', keywords: ['dien tich'], required: false },
    { key: 'thon', label: 'Thôn', keywords: ['thon'], required: false },
    { key: 'tieuMuc', label: 'Tiểu mục', keywords: ['tieu muc'], required: false },
    { key: 'thuePhaiNop', label: 'Tổng số thuế phải nộp', keywords: ['tong so thue phai nop', 'thue phai nop', 'phai nop'], required: true },
    { key: 'daNop', label: 'Số tiền đã nộp', keywords: ['so tien da nop', 'da nop'], required: false },
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
        stt: get('stt'),
        tieuMuc: get('tieuMuc'),
        thuePhaiNop: map.thuePhaiNop === -1 ? null : parseMoney(r[map.thuePhaiNop]),
        daNop: map.daNop === -1 ? null : parseMoney(r[map.daNop]),
      });
    }
    return { ok: true, rows, headerRowIndex: h };
  }

  function matchesRow(r, q, qDigits) {
    const byName = normalizeText(r.ten).includes(q);
    const byId = qDigits !== '' && String(r.cccd || '').replace(/\s+/g, '').includes(qDigits);
    return byName || byId;
  }

  function searchInfo(rows, query, limit) {
    if (limit === undefined) limit = 50;
    const q = normalizeText(query);
    if (!q) return { rows: [], total: 0 };
    const qDigits = q.replace(/\s+/g, '');
    const out = [];
    let total = 0;
    for (const r of rows) {
      if (matchesRow(r, q, qDigits)) {
        total++;
        if (out.length < limit) out.push(r);
      }
    }
    return { rows: out, total };
  }

  function searchRows(rows, query, limit) {
    return searchInfo(rows, query, limit).rows;
  }

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

  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney, parseSheetRows, searchRows, searchInfo, sumThueDat, computeTotals };
});
