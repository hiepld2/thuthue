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

  return { DEFAULT_RATES, normalizeText, parseMoney, formatMoney };
});
