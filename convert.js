/*
 * 損益計算書（月次推移）CSV → 推移表 xlsx 変換ロジック
 * ブラウザ（window.PLConvert）と Node（module.exports）の両方で使えます。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PLConvert = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // 推移表の集計先（データを入れる行）
  const SGA_ITEMS = [
    '人件費', '採用教育費', '荷造運賃', '広告宣伝費', '旅費交通費', '通信費', '包材費',
    '消耗品費', '修繕費', '水道光熱費', '諸会費', '衛生費', '支払手数料', '販売手数料',
    '地代家賃', '賃借料', '租税公課', '減価償却費', '雑費',
  ];
  const ITEMS = ['売上高', '売上原価', ...SGA_ITEMS, '営業外収益', '営業外費用'];
  const EXCLUDE = '除外';

  // 人件費にまとめる勘定科目
  const PERSONNEL = /給与|給料|雑給|賞与|役員報酬|法定福利|福利厚生|通勤|人件費|退職|アルバイト|パート/;
  // 名前が違っても同じ項目に寄せる別名
  const ALIASES = {
    '旅費交通費': ['旅費', '交通費'],
    '水道光熱費': ['水道', '光熱'],
    '地代家賃': ['家賃', '地代'],
    '賃借料': ['リース料', 'レンタル料'],
    '支払手数料': ['手数料'],
    '雑費': ['雑支出'],
  };

  // ---------- CSV ----------
  function decode(buffer) {
    const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
    } catch (e) {
      return new TextDecoder('shift_jis').decode(bytes);
    }
  }

  function parseCSV(text) {
    const rows = [];
    let row = [], field = '', q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') {
          if (text[i + 1] === '"') { field += '"'; i++; } else q = false;
        } else field += ch;
      } else if (ch === '"') q = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += ch;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    return rows.filter(r => r.some(c => c.trim() !== ''));
  }

  const num = s => {
    if (s == null) return 0;
    const t = String(s).replace(/[,\s円]/g, '').replace(/^△/, '-').replace(/^\((.*)\)$/, '-$1');
    const n = Number(t);
    return Number.isFinite(n) ? n : 0;
  };

  /** 損益計算書CSVを読み取る */
  function parsePL(text) {
    const rows = parseCSV(text);
    const hi = rows.findIndex(r => r.some(c => /^\d{1,2}月$/.test(c.trim())));
    if (hi < 0) throw new Error('「10月」「11月」…の見出し行が見つかりません。損益計算書（月次推移）のCSVか確認してください。');
    const head = rows[hi].map(c => c.trim());
    const monthCols = [];
    head.forEach((c, i) => { if (/^\d{1,2}月$/.test(c)) monthCols.push(i); });
    const months = monthCols.map(i => head[i]);
    const adjCol = head.indexOf('決算整理');
    const totalCol = head.indexOf('合計');

    const lines = [];
    const totals = {};
    let section = '';
    for (const r of rows.slice(hi + 1)) {
      const first = (r[0] || '').trim();
      const name = (r[1] || '').trim();
      const sub = (r[2] || '').trim();
      const hasValues = monthCols.some(i => r[i] !== undefined && r[i].trim() !== '');
      if (first && !hasValues) { section = first; continue; }
      const values = monthCols.map(i => num(r[i]));
      const adj = adjCol >= 0 ? num(r[adjCol]) : 0;
      const total = totalCol >= 0 ? num(r[totalCol]) : values.reduce((a, b) => a + b, 0) + adj;
      if (first) totals[first] = { values, adj, total };
      else if (name) lines.push({ section, name, sub, values, adj, total });
    }
    if (!lines.length) throw new Error('勘定科目の行が見つかりませんでした。');
    return { months, hasAdj: adjCol >= 0, lines, totals };
  }

  // ---------- マッピング ----------
  function lineKey(l) { return l.section + '|' + l.name + '|' + l.sub; }

  /** 勘定科目ごとの初期の集計先と符号 */
  function defaultMapping(line) {
    const { section, name } = line;
    const sign = /期末.*棚卸/.test(name) ? -1 : 1; // 期末棚卸高は原価から控除
    if (/^売上高/.test(section) || section === '売上') return { target: '売上高', sign, known: true };
    if (/売上原価|製造原価/.test(section)) return { target: '売上原価', sign, known: true };
    if (/営業外収益/.test(section)) return { target: '営業外収益', sign, known: true };
    if (/営業外費用/.test(section)) return { target: '営業外費用', sign, known: true };
    if (/販売費|一般管理費|販管費/.test(section)) {
      if (SGA_ITEMS.includes(name)) return { target: name, sign, known: true };
      if (PERSONNEL.test(name)) return { target: '人件費', sign, known: true };
      for (const [item, words] of Object.entries(ALIASES)) {
        if (words.some(w => name.includes(w))) return { target: item, sign, known: true };
      }
      return { target: '雑費', sign, known: false };
    }
    // 特別損益・税金など推移表にない区分
    return { target: EXCLUDE, sign, known: true };
  }

  // ---------- 集計 ----------
  /**
   * @param pl parsePL の結果
   * @param overrides { [lineKey]: target } 画面で変更した集計先
   * @param opts { adjMonthIndex: 決算整理を加算する列（-1で加算しない） }
   */
  function aggregate(pl, overrides, opts) {
    overrides = overrides || {};
    const adjIdx = opts && Number.isInteger(opts.adjMonthIndex) ? opts.adjMonthIndex : pl.months.length - 1;
    const n = pl.months.length;
    const out = {};
    ITEMS.forEach(k => { out[k] = new Array(n).fill(0); });
    const mapped = pl.lines.map(l => {
      const d = defaultMapping(l);
      const target = overrides[lineKey(l)] || d.target;
      if (target !== EXCLUDE && out[target]) {
        for (let i = 0; i < n; i++) out[target][i] += d.sign * l.values[i];
        if (adjIdx >= 0 && adjIdx < n) out[target][adjIdx] += d.sign * l.adj;
      }
      return { ...l, key: lineKey(l), target, sign: d.sign, known: d.known, isDefault: target === d.target };
    });
    return { values: out, mapped };
  }

  /** 推移表の計算行（売上総損益・販管費計・営業損益・経常損益） */
  function derive(values, n) {
    const col = i => {
      const sga = SGA_ITEMS.reduce((a, k) => a + values[k][i], 0);
      const gross = values['売上高'][i] - values['売上原価'][i];
      const op = gross - sga;
      return { gross, sga, op, ord: op + values['営業外収益'][i] - values['営業外費用'][i] };
    };
    return Array.from({ length: n }, (_, i) => col(i));
  }

  /** CSV側の合計行と突き合わせ */
  function checks(pl, values, adjIdx) {
    const n = pl.months.length;
    const d = derive(values, n);
    const sum = a => a.reduce((x, y) => x + y, 0);
    const pick = (re) => {
      const k = Object.keys(pl.totals).find(t => re.test(t));
      if (!k) return null;
      const t = pl.totals[k];
      return adjIdx >= 0 ? t.total : sum(t.values);
    };
    const list = [
      ['売上高', /^売上高合計|^売上高計/, sum(values['売上高'])],
      ['売上原価', /^売上原価合計|^売上原価計/, sum(values['売上原価'])],
      ['販売管理費計', /^販売費及び一般管理費合計|^販売費及び一般管理費計|^販管費合計/, sum(d.map(x => x.sga))],
      ['営業損益', /^営業利益|^営業損益|^営業損失/, sum(d.map(x => x.op))],
      ['経常損益', /^経常利益|^経常損益|^経常損失/, sum(d.map(x => x.ord))],
    ];
    return list.map(([label, re, mine]) => {
      const csv = pick(re);
      return { label, csv, mine, diff: csv == null ? null : mine - csv };
    });
  }

  // ---------- xlsx ----------
  const COL = i => String.fromCharCode(66 + i); // B, C, ...

  function monthNumber(label) { return parseInt(label, 10); }

  /**
   * @param ExcelJS ExcelJS ライブラリ
   * @param pl parsePL の結果
   * @param values aggregate().values
   * @param opts { startYear: 最初の月の年 }
   */
  function buildWorkbook(ExcelJS, pl, values, opts) {
    const n = pl.months.length;
    const last = COL(n - 1);
    const T = COL(n); // 合計列
    const startYear = opts.startYear;

    // 各列の年（1月をまたいだら翌年）
    const years = [];
    let y = startYear;
    pl.months.forEach((m, i) => {
      if (i > 0 && monthNumber(m) < monthNumber(pl.months[i - 1])) y++;
      years.push(y);
    });

    const wb = new ExcelJS.Workbook();
    wb.calcProperties.fullCalcOnLoad = true;
    const ws = wb.addWorksheet(String(years[n - 1]), { views: [{ showGridLines: false }] });

    const FONT = { name: 'Yu Gothic UI', size: 11 };
    const NUM = '#,##0_);[Red](#,##0)';
    const PL = '#,##0;"△ "#,##0';
    const PCT = '0.0%';

    ws.getColumn(1).width = 18.55;
    for (let i = 0; i <= n; i++) ws.getColumn(i + 2).width = 16.11;

    // 1行目：年
    ws.getRow(1).height = 19.5;
    years.forEach((yy, i) => {
      if (i === 0 || yy !== years[i - 1]) {
        const c = ws.getCell(COL(i) + '1');
        c.value = yy; c.numFmt = '0_ '; c.font = FONT;
      }
    });

    // 2行目：見出し
    const h = ws.getRow(2);
    h.getCell(1).value = '勘定科目';
    pl.months.forEach((m, i) => { h.getCell(i + 2).value = m; });
    h.getCell(n + 2).value = '合計';
    h.eachCell(c => { c.font = { ...FONT, bold: true }; c.alignment = { horizontal: c.col === 1 ? 'left' : 'center' }; });

    const d = derive(values, n);
    const sum = a => a.reduce((x, y2) => x + y2, 0);
    let r = 3;
    const rowOf = {};

    function dataRow(label, arr, o) {
      o = o || {};
      const row = ws.getRow(r);
      rowOf[label] = r;
      row.getCell(1).value = label;
      row.getCell(1).font = FONT;
      for (let i = 0; i < n; i++) {
        const c = row.getCell(i + 2);
        c.value = arr[i];
        c.numFmt = o.fmt || NUM; c.font = FONT;
      }
      const t = row.getCell(n + 2);
      t.value = { formula: `SUM(B${r}:${last}${r})`, result: sum(arr) };
      t.numFmt = o.fmt || NUM; t.font = { ...FONT, bold: !!o.boldTotal };
      r++;
    }

    function formulaRow(label, fn, results, o) {
      o = o || {};
      const row = ws.getRow(r);
      rowOf[label] = r;
      row.getCell(1).value = label;
      row.getCell(1).font = FONT;
      for (let i = 0; i <= n; i++) {
        const c = row.getCell(i + 2);
        const L = COL(i);
        c.value = { formula: fn(L, i === n), result: results[i] };
        c.numFmt = o.fmt || NUM; c.font = FONT;
        if (o.fill) c.fill = o.fill;
        if (o.top) c.border = { top: { style: 'thin' } };
      }
      if (o.fill) row.getCell(1).fill = o.fill;
      if (o.top) row.getCell(1).border = { top: { style: 'thin' } };
      r++;
    }

    const salesT = sum(values['売上高']);
    const cogsT = sum(values['売上原価']);
    const laborT = sum(values['人件費']);
    const ratio = (a, b) => (b ? a / b : '');
    const withTotal = (arr, total) => arr.concat([total]);

    dataRow('売上高', values['売上高'], { boldTotal: true });           // 3
    dataRow('売上原価', values['売上原価']);                            // 4
    formulaRow('売上総損益', L => `${L}3-${L}4`,
      withTotal(d.map(x => x.gross), salesT - cogsT));                // 5
    formulaRow('原価率', L => `IF(${L}3="","",${L}4/${L}3)`,
      withTotal(values['売上原価'].map((v, i) => ratio(v, values['売上高'][i])), ratio(cogsT, salesT)), { fmt: PCT }); // 6
    formulaRow('人件費率', L => `IF(${L}3="","",${L}9/${L}3)`,
      withTotal(values['人件費'].map((v, i) => ratio(v, values['売上高'][i])), ratio(laborT, salesT)), { fmt: PCT }); // 7
    const fl = (c, l, s) => (s ? c / s + l / s : '');
    formulaRow('FL', L => `IF(${L}3="","",SUM(${L}6:${L}7))`,
      withTotal(values['売上高'].map((s, i) => fl(values['売上原価'][i], values['人件費'][i], s)), fl(cogsT, laborT, salesT)), { fmt: PCT }); // 8
    SGA_ITEMS.forEach(k => dataRow(k, values[k], { boldTotal: k === '人件費' })); // 9-27
    const sgaT = sum(d.map(x => x.sga));
    formulaRow('販売管理費計', L => `SUM(${L}9:${L}27)`, withTotal(d.map(x => x.sga), sgaT), { top: true }); // 28
    const opT = salesT - cogsT - sgaT;
    const band = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };
    formulaRow('営業損益金額', L => `${L}5-${L}28`, withTotal(d.map(x => x.op), opT), { fmt: PL, fill: band }); // 29
    dataRow('営業外収益', values['営業外収益']);                        // 30
    dataRow('営業外費用', values['営業外費用']);                        // 31
    const ordT = opT + sum(values['営業外収益']) - sum(values['営業外費用']);
    formulaRow('経常損益金額', L => `${L}29+${L}30-${L}31`, withTotal(d.map(x => x.ord), ordT), { fmt: PL, fill: band }); // 32

    for (let i = 3; i < r; i++) ws.getRow(i).getCell(1).alignment = { horizontal: 'left' };
    return wb;
  }

  return { ITEMS, SGA_ITEMS, EXCLUDE, decode, parseCSV, parsePL, defaultMapping, lineKey, aggregate, derive, checks, buildWorkbook };
});
