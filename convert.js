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

  // ---------- xlsx（サンプル推移表のテンプレートに値を書き込む） ----------
  const COLS = 'BCDEFGHIJKLM'.split('');
  const TOTAL_COL = 'N';
  const TEMPLATE_MONTHS = ['10月', '11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月'];
  // テンプレートの行番号
  const ROW = { '売上高': 3, '売上原価': 4, '営業外収益': 30, '営業外費用': 31 };
  SGA_ITEMS.forEach((k, i) => { ROW[k] = 9 + i; });

  function monthNumber(label) { return parseInt(label, 10); }

  function columnYears(months, startYear) {
    const years = [];
    let y = startYear;
    months.forEach((m, i) => {
      if (i > 0 && monthNumber(m) < monthNumber(months[i - 1])) y++;
      years.push(y);
    });
    return years;
  }

  /** テンプレートの数式と同じ計算で、全セルの値を求める（'' は空文字） */
  function computeGrid(values) {
    const g = {};
    const blank = v => (v === 0 ? '' : v); // 0 は空欄（テンプレートと同じ表示・#DIV/0! 回避）
    const sum = a => a.reduce((x, y) => x + (typeof y === 'number' ? y : 0), 0);
    const n = v => (typeof v === 'number' ? v : 0);
    for (const [item, r] of Object.entries(ROW)) {
      g[r] = values[item].map(blank);
      g[r].push(sum(values[item]));
    }
    g[5] = []; g[6] = []; g[7] = []; g[8] = []; g[28] = []; g[29] = []; g[32] = [];
    for (let i = 0; i <= 12; i++) {
      const s = g[3][i];
      g[5][i] = n(s) - n(g[4][i]);
      g[6][i] = s === '' || n(s) === 0 ? '' : n(g[4][i]) / n(s);
      g[7][i] = s === '' || n(s) === 0 ? '' : n(g[9][i]) / n(s);
      g[8][i] = s === '' ? '' : n(g[6][i]) + n(g[7][i]);
      let sga = 0;
      for (let r = 9; r <= 27; r++) sga += n(g[r][i]);
      g[28][i] = sga;
      g[29][i] = g[5][i] - sga;
      g[32][i] = g[29][i] + n(g[30][i]) - n(g[31][i]);
    }
    return g;
  }

  const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const numStr = v => (Number.isInteger(v) ? String(v) : String(Math.round(v * 1e15) / 1e15));

  function fillSheet(xml, grid, months, years) {
    const colIndex = c => (c === TOTAL_COL ? 12 : COLS.indexOf(c));
    const sameMonths = months.length === 12 && months.every((m, i) => m === TEMPLATE_MONTHS[i]);

    // 1行目：年（年が変わる列にだけ表示）
    xml = xml.replace(/(<row r="1"[^>]*>)([\s\S]*?)(<\/row>)/, (m, open, inner, close) => {
      const cells = years.map((y, i) => (i === 0 || y !== years[i - 1]) ? `<c r="${COLS[i]}1" s="28"><v>${y}</v></c>` : '').join('');
      return open + cells + close;
    });

    return xml.replace(/<c r="([A-Z]+)(\d+)"([^>]*?)(\/>|>([\s\S]*?)<\/c>)/g, (m, col, rowStr, attrs, _e, body) => {
      const row = +rowStr;
      const ci = colIndex(col);
      if (ci < 0) return m;
      // 2行目：月の見出し（テンプレートと違う並びのときだけ書き換え）
      if (row === 2) {
        if (sameMonths || col === TOTAL_COL || ci >= months.length) return m;
        const a = attrs.replace(/ t="[^"]*"/, '');
        return `<c r="${col}2"${a} t="inlineStr"><is><t>${xmlEsc(months[ci])}</t></is></c>`;
      }
      const vals = grid[row];
      if (!vals) return m;
      const v = vals[ci];
      const a = attrs.replace(/ t="[^"]*"/, '');
      const f = body && (body.match(/<f[^>]*\/>|<f[^>]*>[\s\S]*?<\/f>/) || [])[0];
      if (f) {
        if (v === '' || v == null) return `<c r="${col}${row}"${a} t="str">${f}<v></v></c>`;
        return `<c r="${col}${row}"${a}>${f}<v>${numStr(v)}</v></c>`;
      }
      if (v === '' || v == null) return `<c r="${col}${row}"${a}/>`;
      return `<c r="${col}${row}"${a}><v>${numStr(v)}</v></c>`;
    });
  }

  /**
   * サンプル推移表と同じ書式の xlsx を作る
   * @param JSZip JSZip ライブラリ
   * @param template テンプレート xlsx のバイト列（Uint8Array / ArrayBuffer）
   * @param pl parsePL の結果
   * @param values aggregate().values
   * @param opts { startYear: 最初の月の年 }
   * @returns Promise<Uint8Array>
   */
  const NO_DIRS = { createFolders: false };
  async function buildXlsx(JSZip, template, pl, values, opts) {
    if (pl.months.length !== 12) throw new Error('12か月分の列があるCSVにのみ対応しています（このCSVは' + pl.months.length + '列）。');
    const years = columnYears(pl.months, opts.startYear);
    const sheetName = String(years[11]);
    const zip = await JSZip.loadAsync(template);
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const sheet = await zip.file(sheetPath).async('string');
    zip.file(sheetPath, fillSheet(sheet, computeGrid(values), pl.months, years), NO_DIRS);

    let wbXml = await zip.file('xl/workbook.xml').async('string');
    wbXml = wbXml.replace(/(<sheet [^>]*name=")[^"]*(")/, `$1${sheetName}$2`)
      .replace(/<calcPr([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/ fullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"/>`);
    zip.file('xl/workbook.xml', wbXml, NO_DIRS);

    const app = zip.file('docProps/app.xml');
    if (app) {
      const x = await app.async('string');
      zip.file('docProps/app.xml', x.replace(/(<TitlesOfParts>[\s\S]*?<vt:lpstr>)[^<]*(<\/vt:lpstr>)/, `$1${sheetName}$2`), NO_DIRS);
    }
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function base64ToBytes(b64) {
    if (typeof atob === 'function') {
      const bin = atob(b64);
      const out = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
      return out;
    }
    return new Uint8Array(Buffer.from(b64, 'base64'));
  }

  return { ITEMS, SGA_ITEMS, EXCLUDE, decode, parseCSV, parsePL, defaultMapping, lineKey, aggregate, derive, checks, computeGrid, buildXlsx, base64ToBytes };
});
