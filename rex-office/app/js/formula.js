/* REX Tabelle – Formel-Engine (Excel-kompatibel, deutsche und englische Funktionsnamen) */
(function () {
  'use strict';

  // ---------- Fehlerwerte ----------
  class FErr { constructor(code) { this.code = code; } toString() { return this.code; } }
  const ERR = {
    DIV0: new FErr('#DIV/0!'), VALUE: new FErr('#WERT!'), REF: new FErr('#BEZUG!'), NAME: new FErr('#NAME?'),
    NUM: new FErr('#ZAHL!'), NA: new FErr('#NV'), CIRC: new FErr('#ZIRKEL!'), NULL: new FErr('#NULL!')
  };
  const ERR_EXCEL = { '#DIV/0!': '#DIV/0!', '#WERT!': '#VALUE!', '#BEZUG!': '#REF!', '#NAME?': '#NAME?', '#ZAHL!': '#NUM!', '#NV': '#N/A', '#ZIRKEL!': '#REF!', '#NULL!': '#NULL!' };
  const EXCEL_ERR = { '#DIV/0!': ERR.DIV0, '#VALUE!': ERR.VALUE, '#REF!': ERR.REF, '#NAME?': ERR.NAME, '#NUM!': ERR.NUM, '#N/A': ERR.NA, '#NULL!': ERR.NULL };
  const isErr = v => v instanceof FErr;

  // ---------- Spalten-Hilfen ----------
  function colName(c) { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function colIndex(name) { let c = 0; for (const ch of name.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return c - 1; }
  function addr(r, c) { return colName(c) + (r + 1); }
  function parseAddr(a) {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a);
    return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null;
  }

  // ---------- Deutsche Funktionsnamen ----------
  const DE = {
    SUMME: 'SUM', MITTELWERT: 'AVERAGE', ANZAHL: 'COUNT', ANZAHL2: 'COUNTA', ANZAHLLEEREZELLEN: 'COUNTBLANK', WENN: 'IF', WENNS: 'IFS',
    UND: 'AND', ODER: 'OR', NICHT: 'NOT', XODER: 'XOR', RUNDEN: 'ROUND', AUFRUNDEN: 'ROUNDUP', ABRUNDEN: 'ROUNDDOWN', GANZZAHL: 'INT',
    WURZEL: 'SQRT', POTENZ: 'POWER', REST: 'MOD', ZUFALLSZAHL: 'RAND', ZUFALLSBEREICH: 'RANDBETWEEN', SUMMEWENN: 'SUMIF',
    'ZÄHLENWENN': 'COUNTIF', ZAEHLENWENN: 'COUNTIF', MITTELWERTWENN: 'AVERAGEIF', SUMMEWENNS: 'SUMIFS', 'ZÄHLENWENNS': 'COUNTIFS',
    ZAEHLENWENNS: 'COUNTIFS', SVERWEIS: 'VLOOKUP', WVERWEIS: 'HLOOKUP', VERGLEICH: 'MATCH', XVERWEIS: 'XLOOKUP', WENNFEHLER: 'IFERROR',
    ISTLEER: 'ISBLANK', ISTZAHL: 'ISNUMBER', ISTTEXT: 'ISTEXT', ISTFEHLER: 'ISERROR', 'LÄNGE': 'LEN', LAENGE: 'LEN', LINKS: 'LEFT', RECHTS: 'RIGHT', TEIL: 'MID',
    GROSS: 'UPPER', KLEIN: 'LOWER', GROSS2: 'PROPER', 'GLÄTTEN': 'TRIM', GLAETTEN: 'TRIM', VERKETTEN: 'CONCATENATE', TEXTVERKETTEN: 'TEXTJOIN',
    WERT: 'VALUE', WECHSELN: 'SUBSTITUTE', FINDEN: 'FIND', SUCHEN: 'SEARCH', WIEDERHOLEN: 'REPT', HEUTE: 'TODAY', JETZT: 'NOW',
    DATUM: 'DATE', JAHR: 'YEAR', MONAT: 'MONTH', TAG: 'DAY', WOCHENTAG: 'WEEKDAY', STUNDE: 'HOUR', SEKUNDE: 'SECOND',
    STABW: 'STDEV', 'STABW.S': 'STDEV', VARIANZ: 'VAR', 'VAR.S': 'VAR', PRODUKT: 'PRODUCT', SUMMENPRODUKT: 'SUMPRODUCT',
    'KGRÖSSTE': 'LARGE', KGROESSTE: 'LARGE', KKLEINSTE: 'SMALL', RANG: 'RANK', 'RANG.GLEICH': 'RANK', WAHL: 'CHOOSE', ZEILEN: 'ROWS', SPALTEN: 'COLUMNS',
    OBERGRENZE: 'CEILING', UNTERGRENZE: 'FLOOR', ZEILE: 'ROW', SPALTE: 'COLUMN', 'MODUS.EINF': 'MODE', MODALWERT: 'MODE', ANZAHLEINDEUTIG: 'COUNTUNIQUE',
    MAXWENNS: 'MAXIFS', MINWENNS: 'MINIFS', VORZEICHEN: 'SIGN', KÜRZEN: 'TRUNC', KUERZEN: 'TRUNC', GERADE: 'EVEN', UNGERADE: 'ODD', FAKULTÄT: 'FACT', FAKULTAET: 'FACT',
    ISTGERADE: 'ISEVEN', ISTUNGERADE: 'ISODD', ZEICHEN: 'CHAR', CODE: 'CODE', IDENTISCH: 'EXACT', ERSETZEN: 'REPLACE', TAGE: 'DAYS',
    MONATSENDE: 'EOMONTH', EDATUM: 'EDATE', KALENDERWOCHE: 'WEEKNUM', NETTOARBEITSTAGE: 'NETWORKDAYS', WAHR: 'TRUE', FALSCH: 'FALSE'
  };
  const EN_TO_DE = {};
  for (const [de, en] of Object.entries(DE)) if (!EN_TO_DE[en] && !/AE|OE|UE/.test(de)) EN_TO_DE[en] = de;

  // ---------- Datum ----------
  const EPOCH = Date.UTC(1899, 11, 30);
  const dateToSerial = (d) => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours(), d.getMinutes(), d.getSeconds()) - EPOCH) / 86400000;
  const serialToDate = (s) => { const d = new Date(EPOCH + Math.round(s * 86400000)); return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds()); };

  // ---------- Tokenizer ----------
  const T = { NUM: 'num', STR: 'str', REF: 'ref', FUNC: 'func', OP: 'op', LP: '(', RP: ')', SEP: 'sep', BOOL: 'bool', ERR: 'err', NAME: 'name' };
  function tokenize(src) {
    const toks = [];
    let i = 0;
    const s = src;
    while (i < s.length) {
      const ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, str = '';
        while (j < s.length) {
          if (s[j] === '"') { if (s[j + 1] === '"') { str += '"'; j += 2; continue; } break; }
          str += s[j++];
        }
        toks.push({ t: T.STR, v: str, raw: s.slice(i, j + 1) });
        i = j + 1;
        continue;
      }
      if (ch === '#') {
        const m = /^#(DIV\/0!|WERT!|VALUE!|BEZUG!|REF!|NAME\?|ZAHL!|NUM!|NV|N\/A|NULL!)/i.exec(s.slice(i));
        if (m) { toks.push({ t: T.ERR, v: m[0].toUpperCase(), raw: m[0] }); i += m[0].length; continue; }
      }
      // Bezug mit Tabellenname: 'Name'!A1 oder Name!A1
      let m = /^('(?:[^']|'')+'|[A-Za-z_À-ɏ][\w.À-ɏ]*)!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})/.exec(s.slice(i));
      if (m) {
        const sheet = m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1];
        toks.push({ t: T.REF, sheet, v: m[2], raw: m[0] });
        i += m[0].length;
        continue;
      }
      m = /^\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(])/.exec(s.slice(i));
      if (m) { toks.push({ t: T.REF, v: m[0], raw: m[0] }); i += m[0].length; continue; }
      m = /^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(])/.exec(s.slice(i));
      if (m) { toks.push({ t: T.REF, v: m[0], raw: m[0] }); i += m[0].length; continue; }
      m = /^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?/.exec(s.slice(i)) || /^\.\d+/.exec(s.slice(i));
      if (m) { toks.push({ t: T.NUM, v: parseFloat(m[0]), raw: m[0] }); i += m[0].length; continue; }
      m = /^[A-Za-z_À-ɏ][\w.À-ɏ]*(?=\s*\()/.exec(s.slice(i));
      if (m) { toks.push({ t: T.FUNC, v: m[0].toUpperCase(), raw: m[0] }); i += m[0].length; continue; }
      m = /^[A-Za-z_À-ɏ][\w.À-ɏ]*/.exec(s.slice(i));
      if (m) {
        const up = m[0].toUpperCase();
        if (up === 'TRUE' || up === 'WAHR') toks.push({ t: T.BOOL, v: true, raw: m[0] });
        else if (up === 'FALSE' || up === 'FALSCH') toks.push({ t: T.BOOL, v: false, raw: m[0] });
        else toks.push({ t: T.NAME, v: m[0], raw: m[0] });
        i += m[0].length;
        continue;
      }
      if (s.startsWith('<=', i) || s.startsWith('>=', i) || s.startsWith('<>', i)) { toks.push({ t: T.OP, v: s.slice(i, i + 2), raw: s.slice(i, i + 2) }); i += 2; continue; }
      if ('+-*/^&=<>%'.includes(ch)) { toks.push({ t: T.OP, v: ch, raw: ch }); i++; continue; }
      if (ch === '(') { toks.push({ t: T.LP, raw: ch }); i++; continue; }
      if (ch === ')') { toks.push({ t: T.RP, raw: ch }); i++; continue; }
      if (ch === ',' || ch === ';') { toks.push({ t: T.SEP, raw: ch }); i++; continue; }
      throw new Error('Unerwartetes Zeichen: ' + ch);
    }
    return toks;
  }

  // ---------- Parser ----------
  function parse(src) {
    const toks = tokenize(src);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];
    const isOp = (v) => peek() && peek().t === T.OP && (Array.isArray(v) ? v.includes(peek().v) : peek().v === v);

    function expr() { return comparison(); }
    function comparison() {
      let l = concat();
      while (isOp(['=', '<>', '<', '>', '<=', '>='])) { const op = next().v; l = { k: 'bin', op, l, r: concat() }; }
      return l;
    }
    function concat() {
      let l = additive();
      while (isOp('&')) { next(); l = { k: 'bin', op: '&', l, r: additive() }; }
      return l;
    }
    function additive() {
      let l = mult();
      while (isOp(['+', '-'])) { const op = next().v; l = { k: 'bin', op, l, r: mult() }; }
      return l;
    }
    function mult() {
      let l = power();
      while (isOp(['*', '/'])) { const op = next().v; l = { k: 'bin', op, l, r: power() }; }
      return l;
    }
    function power() {
      let l = unary();
      while (isOp('^')) { next(); l = { k: 'bin', op: '^', l, r: unary() }; }
      return l;
    }
    function unary() {
      if (isOp('-')) { next(); return { k: 'neg', e: unary() }; }
      if (isOp('+')) { next(); return unary(); }
      return postfix();
    }
    function postfix() {
      let e = primary();
      while (isOp('%')) { next(); e = { k: 'pct', e }; }
      return e;
    }
    function primary() {
      const t = next();
      if (!t) throw new Error('Formel unvollständig');
      switch (t.t) {
        case T.NUM: return { k: 'num', v: t.v };
        case T.STR: return { k: 'str', v: t.v };
        case T.BOOL: return { k: 'bool', v: t.v };
        case T.ERR: return { k: 'err', v: EXCEL_ERR[t.v] || Object.values(ERR).find(e => e.code === t.v) || ERR.VALUE };
        case T.REF: return { k: 'ref', sheet: t.sheet || null, v: t.v };
        case T.NAME: return { k: 'name', v: t.v };
        case T.LP: { const e = expr(); if (!peek() || peek().t !== T.RP) throw new Error('„)“ fehlt'); next(); return e; }
        case T.FUNC: {
          next(); // (
          const args = [];
          if (peek() && peek().t === T.RP) { next(); return { k: 'fn', name: DE[t.v] || t.v.replace(/^_XLFN\./, ''), args }; }
          while (true) {
            if (peek() && (peek().t === T.SEP || peek().t === T.RP)) args.push({ k: 'empty' });
            else args.push(expr());
            const n = next();
            if (!n) throw new Error('„)“ fehlt');
            if (n.t === T.RP) break;
            if (n.t !== T.SEP) throw new Error('Unerwartetes Zeichen in Funktion');
          }
          return { k: 'fn', name: DE[t.v] || t.v.replace(/^_XLFN\./, ''), args };
        }
        default: throw new Error('Unerwartet: ' + (t.raw || t.t));
      }
    }
    const ast = expr();
    if (p < toks.length) throw new Error('Unerwartet: ' + toks[p].raw);
    return ast;
  }

  // ---------- Bereiche ----------
  function refRange(v) {
    const [a, b] = v.split(':');
    if (/^\$?[A-Za-z]{1,3}$/.test(a)) {
      return { r1: 0, c1: colIndex(a.replace(/\$/g, '')), r2: Infinity, c2: colIndex(b.replace(/\$/g, '')) };
    }
    const p1 = parseAddr(a), p2 = b ? parseAddr(b) : p1;
    return { r1: Math.min(p1.r, p2.r), c1: Math.min(p1.c, p2.c), r2: Math.max(p1.r, p2.r), c2: Math.max(p1.c, p2.c) };
  }

  // Range-Objekt fuer Funktionen
  class RangeVal {
    constructor(ctx, sheet, rg) { this.ctx = ctx; this.sheet = sheet; this.rg = rg; }
    get rows() { return Math.min(this.rg.r2, this.ctx.maxRow(this.sheet)) - this.rg.r1 + 1; }
    get cols() { return this.rg.c2 - this.rg.c1 + 1; }
    get(i, j) { return this.ctx.value(this.sheet, this.rg.r1 + i, this.rg.c1 + j); }
    *values() {
      const r2 = Math.min(this.rg.r2, this.ctx.maxRow(this.sheet));
      for (let r = this.rg.r1; r <= r2; r++) for (let c = this.rg.c1; c <= this.rg.c2; c++) yield this.ctx.value(this.sheet, r, c);
    }
    toArray() {
      const out = [];
      const r2 = Math.min(this.rg.r2, this.ctx.maxRow(this.sheet));
      for (let r = this.rg.r1; r <= r2; r++) { const row = []; for (let c = this.rg.c1; c <= this.rg.c2; c++) row.push(this.ctx.value(this.sheet, r, c)); out.push(row); }
      return out;
    }
  }

  // ---------- Umwandlungen ----------
  function toNum(v) {
    if (isErr(v)) throw v;
    if (v instanceof RangeVal) return toNum(v.get(0, 0));
    if (v == null || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    const n = parseNumber(String(v));
    if (n == null) throw ERR.VALUE;
    return n;
  }
  function toStr(v) {
    if (isErr(v)) throw v;
    if (v instanceof RangeVal) return toStr(v.get(0, 0));
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'WAHR' : 'FALSCH';
    if (typeof v === 'number') return fmtGeneral(v);
    return String(v);
  }
  function toBool(v) {
    if (isErr(v)) throw v;
    if (v instanceof RangeVal) return toBool(v.get(0, 0));
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v == null || v === '') return false;
    const u = String(v).toUpperCase();
    if (u === 'TRUE' || u === 'WAHR') return true;
    if (u === 'FALSE' || u === 'FALSCH') return false;
    throw ERR.VALUE;
  }
  function parseNumber(s) {
    s = s.trim();
    if (!s) return null;
    let m = /^([+-]?)(\d{1,3}(?:\.\d{3})+|\d+)(?:,(\d+))?\s*(%|€|EUR)?$/i.exec(s);   // deutsch: 1.234,56
    if (m && (s.includes(',') || /\.\d{3}(?!\d)/.test(s) && (s.match(/\./g) || []).length > 1)) {
      let n = parseFloat(m[1] + m[2].replace(/\./g, '') + (m[3] ? '.' + m[3] : ''));
      if (m[4] === '%') n /= 100;
      return n;
    }
    m = /^([+-]?)(\d+(?:\.\d+)?|\.\d+)(?:[eE]([+-]?\d+))?\s*(%|€|EUR)?$/i.exec(s);
    if (m) {
      let n = parseFloat(m[1] + m[2] + (m[3] ? 'e' + m[3] : ''));
      if (m[4] === '%') n /= 100;
      return n;
    }
    return null;
  }
  function fmtGeneral(n) {
    if (!isFinite(n)) return '#ZAHL!';
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    const a = Math.abs(n);
    if (a !== 0 && (a < 1e-9 || a >= 1e15)) return n.toExponential(5).replace(/\.?0+e/, 'e').replace('.', ',').toUpperCase();
    return String(parseFloat(n.toPrecision(11))).replace('.', ',');
  }
  function flat(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof RangeVal) { for (const v of a.values()) out.push({ v, fromRange: true }); }
      else out.push({ v: a, fromRange: false });
    }
    return out;
  }
  function nums(args) {
    const out = [];
    for (const { v, fromRange } of flat(args)) {
      if (isErr(v)) throw v;
      if (typeof v === 'number') out.push(v);
      else if (!fromRange && v !== null && v !== '') out.push(toNum(v));
    }
    return out;
  }

  // Kriterien fuer SUMMEWENN usw.
  function criterion(crit) {
    if (crit instanceof RangeVal) crit = crit.get(0, 0);
    if (typeof crit === 'number' || typeof crit === 'boolean') return (v) => v === crit || (typeof v === 'string' && parseNumber(v) === crit);
    const s = String(crit ?? '');
    const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(s);
    const op = m[1] || '=';
    const rhs = m[2];
    const n = parseNumber(rhs);
    if (n != null) {
      return (v) => {
        const x = typeof v === 'number' ? v : (typeof v === 'string' ? parseNumber(v) : null);
        if (x == null) return op === '<>';
        return { '=': x === n, '<>': x !== n, '<': x < n, '>': x > n, '<=': x <= n, '>=': x >= n }[op];
      };
    }
    const re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
    if (op === '=') return (v) => rhs === '' ? (v == null || v === '') : re.test(toStrSafe(v));
    if (op === '<>') return (v) => rhs === '' ? !(v == null || v === '') : !re.test(toStrSafe(v));
    return (v) => {
      const a = toStrSafe(v).toLowerCase(), b = rhs.toLowerCase();
      return { '<': a < b, '>': a > b, '<=': a <= b, '>=': a >= b }[op];
    };
  }
  const toStrSafe = (v) => { try { return toStr(v); } catch { return ''; } };

  function compare(a, b) {
    const rank = (v) => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : -1;
    if (a == null || a === '') a = typeof b === 'string' ? '' : (typeof b === 'boolean' ? false : 0);
    if (b == null || b === '') b = typeof a === 'string' ? '' : (typeof a === 'boolean' ? false : 0);
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  function lookupEq(a, b) {
    if (typeof a === 'string' && typeof b === 'string') return a.toLowerCase() === b.toLowerCase();
    if (typeof a === 'number' && typeof b === 'string') return parseNumber(b) === a;
    if (typeof b === 'number' && typeof a === 'string') return parseNumber(a) === b;
    return a === b;
  }
  function wildEq(pattern, v) {
    if (typeof pattern !== 'string' || !/[*?]/.test(pattern)) return lookupEq(pattern, v);
    const re = new RegExp('^' + pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
    return re.test(toStrSafe(v));
  }

  const round = (n, d) => { const f = Math.pow(10, d); return Math.round((n + Math.sign(n) * Number.EPSILON) * f) / f; };

  // ---------- Funktionen ----------
  // lazy: Argumente werden als Thunks uebergeben (fuer WENN, WENNFEHLER usw.)
  const LAZY = new Set(['IF', 'IFERROR', 'IFNA', 'IFS', 'AND', 'OR', 'CHOOSE']);
  const F = {
    SUM: (a) => nums(a).reduce((s, x) => s + x, 0),
    AVERAGE: (a) => { const n = nums(a); if (!n.length) throw ERR.DIV0; return n.reduce((s, x) => s + x, 0) / n.length; },
    MIN: (a) => { const n = nums(a); return n.length ? Math.min(...n) : 0; },
    MAX: (a) => { const n = nums(a); return n.length ? Math.max(...n) : 0; },
    COUNT: (a) => flat(a).filter(({ v, fromRange }) => typeof v === 'number' || (!fromRange && parseNumber(String(v ?? '')) != null)).length,
    COUNTA: (a) => flat(a).filter(({ v }) => v != null && v !== '').length,
    COUNTBLANK: (a) => flat(a).filter(({ v }) => v == null || v === '').length,
    COUNTUNIQUE: (a) => new Set(flat(a).filter(({ v }) => v != null && v !== '').map(({ v }) => typeof v === 'string' ? v.toLowerCase() : v)).size,
    PRODUCT: (a) => nums(a).reduce((s, x) => s * x, 1),
    MEDIAN: (a) => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw ERR.NUM; const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
    MODE: (a) => { const n = nums(a); const cnt = new Map(); let best = null, bc = 1; for (const x of n) { const c = (cnt.get(x) || 0) + 1; cnt.set(x, c); if (c > bc) { bc = c; best = x; } } if (best == null) throw ERR.NA; return best; },
    STDEV: (a) => { const n = nums(a); if (n.length < 2) throw ERR.DIV0; const m = n.reduce((s, x) => s + x, 0) / n.length; return Math.sqrt(n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1)); },
    VAR: (a) => { const n = nums(a); if (n.length < 2) throw ERR.DIV0; const m = n.reduce((s, x) => s + x, 0) / n.length; return n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - 1); },
    LARGE: ([r, k]) => { const n = nums([r]).sort((x, y) => y - x); const i = toNum(k) - 1; if (i < 0 || i >= n.length) throw ERR.NUM; return n[i]; },
    SMALL: ([r, k]) => { const n = nums([r]).sort((x, y) => x - y); const i = toNum(k) - 1; if (i < 0 || i >= n.length) throw ERR.NUM; return n[i]; },
    RANK: ([x, r, ord]) => { const v = toNum(x); const n = nums([r]); const asc = ord !== undefined && toNum(ord) !== 0; if (!n.includes(v)) throw ERR.NA; return n.filter(y => asc ? y < v : y > v).length + 1; },
    SUMPRODUCT: (a) => {
      const arrs = a.map(r => r instanceof RangeVal ? r.toArray().flat() : [r]);
      const len = arrs[0].length;
      if (arrs.some(x => x.length !== len)) throw ERR.VALUE;
      let s = 0;
      for (let i = 0; i < len; i++) s += arrs.reduce((p, arr) => p * (typeof arr[i] === 'number' ? arr[i] : 0), 1);
      return s;
    },

    IF: ([c, t, f]) => toBool(c()) ? (t ? t() : true) : (f ? f() : false),
    IFS: (args) => { for (let i = 0; i + 1 < args.length; i += 2) if (toBool(args[i]())) return args[i + 1](); throw ERR.NA; },
    IFERROR: ([v, alt]) => { try { const x = v(); if (isErr(x)) return alt(); return x; } catch (e) { if (isErr(e)) return alt(); throw e; } },
    IFNA: ([v, alt]) => { try { const x = v(); if (x === ERR.NA) return alt(); return x; } catch (e) { if (e === ERR.NA) return alt(); throw e; } },
    AND: (args) => { let r = true; for (const a of args) { const v = a(); if (v instanceof RangeVal) { for (const x of v.values()) if (typeof x === 'boolean' || typeof x === 'number') r = r && toBool(x); } else r = r && toBool(v); } return r; },
    OR: (args) => { let r = false; for (const a of args) { const v = a(); if (v instanceof RangeVal) { for (const x of v.values()) if (typeof x === 'boolean' || typeof x === 'number') r = r || toBool(x); } else r = r || toBool(v); } return r; },
    XOR: (a) => flat(a).filter(({ v }) => v != null && v !== '' && toBool(v)).length % 2 === 1,
    NOT: ([v]) => !toBool(v),
    TRUE: () => true,
    FALSE: () => false,
    CHOOSE: ([i, ...opts]) => { const n = Math.trunc(toNum(i())); if (n < 1 || n > opts.length) throw ERR.VALUE; return opts[n - 1](); },

    ROUND: ([n, d]) => round(toNum(n), d === undefined ? 0 : toNum(d)),
    ROUNDUP: ([n, d]) => { const f = 10 ** (d === undefined ? 0 : toNum(d)); const x = toNum(n); return Math.sign(x) * Math.ceil(round(Math.abs(x) * f, 9)) / f; },
    ROUNDDOWN: ([n, d]) => { const f = 10 ** (d === undefined ? 0 : toNum(d)); const x = toNum(n); return Math.sign(x) * Math.floor(round(Math.abs(x) * f, 9)) / f; },
    TRUNC: ([n, d]) => { const f = 10 ** (d === undefined ? 0 : toNum(d)); return Math.trunc(toNum(n) * f) / f; },
    INT: ([n]) => Math.floor(toNum(n)),
    ABS: ([n]) => Math.abs(toNum(n)),
    SIGN: ([n]) => Math.sign(toNum(n)),
    SQRT: ([n]) => { const x = toNum(n); if (x < 0) throw ERR.NUM; return Math.sqrt(x); },
    POWER: ([a, b]) => Math.pow(toNum(a), toNum(b)),
    MOD: ([a, b]) => { const d = toNum(b); if (d === 0) throw ERR.DIV0; const x = toNum(a); return x - d * Math.floor(x / d); },
    PI: () => Math.PI,
    EXP: ([n]) => Math.exp(toNum(n)),
    LN: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log(x); },
    LOG10: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log10(x); },
    LOG: ([n, b]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log(x) / Math.log(b === undefined ? 10 : toNum(b)); },
    SIN: ([n]) => Math.sin(toNum(n)), COS: ([n]) => Math.cos(toNum(n)), TAN: ([n]) => Math.tan(toNum(n)),
    RADIANS: ([n]) => toNum(n) * Math.PI / 180, DEGREES: ([n]) => toNum(n) * 180 / Math.PI,
    RAND: () => Math.random(),
    RANDBETWEEN: ([a, b]) => { const lo = Math.ceil(toNum(a)), hi = Math.floor(toNum(b)); return lo + Math.floor(Math.random() * (hi - lo + 1)); },
    CEILING: ([n, s]) => { const st = s === undefined ? 1 : toNum(s); return st === 0 ? 0 : Math.ceil(toNum(n) / st) * st; },
    FLOOR: ([n, s]) => { const st = s === undefined ? 1 : toNum(s); if (st === 0) throw ERR.DIV0; return Math.floor(toNum(n) / st) * st; },
    EVEN: ([n]) => { const x = toNum(n); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; },
    ODD: ([n]) => { const x = toNum(n); let r = Math.ceil(Math.abs(x)); if (r % 2 === 0) r++; return x < 0 ? -r : r; },
    FACT: ([n]) => { let x = Math.floor(toNum(n)); if (x < 0) throw ERR.NUM; let r = 1; while (x > 1) r *= x--; return r; },
    ISEVEN: ([n]) => Math.floor(Math.abs(toNum(n))) % 2 === 0,
    ISODD: ([n]) => Math.floor(Math.abs(toNum(n))) % 2 === 1,

    SUMIF: ([r, crit, sr]) => {
      const test = criterion(crit);
      let s = 0;
      for (let i = 0; i < r.rows; i++) for (let j = 0; j < r.cols; j++) {
        if (test(r.get(i, j))) { const v = (sr || r).get(i, j); if (typeof v === 'number') s += v; }
      }
      return s;
    },
    COUNTIF: ([r, crit]) => { const test = criterion(crit); let n = 0; for (const v of r.values()) if (test(v)) n++; return n; },
    AVERAGEIF: ([r, crit, ar]) => {
      const test = criterion(crit);
      let s = 0, n = 0;
      for (let i = 0; i < r.rows; i++) for (let j = 0; j < r.cols; j++) {
        if (test(r.get(i, j))) { const v = (ar || r).get(i, j); if (typeof v === 'number') { s += v; n++; } }
      }
      if (!n) throw ERR.DIV0;
      return s / n;
    },
    SUMIFS: ([sr, ...pairs]) => ifs(sr, pairs, vals => vals.reduce((s, x) => s + x, 0)),
    AVERAGEIFS: ([sr, ...pairs]) => ifs(sr, pairs, vals => { if (!vals.length) throw ERR.DIV0; return vals.reduce((s, x) => s + x, 0) / vals.length; }),
    MAXIFS: ([sr, ...pairs]) => ifs(sr, pairs, vals => vals.length ? Math.max(...vals) : 0),
    MINIFS: ([sr, ...pairs]) => ifs(sr, pairs, vals => vals.length ? Math.min(...vals) : 0),
    COUNTIFS: (pairs) => {
      const r0 = pairs[0];
      let n = 0;
      for (let i = 0; i < r0.rows; i++) for (let j = 0; j < r0.cols; j++) {
        let ok = true;
        for (let k = 0; k + 1 < pairs.length; k += 2) if (!criterion(pairs[k + 1])(pairs[k].get(i, j))) { ok = false; break; }
        if (ok) n++;
      }
      return n;
    },

    VLOOKUP: ([key, table, col, approx]) => {
      const k = key instanceof RangeVal ? key.get(0, 0) : key;
      const ci = Math.trunc(toNum(col)) - 1;
      if (ci < 0 || ci >= table.cols) throw ERR.REF;
      const exact = approx !== undefined && !toBool(approx);
      if (exact) {
        for (let i = 0; i < table.rows; i++) if (wildEq(k, table.get(i, 0))) return table.get(i, ci);
        throw ERR.NA;
      }
      let found = -1;
      for (let i = 0; i < table.rows; i++) { const v = table.get(i, 0); if (v == null || v === '') continue; if (compare(v, k) <= 0) found = i; else break; }
      if (found < 0) throw ERR.NA;
      return table.get(found, ci);
    },
    HLOOKUP: ([key, table, row, approx]) => {
      const k = key instanceof RangeVal ? key.get(0, 0) : key;
      const ri = Math.trunc(toNum(row)) - 1;
      if (ri < 0 || ri >= table.rows) throw ERR.REF;
      const exact = approx !== undefined && !toBool(approx);
      if (exact) {
        for (let j = 0; j < table.cols; j++) if (wildEq(k, table.get(0, j))) return table.get(ri, j);
        throw ERR.NA;
      }
      let found = -1;
      for (let j = 0; j < table.cols; j++) { const v = table.get(0, j); if (compare(v, k) <= 0) found = j; else break; }
      if (found < 0) throw ERR.NA;
      return table.get(ri, found);
    },
    INDEX: ([r, row, col]) => {
      let i = row === undefined ? 0 : Math.trunc(toNum(row)) - 1;
      let j = col === undefined ? 0 : Math.trunc(toNum(col)) - 1;
      if (r.rows === 1 && col === undefined) { j = i; i = 0; }
      if (i < 0 || j < 0 || i >= r.rows || j >= r.cols) throw ERR.REF;
      return r.get(i, j);
    },
    MATCH: ([key, r, type]) => {
      const k = key instanceof RangeVal ? key.get(0, 0) : key;
      const arr = r.toArray().flat();
      const t = type === undefined ? 1 : toNum(type);
      if (t === 0) { const i = arr.findIndex(v => wildEq(k, v)); if (i < 0) throw ERR.NA; return i + 1; }
      let found = -1;
      for (let i = 0; i < arr.length; i++) {
        const c = compare(arr[i], k);
        if (t > 0 ? c <= 0 : c >= 0) found = i; else break;
      }
      if (found < 0) throw ERR.NA;
      return found + 1;
    },
    XLOOKUP: ([key, lr, rr, notFound]) => {
      const k = key instanceof RangeVal ? key.get(0, 0) : key;
      const arr = lr.toArray().flat();
      const i = arr.findIndex(v => wildEq(k, v));
      if (i < 0) { if (notFound !== undefined) return notFound; throw ERR.NA; }
      return lr.cols === 1 ? rr.get(i, 0) : rr.get(0, i);
    },
    ROWS: ([r]) => r instanceof RangeVal ? r.rows : 1,
    COLUMNS: ([r]) => r instanceof RangeVal ? r.cols : 1,
    ROW: ([r], ctx) => r instanceof RangeVal ? r.rg.r1 + 1 : ctx.cur.r + 1,
    COLUMN: ([r], ctx) => r instanceof RangeVal ? r.rg.c1 + 1 : ctx.cur.c + 1,

    ISBLANK: ([v]) => { const x = v instanceof RangeVal ? v.get(0, 0) : v; return x == null || x === ''; },
    ISNUMBER: ([v]) => typeof (v instanceof RangeVal ? v.get(0, 0) : v) === 'number',
    ISTEXT: ([v]) => typeof (v instanceof RangeVal ? v.get(0, 0) : v) === 'string' && (v instanceof RangeVal ? v.get(0, 0) : v) !== '',
    ISERROR: ([v]) => isErr(v instanceof RangeVal ? v.get(0, 0) : v),

    LEN: ([s]) => toStr(s).length,
    LEFT: ([s, n]) => toStr(s).slice(0, n === undefined ? 1 : toNum(n)),
    RIGHT: ([s, n]) => { const t = toStr(s); const k = n === undefined ? 1 : toNum(n); return k <= 0 ? '' : t.slice(-k); },
    MID: ([s, st, n]) => toStr(s).substr(toNum(st) - 1, toNum(n)),
    UPPER: ([s]) => toStr(s).toUpperCase(),
    LOWER: ([s]) => toStr(s).toLowerCase(),
    PROPER: ([s]) => toStr(s).toLowerCase().replace(/(^|[^\p{L}])(\p{L})/gu, (m, a, b) => a + b.toUpperCase()),
    TRIM: ([s]) => toStr(s).trim().replace(/\s+/g, ' '),
    CONCATENATE: (a) => a.map(toStr).join(''),
    CONCAT: (a) => flat(a).map(({ v }) => toStr(v)).join(''),
    TEXTJOIN: ([d, ign, ...a]) => { const skip = toBool(ign); return flat(a).map(({ v }) => toStr(v)).filter(s => !(skip && s === '')).join(toStr(d)); },
    REPT: ([s, n]) => toStr(s).repeat(Math.max(0, toNum(n))),
    SUBSTITUTE: ([s, o, n, inst]) => {
      const t = toStr(s), old = toStr(o), nw = toStr(n);
      if (!old) return t;
      if (inst === undefined) return t.split(old).join(nw);
      let k = toNum(inst), idx = -1;
      while (k-- > 0) { idx = t.indexOf(old, idx + 1); if (idx < 0) return t; }
      return t.slice(0, idx) + nw + t.slice(idx + old.length);
    },
    REPLACE: ([s, st, n, nw]) => { const t = toStr(s); const a = toNum(st) - 1; return t.slice(0, a) + toStr(nw) + t.slice(a + toNum(n)); },
    FIND: ([f, s, st]) => { const i = toStr(s).indexOf(toStr(f), st === undefined ? 0 : toNum(st) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    SEARCH: ([f, s, st]) => { const i = toStr(s).toLowerCase().indexOf(toStr(f).toLowerCase(), st === undefined ? 0 : toNum(st) - 1); if (i < 0) throw ERR.VALUE; return i + 1; },
    EXACT: ([a, b]) => toStr(a) === toStr(b),
    VALUE: ([s]) => toNum(s),
    CHAR: ([n]) => String.fromCharCode(toNum(n)),
    CODE: ([s]) => { const t = toStr(s); if (!t) throw ERR.VALUE; return t.charCodeAt(0); },
    TEXT: ([v, f]) => window.RexFormat ? RexFormat.format(toNum(v), toStr(f)) : toStr(v),

    TODAY: () => Math.floor(dateToSerial(new Date())),
    NOW: () => dateToSerial(new Date()),
    DATE: ([y, m, d]) => dateToSerial(new Date(toNum(y), toNum(m) - 1, toNum(d))),
    YEAR: ([s]) => serialToDate(toNum(s)).getFullYear(),
    MONTH: ([s]) => serialToDate(toNum(s)).getMonth() + 1,
    DAY: ([s]) => serialToDate(toNum(s)).getDate(),
    HOUR: ([s]) => serialToDate(toNum(s)).getHours(),
    MINUTE: ([s]) => serialToDate(toNum(s)).getMinutes(),
    SECOND: ([s]) => serialToDate(toNum(s)).getSeconds(),
    WEEKDAY: ([s, t]) => { const d = serialToDate(toNum(s)).getDay(); const type = t === undefined ? 1 : toNum(t); return type === 2 ? (d + 6) % 7 + 1 : type === 3 ? (d + 6) % 7 : d + 1; },
    WEEKNUM: ([s]) => { const d = serialToDate(toNum(s)); const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const day = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - day); const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1)); return Math.ceil(((t - y0) / 86400000 + 1) / 7); },
    DAYS: ([e, s]) => Math.floor(toNum(e)) - Math.floor(toNum(s)),
    EDATE: ([s, m]) => { const d = serialToDate(toNum(s)); return Math.floor(dateToSerial(new Date(d.getFullYear(), d.getMonth() + toNum(m), d.getDate()))); },
    EOMONTH: ([s, m]) => { const d = serialToDate(toNum(s)); return Math.floor(dateToSerial(new Date(d.getFullYear(), d.getMonth() + toNum(m) + 1, 0))); },
    NETWORKDAYS: ([s, e]) => { let a = Math.floor(toNum(s)), b = Math.floor(toNum(e)), sign = 1; if (a > b) { [a, b] = [b, a]; sign = -1; } let n = 0; for (let x = a; x <= b; x++) { const wd = serialToDate(x).getDay(); if (wd !== 0 && wd !== 6) n++; } return n * sign; }
  };
  function ifs(sr, pairs, agg) {
    const vals = [];
    for (let i = 0; i < sr.rows; i++) for (let j = 0; j < sr.cols; j++) {
      let ok = true;
      for (let k = 0; k + 1 < pairs.length; k += 2) if (!criterion(pairs[k + 1])(pairs[k].get(i, j))) { ok = false; break; }
      if (ok) { const v = sr.get(i, j); if (typeof v === 'number') vals.push(v); }
    }
    return agg(vals);
  }

  // ---------- Auswertung ----------
  function evaluate(ast, ctx) {
    switch (ast.k) {
      case 'num': case 'str': case 'bool': return ast.v;
      case 'err': return ast.v;
      case 'empty': return undefined;
      case 'name': throw ERR.NAME;
      case 'ref': {
        const sheet = ast.sheet == null ? ctx.sheet : ctx.sheetIndex(ast.sheet);
        if (sheet < 0) throw ERR.REF;
        const rg = refRange(ast.v);
        if (!ast.v.includes(':')) return ctx.value(sheet, rg.r1, rg.c1);
        return new RangeVal(ctx, sheet, rg);
      }
      case 'neg': return -toNum(evaluate(ast.e, ctx));
      case 'pct': return toNum(evaluate(ast.e, ctx)) / 100;
      case 'bin': {
        let l = evaluate(ast.l, ctx), r = evaluate(ast.r, ctx);
        if (l instanceof RangeVal) l = l.get(0, 0);
        if (r instanceof RangeVal) r = r.get(0, 0);
        if (isErr(l)) throw l;
        if (isErr(r)) throw r;
        switch (ast.op) {
          case '+': return toNum(l) + toNum(r);
          case '-': return toNum(l) - toNum(r);
          case '*': return toNum(l) * toNum(r);
          case '/': { const d = toNum(r); if (d === 0) throw ERR.DIV0; return toNum(l) / d; }
          case '^': return Math.pow(toNum(l), toNum(r));
          case '&': return toStr(l) + toStr(r);
          case '=': return compare(l, r) === 0;
          case '<>': return compare(l, r) !== 0;
          case '<': return compare(l, r) < 0;
          case '>': return compare(l, r) > 0;
          case '<=': return compare(l, r) <= 0;
          case '>=': return compare(l, r) >= 0;
        }
        throw ERR.VALUE;
      }
      case 'fn': {
        const fn = F[ast.name];
        if (!fn) throw ERR.NAME;
        if (LAZY.has(ast.name)) {
          return fn(ast.args.map(a => () => {
            const v = evaluate(a, ctx);
            return v === undefined ? 0 : v;
          }), ctx);
        }
        const args = ast.args.map(a => evaluate(a, ctx));
        while (args.length && args[args.length - 1] === undefined) args.pop();
        return fn(args, ctx);
      }
    }
    throw ERR.VALUE;
  }

  const astCache = new Map();
  function compute(src, ctx) {
    let ast = astCache.get(src);
    if (!ast) {
      try { ast = parse(src); } catch (e) { ast = { k: 'err', v: ERR.NAME, parseError: e.message }; }
      if (astCache.size > 20000) astCache.clear();
      astCache.set(src, ast);
    }
    try {
      let v = evaluate(ast, ctx);
      if (v instanceof RangeVal) v = v.rows === 1 && v.cols === 1 ? v.get(0, 0) : (v.rg.r1 <= ctx.cur.r && ctx.cur.r <= v.rg.r2 && v.cols === 1 ? v.get(ctx.cur.r - v.rg.r1, 0) : ERR.VALUE);
      if (typeof v === 'number' && !isFinite(v)) return ERR.NUM;
      if (v == null) return 0;
      return v;
    } catch (e) {
      if (isErr(e)) return e;
      console.warn('Formelfehler', src, e);
      return ERR.VALUE;
    }
  }
  function check(src) { try { parse(src); return null; } catch (e) { return e.message; } }

  // ---------- Formel-Umschreibung ----------
  // Relative Bezuege verschieben (Kopieren/Ausfuellen)
  function shiftFormula(src, dr, dc) {
    let toks;
    try { toks = tokenize(src); } catch { return src; }
    let out = '', pos = 0;
    for (const t of toks) {
      const at = src.indexOf(t.raw, pos);
      out += src.slice(pos, at);
      if (t.t === T.REF) {
        const pre = t.sheet != null ? t.raw.slice(0, t.raw.length - t.v.length) : '';
        out += pre + t.v.split(':').map(part => shiftRef(part, dr, dc)).join(':');
      } else out += t.raw;
      pos = at + t.raw.length;
    }
    return out + src.slice(pos);
  }
  function shiftRef(part, dr, dc) {
    const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$/.exec(part);
    if (!m) return part;
    let c = colIndex(m[2]), r = m[4] ? +m[4] - 1 : null;
    if (!m[1]) c += dc;
    if (r != null && !m[3]) r += dr;
    if (c < 0 || (r != null && r < 0)) return '#BEZUG!';
    return m[1] + colName(c) + (m[4] ? m[3] + (r + 1) : '');
  }
  // Zeilen/Spalten einfuegen/loeschen: Bezuege anpassen
  function adjustForInsert(src, axis, at, count, sheetName, formulaSheetName) {
    let toks;
    try { toks = tokenize(src); } catch { return src; }
    let out = '', pos = 0;
    for (const t of toks) {
      const at2 = src.indexOf(t.raw, pos);
      out += src.slice(pos, at2);
      const target = t.sheet != null ? t.sheet : formulaSheetName;
      if (t.t === T.REF && target === sheetName) {
        const pre = t.sheet != null ? t.raw.slice(0, t.raw.length - t.v.length) : '';
        const parts = t.v.split(':');
        const adj = parts.map((part, idx) => {
          const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$/.exec(part);
          if (!m) return part;
          let c = colIndex(m[2]), r = m[4] ? +m[4] - 1 : null;
          if (axis === 'row' && r != null) {
            if (count > 0 && r >= at) r += count;
            else if (count < 0) { if (r >= at - count) r += count; else if (r >= at) r = idx === 0 ? at : Math.max(at - 1, 0); }
          }
          if (axis === 'col') {
            if (count > 0 && c >= at) c += count;
            else if (count < 0) { if (c >= at - count) c += count; else if (c >= at) c = idx === 0 ? at : Math.max(at - 1, 0); }
          }
          return m[1] + colName(c) + (m[4] ? m[3] + (r + 1) : '');
        });
        out += pre + adj.join(':');
      } else out += t.raw;
      pos = at2 + t.raw.length;
    }
    return out + src.slice(pos);
  }
  // In Excel-Schreibweise umwandeln (englische Namen, Komma-Trenner)
  function toExcel(src) {
    let toks;
    try { toks = tokenize(src); } catch { return src; }
    let out = '', pos = 0;
    for (const t of toks) {
      const at = src.indexOf(t.raw, pos);
      out += src.slice(pos, at);
      if (t.t === T.FUNC) out += DE[t.v] || t.v;
      else if (t.t === T.SEP) out += ',';
      else if (t.t === T.BOOL) out += t.v ? 'TRUE' : 'FALSE';
      else if (t.t === T.ERR) out += ERR_EXCEL[t.v] || t.v;
      else out += t.raw;
      pos = at + t.raw.length;
    }
    return out + src.slice(pos);
  }
  // Fuer die Anzeige: deutsche Namen und Semikolon
  function toGerman(src) {
    let toks;
    try { toks = tokenize(src); } catch { return src; }
    let out = '', pos = 0;
    for (const t of toks) {
      const at = src.indexOf(t.raw, pos);
      out += src.slice(pos, at);
      if (t.t === T.FUNC) { const en = DE[t.v] || t.v.replace(/^_XLFN\./, ''); out += EN_TO_DE[en] || en; }
      else if (t.t === T.SEP) out += ';';
      else if (t.t === T.BOOL) out += t.v ? 'WAHR' : 'FALSCH';
      else out += t.raw;
      pos = at + t.raw.length;
    }
    return out + src.slice(pos);
  }
  // Funktionsnamen und Bezuege gross schreiben (wie Excel)
  function normalize(src) {
    let toks;
    try { toks = tokenize(src); } catch { return src; }
    let out = '', pos = 0;
    for (const t of toks) {
      const at = src.indexOf(t.raw, pos);
      out += src.slice(pos, at);
      if (t.t === T.FUNC) out += t.raw.toUpperCase();
      else if (t.t === T.REF) out += t.raw.slice(0, t.raw.length - t.v.length) + t.v.toUpperCase();
      else if (t.t === T.BOOL) out += t.raw.toUpperCase();
      else out += t.raw;
      pos = at + t.raw.length;
    }
    return out + src.slice(pos);
  }
  function refsIn(src) {
    try { return tokenize(src).filter(t => t.t === T.REF); } catch { return []; }
  }

  const FUNCTION_LIST = Object.keys(F).filter(n => n !== 'TRUE' && n !== 'FALSE').map(en => ({ en, de: EN_TO_DE[en] || en }));

  window.RexFormula = {
    compute, check, parse, tokenize, normalize, shiftFormula, adjustForInsert, toExcel, toGerman, refsIn, parseNumber, fmtGeneral,
    colName, colIndex, addr, parseAddr, refRange, dateToSerial, serialToDate, isErr, ERR, EXCEL_ERR, ERR_EXCEL, FErr, FUNCTION_LIST, DE, EN_TO_DE
  };
})();
