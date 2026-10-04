/* REX Tabelle – Excel-Zahlenformate (Anzeige auf Deutsch) */
(function () {
  'use strict';
  const MONTHS = ['Januar', 'Februar', 'März', 'April', 'Mai', 'Juni', 'Juli', 'August', 'September', 'Oktober', 'November', 'Dezember'];
  const DAYS = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag'];

  function splitSections(fmt) {
    const out = [];
    let cur = '', q = false, br = false;
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '"' && !br) q = !q;
      if (ch === '[' && !q) br = true;
      if (ch === ']' && !q) br = false;
      if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
      if (ch === ';' && !q && !br) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  // Formatcode in Literale und Platzhalter zerlegen
  function lex(sec) {
    const parts = [];
    let color = null;
    for (let i = 0; i < sec.length; i++) {
      const ch = sec[i];
      if (ch === '"') { const j = sec.indexOf('"', i + 1); parts.push({ lit: sec.slice(i + 1, j < 0 ? undefined : j) }); i = j < 0 ? sec.length : j; continue; }
      if (ch === '\\') { parts.push({ lit: sec[i + 1] || '' }); i++; continue; }
      if (ch === '_') { parts.push({ lit: ' ' }); i++; continue; }
      if (ch === '*') { i++; continue; }
      if (ch === '[') {
        const j = sec.indexOf(']', i);
        const inner = sec.slice(i + 1, j);
        const cur = /^\$([^-\]]*)/.exec(inner);
        if (cur) parts.push({ lit: cur[1] });
        else if (/^(red|rot)$/i.test(inner)) color = '#c00000';
        else if (/^(blue|blau)$/i.test(inner)) color = '#0000ff';
        else if (/^(green|grün)$/i.test(inner)) color = '#008000';
        else if (/^[hms]+$/i.test(inner)) parts.push({ tok: inner.toLowerCase(), elapsed: true });
        i = j;
        continue;
      }
      parts.push({ ch });
    }
    return { parts, color };
  }

  function isDateFormat(sec) {
    const s = sec.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
    return /[dmyhsDMYHS]|AM\/PM/.test(s) && !/[0#?]/.test(s.replace(/\.0+/, ''));
  }

  function formatDate(serial, sec) {
    const d = RexFormula.serialToDate(serial);
    const { parts, color } = lex(sec);
    let s = '';
    const raw = parts.map(p => p.ch || '').join('');
    const hasAmPm = /AM\/PM/i.test(raw);
    // Token-Folge bilden
    const toks = [];
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i];
      if (p.lit != null) { toks.push({ lit: p.lit }); continue; }
      if (p.tok) { toks.push({ t: p.tok, elapsed: true }); continue; }
      const ch = p.ch.toLowerCase();
      if (/[dmyhs]/.test(ch)) {
        let j = i, t = '';
        while (j < parts.length && parts[j].ch && parts[j].ch.toLowerCase() === ch) { t += ch; j++; }
        toks.push({ t });
        i = j - 1;
      } else if (p.ch.toUpperCase() === 'A' && raw.slice(raw.length - (parts.length - i)).toUpperCase().startsWith('AM/PM')) {
        toks.push({ t: 'ampm' }); i += 4;
      } else if (p.ch === '.' && parts[i + 1] && parts[i + 1].ch === '0') {
        let j = i + 1, z = '';
        while (j < parts.length && parts[j].ch === '0') { z += '0'; j++; }
        toks.push({ t: 'frac', n: z.length });
        i = j - 1;
      } else toks.push({ lit: p.ch });
    }
    // m nach h oder vor s = Minuten
    for (let i = 0; i < toks.length; i++) {
      if (toks[i].t && toks[i].t[0] === 'm' && toks[i].t.length <= 2) {
        const prev = toks.slice(0, i).reverse().find(x => x.t);
        const next = toks.slice(i + 1).find(x => x.t);
        if ((prev && prev.t[0] === 'h') || (next && next.t[0] === 's')) toks[i].min = true;
      }
    }
    const pad = (n, l = 2) => String(n).padStart(l, '0');
    for (const t of toks) {
      if (t.lit != null) { s += t.lit; continue; }
      const k = t.t;
      if (t.elapsed) {
        const tot = serial * 24;
        s += k[0] === 'h' ? Math.floor(tot) : k[0] === 'm' ? Math.floor(tot * 60) : Math.floor(tot * 3600);
        continue;
      }
      if (k === 'ampm') { s += d.getHours() < 12 ? 'AM' : 'PM'; continue; }
      if (k === 'frac') { s += ',' + String((serial * 86400) % 1).slice(2, 2 + t.n).padEnd(t.n, '0'); continue; }
      if (k[0] === 'y') s += k.length <= 2 ? pad(d.getFullYear() % 100) : d.getFullYear();
      else if (k[0] === 'd') s += k.length === 1 ? d.getDate() : k.length === 2 ? pad(d.getDate()) : k.length === 3 ? DAYS[d.getDay()].slice(0, 2) : DAYS[d.getDay()];
      else if (k[0] === 'h') { let h = d.getHours(); if (hasAmPm) h = h % 12 || 12; s += k.length === 1 ? h : pad(h); }
      else if (k[0] === 's') s += k.length === 1 ? d.getSeconds() : pad(d.getSeconds());
      else if (k[0] === 'm') {
        if (t.min) s += k.length === 1 ? d.getMinutes() : pad(d.getMinutes());
        else s += k.length === 1 ? d.getMonth() + 1 : k.length === 2 ? pad(d.getMonth() + 1) : k.length === 3 ? MONTHS[d.getMonth()].slice(0, 3) : k.length === 5 ? MONTHS[d.getMonth()][0] : MONTHS[d.getMonth()];
      }
    }
    return { text: s, color };
  }

  function groupInt(str) { return str.replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

  function formatNumber(v, sec, forceAbs) {
    const { parts, color } = lex(sec);
    // Platzhalterbereich finden
    let first = -1, last = -1;
    for (let i = 0; i < parts.length; i++) {
      if (parts[i].ch && /[0#?]/.test(parts[i].ch)) { if (first < 0) first = i; last = i; }
    }
    if (first >= 0) {
      if (parts[first - 1] && parts[first - 1].ch === '.') first--;
      const e1 = parts[last + 1], e2 = parts[last + 2];
      if (e1 && e1.ch && /[Ee]/.test(e1.ch) && e2 && e2.ch && /[+-]/.test(e2.ch)) {
        last += 2;
        while (parts[last + 1] && parts[last + 1].ch && /[0#]/.test(parts[last + 1].ch)) last++;
      }
      while (parts[last + 1] && parts[last + 1].ch === ',') last++;
    }
    if (first < 0) {
      // nur Text/Literale
      return { text: parts.map(p => p.lit ?? (p.ch === '@' ? '' : p.ch)).join(''), color };
    }
    const pat = parts.slice(first, last + 1).map(p => p.ch || '').join('');
    const pctCount = parts.filter(p => p.ch === '%').length;
    let x = forceAbs ? Math.abs(v) : v;
    x *= Math.pow(100, pctCount);
    let numStr;
    const em = /^([0#?.,]*)[Ee]([+-])(0+)/.exec(pat);
    if (em) {
      const dec = (em[1].split('.')[1] || '').length;
      const [mant, exp] = Math.abs(x).toExponential(dec).split('e');
      const e = parseInt(exp, 10);
      numStr = mant.replace('.', ',') + 'E' + (e < 0 ? '-' : (em[2] === '+' ? '+' : '')) + String(Math.abs(e)).padStart(em[3].length, '0');
    } else {
      const core = pat.replace(/%/g, '');
      const [ip, dp = ''] = core.split('.');
      const grouping = ip.includes(',');
      const minDec = (dp.match(/0/g) || []).length;
      const maxDec = (dp.match(/[0#?]/g) || []).length;
      const minInt = (ip.match(/0/g) || []).length;
      const scaleCommas = (/,+$/.exec(ip) || [''])[0].length;
      if (scaleCommas && ip.replace(/,+$/, '').length) x /= Math.pow(1000, scaleCommas);
      let fixed = Math.abs(x).toFixed(maxDec);
      let [iStr, dStr = ''] = fixed.split('.');
      while (dStr.length > minDec && dStr.endsWith('0')) dStr = dStr.slice(0, -1);
      if (iStr === '0' && minInt === 0) iStr = '';
      iStr = iStr.padStart(minInt, '0');
      if (grouping) iStr = groupInt(iStr);
      numStr = iStr + (dStr.length ? ',' + dStr : '');
      if (!forceAbs && x < 0 && parseFloat(fixed) !== 0) numStr = '-' + numStr;
    }
    const before = parts.slice(0, first).map(p => p.lit ?? (p.ch === '%' ? '%' : p.ch)).join('');
    const after = parts.slice(last + 1).map(p => p.lit ?? p.ch).join('');
    return { text: before + numStr + after, color };
  }

  function format(v, fmt) {
    if (v == null || v === '') return { text: '' };
    if (RexFormula.isErr(v)) return { text: v.code, color: null };
    if (typeof v === 'boolean') return { text: v ? 'WAHR' : 'FALSCH' };
    if (!fmt || /^general$/i.test(fmt) || fmt === 'Standard') {
      if (typeof v === 'number') return { text: RexFormula.fmtGeneral(v) };
      return { text: String(v) };
    }
    const secs = splitSections(fmt);
    if (typeof v === 'string') {
      const ts = secs[3] || (secs.length === 1 && secs[0].includes('@') ? secs[0] : null);
      if (!ts) return { text: v };
      return { text: lex(ts).parts.map(p => p.lit ?? (p.ch === '@' ? v : p.ch)).join('') };
    }
    let sec = secs[0], abs = false;
    if (v < 0 && secs.length > 1 && secs[1] !== '') { sec = secs[1]; abs = true; }
    else if (v === 0 && secs.length > 2 && secs[2] !== '') { sec = secs[2]; abs = true; }
    try {
      if (isDateFormat(sec)) {
        if (v < 0) return { text: '#'.repeat(8) };
        return formatDate(v, sec);
      }
      if (/^@$/.test(sec.trim())) return { text: RexFormula.fmtGeneral(v) };
      if (/general|standard/i.test(sec)) return { text: (abs ? '' : '') + RexFormula.fmtGeneral(abs ? Math.abs(v) : v) };
      return formatNumber(v, sec, abs);
    } catch (e) {
      return { text: RexFormula.fmtGeneral(v) };
    }
  }

  window.RexFormat = { format: (v, f) => format(v, f).text, formatFull: format, isDateFormat, MONTHS, DAYS };
})();
