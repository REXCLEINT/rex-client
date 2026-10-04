/* REX Office – Diagramme als SVG (fuer Tabelle und Praesentation) */
(function () {
  'use strict';
  const COLORS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const fmtNum = (n) => {
    if (!isFinite(n)) return '';
    const a = Math.abs(n);
    if (a >= 1e6) return (n / 1e6).toLocaleString('de-DE', { maximumFractionDigits: 1 }) + ' Mio.';
    return n.toLocaleString('de-DE', { maximumFractionDigits: a < 10 ? 2 : 0 });
  };

  function niceScale(min, max, ticks = 5) {
    if (min === max) { if (max === 0) max = 1; else if (max > 0) min = 0; else max = 0; }
    const range = max - min;
    const rough = range / ticks;
    const mag = Math.pow(10, Math.floor(Math.log10(rough)));
    const norm = rough / mag;
    const step = (norm < 1.5 ? 1 : norm < 3 ? 2 : norm < 7 ? 5 : 10) * mag;
    const lo = Math.floor(min / step) * step;
    const hi = Math.ceil(max / step) * step;
    const out = [];
    for (let v = lo; v <= hi + step / 2; v += step) out.push(Math.round(v / step) * step);
    return { lo, hi, step, ticks: out };
  }

  function legend(series, x, y, w, color, k = 1) {
    const items = series.map((s, i) => ({ name: s.name || `Reihe ${i + 1}`, c: COLORS[i % COLORS.length] }));
    const widths = items.map(it => (22 + Math.min(140, it.name.length * 6.5)) * k);
    const total = widths.reduce((a, b) => a + b, 0);
    let cx = x + Math.max(0, (w - total) / 2);
    return items.map((it, i) => {
      const g = `<rect x="${cx}" y="${y - 8 * k}" width="${10 * k}" height="${10 * k}" fill="${it.c}"/><text x="${cx + 14 * k}" y="${y + k}" font-size="${Math.round(11 * k)}" fill="${color}">${esc(it.name.slice(0, 22))}</text>`;
      cx += widths[i];
      return g;
    }).join('');
  }

  function render(data, w, h, opts = {}) {
    const type = data.type || 'column';
    const k = Math.max(1, Math.min(w / 520, h / 320, 2.2)); // groessere Schrift bei grossen Diagrammen
    const fs = (n) => Math.round(n * k * 10) / 10;
    const fg = opts.dark ? '#ddd' : '#595959';
    const gridC = opts.dark ? '#444' : '#d9d9d9';
    const bg = opts.transparent ? 'none' : (opts.dark ? '#1e1e1e' : '#ffffff');
    const series = (data.series || []).filter(s => s.values && s.values.length);
    const cats = data.categories && data.categories.length ? data.categories : (series[0] ? series[0].values.map((_, i) => String(i + 1)) : []);
    let svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}" width="${w}" height="${h}" font-family="Calibri, Carlito, 'Segoe UI', Arial, sans-serif">`;
    svg += `<rect width="${w}" height="${h}" fill="${bg}"/>`;
    let top = 12;
    if (data.title) {
      svg += `<text x="${w / 2}" y="${fs(26)}" text-anchor="middle" font-size="${fs(16)}" fill="${opts.dark ? '#eee' : '#404040'}">${esc(data.title)}</text>`;
      top = fs(40);
    }
    if (!series.length) {
      svg += `<text x="${w / 2}" y="${h / 2}" text-anchor="middle" font-size="${fs(13)}" fill="${fg}">Keine Daten – wähle Zahlen in der Tabelle aus.</text></svg>`;
      return svg;
    }
    const legendH = fs(26);
    const bottom = h - legendH - 8;

    if (type === 'pie') {
      const s = series[0];
      const vals = s.values.map(v => Math.max(0, +v || 0));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = w / 2, cy = top + (bottom - top) / 2, r = Math.max(10, Math.min(w / 2 - 20, (bottom - top) / 2 - 6));
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + (v / total) * Math.PI * 2;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0), x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
        const c = COLORS[i % COLORS.length];
        if (v / total >= 0.9999) svg += `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${c}" stroke="${bg === 'none' ? '#fff' : bg}" stroke-width="1.5"/>`;
        else if (v > 0) svg += `<path d="M${cx},${cy} L${x0},${y0} A${r},${r} 0 ${large} 1 ${x1},${y1} Z" fill="${c}" stroke="${bg === 'none' ? '#fff' : bg}" stroke-width="1.5"/>`;
        const am = (a0 + a1) / 2;
        if (v / total > 0.04) svg += `<text x="${cx + r * 0.65 * Math.cos(am)}" y="${cy + r * 0.65 * Math.sin(am) + 4}" text-anchor="middle" font-size="${fs(12)}" fill="#fff" font-weight="600">${Math.round(v / total * 100)} %</text>`;
        a0 = a1;
      });
      svg += legend(cats.map(c => ({ name: String(c) })), 0, h - fs(14), w, fg, k);
      return svg + '</svg>';
    }

    const all = series.flatMap(s => s.values.map(v => +v || 0));
    const sc = niceScale(Math.min(0, ...all), Math.max(0, ...all));
    const horizontal = type === 'bar';
    const labelW = horizontal ? Math.min(120, Math.max(30, Math.max(...cats.map(c => String(c).length)) * 6.5 * k + 8)) : Math.max(...sc.ticks.map(t => fmtNum(t).length)) * 6.5 * k + 12;
    const left = labelW + 6, right = w - 16;
    const plotTop = top, plotBottom = bottom - (horizontal ? 18 : 22) * k;
    const pw = Math.max(10, right - left), ph = Math.max(10, plotBottom - plotTop);

    if (!horizontal) {
      const yOf = v => plotBottom - (v - sc.lo) / (sc.hi - sc.lo) * ph;
      for (const t of sc.ticks) {
        const y = yOf(t);
        svg += `<line x1="${left}" x2="${right}" y1="${y}" y2="${y}" stroke="${gridC}" stroke-width="1"/>`;
        svg += `<text x="${left - 6}" y="${y + 4}" text-anchor="end" font-size="${fs(11)}" fill="${fg}">${fmtNum(t)}</text>`;
      }
      const n = cats.length || 1;
      const bw = pw / n;
      const step = Math.max(1, Math.ceil(n / Math.max(1, pw / 60)));
      cats.forEach((c, i) => {
        if (i % step) return;
        svg += `<text x="${left + bw * i + bw / 2}" y="${plotBottom + 15 * k}" text-anchor="middle" font-size="${fs(11)}" fill="${fg}">${esc(String(c).slice(0, 14))}</text>`;
      });
      if (type === 'line') {
        series.forEach((s, si) => {
          const c = COLORS[si % COLORS.length];
          const pts = s.values.map((v, i) => `${left + bw * i + bw / 2},${yOf(+v || 0)}`);
          svg += `<polyline points="${pts.join(' ')}" fill="none" stroke="${c}" stroke-width="2.5" stroke-linejoin="round" stroke-linecap="round"/>`;
          if (n <= 40) pts.forEach(p => { const [x, y] = p.split(','); svg += `<circle cx="${x}" cy="${y}" r="3.5" fill="${c}" stroke="${bg === 'none' ? '#fff' : bg}" stroke-width="1"/>`; });
        });
      } else {
        const groupW = bw * 0.7, barW = groupW / series.length;
        const y0 = yOf(0);
        series.forEach((s, si) => {
          const c = COLORS[si % COLORS.length];
          s.values.forEach((v, i) => {
            const y = yOf(+v || 0);
            const x = left + bw * i + (bw - groupW) / 2 + barW * si;
            svg += `<rect x="${x}" y="${Math.min(y, y0)}" width="${Math.max(1, barW - 1)}" height="${Math.abs(y0 - y)}" fill="${c}"/>`;
          });
        });
      }
      svg += `<line x1="${left}" x2="${right}" y1="${yOf(0)}" y2="${yOf(0)}" stroke="${opts.dark ? '#777' : '#bfbfbf'}"/>`;
    } else {
      const xOf = v => left + (v - sc.lo) / (sc.hi - sc.lo) * pw;
      for (const t of sc.ticks) {
        const x = xOf(t);
        svg += `<line x1="${x}" x2="${x}" y1="${plotTop}" y2="${plotBottom}" stroke="${gridC}"/>`;
        svg += `<text x="${x}" y="${plotBottom + 14 * k}" text-anchor="middle" font-size="${fs(11)}" fill="${fg}">${fmtNum(t)}</text>`;
      }
      const n = cats.length || 1;
      const bh = ph / n;
      const groupH = bh * 0.7, barH = groupH / series.length;
      const x0 = xOf(0);
      cats.forEach((c, i) => {
        svg += `<text x="${left - 6}" y="${plotTop + bh * i + bh / 2 + 4}" text-anchor="end" font-size="${fs(11)}" fill="${fg}">${esc(String(c).slice(0, 18))}</text>`;
      });
      series.forEach((s, si) => {
        const c = COLORS[si % COLORS.length];
        s.values.forEach((v, i) => {
          const x = xOf(+v || 0);
          const y = plotTop + bh * i + (bh - groupH) / 2 + barH * si;
          svg += `<rect x="${Math.min(x, x0)}" y="${y}" width="${Math.abs(x - x0)}" height="${Math.max(1, barH - 1)}" fill="${c}"/>`;
        });
      });
      svg += `<line x1="${x0}" x2="${x0}" y1="${plotTop}" y2="${plotBottom}" stroke="${opts.dark ? '#777' : '#bfbfbf'}"/>`;
    }
    svg += legend(series, 0, h - fs(14), w, fg, k);
    return svg + '</svg>';
  }

  window.RexCharts = { render, COLORS };
})();
