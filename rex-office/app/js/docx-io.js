/* REX Text – Lesen und Schreiben von Word-Dateien (.docx) */
(function () {
  'use strict';

  const PX_TO_TW = 15;              // 1 px = 15 twips (96 dpi)
  const MM_TO_TW = 1440 / 25.4;
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  // Standardformate des Editors (muss zu writer.css passen)
  const TAG_DEFAULTS = {
    p:          { font: 'Calibri', size: 11, color: '000000' },
    title:      { font: 'Calibri Light', size: 28, color: '000000' },
    subtitle:   { font: 'Calibri', size: 11, color: '5A5A5A' },
    h1:         { font: 'Calibri Light', size: 16, color: '2F5496' },
    h2:         { font: 'Calibri Light', size: 13, color: '2F5496' },
    h3:         { font: 'Calibri Light', size: 12, color: '1F3763' },
    h4:         { font: 'Calibri Light', size: 11, color: '2F5496', i: true },
    blockquote: { font: 'Calibri', size: 11, color: '404040', i: true },
    pre:        { font: 'Consolas', size: 10, color: '000000' }
  };

  // ---------- Farb-Hilfen ----------
  function cssToHex(c) {
    if (!c) return null;
    c = c.trim();
    if (c[0] === '#') {
      if (c.length === 4) c = '#' + c[1] + c[1] + c[2] + c[2] + c[3] + c[3];
      return c.slice(1, 7).toUpperCase();
    }
    const m = /rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)(?:[,\s/]+([\d.]+))?/.exec(c);
    if (!m) return null;
    if (m[4] !== undefined && parseFloat(m[4]) === 0) return null;
    return [m[1], m[2], m[3]].map(n => (+n).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  const HIGHLIGHT_COLORS = {
    yellow: 'FFFF00', green: '00FF00', cyan: '00FFFF', magenta: 'FF00FF', blue: '0000FF', red: 'FF0000',
    darkBlue: '000080', darkCyan: '008080', darkGreen: '008000', darkMagenta: '800080', darkRed: '800000',
    darkYellow: '808000', darkGray: '808080', lightGray: 'C0C0C0', black: '000000', white: 'FFFFFF'
  };

  // =====================================================================
  //  EXPORT: Editor-HTML -> .docx
  // =====================================================================
  async function exportDocx(editor, settings, meta = {}) {
    const D = window.docx;
    const content = { tw: (settings.orientation === 'landscape' ? settings.h : settings.w) * MM_TO_TW
      - (settings.margins.left + settings.margins.right) * MM_TO_TW };

    // Bilder vorbereiten (WebP/SVG -> PNG)
    const imgData = new Map();
    for (const img of editor.querySelectorAll('img')) {
      let src = img.getAttribute('src') || '';
      if (!src.startsWith('data:')) {
        try {
          const blob = await (await fetch(src)).blob();
          src = await new Promise(r => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(blob); });
        } catch { continue; }
      }
      src = await Rex.util.normalizeImage(src);
      imgData.set(img, Rex.util.dataURLToBytes(src));
    }

    let olInstance = 0;

    const isBlockEl = (n) => n.nodeType === 1 && /^(P|DIV|H[1-6]|UL|OL|TABLE|BLOCKQUOTE|PRE|HR|LI|SECTION|ARTICLE|HEADER|FOOTER)$/.test(n.tagName);

    function runFormat(el, block) {
      const cs = getComputedStyle(el);
      const bs = getComputedStyle(block);
      const f = {};
      const w = parseInt(cs.fontWeight, 10) >= 600, bw = parseInt(bs.fontWeight, 10) >= 600;
      if (w !== bw) f.bold = w;
      const it = cs.fontStyle === 'italic', bit = bs.fontStyle === 'italic';
      if (it !== bit) f.italics = it;
      if (cs.fontSize !== bs.fontSize) f.size = Math.round(parseFloat(cs.fontSize) * 0.75 * 2);
      const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      const bfam = bs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (fam !== bfam) f.font = fam;
      if (cs.color !== bs.color) { const h = cssToHex(cs.color); if (h) f.color = h; }
      // Nicht vererbte Eigenschaften: Vorfahren bis zum Block pruefen
      for (let n = el; n && n !== block; n = n.parentElement) {
        const ns = n === el ? cs : getComputedStyle(n);
        const dec = ns.textDecorationLine || '';
        if (dec.includes('underline') && f.underline === undefined) f.underline = {};
        if (dec.includes('line-through')) f.strike = true;
        if (!f.shading) {
          const bg = cssToHex(ns.backgroundColor);
          if (bg) f.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: bg };
        }
        if (ns.verticalAlign === 'super') f.superScript = true;
        if (ns.verticalAlign === 'sub') f.subScript = true;
      }
      return f;
    }

    function inlineRuns(node, block, out, preformatted) {
      for (const ch of node.childNodes) {
        if (ch.nodeType === 3) {
          let text = ch.nodeValue;
          if (!preformatted) text = text.replace(/[\t\n\r]+/g, ' ');
          if (!text) continue;
          const f = runFormat(ch.parentElement, block);
          if (preformatted && text.includes('\n')) {
            text.split('\n').forEach((part, i) => {
              if (i > 0) out.push(new D.TextRun({ break: 1 }));
              if (part) out.push(new D.TextRun({ text: part, ...f }));
            });
          } else {
            out.push(new D.TextRun({ text, ...f }));
          }
        } else if (ch.nodeType === 1) {
          const tag = ch.tagName;
          if (tag === 'BR') {
            const isLast = !ch.nextSibling || (ch.nextSibling.nodeType === 3 && !ch.nextSibling.nodeValue.trim() && !ch.nextSibling.nextSibling);
            if (!isLast) out.push(new D.TextRun({ break: 1 }));
          } else if (tag === 'IMG') {
            const data = imgData.get(ch);
            if (!data) continue;
            const w = parseFloat(ch.style.width) || ch.getAttribute('width') || ch.naturalWidth || 200;
            const ratio = ch.naturalWidth ? ch.naturalHeight / ch.naturalWidth : 0.66;
            const h = parseFloat(ch.style.height) || Math.round(w * ratio);
            out.push(new D.ImageRun({ data, transformation: { width: Math.round(w), height: Math.round(h) } }));
          } else if (tag === 'A' && ch.getAttribute('href')) {
            const inner = [];
            inlineRuns(ch, block, inner, preformatted);
            out.push(new D.ExternalHyperlink({ link: ch.getAttribute('href'), children: inner.filter(r => r instanceof D.TextRun) }));
          } else if (isBlockEl(ch)) {
            // verschachtelter Block in Inline-Kontext -> als Zeilenumbruch behandeln
            if (out.length) out.push(new D.TextRun({ break: 1 }));
            inlineRuns(ch, block, out, preformatted);
          } else {
            inlineRuns(ch, block, out, preformatted || ch.tagName === 'CODE');
          }
        }
      }
      return out;
    }

    function paraOptions(el) {
      const cs = getComputedStyle(el);
      const o = {};
      const tag = el.tagName;
      if (el.classList.contains('title')) o.heading = D.HeadingLevel.TITLE;
      else if (el.classList.contains('subtitle')) o.style = 'Subtitle';
      else if (/^H[1-6]$/.test(tag)) o.heading = D.HeadingLevel['HEADING_' + Math.min(+tag[1], 6)];
      else if (tag === 'BLOCKQUOTE') o.style = 'Quote';
      else if (tag === 'PRE' || el.classList.contains('code')) o.style = 'Code';
      const ta = cs.textAlign;
      if (ta === 'center') o.alignment = D.AlignmentType.CENTER;
      else if (ta === 'right' || ta === 'end') o.alignment = D.AlignmentType.RIGHT;
      else if (ta === 'justify') o.alignment = D.AlignmentType.JUSTIFIED;
      const ind = {};
      if (el.style.marginLeft) ind.left = Math.round(parseFloat(cs.marginLeft) * PX_TO_TW);
      if (el.style.marginRight) ind.right = Math.round(parseFloat(cs.marginRight) * PX_TO_TW);
      if (el.style.textIndent) {
        const ti = parseFloat(cs.textIndent) * PX_TO_TW;
        if (ti >= 0) ind.firstLine = Math.round(ti); else ind.hanging = Math.round(-ti);
      }
      if (Object.keys(ind).length) o.indent = ind;
      const sp = {};
      if (el.style.marginTop) sp.before = Math.round(parseFloat(cs.marginTop) * PX_TO_TW);
      if (el.style.marginBottom) sp.after = Math.round(parseFloat(cs.marginBottom) * PX_TO_TW);
      if (el.style.lineHeight) {
        const lh = parseFloat(el.style.lineHeight);
        if (!isNaN(lh) && !/px|pt/.test(el.style.lineHeight)) { sp.line = Math.round(lh * 240); sp.lineRule = 'auto'; }
      }
      if (Object.keys(sp).length) o.spacing = sp;
      if (el.style.backgroundColor) {
        const bg = cssToHex(el.style.backgroundColor);
        if (bg) o.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: bg };
      }
      return o;
    }

    function blocksFrom(container, level = 0) {
      const out = [];
      let pending = null; // anonyme Inline-Inhalte
      const flush = () => {
        if (pending && pending.length) out.push(new D.Paragraph({ children: pending }));
        pending = null;
      };
      for (const ch of container.childNodes) {
        if (ch.nodeType === 3 || (ch.nodeType === 1 && !isBlockEl(ch))) {
          if (ch.nodeType === 3 && !ch.nodeValue.trim() && !pending) continue;
          if (ch.nodeType === 1 && ch.tagName === 'BR' && !pending) { out.push(new D.Paragraph({})); continue; }
          pending = pending || [];
          if (ch.nodeType === 3) {
            const t = ch.nodeValue.replace(/[\t\n\r]+/g, ' ');
            if (t) pending.push(new D.TextRun({ text: t, ...runFormat(container, container) }));
          } else {
            // Inline-Element direkt verarbeiten (Container dient als Block)
            inlineRuns({ childNodes: [ch] }, container, pending, false);
          }
          continue;
        }
        flush();
        out.push(...blockToDocx(ch, level));
      }
      flush();
      return out;
    }

    function hasBlockChildren(el) {
      for (const c of el.children) if (isBlockEl(c)) return true;
      return false;
    }

    function listToDocx(list, level) {
      const out = [];
      const ordered = list.tagName === 'OL';
      const inst = ordered ? ++olInstance : 0;
      const ref = ordered ? 'rex-numbers' : 'rex-bullets';
      for (const li of list.children) {
        if (li.tagName === 'UL' || li.tagName === 'OL') { out.push(...listToDocx(li, level + 1)); continue; }
        if (li.tagName !== 'LI') { out.push(...blockToDocx(li, level)); continue; }
        if (li.style.listStyleType === 'none' && li.children.length === 1 && /^(UL|OL)$/.test(li.firstElementChild.tagName)) {
          out.push(...listToDocx(li.firstElementChild, level + 1));
          continue;
        }
        let runs = [];
        let first = true;
        const emit = (opts = {}) => {
          if (!runs.length && !first) return;
          out.push(new D.Paragraph({
            ...opts,
            numbering: first ? { reference: ref, level: Math.min(level, 8), instance: inst } : undefined,
            indent: first ? undefined : { left: 720 * (level + 1) },
            children: runs
          }));
          runs = [];
          first = false;
        };
        for (const c of li.childNodes) {
          if (c.nodeType === 1 && (c.tagName === 'UL' || c.tagName === 'OL')) {
            emit();
            out.push(...listToDocx(c, level + 1));
          } else if (c.nodeType === 1 && isBlockEl(c)) {
            if (runs.length) emit();
            inlineRuns(c, c, runs, false);
            emit(paraOptions(c));
          } else {
            inlineRuns({ childNodes: [c] }, li, runs, false);
          }
        }
        if (runs.length || first) emit(paraOptions(li));
      }
      return out;
    }

    function tableToDocx(table) {
      const rows = Array.from(table.rows);
      if (!rows.length) return [];
      const gridCount = Math.max(...rows.map(r => Array.from(r.cells).reduce((s, c) => s + (c.colSpan || 1), 0)));
      // Spaltenbreiten aus der Darstellung ableiten
      let widths = null;
      const plainRow = rows.find(r => r.cells.length === gridCount);
      const tw = table.getBoundingClientRect().width || 1;
      if (plainRow) widths = Array.from(plainRow.cells).map(c => c.getBoundingClientRect().width / tw);
      else widths = new Array(gridCount).fill(1 / gridCount);
      const totalTw = Math.round(content.tw * Math.min(1, (parseFloat(table.style.width) || 100) / 100));
      const colTw = widths.map(w => Math.max(200, Math.round(w * totalTw)));

      const docRows = rows.map(tr => new D.TableRow({
        tableHeader: tr.parentElement.tagName === 'THEAD',
        children: Array.from(tr.cells).map(td => {
          let children = hasBlockChildren(td) ? blocksFrom(td) : (() => {
            const runs = inlineRuns(td, td, [], false);
            return [new D.Paragraph({ children: runs, ...paraOptions(td), heading: undefined })];
          })();
          if (!children.length) children = [new D.Paragraph({})];
          const opts = { children };
          if (td.colSpan > 1) opts.columnSpan = td.colSpan;
          if (td.rowSpan > 1) opts.rowSpan = td.rowSpan;
          const bg = cssToHex(td.style.backgroundColor);
          if (bg) opts.shading = { type: D.ShadingType.CLEAR, color: 'auto', fill: bg };
          const va = td.style.verticalAlign;
          if (va === 'middle') opts.verticalAlign = D.VerticalAlign.CENTER;
          else if (va === 'bottom') opts.verticalAlign = D.VerticalAlign.BOTTOM;
          return new D.TableCell(opts);
        })
      }));
      return [new D.Table({ rows: docRows, columnWidths: colTw, width: { size: totalTw, type: D.WidthType.DXA } }), new D.Paragraph({ spacing: { after: 0 }, children: [] })];
    }

    function blockToDocx(el, level) {
      const tag = el.tagName;
      if (el.classList.contains('page-break')) return [new D.Paragraph({ children: [new D.PageBreak()] })];
      if (tag === 'HR') return [new D.Paragraph({ border: { bottom: { color: 'A5A5A5', space: 1, style: D.BorderStyle.SINGLE, size: 6 } }, children: [] })];
      if (tag === 'UL' || tag === 'OL') return listToDocx(el, level);
      if (tag === 'TABLE') return tableToDocx(el);
      if (hasBlockChildren(el) && !/^H[1-6]$/.test(tag)) return blocksFrom(el, level);
      const pre = tag === 'PRE';
      const runs = inlineRuns(el, el, [], pre);
      return [new D.Paragraph({ ...paraOptions(el), children: runs })];
    }

    const children = blocksFrom(editor);

    const numberingLevels = (ordered) => Array.from({ length: 9 }, (_, i) => ({
      level: i,
      format: ordered ? [D.LevelFormat.DECIMAL, D.LevelFormat.LOWER_LETTER, D.LevelFormat.LOWER_ROMAN][i % 3] : D.LevelFormat.BULLET,
      text: ordered ? `%${i + 1}.` : ['•', '◦', '▪'][i % 3],
      alignment: D.AlignmentType.LEFT,
      style: { paragraph: { indent: { left: 720 * (i + 1), hanging: 360 } } }
    }));

    const footers = settings.pageNumbers ? {
      default: new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ children: [D.PageNumber.CURRENT] })] })] })
    } : undefined;

    const document_ = new D.Document({
      creator: 'REX Office',
      title: meta.title || '',
      description: 'Erstellt mit REX Office',
      styles: {
        default: {
          document: { run: { font: 'Calibri', size: 22 }, paragraph: { spacing: { after: 160, line: 276, lineRule: 'auto' } } },
          title: { run: { font: 'Calibri Light', size: 56, color: '000000' }, paragraph: { spacing: { after: 80, line: 240, lineRule: 'auto' } } },
          heading1: { run: { font: 'Calibri Light', size: 32, color: '2F5496' }, paragraph: { spacing: { before: 240, after: 80 }, keepNext: true } },
          heading2: { run: { font: 'Calibri Light', size: 26, color: '2F5496' }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true } },
          heading3: { run: { font: 'Calibri Light', size: 24, color: '1F3763' }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true } },
          heading4: { run: { font: 'Calibri Light', size: 22, color: '2F5496', italics: true }, paragraph: { spacing: { before: 200, after: 80 }, keepNext: true } },
          hyperlink: { run: { color: '0563C1', underline: { type: D.UnderlineType.SINGLE } } }
        },
        paragraphStyles: [
          { id: 'Subtitle', name: 'Subtitle', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { color: '5A5A5A', characterSpacing: 15 } },
          { id: 'Quote', name: 'Quote', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { italics: true, color: '404040' }, paragraph: { alignment: D.AlignmentType.CENTER, indent: { left: 864, right: 864 }, spacing: { before: 200, after: 160 } } },
          { id: 'Code', name: 'Code', basedOn: 'Normal', next: 'Normal', quickFormat: true,
            run: { font: 'Consolas', size: 20 }, paragraph: { shading: { type: D.ShadingType.CLEAR, color: 'auto', fill: 'F4F4F4' }, spacing: { after: 0, line: 240, lineRule: 'auto' } } }
        ]
      },
      numbering: {
        config: [
          { reference: 'rex-bullets', levels: numberingLevels(false) },
          { reference: 'rex-numbers', levels: numberingLevels(true) }
        ]
      },
      sections: [{
        properties: {
          page: {
            size: {
              width: Math.round(settings.w * MM_TO_TW),
              height: Math.round(settings.h * MM_TO_TW),
              orientation: settings.orientation === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT
            },
            margin: {
              top: Math.round(settings.margins.top * MM_TO_TW),
              right: Math.round(settings.margins.right * MM_TO_TW),
              bottom: Math.round(settings.margins.bottom * MM_TO_TW),
              left: Math.round(settings.margins.left * MM_TO_TW)
            }
          }
        },
        footers,
        children: children.length ? children : [new D.Paragraph({})]
      }]
    });
    const blob = await D.Packer.toBlob(document_);
    return new Uint8Array(await blob.arrayBuffer());
  }

  // =====================================================================
  //  IMPORT: .docx -> Editor-HTML
  // =====================================================================
  const kids = (el, name) => el ? Array.from(el.childNodes).filter(n => n.nodeType === 1 && n.localName === name) : [];
  const kid = (el, name) => kids(el, name)[0] || null;
  const wattr = (el, name) => el ? (el.getAttributeNS(W_NS, name) ?? el.getAttribute('w:' + name)) : null;
  const esc = (s) => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  function onOff(el) {
    if (!el) return undefined;
    const v = wattr(el, 'val');
    return !(v === '0' || v === 'false' || v === 'none');
  }

  async function importDocx(bytes) {
    const zip = await JSZip.loadAsync(bytes);
    const read = async (p) => { const f = zip.file(p); return f ? f.async('string') : null; };
    const parse = (s) => s ? new DOMParser().parseFromString(s, 'application/xml') : null;

    const docXml = parse(await read('word/document.xml'));
    if (!docXml) throw new Error('Keine gültige Word-Datei (word/document.xml fehlt).');
    const stylesXml = parse(await read('word/styles.xml'));
    const numXml = parse(await read('word/numbering.xml'));
    const relsXml = parse(await read('word/_rels/document.xml.rels'));
    const themeXml = parse(await read('word/theme/theme1.xml'));

    // Beziehungen
    const rels = {};
    if (relsXml) for (const r of relsXml.getElementsByTagName('Relationship')) {
      rels[r.getAttribute('Id')] = { target: r.getAttribute('Target'), mode: r.getAttribute('TargetMode'), type: r.getAttribute('Type') };
    }

    // Theme (Schriften + Farben)
    const theme = { major: 'Calibri Light', minor: 'Calibri', colors: {} };
    if (themeXml) {
      const maj = themeXml.getElementsByTagNameNS('*', 'majorFont')[0];
      const min = themeXml.getElementsByTagNameNS('*', 'minorFont')[0];
      if (maj) theme.major = maj.getElementsByTagNameNS('*', 'latin')[0]?.getAttribute('typeface') || theme.major;
      if (min) theme.minor = min.getElementsByTagNameNS('*', 'latin')[0]?.getAttribute('typeface') || theme.minor;
      const scheme = themeXml.getElementsByTagNameNS('*', 'clrScheme')[0];
      if (scheme) for (const c of scheme.children) {
        const v = c.firstElementChild;
        if (v) theme.colors[c.localName] = (v.getAttribute('val') || v.getAttribute('lastClr') || '').toUpperCase();
      }
    }

    // Formatvorlagen
    const styles = {};
    let defRPr = null, defPPr = null, defaultParaStyle = null;
    if (stylesXml) {
      const dd = stylesXml.getElementsByTagNameNS(W_NS, 'docDefaults')[0];
      if (dd) {
        defRPr = dd.getElementsByTagNameNS(W_NS, 'rPr')[0] || null;
        defPPr = dd.getElementsByTagNameNS(W_NS, 'pPr')[0] || null;
      }
      for (const s of stylesXml.getElementsByTagNameNS(W_NS, 'style')) {
        const id = wattr(s, 'styleId');
        const st = {
          id, type: wattr(s, 'type'),
          name: (wattr(kid(s, 'name'), 'val') || id || '').toLowerCase(),
          basedOn: wattr(kid(s, 'basedOn'), 'val'),
          pPr: kid(s, 'pPr'), rPr: kid(s, 'rPr')
        };
        styles[id] = st;
        if (st.type === 'paragraph' && (wattr(s, 'default') === '1' || wattr(s, 'default') === 'true')) defaultParaStyle = id;
      }
    }
    const styleChain = (id) => {
      const chain = [];
      let guard = 0;
      while (id && styles[id] && guard++ < 20) { chain.unshift(styles[id]); id = styles[id].basedOn; }
      return chain;
    };

    // Nummerierung
    const nums = {};
    if (numXml) {
      const abstract = {};
      for (const a of numXml.getElementsByTagNameNS(W_NS, 'abstractNum')) {
        const lv = {};
        for (const l of kids(a, 'lvl')) lv[wattr(l, 'ilvl')] = wattr(kid(l, 'numFmt'), 'val') || 'decimal';
        abstract[wattr(a, 'abstractNumId')] = lv;
      }
      for (const n of kids(numXml.documentElement, 'num')) {
        nums[wattr(n, 'numId')] = abstract[wattr(kid(n, 'abstractNumId'), 'val')] || {};
      }
    }

    // Run-Eigenschaften zusammenfuehren
    function applyRPr(target, rPr) {
      if (!rPr) return;
      for (const c of rPr.children) {
        switch (c.localName) {
          case 'b': target.b = onOff(c); break;
          case 'i': target.i = onOff(c); break;
          case 'u': target.u = wattr(c, 'val') && wattr(c, 'val') !== 'none'; break;
          case 'strike': case 'dstrike': target.strike = onOff(c); break;
          case 'caps': target.caps = onOff(c); break;
          case 'vanish': target.hidden = onOff(c); break;
          case 'sz': target.size = parseInt(wattr(c, 'val'), 10) / 2; break;
          case 'color': {
            const v = wattr(c, 'val');
            const tc = wattr(c, 'themeColor');
            if (v && v !== 'auto') target.color = v.toUpperCase();
            else if (tc && theme.colors[tc.replace(/^text1$/, 'dk1').replace(/^text2$/, 'dk2').replace(/^background1$/, 'lt1').replace(/^background2$/, 'lt2')]) {
              target.color = theme.colors[tc.replace(/^text1$/, 'dk1').replace(/^text2$/, 'dk2').replace(/^background1$/, 'lt1').replace(/^background2$/, 'lt2')];
            } else if (v === 'auto') target.color = '000000';
            break;
          }
          case 'rFonts': {
            const a = wattr(c, 'ascii') || wattr(c, 'hAnsi');
            const at = wattr(c, 'asciiTheme') || wattr(c, 'hAnsiTheme');
            if (a) target.font = a;
            else if (at) target.font = /major/i.test(at) ? theme.major : theme.minor;
            break;
          }
          case 'highlight': { const h = HIGHLIGHT_COLORS[wattr(c, 'val')]; if (h) target.bg = h; break; }
          case 'shd': { const f = wattr(c, 'fill'); if (f && f !== 'auto') target.bg = f.toUpperCase(); break; }
          case 'vertAlign': target.va = wattr(c, 'val'); break;
          case 'rStyle': break;
        }
      }
    }
    function baseRun(paraStyleId) {
      const r = { font: theme.minor, size: 10, color: '000000' };
      applyRPr(r, defRPr);
      for (const s of styleChain(paraStyleId || defaultParaStyle)) applyRPr(r, s.rPr);
      return r;
    }

    function blockTagFor(styleId) {
      const chain = styleChain(styleId);
      for (let i = chain.length - 1; i >= 0; i--) {
        const n = chain[i].name;
        if (n === 'title') return 'title';
        if (n === 'subtitle') return 'subtitle';
        const m = /^heading (\d)$/.exec(n);
        if (m) return 'h' + Math.min(4, +m[1]);
        if (n === 'quote' || n === 'intense quote') return 'blockquote';
        if (n === 'code' || n === 'html preformatted') return 'pre';
      }
      return 'p';
    }

    function paraProps(pPr) {
      const out = { jc: null, numId: null, ilvl: 0, style: null, ind: null, spacing: null, shd: null, pageBreakBefore: false };
      const chainApply = (pp) => {
        if (!pp) return;
        const jc = kid(pp, 'jc'); if (jc) out.jc = wattr(jc, 'val');
        const np = kid(pp, 'numPr');
        if (np) {
          const id = wattr(kid(np, 'numId'), 'val');
          if (id != null) out.numId = id;
          const lv = wattr(kid(np, 'ilvl'), 'val');
          if (lv != null) out.ilvl = +lv;
        }
        const shd = kid(pp, 'shd'); if (shd && wattr(shd, 'fill') && wattr(shd, 'fill') !== 'auto') out.shd = wattr(shd, 'fill');
        if (kid(pp, 'pageBreakBefore') && onOff(kid(pp, 'pageBreakBefore'))) out.pageBreakBefore = true;
      };
      out.style = wattr(kid(pPr, 'pStyle'), 'val') || defaultParaStyle;
      chainApply(defPPr);
      for (const s of styleChain(out.style)) chainApply(s.pPr);
      chainApply(pPr);
      // Einzug/Abstand nur aus direkter Formatierung uebernehmen
      const ind = kid(pPr, 'ind');
      if (ind) out.ind = {
        left: +(wattr(ind, 'left') || wattr(ind, 'start') || 0), right: +(wattr(ind, 'right') || wattr(ind, 'end') || 0),
        firstLine: +(wattr(ind, 'firstLine') || 0), hanging: +(wattr(ind, 'hanging') || 0)
      };
      const sp = kid(pPr, 'spacing');
      if (sp) out.spacing = { before: wattr(sp, 'before'), after: wattr(sp, 'after'), line: wattr(sp, 'line'), rule: wattr(sp, 'lineRule') };
      if (out.numId === '0') out.numId = null;
      return out;
    }

    async function imageFromRid(rid) {
      const rel = rels[rid];
      if (!rel || rel.mode === 'External') return null;
      let p = rel.target.replace(/^\//, '');
      if (!p.startsWith('word/')) p = 'word/' + p;
      p = p.replace(/\/[^/]+\/\.\.\//g, '/');
      const f = zip.file(p);
      if (!f) return null;
      const ext = p.split('.').pop().toLowerCase();
      const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', emf: null, wmf: null }[ext];
      if (!mime) return null;
      const b64 = await f.async('base64');
      return `data:${mime};base64,${b64}`;
    }

    async function runHTML(r, base, defaults) {
      const rp = { ...base };
      const rPr = kid(r, 'rPr');
      const rs = wattr(kid(rPr, 'rStyle'), 'val');
      if (rs) for (const s of styleChain(rs)) applyRPr(rp, s.rPr);
      applyRPr(rp, rPr);
      if (rp.hidden) return '';
      let html = '';
      for (const c of r.children) {
        switch (c.localName) {
          case 't': html += esc(c.textContent); break;
          case 'tab': html += '&emsp;'; break;
          case 'br': html += wattr(c, 'type') === 'page' ? '\u0000PB\u0000' : '<br>'; break;
          case 'cr': html += '<br>'; break;
          case 'noBreakHyphen': html += '&#8209;'; break;
          case 'softHyphen': html += '&shy;'; break;
          case 'sym': html += `&#x${wattr(c, 'char') || '25A1'};`; break;
          case 'drawing': {
            const blip = c.getElementsByTagNameNS('*', 'blip')[0];
            const ext = c.getElementsByTagNameNS('*', 'extent')[0];
            const rid = blip && (blip.getAttributeNS(R_NS, 'embed') || blip.getAttribute('r:embed'));
            const src = rid && await imageFromRid(rid);
            if (src) {
              const w = ext ? Math.round(+ext.getAttribute('cx') / 9525) : null;
              html += `<img src="${src}"${w ? ` style="width:${w}px"` : ''}>`;
            }
            break;
          }
          case 'pict': case 'object': {
            const im = c.getElementsByTagNameNS('*', 'imagedata')[0];
            const rid = im && (im.getAttributeNS(R_NS, 'id') || im.getAttribute('r:id'));
            const src = rid && await imageFromRid(rid);
            if (src) html += `<img src="${src}">`;
            break;
          }
          case 'AlternateContent': {
            const fb = kid(c, 'Fallback') || kid(c, 'Choice');
            if (fb) html += await runHTML({ children: fb.children }, rp, defaults);
            break;
          }
        }
      }
      if (!html) return '';
      const css = [];
      if (rp.font && rp.font !== defaults.font) css.push(`font-family:'${rp.font.replace(/'/g, '')}'`);
      if (rp.size && Math.abs(rp.size - defaults.size) > 0.01) css.push(`font-size:${rp.size}pt`);
      if (rp.color && rp.color !== defaults.color) css.push(`color:#${rp.color}`);
      if (rp.bg) css.push(`background-color:#${rp.bg}`);
      if (rp.caps) css.push('text-transform:uppercase');
      if (!!rp.i !== !!defaults.i) css.push(`font-style:${rp.i ? 'italic' : 'normal'}`);
      if (rp.b && !defaults.b) html = `<b>${html}</b>`;
      else if (!rp.b && defaults.b) css.push('font-weight:normal');
      if (rp.u) html = `<u>${html}</u>`;
      if (rp.strike) html = `<s>${html}</s>`;
      if (rp.va === 'superscript') html = `<sup>${html}</sup>`;
      if (rp.va === 'subscript') html = `<sub>${html}</sub>`;
      return css.length ? `<span style="${css.join(';')}">${html}</span>` : html;
    }

    async function inlineHTML(p, base, defaults) {
      let html = '';
      for (const c of p.children) {
        switch (c.localName) {
          case 'r': html += await runHTML(c, base, defaults); break;
          case 'hyperlink': {
            const rid = c.getAttributeNS(R_NS, 'id') || c.getAttribute('r:id');
            const href = rid && rels[rid] ? rels[rid].target : (wattr(c, 'anchor') ? '#' + wattr(c, 'anchor') : null);
            const inner = await inlineHTML(c, base, defaults);
            html += href ? `<a href="${esc(href)}">${inner}</a>` : inner;
            break;
          }
          case 'ins': case 'smartTag': case 'customXml': case 'fldSimple': case 'moveTo':
            html += await inlineHTML(c, base, defaults); break;
          case 'sdt': html += await inlineHTML(kid(c, 'sdtContent') || c, base, defaults); break;
          case 'AlternateContent': {
            const fb = kid(c, 'Choice') || kid(c, 'Fallback');
            if (fb) html += await inlineHTML(fb, base, defaults);
            break;
          }
        }
      }
      return html;
    }

    async function paragraphHTML(p) {
      const pPr = kid(p, 'pPr');
      const pp = paraProps(pPr);
      const kind = blockTagFor(pp.style);
      const defaults = TAG_DEFAULTS[kind];
      const base = baseRun(pp.style);
      // Absatzmarke-Formatierung ignorieren; Inhalte erzeugen
      let inner = await inlineHTML(p, base, defaults);
      const css = [];
      const align = { center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify' }[pp.jc];
      if (align && !(kind === 'blockquote' && align === 'center')) css.push(`text-align:${align}`);
      if (pp.ind && !pp.numId) {
        if (pp.ind.left) css.push(`margin-left:${Math.round(pp.ind.left / PX_TO_TW)}px`);
        if (pp.ind.right) css.push(`margin-right:${Math.round(pp.ind.right / PX_TO_TW)}px`);
        if (pp.ind.firstLine) css.push(`text-indent:${Math.round(pp.ind.firstLine / PX_TO_TW)}px`);
        if (pp.ind.hanging) css.push(`text-indent:-${Math.round(pp.ind.hanging / PX_TO_TW)}px`);
      }
      if (pp.spacing) {
        if (pp.spacing.before != null && kind === 'p') css.push(`margin-top:${+pp.spacing.before / 20}pt`);
        if (pp.spacing.after != null && kind === 'p') css.push(`margin-bottom:${+pp.spacing.after / 20}pt`);
        if (pp.spacing.line && (!pp.spacing.rule || pp.spacing.rule === 'auto')) {
          const lh = Math.round((+pp.spacing.line / 240) * 100) / 100;
          if (Math.abs(lh - 1.15) > 0.08) css.push(`line-height:${lh}`);
        }
      }
      if (pp.shd) css.push(`background-color:#${pp.shd}`);

      // Seitenumbrueche innerhalb des Absatzes aufteilen
      const parts = inner.split('\u0000PB\u0000');
      const tagOpen = (() => {
        const st = css.length ? ` style="${css.join(';')}"` : '';
        if (kind === 'title') return [`<p class="title"${st}>`, '</p>'];
        if (kind === 'subtitle') return [`<p class="subtitle"${st}>`, '</p>'];
        return [`<${kind}${st}>`, `</${kind}>`];
      })();
      let html = pp.pageBreakBefore ? '<div class="page-break" contenteditable="false"></div>' : '';
      parts.forEach((part, i) => {
        if (i > 0) html += '<div class="page-break" contenteditable="false"></div>';
        if (i > 0 && !part.replace(/<[^>]+>/g, '').trim() && !/<img/.test(part)) return;
        html += tagOpen[0] + (part || '<br>') + tagOpen[1];
      });
      return { html, numId: pp.numId, ilvl: pp.ilvl, inner: parts.join('') };
    }

    async function tableHTML(tbl) {
      const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => +wattr(g, 'w') || 0);
      const total = grid.reduce((a, b) => a + b, 0);
      let html = '<table>';
      if (total) html += '<colgroup>' + grid.map(w => `<col style="width:${(w / total * 100).toFixed(2)}%">`).join('') + '</colgroup>';
      const rows = kids(tbl, 'tr');
      const matrix = []; // fuer vertikale Verbindungen
      const cellsOut = [];
      for (let ri = 0; ri < rows.length; ri++) {
        const tr = rows[ri];
        let col = 0;
        const rowCells = [];
        for (const tc of kids(tr, 'tc')) {
          const tcPr = kid(tc, 'tcPr');
          const span = +(wattr(kid(tcPr, 'gridSpan'), 'val') || 1);
          const vm = kid(tcPr, 'vMerge');
          const vmVal = vm ? (wattr(vm, 'val') || 'continue') : null;
          if (vmVal === 'continue') {
            for (let r = ri - 1; r >= 0; r--) {
              const above = matrix[r] && matrix[r][col];
              if (above) { above.rowspan++; break; }
            }
            col += span;
            continue;
          }
          let body = '';
          for (const c of tc.children) {
            if (c.localName === 'p') body += (await paragraphHTML(c)).html;
            else if (c.localName === 'tbl') body += await tableHTML(c);
          }
          const shd = kid(tcPr, 'shd');
          const fill = shd && wattr(shd, 'fill');
          const va = wattr(kid(tcPr, 'vAlign'), 'val');
          const cell = { span, rowspan: 1, body, fill: fill && fill !== 'auto' ? fill : null, va };
          matrix[ri] = matrix[ri] || [];
          matrix[ri][col] = cell;
          rowCells.push(cell);
          col += span;
        }
        cellsOut.push(rowCells);
      }
      for (const row of cellsOut) {
        html += '<tr>' + row.map(c => {
          const st = [];
          if (c.fill) st.push(`background-color:#${c.fill}`);
          if (c.va === 'center') st.push('vertical-align:middle');
          if (c.va === 'bottom') st.push('vertical-align:bottom');
          return `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}${st.length ? ` style="${st.join(';')}"` : ''}>${c.body || '<p><br></p>'}</td>`;
        }).join('') + '</tr>';
      }
      return html + '</table>';
    }

    // Body durchlaufen, Listen gruppieren
    const body = docXml.getElementsByTagNameNS(W_NS, 'body')[0];
    let out = '';
    let listStack = []; // [{numId, tag}]
    const closeLists = (depth = 0) => {
      while (listStack.length > depth) { const l = listStack.pop(); out += `</li></${l.tag}>`; }
    };

    async function walk(container) {
      for (const el of container.children) {
        if (el.localName === 'p') {
          const r = await paragraphHTML(el);
          if (r.numId && nums[r.numId]) {
            const fmt = nums[r.numId][String(r.ilvl)] || 'decimal';
            const tag = fmt === 'bullet' || fmt === 'none' ? 'ul' : 'ol';
            const lvl = Math.min(r.ilvl, 8);
            if (listStack.length > lvl + 1) closeLists(lvl + 1);
            if (listStack.length === lvl + 1) {
              const top = listStack[lvl];
              if (top.tag !== tag || (lvl === 0 && top.numId !== r.numId)) { closeLists(lvl); }
              else out += '</li>';
            }
            while (listStack.length < lvl + 1) {
              const t = listStack.length === lvl ? tag : 'ul';
              out += `<${t}>`;
              listStack.push({ tag: t, numId: r.numId });
              if (listStack.length < lvl + 1) out += '<li>';
            }
            out += `<li>${r.inner || '<br>'}`;
          } else {
            closeLists();
            out += r.html;
          }
        } else if (el.localName === 'tbl') {
          closeLists();
          out += await tableHTML(el);
        } else if (el.localName === 'sdt') {
          await walk(kid(el, 'sdtContent') || el);
        } else if (el.localName === 'customXml' || el.localName === 'ins') {
          await walk(el);
        }
      }
    }
    await walk(body);
    closeLists();

    // Seiteneinstellungen
    const settings = { size: 'A4', w: 210, h: 297, orientation: 'portrait', margins: { top: 25, right: 25, bottom: 25, left: 25 }, pageNumbers: false };
    const sect = kids(body, 'sectPr').pop() || body.getElementsByTagNameNS(W_NS, 'sectPr')[0];
    if (sect) {
      const sz = kid(sect, 'pgSz');
      if (sz) {
        let w = +wattr(sz, 'w') / MM_TO_TW, h = +wattr(sz, 'h') / MM_TO_TW;
        const landscape = wattr(sz, 'orient') === 'landscape' || w > h;
        if (w > h) [w, h] = [h, w];
        settings.w = Math.round(w * 10) / 10;
        settings.h = Math.round(h * 10) / 10;
        settings.orientation = landscape ? 'landscape' : 'portrait';
        settings.size = Math.abs(settings.w - 210) < 2 && Math.abs(settings.h - 297) < 2 ? 'A4'
          : Math.abs(settings.w - 148) < 2 && Math.abs(settings.h - 210) < 2 ? 'A5'
          : Math.abs(settings.w - 215.9) < 2 && Math.abs(settings.h - 279.4) < 2 ? 'Letter'
          : Math.abs(settings.w - 215.9) < 2 && Math.abs(settings.h - 355.6) < 2 ? 'Legal' : 'custom';
      }
      const mg = kid(sect, 'pgMar');
      if (mg) for (const k of ['top', 'right', 'bottom', 'left']) {
        const v = wattr(mg, k);
        if (v != null) settings.margins[k] = Math.round(Math.abs(+v) / MM_TO_TW * 10) / 10;
      }
    }
    for (const name of Object.keys(zip.files)) {
      if (/^word\/footer\d*\.xml$/.test(name)) {
        const s = await zip.file(name).async('string');
        if (/PAGE/.test(s)) settings.pageNumbers = true;
      }
    }

    return { html: out || '<p><br></p>', settings };
  }

  async function importDocxSafe(bytes) {
    try {
      return await importDocx(bytes);
    } catch (err) {
      console.warn('Eigener Import fehlgeschlagen, nutze Mammoth:', err);
      const res = await mammoth.convertToHtml({ arrayBuffer: bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) }, {
        styleMap: ["p[style-name='Title'] => p.title:fresh", "p[style-name='Subtitle'] => p.subtitle:fresh",
          "p[style-name='Quote'] => blockquote:fresh", 'u => u', 'strike => s']
      });
      return { html: res.value, settings: null };
    }
  }

  window.RexDocx = { exportDocx, importDocx: importDocxSafe, cssToHex };
})();
