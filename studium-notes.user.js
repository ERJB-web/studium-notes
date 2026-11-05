// ==UserScript==
// @name         Notes StudiUM
// @namespace    studium-notes
// @version      1.4
// @description  Remplit les notes StudiUM (Moodle) à partir d'une liste ou d'un fichier Excel, et affiche les statistiques de la classe.
// @match        https://studium.umontreal.ca/*
// @match        https://*.studium.umontreal.ca/*
// @grant        none
// @run-at       document-idle
// ==/UserScript==

(function () {
  'use strict';

  const IN_BROWSER = typeof document !== 'undefined';
  const BUILD = '1.4-20261105';
  let legacyText = '', legacyColumn = '';
  function removeLegacy() {
    const old = document.getElementById('sn-panel');
    if (old && old.querySelector('#sn-text') && /Saisie rapide des notes/.test(old.textContent)) {
      legacyText = old.querySelector('#sn-text').value || legacyText;
      legacyColumn = old.querySelector('#sn-col')?.value || legacyColumn;
      old.remove();
      const button = document.getElementById('sn-btn');
      if (button) button.remove();
    }
  }
  if (IN_BROWSER) {
    try { localStorage.removeItem('studium-notes-texte'); } catch (_) {}
    removeLegacy();
    const existing = document.getElementById('studium-notes-host');
    if (existing && window.__snBuild === BUILD && typeof window.__snShow === 'function') {
      const currentText = existing.shadowRoot?.getElementById('text');
      if (legacyText && currentText && currentText.value !== legacyText) { currentText.value = [currentText.value, legacyText].filter(Boolean).join('\n'); }
      window.__snShow();
      return;
    }
    if (existing) {
      legacyText = existing.shadowRoot?.getElementById('text')?.value || legacyText;
      legacyColumn = existing.shadowRoot?.getElementById('col')?.value || legacyColumn;
      if (typeof window.__snDestroy === 'function') window.__snDestroy();
      else existing.remove();
    }
  }

  const INPUT_SELECTOR = [
    'input[name^="grade["]',
    'select[name^="grade["]',
    'input[name^="quickgrade_"]',
    'select[name^="quickgrade_"]',
    'input[name^="finalgrade_"]',
    'select[name^="finalgrade_"]',
  ].join(',');

  const UI_KEY = 'studium-notes-fenetre';
  const EMAIL_RE = /[\w.+-]+@[\w-]+(?:\.[\w-]+)+/;

  const norm = (s) =>
    (s || '')
      .normalize('NFD')
      .replace(/[̀-ͯ]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9@._]+/g, ' ')
      .trim();
  const tokens = (s) => norm(s).split(/\s+/).filter(Boolean);
  const escapeHtml = (s) =>
    String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  let LOCALE = 'fr-CA';
  let NA = 'n.d.';
  const fmt = (x, d = 1) =>
    Number.isFinite(x) ? x.toLocaleString(LOCALE, { maximumFractionDigits: d }) : NA;
  const fmtGrade = (x) => String(Math.round(x * 10000) / 10000).replace('.', ',');

  const GRADE_ONLY = /^-?\d+(?:[.,]\d+)?\s*(?:%|\/\s*\d+(?:[.,]\d+)?)?$/;
  const HAS_GRADE = /(?:^-?\d+(?:[.,]\d+)?[\s,\t:])|(?:[\s,\t:]-?\d+(?:[.,]\d+)?\s*(?:%|\/\s*\d+(?:[.,]\d+)?)?$)/;
  function splitEntries(text) {
    const out = [];
    text.split(/\r?\n|;/).map((s) => s.trim()).filter(Boolean).forEach((p) => {
      const last = out.length - 1;
      if (GRADE_ONLY.test(p) && last >= 0 && !HAS_GRADE.test(out[last])) out[last] += ' ' + p;
      else out.push(p);
    });
    return out;
  }

  function parseList(text) {
    const out = [];
    const bad = [];
    splitEntries(text).forEach((raw) => {
      const line = raw.trim();
      if (!line) return;
      let m = line.match(/^(.*?)[\s;,\t:]+(-?\d+(?:[.,]\d+)?)\s*(?:%|\/\s*\d+(?:[.,]\d+)?)?\s*$/);
      if (m && /[a-zA-ZÀ-ÿ0-9]/.test(m[1])) {
        out.push({ name: m[1].trim(), grade: m[2].replace(',', '.'), line });
        return;
      }
      m = line.match(/^(-?\d+(?:[.,]\d+)?)\s*(?:%|\/\s*\d+)?[\s;,\t:]+(.+)$/);
      if (m) {
        out.push({ name: m[2].trim(), grade: m[1].replace(',', '.'), line });
        return;
      }
      bad.push(line);
    });
    return { entries: out, bad };
  }

  function splitName(full) {
    const s = String(full || '').replace(/\s+/g, ' ').trim();
    if (s.includes(',')) {
      const [nom, ...rest] = s.split(',');
      return { prenom: rest.join(',').trim(), nom: nom.trim() };
    }
    const i = s.indexOf(' ');
    return i < 0 ? { prenom: '', nom: s } : { prenom: s.slice(0, i), nom: s.slice(i + 1) };
  }

  function toNumber(v) {
    if (typeof v === 'number') return v;
    if (v === null || v === undefined) return NaN;
    const s = String(v).replace(/\s+/g, '').replace(',', '.').replace(/%$/, '');
    return /^-?\d+(?:\.\d+)?$/.test(s) ? parseFloat(s) : NaN;
  }
  function quantile(sorted, p) {
    if (!sorted.length) return NaN;
    const h = (sorted.length - 1) * p, lo = Math.floor(h), hi = Math.ceil(h);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (h - lo);
  }
  function computeStats(values) {
    const v = values.filter(Number.isFinite).slice().sort((a, b) => a - b);
    const n = v.length;
    if (!n) return { n: 0, sorted: [] };
    const mean = v.reduce((s, x) => s + x, 0) / n;
    const sd = n > 1 ? Math.sqrt(v.reduce((s, x) => s + (x - mean) ** 2, 0) / (n - 1)) : 0;
    const q1 = quantile(v, 0.25), q3 = quantile(v, 0.75);
    const p10 = quantile(v, 0.1), p90 = quantile(v, 0.9);
    return {
      n, sorted: v, mean, sd, min: v[0], max: v[n - 1], median: quantile(v, 0.5),
      q1, q3, iqr: q3 - q1, p10, p90, idr: p90 - p10,
    };
  }
  const LETTER_BANDS = [['C', 60], ['C+', 65], ['B-', 70], ['B', 73], ['B+', 77], ['A-', 80], ['A', 85], ['A+', 90]];
  function countLetterGrades(values, outOf) {
    if (!Number.isFinite(outOf) || outOf <= 0) return null;
    const counts = Object.fromEntries(LETTER_BANDS.map(([letter]) => [letter, 0]));
    counts.fail = 0;
    values.forEach((value) => {
      if (!Number.isFinite(value)) return;
      const pct = value / outOf * 100;
      const band = LETTER_BANDS.findLast(([, minimum]) => pct + 1e-10 >= minimum);
      counts[band ? band[0] : 'fail']++;
    });
    return counts;
  }
  const SCALES = [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 100];
  function guessMax(max) {
    if (!(max > 0)) return 100;
    return SCALES.find((s) => s >= max) || Math.ceil(max / 10) * 10;
  }
  function niceStep(span, target = 10) {
    if (!(span > 0)) return 1;
    const raw = span / target;
    const pow = 10 ** Math.floor(Math.log10(raw));
    const f = raw / pow;
    return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * pow;
  }
  function histogram(values, step, lo, hi) {
    const nb = Math.max(1, Math.ceil((hi - lo) / step - 1e-9));
    const bins = [];
    for (let i = 0; i < nb; i++) bins.push({ from: lo + i * step, to: Math.min(hi, lo + (i + 1) * step), count: 0 });
    values.forEach((x) => {
      if (!Number.isFinite(x)) return;
      let i = Math.floor((x - lo) / step + 1e-9);
      if (i < 0) i = 0;
      if (i >= nb) i = nb - 1;
      bins[i].count++;
    });
    return bins;
  }

  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(bytes) {
    let c = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }
  function zip(files) {
    const enc = new TextEncoder();
    const parts = [], central = [];
    let offset = 0;
    const d = new Date();
    const time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    files.forEach((f) => {
      const name = enc.encode(f.name);
      const data = typeof f.data === 'string' ? enc.encode(f.data) : f.data;
      const method = f.deflated ? 8 : 0;
      const size = f.deflated ? f.size : data.length;
      const crc = f.deflated ? f.crc : crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x0800, true);
      lh.setUint16(8, method, true);
      lh.setUint16(10, time, true);
      lh.setUint16(12, date, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, size, true);
      lh.setUint16(26, name.length, true);
      parts.push(new Uint8Array(lh.buffer), name, data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true);
      ch.setUint16(10, method, true);
      ch.setUint16(12, time, true);
      ch.setUint16(14, date, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true);
      ch.setUint32(24, size, true);
      ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    });
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    const all = [...parts, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(all.reduce((s, c) => s + c.length, 0));
    let p = 0;
    all.forEach((c) => { out.set(c, p); p += c.length; });
    return out;
  }
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream === 'undefined')
      throw new Error(IN_BROWSER ? t('oldBrowser') : 'Ce navigateur est trop ancien pour lire les fichiers Excel. Enregistrez le fichier en CSV ou mettez votre navigateur à jour.');
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  function unzip(bytes) {
    const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    let e = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { e = i; break; }
    }
    if (e < 0) throw new Error(IN_BROWSER ? t('invalidXlsx') : 'Ce fichier n’est pas un fichier Excel .xlsx valide.');
    const count = dv.getUint16(e + 10, true);
    let p = dv.getUint32(e + 16, true);
    const dec = new TextDecoder();
    const files = new Map();
    for (let k = 0; k < count && dv.getUint32(p, true) === 0x02014b50; k++) {
      const method = dv.getUint16(p + 10, true);
      const csize = dv.getUint32(p + 20, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      const off = dv.getUint32(p + 42, true);
      const name = dec.decode(bytes.subarray(p + 46, p + 46 + nlen));
      p += 46 + nlen + xlen + clen;
      const start = off + 30 + dv.getUint16(off + 26, true) + dv.getUint16(off + 28, true);
      files.set(name, { method, data: bytes.subarray(start, start + csize) });
    }
    return {
      async text(name) {
        const f = files.get(name);
        if (!f) return null;
        if (f.method === 0) return dec.decode(f.data);
        if (f.method === 8) return dec.decode(await inflateRaw(f.data));
        throw new Error(IN_BROWSER ? t('unsupportedZip') : 'Compression du fichier Excel non prise en charge.');
      },
    };
  }

  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  const xmlEsc = (s) =>
    String(s)
      .replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]))
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
  const xmlUnesc = (s) =>
    s.replace(/&(#x[0-9a-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (_, e) => {
      const l = e.toLowerCase();
      if (l[0] === '#') return String.fromCodePoint(l[1] === 'x' ? parseInt(l.slice(2), 16) : parseInt(l.slice(1), 10));
      return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }[l];
    });
  function colName(i) {
    let s = '';
    for (i++; i; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
    return s;
  }
  function colIndex(letters) {
    let n = 0;
    for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
    return n - 1;
  }

  function buildXlsx(sheets) {
    const sheetXml = ({ rows, widths }) => {
      const cols = widths && widths.length
        ? '<cols>' + widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('') + '</cols>'
        : '';
      const data = rows
        .map((row, r) =>
          `<row r="${r + 1}">` +
          row
            .map((v, c) => {
              if (v === null || v === undefined || v === '') return '';
              const ref = colName(c) + (r + 1), s = r === 0 ? ' s="1"' : '';
              if (typeof v === 'number' && Number.isFinite(v)) return `<c r="${ref}"${s}><v>${v}</v></c>`;
              return `<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
            })
            .join('') +
          '</row>'
        )
        .join('');
      return (
        XML_HEAD +
        `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheetViews><sheetView workbookViewId="0">` +
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>' +
        `${cols}<sheetData>${data}</sheetData></worksheet>`
      );
    };
    const safe = (n, i) => (String(n || 'Feuille' + (i + 1)).replace(/[\[\]:*?\/\\]/g, ' ').slice(0, 31) || 'Feuille' + (i + 1));
    const files = [
      {
        name: '[Content_Types].xml',
        data:
          XML_HEAD +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
          sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('') +
          '</Types>',
      },
      {
        name: '_rels/.rels',
        data:
          XML_HEAD +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
          '</Relationships>',
      },
      {
        name: 'xl/workbook.xml',
        data:
          XML_HEAD +
          `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_REL}"><sheets>` +
          sheets.map((s, i) => `<sheet name="${xmlEsc(safe(s.name, i))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('') +
          '</sheets></workbook>',
      },
      {
        name: 'xl/_rels/workbook.xml.rels',
        data:
          XML_HEAD +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="${NS_REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('') +
          `<Relationship Id="rId${sheets.length + 1}" Type="${NS_REL}/styles" Target="styles.xml"/>` +
          '</Relationships>',
      },
      {
        name: 'xl/styles.xml',
        data:
          XML_HEAD +
          `<styleSheet xmlns="${NS_MAIN}">` +
          '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
          '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
          '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>' +
          '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
          '</styleSheet>',
      },
      ...sheets.map((s, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, data: sheetXml(s) })),
    ];
    return zip(files);
  }

  async function readXlsx(bytes) {
    const z = unzip(bytes);
    let sheetPath = 'xl/worksheets/sheet1.xml';
    const wb = await z.text('xl/workbook.xml');
    const rels = await z.text('xl/_rels/workbook.xml.rels');
    if (wb && rels) {
      const tag = (/<(?:\w+:)?sheet\b[^>]*>/.exec(wb) || [''])[0];
      const rid = (/\s\w+:id="([^"]+)"/.exec(tag) || [])[1];
      (rels.match(/<(?:\w+:)?Relationship\b[^>]*>/g) || []).forEach((r) => {
        if ((/\bId="([^"]+)"/.exec(r) || [])[1] !== rid) return;
        const target = (/\bTarget="([^"]+)"/.exec(r) || [])[1] || '';
        sheetPath = target.startsWith('/') ? target.slice(1) : 'xl/' + target.replace(/^\.\//, '');
      });
    }
    const xml = await z.text(sheetPath);
    if (!xml) throw new Error(IN_BROWSER ? t('missingSheet') : 'Impossible de trouver la feuille de calcul dans ce fichier.');
    const allText = (s) =>
      xmlUnesc(
        s.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, '')
          .match(/<(?:\w+:)?t\b[^>]*>[\s\S]*?<\/(?:\w+:)?t>/g)?.map((t) => t.replace(/<[^>]+>/g, '')).join('') || ''
      );
    const sstXml = (await z.text('xl/sharedStrings.xml')) || '';
    const shared = (sstXml.match(/<(?:\w+:)?si\b[^>]*>[\s\S]*?<\/(?:\w+:)?si>/g) || []).map(allText);

    const rows = [];
    const rowRe = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
    let m, nextRow = 0;
    while ((m = rowRe.exec(xml))) {
      const rAttr = /\br="(\d+)"/.exec(m[1]);
      const r = rAttr ? +rAttr[1] - 1 : nextRow;
      nextRow = r + 1;
      const row = [];
      const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
      let c, nextCol = 0;
      while ((c = cellRe.exec(m[2] || ''))) {
        const attrs = c[1], body = c[2] || '';
        const ref = /\br="([A-Z]+)\d*"/.exec(attrs);
        const ci = ref ? colIndex(ref[1]) : nextCol;
        nextCol = ci + 1;
        const t = (/\bt="(\w+)"/.exec(attrs) || [])[1];
        const v = (/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body) || [])[1];
        let val = '';
        if (t === 's') val = v !== undefined ? shared[+v] ?? '' : '';
        else if (t === 'inlineStr') val = allText(body);
        else if (t === 'str' || t === 'e') val = v !== undefined ? xmlUnesc(v) : '';
        else if (t === 'b') val = v === '1' ? 'VRAI' : v === '0' ? 'FAUX' : '';
        else if (v !== undefined && v !== '') val = Number(v);
        row[ci] = val;
      }
      rows[r] = row;
    }
    const width = Math.max(0, ...rows.filter(Boolean).map((r) => r.length));
    return rows
      .filter(Boolean)
      .map((r) => Array.from({ length: width }, (_, i) => (r[i] === undefined ? '' : r[i])))
      .filter((r) => r.some((x) => String(x).trim() !== ''));
  }

  function decodeText(bytes) {
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (_) {
      return new TextDecoder('windows-1252').decode(bytes);
    }
  }
  function parseCsv(text) {
    text = text.replace(/^﻿/, '');
    const first = text.split(/\r?\n/, 1)[0];
    const delim = [';', '\t', ','].map((d) => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
    const rows = [];
    let row = [], cell = '', quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (quoted) {
        if (ch === '"') {
          if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false;
        } else cell += ch;
      } else if (ch === '"' && cell === '') quoted = true;
      else if (ch === delim) { row.push(cell); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(cell); rows.push(row); row = []; cell = '';
      } else cell += ch;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter((r) => r.some((c) => String(c).trim() !== ''));
  }

  function detectColumns(rows) {
    const N = (s) => norm(String(s ?? ''));
    const width = Math.max(0, ...rows.map((r) => r.length));
    let map = { header: -1, nom: -1, prenom: -1, note: -1 };
    for (let r = 0; r < Math.min(5, rows.length); r++) {
      const m = { header: r, nom: -1, prenom: -1, note: -1 };
      let complet = -1;
      rows[r].map(N).forEach((c, i) => {
        if (m.prenom < 0 && /^(prenoms?|first ?name|given ?name)$/.test(c)) m.prenom = i;
        else if (m.nom < 0 && /^(noms?|nom de famille|last ?name|surname|family ?name)$/.test(c)) m.nom = i;
        else if (complet < 0 && /^(nom complet|noms? et prenoms?|nom prenom|prenom nom|etudiante?s?|eleves?|name|full ?name|students?)$/.test(c)) complet = i;
        else if (m.note < 0 && /\b(notes?|resultats?|grades?|scores?|points?|total|reel)\b/.test(c)) m.note = i;
      });
      if (m.nom < 0 && complet >= 0) { m.nom = complet; m.prenom = -1; }
      if (m.nom >= 0 || m.prenom >= 0) {
        if (m.nom < 0) { m.nom = m.prenom; m.prenom = -1; }
        map = m;
        break;
      }
    }
    const data = rows.slice(map.header + 1);
    const isNum = (x) => Number.isFinite(toNumber(x));
    const share = (i, test) => data.filter((r) => test(r[i])).length / Math.max(1, data.length);
    const headerOf = (i) => (map.header >= 0 ? N(rows[map.header][i]) : '');
    const idLike = (i) => /(numero|matricule|identifi|\bid\b|code|telephone|courriel|email)/.test(headerOf(i));
    if (map.note < 0) {
      let best = -1, bestShare = 0;
      for (let i = 0; i < width; i++) {
        if (i === map.nom || i === map.prenom || idLike(i)) continue;
        const s = share(i, isNum);
        if (s > bestShare + 1e-9) { best = i; bestShare = s; }
      }
      if (bestShare >= 0.3) map.note = best;
    }
    if (map.nom < 0) {
      for (let i = 0; i < width; i++) {
        if (i === map.note) continue;
        if (share(i, (x) => /[a-zA-ZÀ-ÿ]/.test(String(x ?? '')) && !EMAIL_RE.test(String(x))) >= 0.5) { map.nom = i; break; }
      }
    }
    return map;
  }

  function rowsToText(rows, map) {
    const clean = (s) => String(s ?? '').replace(/[;\t\r\n]+/g, ' ').replace(/\s+/g, ' ').trim();
    const lines = [];
    let skipped = 0;
    rows.slice(map.header + 1).forEach((r) => {
      const name = clean([r[map.nom], map.prenom >= 0 ? r[map.prenom] : ''].filter((x) => clean(x)).join(' '));
      const g = toNumber(r[map.note]);
      if (!name) return;
      if (!Number.isFinite(g)) { skipped++; return; }
      lines.push(`${name} ${fmtGrade(g)}`);
    });
    return { text: lines.join('\n'), count: lines.length, skipped };
  }

  function similarity(a, b) {
    if (a === b) return 1;
    const m = a.length, n = b.length;
    if (!m || !n) return 0;
    let prev = Array.from({ length: n + 1 }, (_, j) => j);
    for (let i = 1; i <= m; i++) {
      const cur = [i];
      for (let j = 1; j <= n; j++) cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
      prev = cur;
    }
    return 1 - prev[n] / Math.max(m, n);
  }
  function phonetic(word) {
    let s = norm(word).replace(/[^a-z]/g, '');
    if (s.length < 2) return s;
    s = s
      .replace(/sch|ch|sh/g, 'S')
      .replace(/ph/g, 'f').replace(/th/g, 't')
      .replace(/qu|q/g, 'k').replace(/c(?=[eiy])/g, 's').replace(/ck|c/g, 'k')
      .replace(/gn/g, 'ni').replace(/gu(?=[eiy])/g, 'g').replace(/g(?=[eiy])/g, 'j')
      .replace(/x/g, 'ks').replace(/h/g, '').replace(/w/g, 'v').replace(/y/g, 'i').replace(/z/g, 's')
      .replace(/eaux?|aux?|au/g, 'o').replace(/ou/g, 'u').replace(/oi/g, 'oa')
      .replace(/ai|ei/g, 'e')
      .replace(/(?:e|a)[nm](?![aeiou])/g, 'A')
      .replace(/(?:ai|ei|i|u)[nm](?![aeiou])/g, 'I')
      .replace(/o[nm](?![aeiou])/g, 'O')
      .replace(/(.)\1+/g, '$1')
      .replace(/(?:er|ez|et)$/, 'e')
      .replace(/(.{2,}?)[stdxpe]+$/, '$1');
    return s;
  }
  function wordSim(w, t) {
    if (w === t) return 1;
    let s = Math.max(similarity(w, t), 0.97 * similarity(phonetic(w), phonetic(t)));
    if (t.startsWith(w)) s = Math.max(s, w.length >= 3 ? 0.93 : w.length === 2 ? 0.85 : 0.8);
    return s;
  }
  function nameScore(want, nameToks) {
    if (!nameToks.length || !want.length) return 0;
    const joinedW = want.join(''), joinedN = nameToks.join(''), joinedR = nameToks.slice().reverse().join('');
    const perToken = want.reduce((s, w) => s + Math.max(...nameToks.map((t) => wordSim(w, t))), 0) / want.length;
    return Math.max(perToken, wordSim(joinedW, joinedN), wordSim(joinedW, joinedR));
  }
  function bestMatch(name, roster) {
    const want = tokens(name);
    const scored = roster.map((r) => ({ r, s: nameScore(want, r.toks) })).sort((a, b) => b.s - a.s);
    return { best: scored[0] ? scored[0].r : null, score: scored[0] ? scored[0].s : 0, second: scored[1] ? scored[1].s : 0 };
  }
  function compareAssessments(first, second, maxFirst, maxSecond) {
    const index = (rows) => {
      const byId = new Map();
      rows.forEach((r) => {
        if (!r.id) return;
        const previous = byId.get(r.id);
        if (previous && !Object.is(previous.value, r.value)) byId.set(r.id, { ...r, value: NaN });
        else if (!previous) byId.set(r.id, r);
      });
      return byId;
    };
    const a = index(first), b = index(second);
    const scaled = Number.isFinite(maxFirst) && maxFirst > 0 && Number.isFinite(maxSecond) && maxSecond > 0;
    const rows = [...new Set([...a.keys(), ...b.keys()])].map((id) => {
      const x = a.get(id), y = b.get(id);
      const from = x ? x.value : NaN, to = y ? y.value : NaN;
      const paired = Number.isFinite(from) && Number.isFinite(to);
      const fromPct = scaled && Number.isFinite(from) ? from / maxFirst * 100 : NaN;
      const toPct = scaled && Number.isFinite(to) ? to / maxSecond * 100 : NaN;
      const delta = scaled && paired ? Math.round((toPct - fromPct) * 1e10) / 1e10 : NaN;
      return {id, name: (x && x.name) || (y && y.name) || '', from, to, paired, fromPct, toPct, delta};
    });
    const pairs = rows.filter((r) => r.paired);
    return {
      rows, n: pairs.length, excluded: rows.length - pairs.length, scaled,
      first: computeStats(pairs.map((r) => r.fromPct)),
      second: computeStats(pairs.map((r) => r.toPct)),
      change: computeStats(pairs.map((r) => r.delta)),
      up: pairs.filter((r) => r.delta > 0).length,
      down: pairs.filter((r) => r.delta < 0).length,
      same: pairs.filter((r) => r.delta === 0).length,
    };
  }

  if (!IN_BROWSER) {
    if (typeof module !== 'undefined') {
      module.exports = {
        similarity, phonetic, nameScore, bestMatch,
        norm, tokens, splitEntries, parseList, splitName,
        compareAssessments, countLetterGrades, toNumber, quantile, computeStats, guessMax, niceStep, histogram,
        buildXlsx, readXlsx, parseCsv, decodeText, detectColumns, rowsToText,
        _zip: zip, _crc32: crc32,
      };
    }
    return;
  }

  const LANG_KEY = 'studium-notes-langue';
  const pl = (n, one, many) => (n > 1 ? many : one);
  const I18N = {
    fr: {
      tMinSub: "la plus basse note",
      tMaxSub: "la plus haute note",
      progressTitle: "Progression entre deux évaluations",
      progressNeedTwo: "Affichez au moins deux évaluations avec des champs de notes sur la même page du carnet de notes, en mode édition.",
      progressIntro: "Choisissez les deux évaluations à comparer. Ce choix sert uniquement aux statistiques ; il ne change pas l’évaluation à remplir en haut.",
      progressFrom: "Première évaluation",
      progressTo: "Deuxième évaluation",
      progressMaxFrom: "Première évaluation : note sur",
      progressMaxTo: "Deuxième évaluation : note sur",
      progressMaxHint: "Ex. 20 ou 100",
      progressDifferent: "Choisissez deux évaluations différentes.",
      progressSetMax: "Indiquez le barème de chaque évaluation pour calculer la progression sur une échelle commune de 0 à 100. Les notes brutes restent visibles ci-dessous.",
      progressNoPairs: "Aucun étudiant n’a encore deux notes numériques comparables.",
      progressClass: "Pour la classe, avec les mêmes étudiants aux deux évaluations",
      progressMeanFrom: "Moyenne à la première",
      progressMeanTo: "Moyenne à la deuxième",
      progressMeanChange: "Progression moyenne",
      progressMedianChange: "Progression médiane",
      progressMeaning: "pp = points de pourcentage. Par exemple, passer de 60 % à 75 % correspond à +15 pp. Cela montre l’évolution des notes. Les deux évaluations peuvent toutefois être de difficulté différente.",
      progressStudents: "Par étudiant",
      progressName: "Étudiant",
      progressChange: "Progression",
      progressMissing: "n.d. veut dire que la note manque, n’est pas un nombre ou ne peut pas être comparée. Les moyennes comparées excluent les étudiants sans deux notes numériques. Seuls les étudiants présents sur cette page sont inclus.",
      progressCoverage: (n, excluded) => `${n} étudiant(s) avec deux notes numériques ; ${excluded} exclu(s) de la comparaison.`,
      progressCounts: (up, same, down) => `${up} en hausse · ${same} stable(s) · ${down} en baisse`,
      viewOverview: "Résumé",
      viewCharts: "Graphiques",
      viewClass: "Progression classe",
      viewStudents: "Par étudiant",
      scaleShort: "Sur",
      searchStudent: "Chercher un étudiant",
      howRead: "Comment lire ces résultats",
      scalePrompt: "Renseignez les deux barèmes « Sur » pour calculer la progression en points de pourcentage.",
      upLabel: "En hausse",
      sameLabel: "Stables",
      downLabel: "En baisse",
      previous: "Précédent",
      next: "Suivant",
      noSearchResults: "Aucun étudiant trouvé.",
      closeTitle: "Fermer (le favori rouvre l’outil)",
      noFieldsHelp: "Aucun champ modifiable détecté. Dans StudiUM, activez « Mode d’édition », puis cliquez sur le bouton Relire la page (⟳). Si les cases restent absentes, activez l’évaluation rapide dans les préférences du rapport.",
      noFieldsStats: "Les statistiques seront disponibles lorsque les champs de notes seront affichés.",
      pageCount: (p, n, total) => `Page ${p}/${n} · ${total} étudiant(s)`,
      langName: 'Français', locale: 'fr-CA', na: 'n.d.',
      pillTitle: 'Ouvrir la fenêtre (vous pouvez aussi la glisser)', headTitle: 'Cliquer-glisser pour déplacer',
      minTitle: 'Réduire', langTitle: 'Langue', evalLabel: 'Évaluation :', colTitle: 'Évaluation à remplir', rescanTitle: 'Relire la page',
      cTotal: 'étudiants', cDone: 'avec note', cRest: 'restants ▾', cRestTitle: 'Voir les étudiants sans note',
      noName: '(sans nom)', allDone: 'Tous les étudiants ont une note.',
      tabHome: 'Accueil', tabEnter: 'Saisir', tabStats: 'Statistiques', tabExport: 'Exporter', tabImport: 'Importer',
      hello: 'Que voulez-vous faire ?',
      m1t: 'Saisir des notes', m1d: 'Collez une liste « Nom Prénom note »',
      m2t: 'Voir les statistiques', m2d: 'Moyenne, minimum, maximum et progression',
      m3t: 'Exporter vers Excel', m3d: 'Téléchargez les noms, prénoms et notes',
      m4t: 'Importer un fichier Excel', m4d: 'Remplissez StudiUM à partir d’un fichier',
      howTitle: 'Comment ça marche',
      how1: 'En haut, choisissez l’<b>évaluation</b> à remplir.',
      how2: 'Saisissez ou importez les notes : elles s’inscrivent dans StudiUM.',
      how3: 'Vérifiez, puis cliquez sur <b>« Enregistrer »</b> dans StudiUM.',
      aboutTitle: 'À propos',
      aboutCheck: "Il vous appartient de vérifier l’exactitude des notes inscrites avant d’enregistrer.",
      aboutUsage: "Usage non commercial uniquement. L’outil ne doit pas être vendu ni servir à contourner une fonction payante.",
      aboutTech: "Outil en HTML, CSS et JavaScript, exécuté uniquement dans votre navigateur. L’outil ne garde rien et n’envoie rien. À chaque ouverture, il lit la page StudiUM affichée à l’écran : les noms et les notes viennent de cette page. Ce que vous saisissez est effacé quand vous fermez l’outil ou la page. Aucune IA ne traite les données. Il fonctionne sur StudiUM (studium.(université concernée), par exemple studium.umontreal.ca).",
      aboutPrivacy: "L’auteur n’a aucun accès à l’utilisation de l’outil : pas de statistiques, pas de suivi. Une fois installé, l’outil fonctionne entièrement en local : il ne charge rien depuis ce site, ne communique avec aucun serveur en dehors de StudiUM et ne se met jamais à jour tout seul.",
      aboutBrowser: "Conçu pour Edge et Chrome (Windows et Mac).",
      aboutText: 'V{v} · Dernière mise à jour par Eric Bernier le 05/11/26.',
      hint: 'Une ligne par étudiant : <code>Nom Prénom note</code>. Un nom incomplet suffit. Vous pouvez aussi coller deux colonnes copiées d’Excel.',
      paste: 'Coller', clear: 'Vider',
      placeholder: 'Tremblay Jean 30\nGagnon Marie-Ève 27,5\nRoy Luc 25',
      over: 'Remplacer les notes déjà présentes', fillBtn: 'Remplir les notes', undoBtn: 'Annuler les modifications',
      changesTitle: (n) => `Modifications à vérifier (${n})`, changeBefore: 'Avant', changeAfter: 'Après',
      changesHelp: 'Toutes les notes changées par l’outil sur cette évaluation. Les lignes en bleu viennent d’un nom approchant. Relisez avant de cliquer sur « Enregistrer » dans StudiUM.',
      undone: (n) => `Modifications annulées : ${n} note(s) remise(s) comme avant.`,
      approxLine: (typed, name, grade) => `Nom approchant : ${typed} → ${name}, ${grade}`,
      notice: 'Rien n’est enregistré automatiquement : vérifiez, puis cliquez sur « Enregistrer » dans StudiUM.',
      expTitle: 'Exporter vers Excel',
      expDesc: 'Vous obtenez un fichier Excel avec trois colonnes (<b>Nom</b>, <b>Prénom</b>, <b>Note</b>) pour chaque étudiant de l’évaluation choisie en haut. Une deuxième feuille contient les statistiques. Ce même fichier peut être réimporté.',
      expBtn: 'Télécharger le fichier Excel', expWhere: 'Le fichier arrive dans votre dossier « Téléchargements ».',
      impTitle: 'Importer un fichier Excel',
      impDesc: 'Le plus simple : trois colonnes <b>Nom</b>, <b>Prénom</b>, <b>Note</b>, une ligne par étudiant (fichier .xlsx ou .csv). Un fichier exporté ici fonctionne directement.',
      dropMain: 'Glissez votre fichier ici', dropSub: 'ou cliquez pour le choisir',
      noFields: 'Aucun champ de note sur cette page',
      colOption: (label, n) => `${label} (${n} étudiants)`,
      pageGrade: 'Note de la page', item: (k) => `Élément ${k}`,
      repOk: (n, total) => `${n} note(s) remplie(s) sur ${total} étudiants`,
      repApprox: 'Noms approchants à vérifier (encadrés en bleu)', repNotFound: 'Non trouvés',
      repAmbig: 'Ambigus (non remplis)', repSkipped: 'Ignorés (note déjà présente)', repInvalid: 'Valeur refusée', repBad: 'Lignes illisibles',
      ambigMany: (line, n, names) => `${line} : ${n} étudiants possibles (${names}). Ajoutez un mot du nom.`,
      ambigUsed: (line) => `${line} : étudiant déjà rempli par une autre ligne.`,
      skippedLine: (line, v) => `${line}  (note existante : ${v})`,
      invalidLine: (line) => `${line}  (valeur refusée par la liste déroulante)`,
      alertClipboard: 'Le navigateur a bloqué l’accès au presse-papiers. Cliquez dans le cadre blanc puis faites Ctrl+V (⌘+V sur Mac).',
      alertNoFields: 'Aucun champ de note sur cette page. Activez le mode édition du carnet de notes, ou l’« Évaluation rapide » du devoir.',
      alertNoExport: 'Aucun champ de note sur cette page à exporter.',
      noStatsText: 'Les notes de cette évaluation ne sont pas des nombres : pas de statistiques possibles.',
      noStatsYet: 'Aucune note pour l’instant.<br>Les statistiques apparaîtront dès les premières notes.',
      letterCounts: 'Effectifs par lettre', belowCShort: 'Sous C', belowCTitle: 'Moins de 60 % (sous C)',
      failuresShort: 'Échecs', failuresTitle: 'Échecs : moins de 60 % (sous C, la note de passage)',
      letterScale: 'Barème de cette évaluation : note maximale', letterScaleNeeded: 'Indiquez le barème pour compter les notes par lettre.',
      range: (a, b) => `${a} et ${b}`, to: 'à', pct: (x) => `${x} %`, andBetween: '</b>, et entre <b>',
      partial: (n) => ` (${n} ${pl(n, 'étudiant', 'étudiants')} sans note pour l’instant)`,
      brief: (n, partial, mean, median, tops) =>
        `Sur <b>${n}</b> ${pl(n, 'note', 'notes')}${partial}, la moyenne de la classe est <b>${mean}</b>. ` +
        `La moitié des étudiants a plus de <b>${median}</b>. La plupart des notes se trouvent entre <b>${tops}</b>.`,
      briefTitle: 'En bref', centerTitle: 'Le centre de la classe', spreadTitle: 'L’écart entre les étudiants',
      tMean: 'Moyenne', tMeanSub: 'total des notes ÷ nombre d’étudiants',
      tMedian: 'Médiane', tMedianSub: 'la moitié est au-dessus, la moitié en dessous',
      tSd: 'Écart-type', tSdSub: 'distance moyenne à la moyenne',
      tIqr: 'Interquartile (Q1 à Q3)', tIqrSub: (x) => `les 50 % du milieu · écart ${x}`,
      tIdr: 'Interdécile (P10 à P90)', tIdrSub: (x) => `80 % des notes · écart ${x}`,
      histTitle: 'Répartition des notes',
      histCap: 'Chaque barre indique combien d’étudiants ont une note dans cette tranche. La plus haute est la plus fréquente.',
      settings: 'Réglages du graphique', outOf: 'Note sur', barWidth: 'Largeur des barres',
      auto: (x) => `automatique (${x})`, points: (x) => `${x} points`,
      bandTitle: 'Où se situe la majorité de la classe',
      bandCap: 'Les notes les plus basses et les plus hautes sont écartées pour montrer où se trouve la majorité de la classe.',
      band50: (a, b) => `<b>50 %</b> des étudiants ont entre <b>${a}</b> et <b>${b}</b>`,
      band80: (a, b) => `<b>80 %</b> des étudiants ont entre <b>${a}</b> et <b>${b}</b>`,
      svgMean: 'moyenne', svgMedian: 'médiane', svgBin: (a, b, n, p) => `${a} à ${b} : ${n} note(s), soit ${p} %`,
      svgHistLabel: 'Répartition des notes', svgBandLabel: 'Où se situe la majorité de la classe',
      svg80: (a, b) => `80 % des notes : ${a} à ${b}`, svg50: (a, b) => `50 % des notes : ${a} à ${b}`,
      xNom: 'Nom', xPrenom: 'Prénom', xNote: 'Note', xSheetNotes: 'Notes', xSheetStats: 'Statistiques',
      xStat: 'Statistique', xValue: 'Valeur', xItem: 'Évaluation', xDate: 'Exporté le', xStudents: 'Étudiants',
      xDone: 'Avec note', xRest: 'Restants (sans note)', xMean: 'Moyenne', xMedian: 'Médiane', xSd: 'Écart-type',
      xMin: 'Minimum', xMax: 'Maximum', xQ1: '1er quartile (Q1)', xQ3: '3e quartile (Q3)', xIqr: 'Écart interquartile',
      xP10: '10e centile (P10)', xP90: '90e centile (P90)', xIdr: 'Écart interdécile', xFile: 'Notes StudiUM',
      oldBrowser: 'Ce navigateur est trop ancien pour lire les fichiers Excel. Enregistrez le fichier en CSV ou mettez votre navigateur à jour.',
      invalidXlsx: 'Ce fichier n’est pas un fichier Excel .xlsx valide.',
      unsupportedZip: 'Compression du fichier Excel non prise en charge.',
      missingSheet: 'Impossible de trouver la feuille de calcul dans ce fichier.',
      reading: 'Lecture du fichier…', cantRead: 'Impossible de lire ce fichier', emptyFile: 'Le fichier est vide.',
      oldXls: 'Les anciens fichiers .xls ne sont pas pris en charge. Dans Excel, choisissez Fichier, puis Enregistrer sous, puis « Classeur Excel (.xlsx) ».',
      none: '(aucune)', lines: (n) => `${n} ligne(s)`,
      mapNom: 'Nom (ou nom complet)', mapPrenom: 'Prénom', mapNote: 'Note', firstRowHeaders: 'La première ligne contient les titres',
      importGo: (n) => `Remplir StudiUM avec ${n} note(s)`, importSkipped: (n) => `${n} ligne(s) sans note numérique seront ignorées.`,
    },
    en: {
      tMinSub: "the lowest grade",
      tMaxSub: "the highest grade",
      progressTitle: "Progress between two assessments",
      progressNeedTwo: "Show at least two assessments with grade fields on the same gradebook page, with editing enabled.",
      progressIntro: "Choose the two assessments to compare. This selection is only for statistics; it does not change the assessment to fill in at the top.",
      progressFrom: "First assessment",
      progressTo: "Second assessment",
      progressMaxFrom: "First assessment: out of",
      progressMaxTo: "Second assessment: out of",
      progressMaxHint: "E.g. 20 or 100",
      progressDifferent: "Choose two different assessments.",
      progressSetMax: "Enter the maximum grade for each assessment to calculate progress on a common scale from 0 to 100. Raw grades remain visible below.",
      progressNoPairs: "No student has two comparable numeric grades yet.",
      progressClass: "For the class, with the same students in both assessments",
      progressMeanFrom: "First assessment average",
      progressMeanTo: "Second assessment average",
      progressMeanChange: "Average progress",
      progressMedianChange: "Median progress",
      progressMeaning: "pp = percentage points. For example, going from 60% to 75% is +15 pp. This shows how grades changed. The two assessments may still differ in difficulty.",
      progressStudents: "By student",
      progressName: "Student",
      progressChange: "Progress",
      progressMissing: "n/a means the grade is missing, is not a number or cannot be compared. Compared averages exclude students without two numeric grades. Only students on this page are included.",
      progressCoverage: (n, excluded) => `${n} student(s) with two numeric grades; ${excluded} excluded from the comparison.`,
      progressCounts: (up, same, down) => `${up} improved · ${same} unchanged · ${down} decreased`,
      viewOverview: "Overview",
      viewCharts: "Charts",
      viewClass: "Class progress",
      viewStudents: "By student",
      scaleShort: "Out of",
      searchStudent: "Find a student",
      howRead: "How to read these results",
      scalePrompt: "Enter both maximum grades to calculate progress in percentage points.",
      upLabel: "Improved",
      sameLabel: "Unchanged",
      downLabel: "Decreased",
      previous: "Previous",
      next: "Next",
      noSearchResults: "No student found.",
      closeTitle: "Close (the bookmark reopens the tool)",
      noFieldsHelp: "No editable grade fields found. Turn on “Edit mode” in StudiUM, then click the Reread the page button (⟳). If fields are still missing, enable quick grading in the report preferences.",
      noFieldsStats: "Statistics will be available when the grade fields are displayed.",
      pageCount: (p, n, total) => `Page ${p}/${n} · ${total} student(s)`,
      langName: 'English', locale: 'en-CA', na: 'n/a',
      pillTitle: 'Open the window (you can also drag it)', headTitle: 'Click and drag to move',
      minTitle: 'Minimize', langTitle: 'Language', evalLabel: 'Assessment:', colTitle: 'Assessment to fill in', rescanTitle: 'Re-read the page',
      cTotal: 'students', cDone: 'graded', cRest: 'remaining ▾', cRestTitle: 'Show students without a grade',
      noName: '(no name)', allDone: 'Every student has a grade.',
      tabHome: 'Home', tabEnter: 'Enter', tabStats: 'Statistics', tabExport: 'Export', tabImport: 'Import',
      hello: 'What would you like to do?',
      m1t: 'Enter grades', m1d: 'Paste a list “Last name First name grade”',
      m2t: 'See the statistics', m2d: 'Average, minimum, maximum and progress',
      m3t: 'Export to Excel', m3d: 'Download last names, first names and grades',
      m4t: 'Import an Excel file', m4d: 'Fill in StudiUM from a file',
      howTitle: 'How it works',
      how1: 'At the top, choose the <b>assessment</b> to fill in.',
      how2: 'Type or import the grades: they are entered into StudiUM.',
      how3: 'Check them, then click <b>“Save”</b> in StudiUM.',
      aboutTitle: 'About',
      aboutCheck: "You are responsible for checking that the grades entered are correct before saving.",
      aboutUsage: "Non-commercial use only. The tool must not be sold or used to get around a paid feature.",
      aboutTech: "Built with HTML, CSS and JavaScript, and runs only in your browser. The tool keeps nothing and sends nothing. Each time it opens, it reads the StudiUM page on screen: the names and grades come from that page. What you type is cleared when you close the tool or the page. No AI processes the data. It works on StudiUM (studium.(your university), for example studium.umontreal.ca).",
      aboutPrivacy: "The author has no access to how the tool is used: no statistics, no tracking. Once installed, the tool runs entirely on your computer: it loads nothing from this site, does not talk to any server other than StudiUM and never updates itself.",
      aboutBrowser: "Designed for Edge and Chrome (Windows and Mac).",
      aboutText: 'V{v} · Last updated by Eric Bernier on Nov. 5, 2026.',
      hint: 'One line per student: <code>Last First grade</code>. Part of the name is enough. You can also paste two columns copied from Excel.',
      paste: 'Paste', clear: 'Clear',
      placeholder: 'Tremblay Jean 30\nGagnon Marie-Ève 27.5\nRoy Luc 25',
      over: 'Replace grades already entered', fillBtn: 'Fill in the grades', undoBtn: 'Undo changes',
      changesTitle: (n) => `Changes to check (${n})`, changeBefore: 'Before', changeAfter: 'After',
      changesHelp: 'All grades changed by the tool for this assessment. Rows in blue come from an approximate name. Check them before clicking “Save” in StudiUM.',
      undone: (n) => `Changes undone: ${n} grade(s) put back as they were.`,
      approxLine: (typed, name, grade) => `Approximate name: ${typed} → ${name}, ${grade}`,
      notice: 'Nothing is saved automatically: check, then click “Save” in StudiUM.',
      expTitle: 'Export to Excel',
      expDesc: 'You get an Excel file with three columns (<b>Last name</b>, <b>First name</b>, <b>Grade</b>) for every student in the assessment chosen at the top. A second sheet holds the statistics. The same file can be imported back.',
      expBtn: 'Download the Excel file', expWhere: 'The file goes to your “Downloads” folder.',
      impTitle: 'Import an Excel file',
      impDesc: 'Easiest: 3 columns <b>Last name</b>, <b>First name</b>, <b>Grade</b>, one row per student (.xlsx or .csv file). A file exported here works as is.',
      dropMain: 'Drop your file here', dropSub: 'or click to choose it',
      noFields: 'No grade fields on this page',
      colOption: (label, n) => `${label} (${n} students)`,
      pageGrade: 'Grade on this page', item: (k) => `Item ${k}`,
      repOk: (n, total) => `${n} grade(s) filled in out of ${total} students`,
      repApprox: 'Approximate names to check (outlined in blue)', repNotFound: 'Not found',
      repAmbig: 'Ambiguous (not filled in)', repSkipped: 'Skipped (grade already entered)', repInvalid: 'Value rejected', repBad: 'Unreadable lines',
      ambigMany: (line, n, names) => `${line}: ${n} possible students (${names}). Add a word of the name.`,
      ambigUsed: (line) => `${line}: student already filled in by another line.`,
      skippedLine: (line, v) => `${line}  (existing grade: ${v})`,
      invalidLine: (line) => `${line}  (value rejected by the drop-down list)`,
      alertClipboard: 'The browser blocked clipboard access. Click in the white box, then press Ctrl+V (⌘+V on Mac).',
      alertNoFields: 'No grade fields on this page. Turn editing on in the gradebook, or use the assignment’s “Quick grading”.',
      alertNoExport: 'No grade fields on this page to export.',
      noStatsText: 'The grades for this assessment are not numbers: no statistics available.',
      noStatsYet: 'No grades yet.<br>Statistics will appear as soon as the first grades are entered.',
      letterCounts: 'Students per letter grade', belowCShort: 'Below C', belowCTitle: 'Below 60% (below C)',
      failuresShort: 'Fail', failuresTitle: 'Failures: below 60% (below C, the passing grade)',
      letterScale: 'Assessment scale: maximum possible grade', letterScaleNeeded: 'Enter the maximum possible grade to count letter grades.',
      range: (a, b) => `${a} and ${b}`, to: 'to', pct: (x) => `${x}%`, andBetween: '</b>, and between <b>',
      partial: (n) => ` (${n} ${pl(n, 'student', 'students')} without a grade so far)`,
      brief: (n, partial, mean, median, tops) =>
        `Out of <b>${n}</b> ${pl(n, 'grade', 'grades')}${partial}, the class average is <b>${mean}</b>. ` +
        `Half of the students scored above <b>${median}</b>. Most grades fall between <b>${tops}</b>.`,
      briefTitle: 'In short', centerTitle: 'The middle of the class', spreadTitle: 'How spread out the class is',
      tMean: 'Average', tMeanSub: 'sum of grades ÷ number of students',
      tMedian: 'Median', tMedianSub: 'half above, half below',
      tSd: 'Standard deviation', tSdSub: 'typical distance from the average',
      tIqr: 'Interquartile (Q1 to Q3)', tIqrSub: (x) => `the middle 50% · spread ${x}`,
      tIdr: 'Interdecile (P10 to P90)', tIdrSub: (x) => `80% of grades · spread ${x}`,
      histTitle: 'How grades are distributed',
      histCap: 'Each bar shows how many students have a grade in that range. The tallest bar is the most common.',
      settings: 'Chart settings', outOf: 'Out of', barWidth: 'Bar width',
      auto: (x) => `automatic (${x})`, points: (x) => `${x} points`,
      bandTitle: 'Where most of the class stands',
      bandCap: 'The lowest and highest grades are left out to show where most of the class is.',
      band50: (a, b) => `<b>50%</b> of students scored between <b>${a}</b> and <b>${b}</b>`,
      band80: (a, b) => `<b>80%</b> of students scored between <b>${a}</b> and <b>${b}</b>`,
      svgMean: 'average', svgMedian: 'median', svgBin: (a, b, n, p) => `${a} to ${b}: ${n} grade(s), or ${p}%`,
      svgHistLabel: 'Grade distribution', svgBandLabel: 'Where most of the class stands',
      svg80: (a, b) => `80% of grades: ${a} to ${b}`, svg50: (a, b) => `50% of grades: ${a} to ${b}`,
      xNom: 'Last name', xPrenom: 'First name', xNote: 'Grade', xSheetNotes: 'Grades', xSheetStats: 'Statistics',
      xStat: 'Statistic', xValue: 'Value', xItem: 'Assessment', xDate: 'Exported on', xStudents: 'Students',
      xDone: 'Graded', xRest: 'Remaining (no grade)', xMean: 'Average', xMedian: 'Median', xSd: 'Standard deviation',
      xMin: 'Minimum', xMax: 'Maximum', xQ1: '1st quartile (Q1)', xQ3: '3rd quartile (Q3)', xIqr: 'Interquartile range',
      xP10: '10th percentile (P10)', xP90: '90th percentile (P90)', xIdr: 'Interdecile range', xFile: 'StudiUM grades',
      oldBrowser: 'This browser is too old to read Excel files. Save the file as CSV or update your browser.',
      invalidXlsx: 'This is not a valid Excel .xlsx file.',
      unsupportedZip: 'The compression used by this Excel file is not supported.',
      missingSheet: 'Cannot find a worksheet in this file.',
      reading: 'Reading the file…', cantRead: 'Cannot read this file', emptyFile: 'The file is empty.',
      oldXls: 'Old .xls files are not supported. In Excel, choose File, then Save As, then “Excel Workbook (.xlsx)”.',
      none: '(none)', lines: (n) => `${n} row(s)`,
      mapNom: 'Last name (or full name)', mapPrenom: 'First name', mapNote: 'Grade', firstRowHeaders: 'The first row contains the headings',
      importGo: (n) => `Fill in StudiUM with ${n} grade(s)`, importSkipped: (n) => `${n} row(s) without a numeric grade will be skipped.`,
    },
  };
  let LANG = 'fr';
  try {
    LANG = localStorage.getItem(LANG_KEY) || ((navigator.language || '').toLowerCase().startsWith('en') ? 'en' : 'fr');
  } catch (_) {}
  if (!I18N[LANG]) LANG = 'fr';
  LOCALE = I18N[LANG].locale;
  NA = I18N[LANG].na;
  const VERSION = '1.4';
  function t(key, ...args) {
    const v = I18N[LANG][key] !== undefined ? I18N[LANG][key] : I18N.fr[key];
    if (typeof v === 'function') return v(...args);
    return key === 'aboutText' ? v.replace('{v}', VERSION) : v;
  }
  let flagSeq = 0;
  function flagSvg(lang) {
    if (lang === 'fr')
      return '<svg class="flag" viewBox="0 0 3 2" aria-hidden="true"><rect width="1" height="2" fill="#002654"/><rect x="1" width="1" height="2" fill="#fff"/><rect x="2" width="1" height="2" fill="#ce1126"/></svg>';
    const id = 'ukc' + ++flagSeq;
    return (
      `<svg class="flag" viewBox="0 0 60 30" aria-hidden="true"><clipPath id="${id}"><path d="M30,15 h30 v15 z v15 h-30 z h-30 v-15 z v-15 h30 z"/></clipPath>` +
      '<rect width="60" height="30" fill="#012169"/><path d="M0,0 L60,30 M60,0 L0,30" stroke="#fff" stroke-width="6"/>' +
      `<path d="M0,0 L60,30 M60,0 L0,30" clip-path="url(#${id})" stroke="#C8102E" stroke-width="4"/>` +
      '<path d="M30,0 v30 M0,15 h60" stroke="#fff" stroke-width="10"/><path d="M30,0 v30 M0,15 h60" stroke="#C8102E" stroke-width="6"/></svg>'
    );
  }

  function columnKey(input) {
    const m = input.name.match(/^grade\[\d+\]\[(\d+)\]$/) || input.name.match(/^finalgrade_\d+_(\d+)$/);
    return m ? m[1] : '__unique__';
  }
  function columnLabel(key, sample) {
    if (key === '__unique__') return t('pageGrade');
    const h =
      document.querySelector(`[data-itemid="${key}"] .gradeitemheader`) ||
      document.querySelector(`th[data-itemid="${key}"]`) ||
      document.querySelector(`[data-itemid="${key}"]`);
    let label = h ? h.textContent.replace(/\s+/g, ' ').trim() : '';
    if (!label && sample) {
      const cell = sample.closest('td,th');
      const table = sample.closest('table');
      if (cell && table && table.tHead) {
        const hr = table.tHead.rows[table.tHead.rows.length - 1];
        const hc = hr && hr.cells[cell.cellIndex];
        if (hc) label = hc.textContent.replace(/\s+/g, ' ').trim();
      }
    }
    return label || t('item', key);
  }
  function scanPage() {
    const inputs = [...document.querySelectorAll(INPUT_SELECTOR)].filter((el) => !el.disabled && el.type !== 'hidden');
    const columns = new Map();
    inputs.forEach((el) => {
      const k = columnKey(el);
      if (!columns.has(k)) columns.set(k, []);
      columns.get(k).push(el);
    });
    return columns;
  }
  function assessmentMaximum(sample) {
    if (!sample) return null;
    const explicit = toNumber(sample.getAttribute('max'));
    if (explicit > 0) return explicit;
    const cell = sample.closest('td,th'), table = sample.closest('table');
    if (!cell || !table) return null;
    const column = [...cell.parentElement.cells].slice(0, cell.cellIndex).reduce((sum, c) => sum + c.colSpan, 0);
    for (const row of table.rows) {
      if (!/^(?:valeurs possibles|étendue|range)$/i.test((row.cells[0]?.textContent || '').trim())) continue;
      let start = 0;
      for (const rangeCell of row.cells) {
        if (start === column) {
          const match = rangeCell.textContent.trim().match(/^-?[\d\s.,]+\s*[–—−-]\s*([\d\s.,]+)$/);
          const maximum = match ? toNumber(match[1]) : NaN;
          return maximum > 0 ? maximum : null;
        }
        start += rangeCell.colSpan;
      }
    }
    return null;
  }
  const setupCache = { maxima: null, done: false, loading: false, waiting: [] };
  function pageCourseId() {
    const fromCfg = Number(window.M && window.M.cfg && window.M.cfg.courseId);
    if (fromCfg > 1) return fromCfg;
    if (/\/grade\/report\//.test(location.pathname)) return Number(new URLSearchParams(location.search).get('id')) || null;
    return null;
  }
  function parseSetupMaxima(doc) {
    const out = {};
    for (const table of doc.querySelectorAll('table')) {
      const head = table.tHead ? table.tHead.rows[table.tHead.rows.length - 1] : table.rows[0];
      if (!head) continue;
      let column = -1, start = 0;
      for (const cell of head.cells) {
        if (/note maximale|max(?:imum)? grade|grade max/i.test(cell.textContent)) { column = start; break; }
        start += cell.colSpan;
      }
      if (column < 0) continue;
      for (const row of table.rows) {
        if (row === head) continue;
        const id = row.dataset.itemid || (row.id.match(/item[-_](\d+)/) || [])[1] ||
          row.querySelector('[data-itemid]')?.dataset.itemid;
        if (!id || out[id]) continue;
        let position = 0, target = null;
        for (const cell of row.cells) { if (position === column) { target = cell; break; } position += cell.colSpan; }
        const value = target ? toNumber(target.textContent.trim()) : NaN;
        if (value > 0) out[id] = value;
      }
    }
    return out;
  }
  function setupMaxima(onReady) {
    if (setupCache.done) return setupCache;
    if (onReady) setupCache.waiting.push(onReady);
    if (setupCache.loading) return setupCache;
    const id = pageCourseId();
    const finish = (maxima) => {
      setupCache.maxima = maxima; setupCache.done = true; setupCache.loading = false;
      const waiting = setupCache.waiting.splice(0);
      waiting.forEach((callback) => { try { callback(); } catch (e) {} });
    };
    if (!id || typeof fetch !== 'function' || typeof DOMParser !== 'function') { finish(null); return setupCache; }
    setupCache.loading = true;
    fetch(`/grade/edit/tree/index.php?id=${id}`, { credentials: 'same-origin' })
      .then((response) => (response.ok ? response.text() : ''))
      .then((html) => finish(html ? parseSetupMaxima(new DOMParser().parseFromString(html, 'text/html')) : null))
      .catch(() => finish(null));
    return setupCache;
  }
  function rowNameText(input) {
    const row = input.closest('tr') || input.closest('[data-region="user"], .row, li');
    if (!row) return '';
    return studentName(input) + ' ' + row.textContent;
  }
  const studentLabelCache = new WeakMap();
  function studentLabel(element) {
    if (!element) return '';
    const raw = element.textContent || '';
    const cached = studentLabelCache.get(element);
    if (cached && cached.raw === raw) return cached.name;
    const copy = element.cloneNode(true);
    copy.querySelectorAll('.userinitials, .userpicture, img, svg, button, [aria-hidden="true"]').forEach((node) => node.remove());
    const name = (copy.textContent || '').replace(/\s+/g, ' ').trim();
    studentLabelCache.set(element, { raw, name });
    return name;
  }
  function studentName(input) {
    const row = input.closest('tr') || input.closest('[data-region="user"], .row, li');
    if (!row) return '';
    for (const link of row.querySelectorAll('a[href*="/user/view.php"], a[href*="/user/profile.php"]')) {
      const name = studentLabel(link);
      if (name) return name;
    }
    return studentLabel(row.querySelector('th'));
  }
  function gradeValue(el) {
    if (el.tagName === 'SELECT') {
      const o = el.options[el.selectedIndex];
      return o && el.value !== '-1' && el.value !== '' ? o.textContent.trim() : '';
    }
    const v = (el.value || '').trim();
    return v === '-' ? '' : v;
  }
  function pageRows(key) {
    return (scanPage().get(key) || []).map((el) => {
      const raw = gradeValue(el);
      return { el, name: studentName(el), raw, value: toNumber(raw) };
    });
  }

  function writeField(el, raw) {
    if (el.tagName === 'SELECT') el.value = raw;
    else Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(el, raw);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('blur', { bubbles: true }));
  }
  function setValue(el, value) {
    if (el.tagName !== 'SELECT') {
      writeField(el, value);
      return true;
    }
    const opt = [...el.options].find((o) => o.value === value || norm(o.textContent) === norm(value));
    if (!opt) return false;
    writeField(el, opt.value);
    return true;
  }

  const originals = new Map();
  function remember(el, raw, approx) {
    const saved = originals.get(el);
    if (saved) Object.assign(saved, { written: el.value, approx });
    else originals.set(el, { raw, before: '', written: el.value, approx });
  }
  function changedFields(fields) {
    return fields.filter((el) => {
      const saved = originals.get(el);
      if (!saved) return false;
      if (el.isConnected && el.value !== saved.raw) return true;
      originals.delete(el);
      mark(el, '');
      return false;
    });
  }
  function undoChanges(fields) {
    let restored = 0;
    changedFields(fields).forEach((el) => {
      const saved = originals.get(el);
      originals.delete(el);
      mark(el, '');
      if (el.value !== saved.written) return;
      writeField(el, saved.raw);
      restored++;
    });
    return restored;
  }
  function mark(el, color) {
    const cell = el.closest('td,th') || el;
    cell.style.outline = color ? `3px solid ${color}` : '';
    cell.style.outlineOffset = color ? '-3px' : '';
    cell.dataset.studiumNotes = color ? '1' : '';
  }

  function fill(text, key, overwrite, choices = new Map()) {
    const { entries, bad } = parseList(text);
    const inputs = scanPage().get(key) || [];
    const rows = inputs.map((el) => ({ el, toks: new Set(tokens(rowNameText(el))), nameToks: tokens(studentName(el)) }));
    inputs.forEach((el) => { if (!originals.has(el)) mark(el, ''); });

    const report = { ok: [], approx: [], notFound: [], ambiguous: [], skipped: [], invalid: [], bad, choices: [], changes: [] };
    const used = new Set();
    entries.forEach((e) => {
      const want = tokens(e.name);
      let approx = false;
      let matches = rows.filter((r) => want.every((t) => r.toks.has(t)));
      if (matches.length === 0 && want.length > 1) {
        matches = rows.filter((r) => want.filter((t) => r.toks.has(t)).length >= want.length - 1);
        if (matches.length === 1) approx = true;
      }
      if (matches.length === 0 && want.some((t) => t.length >= 3)) {
        const pre = rows.filter((r) => want.every((t) => r.nameToks.some((n) => n.startsWith(t))));
        if (pre.length) {
          matches = pre;
          if (pre.length === 1) approx = true;
        }
      }
      if (matches.length > 1) {
        const exact = matches.filter((r) => r.nameToks.length === want.length && want.every((t) => r.nameToks.includes(t)));
        if (exact.length === 1) matches = exact;
      }
      if (matches.length === 0 && want.length) {
        const scored = rows.map((r) => ({ r, s: nameScore(want, r.nameToks) })).sort((a, b) => b.s - a.s);
        matches = scored.filter((x) => x.s >= 0.72).map((x) => x.r);
        if (matches.length === 1) approx = true;
      }
      const choiceKey = key + '\u0000' + norm(e.name);
      const chosen = choices.get(choiceKey);
      if (chosen) {
        const selected = matches.find((r) => r.el.name === chosen);
        if (selected) { matches = [selected]; approx = false; }
      }
      if (matches.length === 0) return report.notFound.push(e.line);
      if (matches.length > 1) {
        report.choices.push({ key: choiceKey, line: e.line, candidates: matches.map((r) => ({
          id: r.el.name, name: studentName(r.el),
          detail: (rowNameText(r.el).match(EMAIL_RE) || [])[0] || r.el.name.match(/\d+/)?.[0] || ''
        })) });
        return;
      }
      const { el } = matches[0];
      if (used.has(el)) return report.ambiguous.push(t('ambigUsed', e.line));
      used.add(el);
      const before = gradeValue(el);
      if (!overwrite && before !== '') {
        mark(el, '#f0ad4e');
        return report.skipped.push(t('skippedLine', e.line, before));
      }
      const name = studentName(el);
      if (toNumber(before) !== toNumber(e.grade)) {
        const raw = el.value;
        if (!setValue(el, e.grade)) {
          mark(el, '#d9534f');
          return report.invalid.push(t('invalidLine', e.line));
        }
        if (!originals.has(el)) {
          remember(el, raw, approx);
          originals.get(el).before = before;
        } else remember(el, raw, approx);
      }
      if (approx) {
        mark(el, '#1e90ff');
        return report.approx.push(t('approxLine', e.name, name, e.grade.replace('.', ',')));
      }
      mark(el, '#5cb85c');
      report.ok.push(`${e.line} → ${name}`);
    });
    report.changes = changedFields(inputs).map((el) => {
      const saved = originals.get(el);
      return { name: studentName(el), before: saved.before, after: gradeValue(el), approx: saved.approx };
    });
    return { report, total: rows.length, key, text };
  }

  const CSS = `
  :host{all:initial}
  .sn{--bg:#ffffff;--bg2:#f4f6fa;--bg3:#eaeef5;--fg:#18202e;--muted:#5d6678;--line:#e1e6ee;
      --accent:#0057ac;--accent-2:#0a6fd6;--accent-fg:#fff;--accent-soft:#e7f0fb;
      --ok:#1d8048;--ok-soft:#e6f4ec;--warn:#a86200;--warn-soft:#fdf2e1;--bad:#c0392b;--bad-soft:#fbe9e7;
      --info:#1e6fd9;--info-soft:#e6effc;--mean:#d97706;
      --shadow:0 18px 50px rgba(15,23,42,.22),0 3px 10px rgba(15,23,42,.10);
      font:15.5px/1.45 system-ui,-apple-system,"Segoe UI",Roboto,"Helvetica Neue",Arial,sans-serif;color:var(--fg);
      -webkit-font-smoothing:antialiased}
  @media (prefers-color-scheme:dark){.sn{--bg:#161a22;--bg2:#1e2430;--bg3:#283040;--fg:#e9edf4;--muted:#9aa4b6;--line:#2d3545;
      --accent:#4d8eff;--accent-2:#77a9ff;--accent-fg:#0b1220;--accent-soft:#1c2a45;--ok:#4cc483;--ok-soft:#17301f;
      --warn:#f0b04a;--warn-soft:#352a14;--bad:#ff7a6b;--bad-soft:#3a1e1b;--info:#6fa8ff;--info-soft:#1b2944;--mean:#f5a524;
      --shadow:0 18px 50px rgba(0,0,0,.55),0 3px 10px rgba(0,0,0,.35)}}
  *{box-sizing:border-box}
  [hidden]{display:none !important}
  button,select,input,textarea{font:inherit;color:inherit}
  .panel{position:fixed;z-index:2147483000;display:flex;flex-direction:column;background:var(--bg);border:1px solid var(--line);
      border-radius:14px;box-shadow:none;overflow:hidden;resize:horizontal;min-width:340px;min-height:0;
      max-width:calc(100vw - 8px);max-height:calc(100vh - 8px)}
  .head{display:flex;align-items:center;gap:8px;padding:10px 10px 10px 14px;background:var(--accent);color:#fff;cursor:grab;
      user-select:none;touch-action:none;flex:none}
  .head:active{cursor:grabbing}
  .head .title{flex:1;font-weight:650;font-size:16px;letter-spacing:.1px;display:flex;align-items:center;gap:8px}
  .head .grip{opacity:.7;font-size:14px}
  .iconbtn{width:28px;height:28px;border-radius:8px;border:0;background:rgba(255,255,255,.16);color:#fff;cursor:pointer;
      display:grid;place-items:center;font-size:17px;line-height:1}
  .iconbtn:hover{background:rgba(255,255,255,.28)}
  .summary{padding:12px 14px 10px;border-bottom:1px solid var(--line);background:var(--bg2);flex:none}
  .colrow{display:flex;gap:6px;align-items:center}
  .colrow select{flex:1;min-width:0}
  .counts{display:flex;gap:6px;margin-top:10px}
  .count{flex:1;background:var(--bg);border:1px solid var(--line);border-radius:10px;padding:6px 8px;text-align:center;min-width:0}
  .count b{display:block;font-size:21px;font-weight:700;line-height:1.15;font-variant-numeric:tabular-nums}
  .count span{font-size:13.5px;color:var(--muted)}
  .count.rest{cursor:pointer}
  .count.rest:hover{border-color:var(--accent)}
  .count.rest b{color:var(--warn)}
  .count.done b{color:var(--ok)}
  .bar{height:6px;border-radius:99px;background:var(--bg3);margin-top:8px;overflow:hidden}
  .bar i{display:block;height:100%;background:var(--ok);border-radius:99px;transition:none}
  .restlist{margin-top:8px;max-height:120px;overflow:auto;font-size:14.5px;background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:6px 10px}
  .tabs{display:flex;gap:2px;padding:6px 6px 0;border-bottom:1px solid var(--line);flex:none}
  .tab{flex:1;min-width:0;border:0;background:none;padding:5px 2px 7px;border-bottom:2.5px solid transparent;cursor:pointer;color:var(--muted);
      font-weight:600;font-size:13.5px;display:flex;flex-direction:column;align-items:center;gap:1px;border-radius:8px 8px 0 0}
  .tab span{font-size:19px;line-height:1.2}
  .tab:hover{color:var(--fg);background:var(--bg2)}
  .tab[aria-selected="true"]{color:var(--accent);border-bottom-color:var(--accent)}
  .hello{margin:2px 0 12px;font-size:19px}
  .menu{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px}
  .mcard{display:flex;align-items:center;gap:8px;text-align:left;padding:10px;border:1px solid var(--line);border-radius:12px;
      background:var(--bg);cursor:pointer;transition:none}
  .mcard:hover{border-color:var(--accent);background:var(--accent-soft);transform:none}
  .mcard .mi{flex:none;width:28px;height:28px;border-radius:11px;background:var(--accent-soft);display:grid;place-items:center;font-size:23px}
  .mcard b{display:block;font-size:16px}
  .mcard small{display:block;color:var(--muted);font-size:14px;margin-top:1px}
  .steps{margin-top:14px;padding:12px 14px;border-radius:12px;background:var(--bg2);font-size:14.5px}
  .steps ol{margin:6px 0 0;padding-left:20px;display:grid;gap:3px}
  .brief{background:var(--accent-soft);border-radius:12px;padding:12px 14px;font-size:15.5px;line-height:1.55}
  .brief h3{margin:0 0 4px;font-size:15px;color:var(--accent);text-transform:uppercase;letter-spacing:.5px}
  .brief p{margin:0}
  .brief b{font-variant-numeric:tabular-nums}
  .group{margin-top:14px}
  .group > h4{margin:0 0 6px;font-size:14px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
  .tiles.two{grid-template-columns:repeat(2,minmax(0,1fr))}
  .tile.hl{background:var(--accent);color:var(--accent-fg)}
  .tile.hl span,.tile.hl small{color:inherit;opacity:.85}
  details.settings{margin-top:10px;font-size:14.5px;color:var(--muted)}
  details.settings summary{cursor:pointer;user-select:none}
  details.settings .ctrl{margin:8px 0 0}
  .cap{font-size:14px;color:var(--muted);margin:-2px 0 8px}
  .bandleg{display:grid;gap:4px;margin-top:6px;font-size:14.5px}
  .bandleg i{display:inline-block;width:12px;height:12px;border-radius:3px;vertical-align:-1px;margin-right:6px}
  svg .track{fill:var(--bg3)}
  svg .b80{fill:var(--accent-soft);stroke:var(--accent);stroke-opacity:.35}
  svg .b50{fill:var(--accent)}
  svg .tickv{fill:var(--fg);font-weight:600}
  svg .mlab{fill:var(--mean);font-weight:700}
  svg .dlab{fill:var(--fg);font-weight:700}
  .body{flex:1;overflow:auto;overflow-x:hidden;padding:14px}
  .page[hidden]{display:none}
  .muted{color:var(--muted)}
  .small{font-size:14px}
  select,input[type=number]{background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:6px 8px;height:34px}
  select:focus,input:focus,textarea:focus{outline:2px solid var(--accent-2);outline-offset:0;border-color:transparent}
  textarea{width:100%;min-height:90px;resize:vertical;background:var(--bg);border:1px solid var(--line);border-radius:10px;
      padding:10px;font:14.5px/1.5 ui-monospace,SFMono-Regular,Consolas,monospace}
  .btn{display:inline-flex;align-items:center;justify-content:center;gap:6px;height:34px;padding:0 12px;border-radius:9px;
      border:1px solid var(--line);background:var(--bg);cursor:pointer;font-weight:600;white-space:nowrap}
  .btn:hover{border-color:var(--accent);color:var(--accent)}
  .btn.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-fg)}
  .btn.primary:hover{filter:none;color:var(--accent-fg)}
  .btn.big{height:40px;width:100%;font-size:16px}
  .btn.sq{width:34px;padding:0}
  .tools{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
  .hint{margin:0 0 8px;font-size:14.5px;color:var(--muted)}
  .hint code{background:var(--bg2);border-radius:5px;padding:1px 5px}
  .check{display:flex;align-items:center;gap:8px;margin:10px 0;cursor:pointer;font-size:15px}
  .check input{width:16px;height:16px;accent-color:var(--accent)}
  .notice{margin-top:10px;padding:8px 10px;border-radius:8px;background:var(--info-soft);color:var(--info);font-size:14.5px}
  .report{margin-top:12px;display:grid;gap:8px}
  .candidate-list{display:flex;flex-wrap:wrap;gap:6px;margin-top:6px}
  .rep{border-radius:9px;padding:8px 10px;border-left:4px solid;font-size:14.5px}
  .rep b{display:block;margin-bottom:2px}
  .rep div{white-space:pre-wrap;word-break:break-word}
  .rep.ok{background:var(--ok-soft);border-color:var(--ok)} .rep.ok b{color:var(--ok)}
  .rep.info{background:var(--info-soft);border-color:var(--info)} .rep.info b{color:var(--info)}
  .rep.warn{background:var(--warn-soft);border-color:var(--warn)} .rep.warn b{color:var(--warn)}
  .rep.bad{background:var(--bad-soft);border-color:var(--bad)} .rep.bad b{color:var(--bad)}
  .rep.verify{background:var(--bg2);border-color:var(--accent)} .rep.verify summary{color:var(--accent);font-weight:700;cursor:pointer}
  .rep.verify p{margin:6px 0 0;font-size:13.5px;color:var(--muted)}
  table.changes{width:100%;border-collapse:collapse;margin-top:4px;font-size:14px}
  table.changes th,table.changes td{padding:3px 6px;border-bottom:1px solid var(--line);text-align:left}
  table.changes thead th{position:static;background:none;color:var(--muted);font-weight:600}
  table.changes tr.approx th,table.changes tr.approx td{background:var(--info-soft);color:var(--info)}
  .fill-row{display:flex;gap:8px;align-items:stretch}
  .fill-row #fill{flex:1}
  .tiles{display:grid;grid-template-columns:repeat(auto-fill,minmax(118px,1fr));gap:8px}
  .tile{background:var(--bg2);border-radius:10px;padding:9px 10px;min-width:0}
  .tile span{display:block;font-size:13.5px;color:var(--muted)}
  .tile b{display:block;font-size:20px;font-weight:700;font-variant-numeric:tabular-nums;line-height:1.25;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
  .tile small{display:block;font-size:13.5px;color:var(--muted);line-height:1.3;margin-top:2px}
  .card{border:1px solid var(--line);border-radius:12px;padding:12px;margin-top:12px}
  .card h3{margin:0 0 4px;font-size:16px}
  .card p{margin:0 0 10px}
  .ctrl{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:8px;font-size:14.5px;color:var(--muted)}
  .ctrl label{display:flex;align-items:center;gap:6px}
  .ctrl input{width:76px}
  .conc{font-size:15px;margin:4px 0 10px}
  .conc b{color:var(--accent)}
  svg{display:block;width:100%;height:auto;overflow:visible}
  svg text{fill:var(--muted);font:12.5px system-ui,-apple-system,"Segoe UI",sans-serif}
  svg .lbl{fill:var(--fg);font-weight:600}
  svg .barr{fill:var(--accent)}
  svg .barr.top{fill:var(--accent-2)}
  svg .barr:hover{opacity:.8}
  svg .grid{stroke:var(--line)}
  svg .axis{stroke:var(--muted)}
  svg .mean{stroke:var(--mean);stroke-width:2;stroke-dasharray:5 4}
  svg .median{stroke:var(--fg);stroke-width:2}
  svg .box{fill:var(--accent-soft);stroke:var(--accent);stroke-width:1.5}
  svg .whisk{stroke:var(--accent);stroke-width:1.5}
  svg .out{fill:var(--bad)}
  .legend{display:flex;flex-wrap:wrap;gap:12px;font-size:13.5px;color:var(--muted);margin-top:6px}
  .legend i{display:inline-block;width:16px;height:0;border-top:2px solid;vertical-align:middle;margin-right:5px}
  .empty{padding:24px 10px;text-align:center;color:var(--muted)}
  .drop{display:block;border:2px dashed var(--line);border-radius:12px;padding:18px 12px;text-align:center;cursor:pointer;background:var(--bg2)}
  .drop:hover,.drop.over{border-color:var(--accent);background:var(--accent-soft)}
  .drop b{color:var(--accent)}
  .maprow{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin:10px 0}
  .maprow label{font-size:13.5px;color:var(--muted);display:grid;gap:3px}
  .maprow select{width:100%}
  .prev{overflow:auto;max-height:170px;border:1px solid var(--line);border-radius:8px}
  table{border-collapse:collapse;width:100%;font-size:14px}
  th,td{padding:4px 8px;border-bottom:1px solid var(--line);text-align:left;white-space:nowrap}
  th{background:var(--bg2);position:sticky;top:0}
  td.on,th.on{background:var(--accent-soft)}
  .pill{position:fixed;z-index:2147483000;display:flex;align-items:center;gap:8px;padding:9px 14px 9px 12px;border-radius:99px;
      background:var(--accent);color:#fff;box-shadow:none;cursor:pointer;font-weight:650;user-select:none;touch-action:none;border:0}
  .pill:hover{filter:none}
  .pill .pc{background:rgba(255,255,255,.2);border-radius:99px;padding:1px 8px;font-size:14px;font-variant-numeric:tabular-nums}
  .flag{width:20px;height:13px;border-radius:2px;box-shadow:0 0 0 1px rgba(0,0,0,.18);flex:none;display:inline-block;vertical-align:middle}
  .langwrap{position:relative}
  .langbtn{height:28px;display:flex;align-items:center;gap:5px;padding:0 8px;border:0;border-radius:8px;background:rgba(255,255,255,.16);
      color:#fff;cursor:pointer;font-weight:600;font-size:14px}
  .langbtn:hover{background:rgba(255,255,255,.28)}
  .langmenu{position:absolute;right:0;top:34px;z-index:5;min-width:150px;background:var(--bg);color:var(--fg);border:1px solid var(--line);
      border-radius:10px;box-shadow:none;padding:4px;cursor:default}
  .langmenu button{display:flex;align-items:center;gap:10px;width:100%;padding:8px 10px;border:0;background:none;border-radius:7px;cursor:pointer;text-align:left;font-size:15px}
  .langmenu button:hover{background:var(--bg2)}
  .langmenu button[aria-checked="true"]{font-weight:700;color:var(--accent)}
  .langmenu button[aria-checked="true"]::after{content:"✓";margin-left:auto}
  .tab i{font-style:normal}

  .progress>summary{font-weight:650;cursor:pointer;list-style:revert;padding:2px 0}
  .progress-controls{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:12px 0}
  .progress-controls label{display:flex;flex-direction:column;gap:5px;min-width:0;font-size:14px}
  .progress-controls select,.progress-controls input{width:100%;min-width:0;padding:7px;border:1px solid var(--line);border-radius:7px;background:var(--bg);color:var(--fg);font:inherit}
  .progress-table{max-height:330px;overflow:auto;border:1px solid var(--line);border-radius:8px}
  .progress table{width:100%;border-collapse:collapse;font-size:13px}
  .progress th,.progress td{padding:7px 5px;border-bottom:1px solid var(--line);text-align:right;overflow-wrap:anywhere}
  .progress th:first-child{text-align:left;font-weight:500}
  .progress thead th{position:sticky;top:0;background:var(--bg2);font-weight:650}
  .progress td{white-space:nowrap}.progress td small{display:block;color:var(--muted)}

  .panel{container-type:inline-size;contain:layout paint style;isolation:isolate;transform:translateZ(0)}
  @container (max-width:480px){.tab span{display:none}}
  .head,.summary,.tabs{flex-shrink:0}
  .body{min-height:0;overscroll-behavior:contain}
  .summary{padding:8px 14px}
  .summary>.small{display:none}
  .counts{margin-top:6px;gap:8px}
  .count{padding:3px 6px;border:0;background:none;display:flex;align-items:baseline;justify-content:center;gap:6px}
  .count b{font-size:18px}.bar{height:3px;margin-top:3px}
  .tab{flex-direction:row;align-items:center;justify-content:center;gap:7px;padding:8px 2px}
  .stats-mode .body{padding:10px 14px}
  .stats-nav{position:sticky;top:-10px;z-index:3;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:5px;background:var(--bg);padding:0 0 10px}
  .stats-nav button{border:1px solid var(--line);background:var(--bg2);color:var(--fg);padding:8px 5px;border-radius:8px;font:600 14px system-ui;cursor:pointer}
  .stats-nav button[aria-pressed=true]{background:var(--accent);color:#fff;border-color:var(--accent)}
  .stats-mode .brief{font-size:15px;padding:9px 12px;line-height:1.4}
  .stats-mode .group{margin-top:10px}.stats-mode h4{margin:8px 0}
  .stats-mode .tile{padding:6px 8px}.stats-mode .tile b{font-size:18px}.stats-mode .tile small{display:none}
  .charts-grid>.card{margin-top:0;min-width:0}
  .progress-controls{margin:0 0 8px;gap:14px}
  .assessment-pair{display:grid;grid-template-columns:minmax(0,1fr) 76px;gap:8px}
  .progress-table{max-height:none;overflow:visible;border:1px solid var(--line)}
  .progress th,.progress td{padding:7px;font-size:14px}
  .progress thead th{position:static}.progress table{table-layout:fixed}
  .progress th:first-child{width:34%}.progress td{white-space:normal}
  .coverage{font-size:14px;color:var(--muted);margin:8px 0}
  .overview-tiles{grid-template-columns:repeat(3,minmax(0,1fr))}.comparison-tiles{grid-template-columns:repeat(2,minmax(0,1fr))}
  .letter-counts{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px;margin:8px 0 0;padding:8px 10px 6px;background:var(--bg2);border:1px solid var(--line);border-radius:10px;line-height:1.2}
  .lc-bands{display:grid;grid-template-columns:auto repeat(8,minmax(0,1fr));gap:3px;align-items:end}
  .lc{display:flex;flex-direction:column;align-items:center;gap:2px;min-width:0;cursor:default}
  .lc b{font-size:14px;font-variant-numeric:tabular-nums}
  .lc-bar{display:flex;align-items:flex-end;width:100%;height:26px;border-bottom:1px solid var(--line)}
  .lc-bar i{display:block;width:100%;max-width:26px;margin:0 auto;background:var(--accent);opacity:.8;border-radius:3px 3px 0 0}
  .lc-l{font-size:12px;color:var(--muted);white-space:nowrap}
  .lc.fail b,.lc.fail .lc-l{color:var(--warn)}.lc.fail .lc-bar i{background:var(--warn)}
  .lc.zero b{color:var(--muted);font-weight:500}
  .lc-total{display:flex;flex-direction:column;align-items:center;justify-content:center;min-width:64px;padding:4px 8px;border-radius:8px;background:var(--warn-soft);color:var(--warn);cursor:default}
  .lc-total .lc-l{color:inherit;font-weight:650}.lc-total b{font-size:22px;line-height:1.1;font-variant-numeric:tabular-nums}.lc-total small{font-size:12px}
  .letter-counts.unscaled .lc-total{background:var(--bg3);color:var(--muted)}
  .lc-scale{grid-column:1/-1;display:flex;flex-wrap:wrap;align-items:center;gap:4px 6px;margin:0;font-size:13px;color:var(--muted)}
  .letter-counts input{width:64px;padding:2px 4px;font:inherit;color:var(--fg);background:var(--bg);border:1px solid var(--line);border-radius:4px}
  .trend-counts{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:8px;margin-top:12px}
  .student-search input{width:100%;padding:8px 10px;background:var(--bg);color:var(--fg);border:1px solid var(--line);border-radius:8px;font:15px system-ui}
  .pagination{display:flex;justify-content:space-between;align-items:center;gap:8px;margin-top:8px;font-size:14px}
  .method{margin-top:10px;font-size:14px;color:var(--muted)}.method summary{cursor:pointer}
  @container (min-width:650px){
    .stats-mode .tiles.two{grid-template-columns:repeat(4,minmax(0,1fr))}
    .stats-mode .charts-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
    .stats-mode .overview-tiles,.stats-mode .comparison-tiles{grid-template-columns:repeat(4,minmax(0,1fr))}
  }
  @container (max-width:500px){
    .stats-nav{grid-template-columns:repeat(2,minmax(0,1fr))}
    .progress-controls{grid-template-columns:1fr;gap:8px}
    .stats-mode .tile b{font-size:20px}.stats-mode .tab{gap:3px;font-size:13px}
    .pagination{flex-wrap:wrap}.progress th,.progress td{padding:6px 4px;font-size:13px}
  }
  details.about{margin-top:16px;font-size:13.5px;color:var(--muted)}
  details.about summary{cursor:pointer;user-select:none;width:max-content;opacity:.8}
  details.about p{margin:6px 0 0;line-height:1.45}
  `;

  const HTML = `
  <div class="sn">
    <button class="pill" id="pill" data-it="pillTitle" hidden>📝 Notes <span class="pc" id="pill-count"></span></button>
    <section class="panel" id="panel" role="dialog" aria-label="Notes StudiUM">
      <header class="head" id="head" data-it="headTitle">
        <div class="title"><span class="grip">⠿</span>📝 Notes StudiUM</div>
        <div class="langwrap">
          <button class="langbtn" id="langbtn" data-it="langTitle" aria-haspopup="menu" aria-expanded="false"></button>
          <div class="langmenu" id="langmenu" role="menu" hidden>
            <button role="menuitemradio" data-lang="fr">${flagSvg('fr')}Français</button>
            <button role="menuitemradio" data-lang="en">${flagSvg('en')}English</button>
          </div>
        </div>
        <button class="iconbtn" id="min" data-it="minTitle">−</button>
        <button class="iconbtn" id="close" data-it="closeTitle">×</button>
      </header>
      <div class="summary">
        <div class="small muted" style="margin:0 0 4px" data-i="evalLabel"></div>
        <div class="colrow">
          <select id="col" data-it="colTitle"></select>
          <button class="btn sq" id="rescan" data-it="rescanTitle">⟳</button>
        </div>
        <div class="counts">
          <div class="count"><b id="c-total">0</b><span data-i="cTotal"></span></div>
          <div class="count done"><b id="c-done">0</b><span data-i="cDone"></span></div>
          <div class="count rest" id="c-rest-box" data-it="cRestTitle"><b id="c-rest">0</b><span data-i="cRest"></span></div>
        </div>
        <div class="bar"><i id="c-bar" style="width:0"></i></div>
        <div class="restlist" id="restlist" hidden></div>
      </div>
      <nav class="tabs" role="tablist">
        <button class="tab" role="tab" data-tab="accueil" aria-selected="true"><span>🏠</span><i data-i="tabHome"></i></button>
        <button class="tab" role="tab" data-tab="saisie" aria-selected="false"><span>✍️</span><i data-i="tabEnter"></i></button>
        <button class="tab" role="tab" data-tab="stats" aria-selected="false"><span>📊</span><i data-i="tabStats"></i></button>
        <button class="tab" role="tab" data-tab="export" aria-selected="false"><span>⬇️</span><i data-i="tabExport"></i></button>
        <button class="tab" role="tab" data-tab="import" aria-selected="false"><span>⬆️</span><i data-i="tabImport"></i></button>
      </nav>
      <div class="body">
        <div id="no-fields" class="notice" data-i="noFieldsHelp" hidden></div>
        <div class="page" data-page="accueil">
          <h2 class="hello" data-i="hello"></h2>
          <div class="menu">
            <button class="mcard" data-go="saisie"><span class="mi">✍️</span><span><b data-i="m1t"></b><small data-i="m1d"></small></span></button>
            <button class="mcard" data-go="stats"><span class="mi">📊</span><span><b data-i="m2t"></b><small data-i="m2d"></small></span></button>
            <button class="mcard" data-go="export"><span class="mi">⬇️</span><span><b data-i="m3t"></b><small data-i="m3d"></small></span></button>
            <button class="mcard" data-go="import"><span class="mi">⬆️</span><span><b data-i="m4t"></b><small data-i="m4d"></small></span></button>
          </div>
          <details class="about">
            <summary data-i="aboutTitle"></summary>
            <p data-i="aboutText"></p>
            <p data-i="aboutTech"></p>
            <p data-i="aboutPrivacy"></p>
            <p data-i="aboutCheck"></p>
            <p data-i="aboutUsage"></p>
            <p data-i="aboutBrowser"></p>
          </details>
        </div>
        <div class="page" data-page="saisie" hidden>
          <p class="hint" data-i="hint"></p>
          <div class="tools">
            <button class="btn" id="paste" data-i="paste"></button>
            <button class="btn" id="clear" data-i="clear"></button>
          </div>
          <textarea id="text" spellcheck="false" data-ip="placeholder"></textarea>
          <div class="report" id="report" role="status" aria-live="polite"></div>
          <label class="check"><input type="checkbox" id="over" checked> <span data-i="over"></span></label>
          <div class="fill-row">
            <button class="btn primary big" id="fill" data-i="fillBtn"></button>
            <button class="btn" id="undo" data-i="undoBtn" disabled></button>
          </div>
          <div class="notice" data-i="notice"></div>
        </div>
        <div class="page" data-page="stats" hidden>
          <nav class="stats-nav" aria-label="Statistics">
            <button data-st-view="overview" data-i="viewOverview" aria-pressed="true"></button>
            <button data-st-view="charts" data-i="viewCharts" aria-pressed="false"></button>
            <button data-st-view="class" data-i="viewClass" aria-pressed="false"></button>
            <button data-st-view="students" data-i="viewStudents" aria-pressed="false"></button>
          </nav><div id="st-progress" hidden></div><div id="st-body"></div>
        </div>
        <div class="page" data-page="export" hidden>
          <div class="card" style="margin-top:0">
            <h3 data-i="expTitle"></h3>
            <p class="muted small" data-i="expDesc"></p>
            <button class="btn primary big" id="x-export" data-i="expBtn"></button>
            <p class="muted small" style="margin:8px 0 0" data-i="expWhere"></p>
          </div>
        </div>
        <div class="page" data-page="import" hidden>
          <div class="card" style="margin-top:0">
            <h3 data-i="impTitle"></h3>
            <p class="muted small" data-i="impDesc"></p>
            <label class="drop" id="x-drop">
              <input type="file" id="x-file" accept=".xlsx,.csv,.tsv,.txt" hidden>
              📂 <b data-i="dropMain"></b><br><span class="muted small" data-i="dropSub"></span>
            </label>
            <div id="x-map"></div>
          </div>
        </div>
      </div>
    </section>
  </div>`;

  function buildUI() {
    let destroyed = false;
    const startupTimers = [];
    const host = document.createElement('div');
    host.id = 'studium-notes-host';
    const root = host.attachShadow({ mode: 'open' });
    root.innerHTML = `<style>${CSS}</style>${HTML}`;
    document.body.appendChild(host);
    const $ = (id) => root.getElementById(id);
    const panel = $('panel'), pill = $('pill');

    ['click', 'dblclick', 'mousedown', 'mouseup', 'pointerdown', 'pointerup',
      'keydown', 'keyup', 'keypress', 'paste', 'copy', 'cut', 'input', 'change'].forEach((t) =>
      host.addEventListener(t, (e) => e.stopPropagation())
    );

    let ui = { left: null, top: null, width: 420, height: 480 };
    try { Object.assign(ui, JSON.parse(localStorage.getItem(UI_KEY) || '{}')); } catch (_) {}
    ui.width = Math.min(Number(ui.width) || 420, 440);
    ui.height = Math.min(Number(ui.height) || 480, 500);
    ui.top = Math.max(8, Math.min(Number(ui.top) || 16, 40));
    const saveUi = () => { try { localStorage.setItem(UI_KEY, JSON.stringify(ui)); } catch (_) {} };
    function applyPanel() {
      const W = document.documentElement.clientWidth || window.innerWidth, H = window.innerHeight;
      const w = Math.min(Math.max(ui.width, 360), W - 16);
      if (ui.left === null) ui.left = W - w - 20;
      const left = Math.max(4, Math.min(ui.left, W - w - 4)), top = Math.max(8, Math.min(ui.top, H - Math.min(300, H - 16) - 8));
      Object.assign(panel.style, { left: left + 'px', top: top + 'px', width: w + 'px', height: 'auto', maxHeight: (H - top - 8) + 'px' });
      Object.assign(pill.style, { left: Math.max(4, Math.min(ui.left + w - 150, W - 160)) + 'px', top: Math.max(4, Math.min(ui.top, H - 48)) + 'px' });
    }
    function draggable(handle, onClick) {
      handle.addEventListener('pointerdown', (ev) => {
        if (ev.button !== 0 || ev.target.closest('.langwrap')) return;
        if (ev.target !== handle && ev.target.closest('button,select,input') && ev.target.closest('button') !== handle) return;
        const startX = ev.clientX, startY = ev.clientY, startL = ui.left, startT = ui.top;
        let moved = false;
        ev.preventDefault();
        handle.setPointerCapture(ev.pointerId);
        const move = (e) => {
          if (!moved && Math.hypot(e.clientX - startX, e.clientY - startY) < 5) return;
          moved = true;
          ui.left = startL + e.clientX - startX;
          ui.top = startT + e.clientY - startY;
          applyPanel();
        };
        const up = () => {
          handle.removeEventListener('pointermove', move);
          handle.removeEventListener('pointerup', up);
          handle.removeEventListener('pointercancel', up);
          if (moved) {
            ui.left = parseFloat(panel.style.left) || ui.left;
            ui.top = parseFloat(panel.style.top) || ui.top;
            saveUi();
          } else if (onClick) onClick();
        };
        handle.addEventListener('pointermove', move);
        handle.addEventListener('pointerup', up);
        handle.addEventListener('pointercancel', up);
      });
    }
    let userTouched = false;
    function setOpen(open) {
      panel.hidden = !open;
      pill.hidden = open;
      applyPanel();
    }
    draggable($('head'));
    draggable(pill, () => { userTouched = true; setOpen(true); refreshColumns(); });
    $('min').onclick = () => { userTouched = true; clearTimeout(autoTimer); setOpen(false); };
    $('close').onclick = () => {
      userTouched = true;
      clearTimeout(autoTimer);
      clearTimeout(updTimer);
      forgetData();
      originals.clear();
      panel.hidden = true;
      pill.hidden = true;
    };
    window.addEventListener('resize', applyPanel);
    panel.addEventListener('pointerup', () => {
      if (panel.hidden) return;
      const r = panel.getBoundingClientRect();
      if (r.width && Math.round(r.width) !== ui.width) {
        ui.width = Math.round(r.width); saveUi();
      }
    });
    applyPanel();

    let currentTab = 'accueil';
    function showTab(name) {
      currentTab = name;
      panel.classList.toggle('stats-mode', name === 'stats');
      applyPanel();
      root.querySelectorAll('.tab').forEach((t) => t.setAttribute('aria-selected', String(t.dataset.tab === name)));
      root.querySelectorAll('.page').forEach((p) => { p.hidden = p.dataset.page !== name; });
      root.querySelector('.body').scrollTop = 0;
      if (name === 'stats') renderStats();
      if (name === 'saisie') setTimeout(() => $('text').focus(), 0);
    }
    root.querySelectorAll('.tab').forEach((t) => { t.onclick = () => showTab(t.dataset.tab); });
    root.querySelectorAll('[data-go]').forEach((b) => { b.onclick = () => showTab(b.dataset.go); });

    const sel = $('col');
    function refreshColumns() {
      const cols = scanPage();
      const prev = sel.value || legacyColumn;
      sel.innerHTML = '';
      if (cols.size === 0) {
        sel.innerHTML = `<option value="">${escapeHtml(t('noFields'))}</option>`;
      } else {
        cols.forEach((list, key) => {
          const o = document.createElement('option');
          o.value = key;
          o.textContent = t('colOption', columnLabel(key, list[0]), list.length);
          sel.appendChild(o);
        });
        if (prev && cols.has(prev)) sel.value = prev;
      }
      $('no-fields').hidden = cols.size > 0;
      update();
      return cols.size;
    }
    sel.onchange = update;
    $('rescan').onclick = refreshColumns;

    function update() {
      const rows = sel.value ? pageRows(sel.value) : [];
      const done = rows.filter((r) => r.raw !== '').length;
      const rest = rows.length - done;
      $('c-total').textContent = rows.length;
      $('c-done').textContent = done;
      $('c-rest').textContent = rest;
      $('c-bar').style.width = rows.length ? (100 * done) / rows.length + '%' : '0';
      $('pill-count').textContent = rows.length ? `${done}/${rows.length}` : '';
      const missing = rows.filter((r) => r.raw === '');
      $('restlist').innerHTML = missing.length
        ? missing.map((r) => escapeHtml(r.name || t('noName'))).join('<br>')
        : `<span class="muted">${t('allDone')}</span>`;
      $('undo').disabled = !changedFields(rows.map((r) => r.el)).length;
      if (currentTab === 'stats') renderStats(rows);
    }
    $('c-rest-box').onclick = () => { $('restlist').hidden = !$('restlist').hidden; };
    let updTimer;
    const scheduleUpdate = () => { clearTimeout(updTimer); updTimer = setTimeout(() => { if (!destroyed && !panel.hidden) update(); }, 200); };
    const onPageGradeChange = (e) => {
      if (!destroyed && !panel.hidden && e.target?.matches?.(INPUT_SELECTOR)) scheduleUpdate();
    };
    ['input', 'change'].forEach((t) => document.addEventListener(t, onPageGradeChange, true));

    const ta = $('text');
    $('paste').onclick = async () => {
      try {
        ta.value = await navigator.clipboard.readText();
        ta.dispatchEvent(new Event('input'));
      } catch (_) {
        ta.focus();
        alert(t('alertClipboard'));
      }
    };
    function forgetData() {
      studentChoices.clear();
      ta.value = '';
      $('report').innerHTML = '';
      lastReport = null;
      $('x-map').innerHTML = '';
      currentMapping = null;
      comparison.search = '';
      comparison.page = 0;
    }
    $('clear').onclick = () => { forgetData(); ta.dispatchEvent(new Event('input')); };
    ta.value = legacyText;
    let autoTimer;
    ta.addEventListener('input', () => {
      clearTimeout(autoTimer);
      autoTimer = setTimeout(() => { if (sel.value && ta.value.trim()) runFill(); }, 600);
    });

    let lastReport = null;
    const studentChoices = new Map();
    function renderReport({ report: r, total, key, text }) {
      const block = (cls, title, body) =>
        `<div class="rep ${cls}"><b>${title}</b><div>${body}</div></div>`;
      const sec = (cls, title, arr) =>
        arr.length ? block(cls, `${title} (${arr.length})`, arr.map(escapeHtml).join('\n')) : '';
      const filled = r.ok.length + r.approx.length;
      const chooseLabel = LANG === 'fr' ? 'Choisissez l’étudiant' : 'Choose the student';
      const choiceHtml = (choice, i) => {
        const buttons = choice.candidates.map((candidate, j) => {
          const duplicate = choice.candidates.filter((c) => c.name === candidate.name).length > 1;
          const detail = duplicate ? ` · ${escapeHtml(candidate.detail)}` : '';
          return `<button type="button" class="btn" data-choice="${i}" data-candidate="${j}">` +
            `${escapeHtml(candidate.name)}${detail}</button>`;
        }).join('');
        return `<div class="rep warn"><b>${chooseLabel} : ${escapeHtml(choice.line)}</b>` +
          `<div class="candidate-list">${buttons}</div></div>`;
      };
      const changes = r.changes || [];
      const changeRows = changes.map((c) => `<tr${c.approx ? ' class="approx"' : ''}><th scope="row">${escapeHtml(c.name)}</th>` +
        `<td>${escapeHtml(c.before || NA)}</td><td>${escapeHtml(c.after)}</td></tr>`).join('');
      const changesOpen = changes.length <= 12 || $('report').querySelector('details.verify')?.open;
      const changesHtml = changes.length
        ? `<details class="rep verify"${changesOpen ? ' open' : ''}><summary>${t('changesTitle', changes.length)}</summary>` +
          `<table class="changes"><thead><tr><th>${t('progressName')}</th><th>${t('changeBefore')}</th><th>${t('changeAfter')}</th></tr></thead>` +
          `<tbody>${changeRows}</tbody></table><p>${t('changesHelp')}</p></details>`
        : '';
      $('report').innerHTML =
        (filled ? block('ok', t('repOk', filled, total), r.ok.map(escapeHtml).join('\n')) : '') +
        changesHtml +
        (r.choices || []).map(choiceHtml).join('') +
        sec('info', t('repApprox'), r.approx) +
        sec('bad', t('repNotFound'), r.notFound) +
        sec('warn', t('repAmbig'), r.ambiguous) +
        sec('warn', t('repSkipped'), r.skipped) +
        sec('bad', t('repInvalid'), r.invalid) +
        sec('bad', t('repBad'), r.bad);
      $('report').querySelectorAll('[data-choice]').forEach((button) => {
        button.onclick = () => {
          if (sel.value !== key || ta.value !== text) { runFill(); return; }
          const choice = r.choices[Number(button.dataset.choice)];
          studentChoices.set(choice.key, choice.candidates[Number(button.dataset.candidate)].id);
          runFill();
        };
      });
    }
    function runFill() {
      clearTimeout(autoTimer);
      const key = sel.value;
      if (!key) return alert(t('alertNoFields'));
      lastReport = fill(ta.value, key, $('over').checked, studentChoices);
      renderReport(lastReport);
      update();
    }
    $('fill').onclick = runFill;
    $('undo').onclick = () => {
      const n = undoChanges(scanPage().get(sel.value) || []);
      lastReport = null;
      $('report').innerHTML = `<div class="rep info"><b>${t('undone', n)}</b></div>`;
      update();
    };

    const maxByCol = {};
    let stepChoice = 'auto';
    let statsView = 'overview';
    const comparison = { from: '', to: '', maxima: {}, page: 0, search: '' };
    const PROGRESS_VIEWS = ['class', 'students'];
    function applyStatsView() {
      const progress = PROGRESS_VIEWS.includes(statsView);
      root.querySelectorAll('[data-st-view]').forEach((b) => {
        b.setAttribute('aria-pressed', String(b.dataset.stView === statsView));
      });
      $('st-progress').hidden = !progress;
      $('st-body').hidden = progress;
      root.querySelectorAll('[data-st-section]').forEach((el) => { el.hidden = el.dataset.stSection !== statsView; });
    }
    root.querySelectorAll('[data-st-view]').forEach((b) => {
      b.onclick = () => {
        statsView = b.dataset.stView;
        renderStats();
        root.querySelector('.body').scrollTop = 0;
      };
    });

    function renderProgress() {
      const target = $('st-progress');
      if (!PROGRESS_VIEWS.includes(statsView)) return;
      const cols = scanPage();
      const keys = [...cols.keys()];
      const labels = new Map(keys.map((k) => [k, columnLabel(k, cols.get(k)[0])]));
      if (!cols.has(comparison.from)) {
        comparison.from = keys.find((k) => /intra|midterm/i.test(labels.get(k))) || keys[0] || '';
      }
      if (!cols.has(comparison.to)) {
        const others = keys.filter((k) => k !== comparison.from);
        comparison.to = others.find((k) => /final/i.test(labels.get(k))) || others[0] || '';
      }
      if (keys.length < 2) {
        target.innerHTML = `<p class="notice">${t('progressNeedTwo')}</p>`;
        return;
      }

      const options = (selected) => keys
        .map((k) => `<option value="${escapeHtml(k)}"${k === selected ? ' selected' : ''}>${escapeHtml(labels.get(k))}</option>`)
        .join('');
      const group = (side, title, key) => {
        const maxLabel = t(side === 'from' ? 'progressMaxFrom' : 'progressMaxTo');
        return '<div class="assessment-pair">' +
          `<label>${t(title)}<select id="p-${side}">${options(key)}</select></label>` +
          `<label>${t('scaleShort')}<input aria-label="${maxLabel}" id="p-max-${side}" type="number"` +
          ` min="0.01" step="any" placeholder="100" value="${comparison.maxima[key] || ''}"></label>` +
          '</div>';
      };
      const search = statsView === 'students'
        ? '<label class="student-search"><input type="search" id="p-search"' +
          ` aria-label="${t('searchStudent')}" placeholder="${t('searchStudent')}" value="${escapeHtml(comparison.search)}"></label>`
        : '';
      target.innerHTML = '<section class="progress"><div class="progress-controls">' +
        group('from', 'progressFrom', comparison.from) +
        group('to', 'progressTo', comparison.to) +
        '</div>' + search +
        '<div id="p-results"></div>' +
        `<details class="method"><summary>${t('howRead')}</summary>` +
        `<p>${t('progressIntro')}</p><p>${t('progressMeaning')}</p><p>${t('progressMissing')}</p>` +
        '</details></section>';

      const identify = (key) => pageRows(key).map((r) => {
        const m = r.el.name.match(/^grade\[(\d+)\]/) || r.el.name.match(/^(?:quickgrade_|finalgrade_)(\d+)(?:_|$)/);
        return { id: m && m[1], name: r.name, value: r.value };
      });
      const signed = (n) => (Number.isFinite(n) ? (n > 0 ? '+' : n < 0 ? '−' : '') + fmt(Math.abs(n), 2) : NA);
      const tile = (label, value) => `<div class="tile"><span>${label}</span><b>${value}</b></div>`;

      function classSummary(result) {
        if (!result.scaled) return '';
        if (!result.n) return `<p class="notice">${t('progressNoPairs')}</p>`;
        return '<div class="tiles comparison-tiles">' +
          tile(t('progressMeanFrom'), fmt(result.first.mean, 2) + ' %') +
          tile(t('progressMeanTo'), fmt(result.second.mean, 2) + ' %') +
          tile(t('progressMeanChange'), signed(result.change.mean) + ' pp') +
          tile(t('progressMedianChange'), signed(result.change.median) + ' pp') +
          '</div><div class="trend-counts">' +
          tile(t('upLabel'), result.up) +
          tile(t('sameLabel'), result.same) +
          tile(t('downLabel'), result.down) +
          '</div>';
      }

      function studentTable(result) {
        const filtered = result.rows.filter((r) => norm(r.name).includes(norm(comparison.search)));
        const size = window.innerHeight < 650 ? 3 : 6;
        const pages = Math.max(1, Math.ceil(filtered.length / size));
        comparison.page = Math.max(0, Math.min(comparison.page, pages - 1));
        const visible = filtered.slice(comparison.page * size, (comparison.page + 1) * size);
        const value = (raw, pct) => {
          if (!Number.isFinite(raw)) return NA;
          const title = Number.isFinite(pct) ? ` title="${fmt(pct, 2)} %"` : '';
          return `<span${title}>${fmt(raw, 2)}</span>`;
        };
        const head = `<tr><th>${t('progressName')}</th>` +
          `<th>${escapeHtml(labels.get(comparison.from))}</th>` +
          `<th>${escapeHtml(labels.get(comparison.to))}</th>` +
          `<th>${t('progressChange')} (pp)</th></tr>`;
        const body = visible.map((r) => `<tr><th scope="row">${escapeHtml(r.name || t('noName'))}</th>` +
          `<td>${value(r.from, r.fromPct)}</td><td>${value(r.to, r.toPct)}</td><td>${signed(r.delta)}</td></tr>`).join('');
        const first = comparison.page === 0 ? ' disabled' : '';
        const last = comparison.page === pages - 1 ? ' disabled' : '';
        return `<div class="progress-table"><table><thead>${head}</thead><tbody>${body}</tbody></table></div>` +
          (visible.length ? '' : `<p>${t('noSearchResults')}</p>`) +
          '<div class="pagination">' +
          `<button class="btn" id="p-prev"${first}>${t('previous')}</button>` +
          `<span>${t('pageCount', comparison.page + 1, pages, filtered.length)}</span>` +
          `<button class="btn" id="p-next"${last}>${t('next')}</button>` +
          '</div>';
      }

      function draw() {
        const out = $('p-results');
        if (comparison.from === comparison.to) {
          out.innerHTML = `<p class="notice">${t('progressDifferent')}</p>`;
          return;
        }
        const result = compareAssessments(
          identify(comparison.from),
          identify(comparison.to),
          comparison.maxima[comparison.from],
          comparison.maxima[comparison.to]
        );
        out.innerHTML =
          `<p class="coverage">${t('progressCoverage', result.n, result.excluded)}</p>` +
          (result.scaled ? '' : `<p class="notice">${t('scalePrompt')}</p>`) +
          (statsView === 'class' ? classSummary(result) : studentTable(result));
        if ($('p-prev')) $('p-prev').onclick = () => { comparison.page--; draw(); };
        if ($('p-next')) $('p-next').onclick = () => { comparison.page++; draw(); };
      }

      const pickAssessment = (side) => (e) => {
        comparison[side] = e.target.value;
        comparison.page = 0;
        $(`p-max-${side}`).value = comparison.maxima[comparison[side]] || '';
        draw();
      };
      const scaleChanged = (side) => (e) => {
        const key = comparison[side];
        const n = toNumber(e.target.value);
        if (Number.isFinite(n) && n > 0) comparison.maxima[key] = n;
        else delete comparison.maxima[key];
        draw();
      };
      $('p-from').onchange = pickAssessment('from');
      $('p-to').onchange = pickAssessment('to');
      $('p-max-from').oninput = scaleChanged('from');
      $('p-max-to').oninput = scaleChanged('to');
      if ($('p-search')) {
        $('p-search').oninput = (e) => {
          comparison.search = e.target.value;
          comparison.page = 0;
          draw();
        };
      }
      draw();
    }

    function renderStats(rows) {
      renderProgress();
      applyStatsView();
      if (statsView === 'class' || statsView === 'students') return;
      rows = rows || (sel.value ? pageRows(sel.value) : []);
      const box = $('st-body');
      const s = computeStats(rows.map((r) => r.value));
      const withText = rows.filter((r) => r.raw !== '').length;
      if (!s.n) {
        box.innerHTML = `<div class="empty">${!rows.length ? t('noFieldsStats') : withText ? t('noStatsText') : t('noStatsYet')}</div>`;
        return;
      }
      const setupState = setupMaxima(() => renderStats());
      const declaredMax = assessmentMaximum(rows[0]?.el) || (setupState.maxima && setupState.maxima[sel.value]) || null;
      const autoMax = declaredMax || guessMax(s.max);
      const letterMax = maxByCol[sel.value] || declaredMax;
      const max = Math.max(maxByCol[sel.value] || autoMax, s.max);
      const coarse = niceStep(max);
      let lo = s.min < 0 ? Math.floor(s.min) : Math.max(0, Math.floor(s.min / coarse) * coarse - coarse);
      const autoStep = niceStep(max - lo, 12);
      const step = stepChoice === 'auto' ? autoStep : Number(stepChoice);
      lo = Math.floor(lo / step) * step;
      const bins = histogram(s.sorted, step, lo, max);
      const top = Math.max(...bins.map((b) => b.count));
      const tops = bins.filter((b) => b.count === top);
      const rng = (b) => t('range', fmt(b.from), fmt(b.to));
      const sur = ` / ${fmt(max)}`;
      const tile = (label, value, sub, cls = '') =>
        `<div class="tile ${cls}"><span>${label}</span><b title="${escapeHtml(value)}">${value}</b>${sub ? `<small>${sub}</small>` : ''}</div>`;
      const partial = s.n < rows.length ? t('partial', rows.length - s.n) : '';
      const letterCounts = countLetterGrades(s.sorted, letterMax);
      const scaleNote = letterMax ? ` · ${t('outOf')} ${fmt(letterMax)}` : '';
      const letterBands = [['fail', t('belowCShort'), t('belowCTitle')]].concat(LETTER_BANDS.map(([letter, min], index) => {
        const limit = LETTER_BANDS[index + 1]?.[1];
        return [letter, letter, limit ? `${min} ≤ % < ${limit}` : `≥ ${min} %`];
      }));
      const letterTop = letterCounts ? Math.max(1, ...letterBands.map(([key]) => letterCounts[key])) : 1;
      const share = (count) => (s.n ? ` · ${fmt(count / s.n * 100)} %` : '');
      const letterHtml = `<div class="letter-counts${letterCounts ? '' : ' unscaled'}" role="group" aria-label="${t('letterCounts')}">` +
        '<div class="lc-bands">' +
        letterBands.map(([key, label, range]) => {
          const count = letterCounts ? letterCounts[key] : null;
          const height = count ? Math.max(8, Math.round(count / letterTop * 100)) : 0;
          const cls = key === 'fail' ? ' fail' : '';
          return `<div class="lc${cls}${count === 0 ? ' zero' : ''}" title="${escapeHtml(label)} : ${range}${scaleNote}${count != null ? share(count) : ''}">` +
            `<b>${count != null ? count : NA}</b><span class="lc-bar"><i style="height:${height}%"></i></span><span class="lc-l">${label}</span></div>`;
        }).join('') + '</div>' +
        `<div class="lc-total" title="${t('failuresTitle')}${scaleNote}">` +
        `<span class="lc-l">${t('failuresShort')}</span><b>${letterCounts ? letterCounts.fail : NA}</b>` +
        `<small>${letterCounts && s.n ? fmt(letterCounts.fail / s.n * 100) + ' %' : '&nbsp;'}</small></div>` +
        (!letterMax && setupState.done
          ? `<label class="lc-scale" title="${t('letterScaleNeeded')}">${t('outOf')} ` +
            `<input id="letters-max" aria-label="${t('letterScale')}" type="number" min="0.01" step="any" placeholder="100"></label>`
          : '') +
        '</div>';
      const overviewHtml =
        `<div data-st-section="overview"><div class="brief"><h3>${t('briefTitle')}</h3><p>` +
        t('brief', s.n, partial, fmt(s.mean) + sur, fmt(s.median), tops.map(rng).join(t('andBetween'))) +
        '</p></div>' + letterHtml +
        '<div class="group"><div class="tiles overview-tiles">' +
        tile(t('tMean'), fmt(s.mean) + sur, t('tMeanSub'), 'hl') +
        tile(t('tMedian'), fmt(s.median) + sur, t('tMedianSub'), 'hl') +
        tile(t('xMin'), fmt(s.min), t('tMinSub')) +
        tile(t('xMax'), fmt(s.max), t('tMaxSub')) +
        tile(t('tSd'), fmt(s.sd), t('tSdSub')) +
        tile(t('tIqr'), `${fmt(s.q1)} ${t('to')} ${fmt(s.q3)}`, t('tIqrSub', fmt(s.iqr))) +
        tile(t('tIdr'), `${fmt(s.p10)} ${t('to')} ${fmt(s.p90)}`, t('tIdrSub', fmt(s.idr))) +
        '</div></div></div>';
      if (statsView === 'overview') {
        box.innerHTML = overviewHtml;
        const scale = box.querySelector('#letters-max');
        if (scale) scale.onchange = () => {
          const value = toNumber(scale.value);
          if (Number.isFinite(value) && value > 0) { maxByCol[sel.value] = value; renderStats(); }
        };
        return;
      }
      box.innerHTML = '<div data-st-section="charts" class="charts-grid">' +
        `<div class="card"><h3>${t('histTitle')}</h3>` +
        `<p class="cap">${t('histCap')}</p>` +
        histSvg(bins, s, lo, max, top) +
        `<details class="settings"><summary>${t('settings')}</summary><div class="ctrl">` +
        `<label>${t('outOf')} <input type="number" id="st-max" min="1" step="any" value="${maxByCol[sel.value] || ''}" placeholder="${fmt(autoMax)}"></label>` +
        `<label>${t('barWidth')} <select id="st-step">${['auto', '1', '2', '2.5', '5', '10', '20']
          .map((v) => `<option value="${v}"${v === stepChoice ? ' selected' : ''}>${v === 'auto' ? t('auto', fmt(autoStep)) : t('points', fmt(Number(v)))}</option>`)
          .join('')}</select></label>` +
        '</div></details></div>' +
        `<div class="card"><h3>${t('bandTitle')}</h3>` +
        `<p class="cap">${t('bandCap')}</p>` +
        bandSvg(s, lo, max) +
        '<div class="bandleg">' +
        `<span><i style="background:var(--accent)"></i>${t('band50', fmt(s.q1), fmt(s.q3))}</span>` +
        `<span><i style="background:var(--accent-soft);border:1px solid var(--accent)"></i>${t('band80', fmt(s.p10), fmt(s.p90))}</span>` +
        '</div></div></div>';
      applyStatsView();
      const mx = box.querySelector('#st-max');
      mx.onchange = () => {
        const v = toNumber(mx.value);
        if (v > 0) maxByCol[sel.value] = v; else delete maxByCol[sel.value];
        renderStats();
      };
      box.querySelector('#st-step').onchange = (e) => { stepChoice = e.target.value; renderStats(); };
    }

    const W = 420, PL = 30, PR = 10;
    const xOf = (v, lo, hi) => PL + ((v - lo) / (hi - lo || 1)) * (W - PL - PR);
    function histSvg(bins, s, lo, hi, top) {
      const H = 210, PT = 34, PB = 24, ih = H - PT - PB;
      const yMax = Math.max(1, top), yStep = niceStep(yMax, 4) < 1 ? 1 : niceStep(yMax, 4);
      let g = '';
      for (let y = 0; y <= yMax + 1e-9; y += yStep) {
        const yy = PT + ih - (y / yMax) * ih;
        g += `<line class="grid" x1="${PL}" x2="${W - PR}" y1="${yy}" y2="${yy}"/><text x="${PL - 6}" y="${yy + 3.5}" text-anchor="end">${fmt(y, 0)}</text>`;
      }
      const every = Math.ceil(bins.length / 10);
      bins.forEach((b, i) => {
        const x0 = xOf(b.from, lo, hi), x1 = xOf(b.to, lo, hi), h = (b.count / yMax) * ih;
        const pc = fmt((100 * b.count) / s.n, 0);
        if (b.count) {
          g += `<rect class="barr${b.count === top ? ' top' : ''}" x="${x0 + 1}" y="${PT + ih - h}" width="${Math.max(1, x1 - x0 - 2)}" height="${h}" rx="3">` +
            `<title>${t('svgBin', fmt(b.from), fmt(b.to), b.count, pc)}</title></rect>`;
          if (x1 - x0 > 12) g += `<text class="lbl" x="${(x0 + x1) / 2}" y="${PT + ih - h - 4}" text-anchor="middle">${b.count}</text>`;
        }
        if (i % every === 0) g += `<text x="${x0}" y="${H - 8}" text-anchor="middle">${fmt(b.from)}</text>`;
      });
      g += `<text x="${xOf(hi, lo, hi)}" y="${H - 8}" text-anchor="middle">${fmt(hi)}</text>`;
      g += `<line class="axis" x1="${PL}" x2="${W - PR}" y1="${PT + ih}" y2="${PT + ih}"/>`;
      const xm = xOf(s.mean, lo, hi), xd = xOf(s.median, lo, hi);
      const meanLeft = s.mean <= s.median;
      g += `<line class="mean" x1="${xm}" x2="${xm}" y1="${PT - 30}" y2="${PT + ih}"><title>${t('tMean')} : ${fmt(s.mean)}</title></line>`;
      g += `<line class="median" x1="${xd}" x2="${xd}" y1="${PT - 30}" y2="${PT + ih}"><title>${t('tMedian')} : ${fmt(s.median)}</title></line>`;
      g += `<text class="mlab" x="${xm + (meanLeft ? -5 : 5)}" y="${PT - 20}" text-anchor="${meanLeft ? 'end' : 'start'}">${t('svgMean')} ${fmt(s.mean)}</text>`;
      g += `<text class="dlab" x="${xd + (meanLeft ? 5 : -5)}" y="${PT - 20}" text-anchor="${meanLeft ? 'start' : 'end'}">${t('svgMedian')} ${fmt(s.median)}</text>`;
      return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="${t('svgHistLabel')}">${g}</svg>`;
    }
    function bandSvg(s, lo, hi) {
      const H = 84, y = 30, bh = 20;
      const X = (v) => xOf(v, lo, hi);
      let g = `<rect class="track" x="${PL}" y="${y + 6}" width="${W - PL - PR}" height="${bh - 12}" rx="4"/>`;
      const band = (cls, from, to, title) =>
        `<rect class="${cls}" x="${X(from)}" y="${y}" width="${Math.max(2, X(to) - X(from))}" height="${bh}" rx="6">` +
        `<title>${title}</title></rect>`;
      g += band('b80', s.p10, s.p90, t('svg80', fmt(s.p10), fmt(s.p90)));
      g += band('b50', s.q1, s.q3, t('svg50', fmt(s.q1), fmt(s.q3)));
      g += `<line class="median" x1="${X(s.median)}" x2="${X(s.median)}" y1="${y - 6}" y2="${y + bh + 6}" style="stroke-width:3"/>`;
      g += `<text class="dlab" x="${X(s.median)}" y="${y - 10}" text-anchor="middle">${t('svgMedian')} ${fmt(s.median, 1)}</text>`;
      g += `<text class="tickv" x="${X(s.q1)}" y="${y + bh + 14}" text-anchor="middle">${fmt(s.q1, 1)}</text>`;
      g += `<text class="tickv" x="${X(s.q3)}" y="${y + bh + 14}" text-anchor="middle">${fmt(s.q3, 1)}</text>`;
      g += `<text x="${X(s.p10)}" y="${y + bh + 28}" text-anchor="middle">${fmt(s.p10, 1)}</text>`;
      g += `<text x="${X(s.p90)}" y="${y + bh + 28}" text-anchor="middle">${fmt(s.p90, 1)}</text>`;
      g += `<text x="${PL}" y="${y + bh / 2 + 4}" text-anchor="end" dx="-6">${fmt(lo)}</text>`;
      g += `<text x="${W - PR}" y="${y + bh / 2 + 4}" text-anchor="start" dx="4">${fmt(hi)}</text>`;
      return `<svg viewBox="0 0 ${W + 16} ${H}" role="img" aria-label="${t('svgBandLabel')}">${g}</svg>`;
    }

    function download(bytes, filename, type) {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(new Blob([bytes], { type }));
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
    }
    $('x-export').onclick = () => {
      const key = sel.value;
      if (!key) return alert(t('alertNoExport'));
      const rows = pageRows(key);
      const label = columnLabel(key, rows[0] && rows[0].el);
      const s = computeStats(rows.map((r) => r.value));
      const notes = [[t('xNom'), t('xPrenom'), t('xNote')]].concat(
        rows.map((r) => {
          const { prenom, nom } = splitName(r.name);
          return [nom, prenom, Number.isFinite(r.value) ? r.value : r.raw];
        })
      );
      const r2 = (x) => (Number.isFinite(x) ? Math.round(x * 100) / 100 : '');
      const done = rows.filter((r) => r.raw !== '').length;
      const stats = [
        [t('xStat'), t('xValue')],
        [t('xItem'), label],
        [t('xDate'), new Date().toLocaleString(LOCALE)],
        [t('xStudents'), rows.length],
        [t('xDone'), done],
        [t('xRest'), rows.length - done],
        [t('xMean'), r2(s.mean)],
        [t('xMedian'), r2(s.median)],
        [t('xSd'), r2(s.sd)],
        [t('xMin'), r2(s.min)],
        [t('xMax'), r2(s.max)],
        [t('xQ1'), r2(s.q1)],
        [t('xQ3'), r2(s.q3)],
        [t('xIqr'), r2(s.iqr)],
        [t('xP10'), r2(s.p10)],
        [t('xP90'), r2(s.p90)],
        [t('xIdr'), r2(s.idr)],
      ];
      const bytes = buildXlsx([
        { name: t('xSheetNotes'), rows: notes, widths: [26, 22, 10] },
        { name: t('xSheetStats'), rows: stats, widths: [24, 30] },
      ]);
      const date = new Date().toISOString().slice(0, 10);
      const safe = label.replace(/[\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
      download(bytes, `${t('xFile')} - ${safe} - ${date}.xlsx`, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    };

    const drop = $('x-drop');
    ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, () => drop.classList.remove('over')));
    drop.addEventListener('drop', (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]); });
    $('x-file').onchange = (e) => { if (e.target.files[0]) loadFile(e.target.files[0]); e.target.value = ''; };

    async function loadFile(file) {
      const out = $('x-map');
      currentMapping = null;
      out.innerHTML = `<p class="muted small" style="margin-top:10px">${t('reading')}</p>`;
      try {
        const bytes = new Uint8Array(await file.arrayBuffer());
        let rows;
        if (/\.xls$/i.test(file.name)) throw new Error(t('oldXls'));
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) rows = await readXlsx(bytes);
        else rows = parseCsv(decodeText(bytes));
        if (!rows.length) throw new Error(t('emptyFile'));
        showMapping(file.name, rows, detectColumns(rows));
      } catch (err) {
        out.innerHTML = `<div class="rep bad" style="margin-top:10px"><b>${t('cantRead')}</b><div>${escapeHtml(err.message || String(err))}</div></div>`;
      }
    }
    let currentMapping = null;
    function showMapping(fileName, rows, map) {
      currentMapping = () => showMapping(fileName, rows, get());
      const out = $('x-map');
      const width = Math.max(...rows.map((r) => r.length));
      const head = (i) => (map.header >= 0 ? String(rows[map.header][i] ?? '') : '');
      const colOpts = (cur, optional) =>
        (optional ? `<option value="-1"${cur < 0 ? ' selected' : ''}>${t('none')}</option>` : '') +
        Array.from({ length: width }, (_, i) => {
          const label = head(i) ? ' · ' + escapeHtml(head(i).slice(0, 24)) : '';
          return `<option value="${i}"${i === cur ? ' selected' : ''}>${colName(i)}${label}</option>`;
        }).join('');
      out.innerHTML =
        `<p class="small" style="margin:12px 0 0"><b>${escapeHtml(fileName)}</b> : ${t('lines', rows.length - (map.header + 1))}</p>` +
        '<div class="maprow">' +
        `<label>${t('mapNom')}<select id="m-nom">${colOpts(map.nom)}</select></label>` +
        `<label>${t('mapPrenom')}<select id="m-prenom">${colOpts(map.prenom, true)}</select></label>` +
        `<label>${t('mapNote')}<select id="m-note">${colOpts(map.note)}</select></label>` +
        '</div>' +
        `<label class="check" style="margin-top:0"><input type="checkbox" id="m-head"${map.header >= 0 ? ' checked' : ''}> ${t('firstRowHeaders')}</label>` +
        '<div class="prev" id="m-prev"></div>' +
        '<button class="btn primary big" id="m-go" style="margin-top:10px"></button>' +
        '<p class="muted small" id="m-info" style="margin:6px 0 0"></p>';
      const get = () => ({
        header: out.querySelector('#m-head').checked ? Math.max(0, map.header) : -1,
        nom: +out.querySelector('#m-nom').value,
        prenom: +out.querySelector('#m-prenom').value,
        note: +out.querySelector('#m-note').value,
      });
      function preview() {
        const m = get();
        const on = (i) => (i === m.nom || i === m.prenom || i === m.note ? ' class="on"' : '');
        const body = rows.slice(m.header + 1, m.header + 7);
        out.querySelector('#m-prev').innerHTML =
          '<table><thead><tr>' +
          Array.from({ length: width }, (_, i) => `<th${on(i)}>${m.header >= 0 ? escapeHtml(String(rows[m.header][i] ?? '')) : colName(i)}</th>`).join('') +
          '</tr></thead><tbody>' +
          body.map((r) => '<tr>' + Array.from({ length: width }, (_, i) => `<td${on(i)}>${escapeHtml(String(r[i] ?? ''))}</td>`).join('') + '</tr>').join('') +
          '</tbody></table>';
        const { count, skipped } = rowsToText(rows, m);
        out.querySelector('#m-go').textContent = t('importGo', count);
        out.querySelector('#m-go').disabled = !count;
        out.querySelector('#m-info').textContent = skipped ? t('importSkipped', skipped) : '';
      }
      out.querySelectorAll('select,input').forEach((el) => { el.onchange = preview; });
      out.querySelector('#m-go').onclick = () => {
        const { text } = rowsToText(rows, get());
        ta.value = text;
        showTab('saisie');
        runFill();
      };
      preview();
    }

    function applyLang() {
      LOCALE = I18N[LANG].locale;
      NA = I18N[LANG].na;
  NA = I18N[LANG].na;
      if (currentMapping) currentMapping();
      root.host.lang = LANG;
      root.querySelectorAll('[data-i]').forEach((el) => { el.innerHTML = t(el.dataset.i); });
      root.querySelectorAll('[data-it]').forEach((el) => { el.title = t(el.dataset.it); });
      root.querySelectorAll('[data-ip]').forEach((el) => { el.placeholder = t(el.dataset.ip); });
      $('langbtn').innerHTML = flagSvg(LANG) + LANG.toUpperCase() + ' ▾';
      root.querySelectorAll('[data-lang]').forEach((b) => b.setAttribute('aria-checked', String(b.dataset.lang === LANG)));
        if (lastReport) renderReport(lastReport);
      return refreshColumns();
    }
    $('langbtn').onclick = () => {
      const m = $('langmenu');
      m.hidden = !m.hidden;
      $('langbtn').setAttribute('aria-expanded', String(!m.hidden));
    };
    root.querySelectorAll('[data-lang]').forEach((b) => {
      b.onclick = () => {
        LANG = b.dataset.lang;
        try { localStorage.setItem(LANG_KEY, LANG); } catch (_) {}
        $('langmenu').hidden = true;
        $('langbtn').setAttribute('aria-expanded', 'false');
        applyLang();
      };
    });
    root.addEventListener('pointerdown', (e) => {
      if (e.composedPath().some((n) => n.classList && n.classList.contains('langwrap'))) return;
      $('langmenu').hidden = true;
      $('langbtn').setAttribute('aria-expanded', 'false');
    });
    const initialColumns = applyLang();

    window.__snShow = () => { userTouched = true; removeLegacy(); setOpen(true); refreshColumns(); };
    window.__snBuild = BUILD;
    window.__snDestroy = () => {
      destroyed = true;
      clearTimeout(updTimer); clearTimeout(autoTimer);
      startupTimers.forEach(clearTimeout);
      ['input', 'change'].forEach((t) => document.removeEventListener(t, onPageGradeChange, true));
      window.removeEventListener('resize', applyPanel);
      host.remove();
      delete window.__snShow; delete window.__snBuild; delete window.__snDestroy;
    };
    const fromUserscript = typeof GM_info !== 'undefined';
    const n = initialColumns;
    setOpen(!fromUserscript || n > 0);
    if (!n) {
      let found = false;
      [1500, 4000].forEach((ms) =>
        startupTimers.push(setTimeout(() => {
          if (found || !host.isConnected) return;
          const k = refreshColumns();
          found = k > 0;
          if (found && !userTouched && panel.hidden) setOpen(true);
        }, ms))
      );
    }
  }

  buildUI();
})();
