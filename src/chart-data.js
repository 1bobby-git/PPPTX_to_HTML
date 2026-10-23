'use strict';

// Read only cached chart values in the supplied, already local PPTX package.
// JSZip is injected so the browser worker reuses its existing ZIP library.
const {Parser} = require('xml2js');
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const R = 'http://schemas.openxmlformats.org/package/2006/relationships';
const LIMITS = Object.freeze({xmlBytes: 2 * 1024 * 1024, totalXmlBytes: 16 * 1024 * 1024,
  charts: 100, rows: 100, series: 16, text: 256, metadataBytes: 1024 * 1024});
function children(node, local, uri) {
  return Object.values(node || {}).filter(Array.isArray).flat().filter(n => n && n.$ns?.local === local && (!uri || n.$ns.uri === uri));
}
function first(node, local, uri) { return children(node, local, uri)[0]; }
function attr(node, name) { return node?.$?.[name]?.value; }
function text(node) { return typeof node?._ === 'string' ? node._ : ''; }
function short(value) { return String(value || '').slice(0, LIMITS.text); }
function resolve(owner, target) {
  if (typeof target !== 'string' || /[\\?#\x00-\x20]/.test(target) || /^[A-Za-z][\w+.-]*:|^\/\//.test(target)) return null;
  const parts = target.startsWith('/') ? [] : owner.split('/').slice(0, -1);
  for (const p of target.split('/')) {
    if (!p || p === '.') continue;
    if (p === '..') { if (!parts.length) return null; parts.pop(); }
    else parts.push(p);
  }
  return parts.join('/');
}
function relsPath(owner) { const p = owner.split('/'); const base = p.pop(); return p.concat('_rels', base + '.rels').join('/'); }
function cache(container) {
  const multi = first(first(container, 'multiLvlStrRef', C), 'multiLvlStrCache', C);
  if (multi) {
    const levels = children(multi, 'lvl', C);
    if (levels.length > LIMITS.series) throw new Error('large-cache');
    const values = {}; let count = 0;
    for (const level of levels) {
      const wrapped = {cache: [{...level, $ns: {local: 'strLit', uri: C}}]};
      const current = cache(wrapped);
      count = Math.max(count, current.count);
      for (const idx of Object.keys(current.values)) values[idx] = short([values[idx], current.values[idx]].filter(Boolean).join(' / '));
    }
    const declared = Number(attr(first(multi, 'ptCount', C), 'val'));
    if (Number.isFinite(declared) && declared > LIMITS.rows) throw new Error('large-cache');
    return {values, count};
  }
  const ref = first(container, 'strRef', C) || first(container, 'numRef', C);
  const node = first(ref, 'strCache', C) || first(ref, 'numCache', C) || first(container, 'strLit', C) || first(container, 'numLit', C);
  const points = children(node, 'pt', C);
  if (!node || !points.length) return {values: {}, count: 0};
  const declared = Number(attr(first(node, 'ptCount', C), 'val'));
  if (points.length > LIMITS.rows || (Number.isFinite(declared) && declared > LIMITS.rows)) throw new Error('large-cache');
  const values = {};
  let count = Number.isSafeInteger(declared) && declared >= 0 ? declared : 0;
  for (const point of points) {
    const raw = attr(point, 'idx');
    if (!/^\d+$/.test(raw || '')) throw new Error('invalid-cache');
    const idx = Number(raw);
    if (!Number.isSafeInteger(idx) || idx >= LIMITS.rows || Object.hasOwn(values, idx)) throw new Error('invalid-cache');
    values[idx] = short(text(first(point, 'v', C)));
    count = Math.max(count, idx + 1);
  }
  return {values, count};
}
function chartRows(root) {
  const chart = first(root, 'chart', C), plot = first(chart, 'plotArea', C);
  const series = Object.values(plot || {}).filter(Array.isArray).flat().flatMap(n => children(n, 'ser', C));
  if (series.length > LIMITS.series) return {status: 'limit', columns: [], rows: []};
  const columns = [], data = [];
  let count = 0;
  try {
    for (let i = 0; i < series.length; i++) {
      const s = series[i], tx = first(s, 'tx', C);
      const title = text(first(tx, 'v', C)) || cache(tx).values[0] || '계열 ' + (i + 1);
      const x = first(s, 'xVal', C), y = first(s, 'yVal', C);
      if (x || y) {
        for (const [name, container] of [[' X', x], [' Y', y]]) {
          const values = cache(container); columns.push(short(title + name)); data.push(values.values); count = Math.max(count, values.count);
        }
      } else {
        const category = cache(first(s, 'cat', C)), value = cache(first(s, 'val', C));
        // Separate categories per series avoid assuming that different axes align.
        columns.push(short(title + ' · 항목'), short(title + ' · 값'));
        data.push(category.values, value.values); count = Math.max(count, category.count, value.count);
      }
    }
  } catch (error) { return {status: error.message === 'large-cache' ? 'limit' : 'invalid', columns: [], rows: []}; }
  if (!count) return {status: 'empty', columns: [], rows: []};
  return {status: 'cached', columns, rows: Array.from({length: count}, (_, idx) => data.map(values => values[idx] ?? ''))};
}
async function inspect(bytes, JSZip, fallbacks, renderedHtml) {
  const records = [];
  if (Array.isArray(fallbacks) && fallbacks.length === 0) return records;
  try {
    if (!bytes || bytes.byteLength > 64 * 1024 * 1024) return records;
    const zip = await JSZip.loadAsync(bytes);
    let total = 0, metadata = 0;
    async function xml(path) {
      const entry = zip.file(path), size = entry?._data?.uncompressedSize;
      if (!entry || !Number.isSafeInteger(size) || size > LIMITS.xmlBytes || (total += size) > LIMITS.totalXmlBytes) throw new Error('xml-limit');
      const source = await entry.async('string');
      if (/<!DOCTYPE|<!ENTITY/i.test(source) || source.includes('\0')) throw new Error('xml-invalid');
      const parsed = await new Parser({xmlns: true, explicitArray: true, explicitRoot: false, strict: true}).parseStringPromise(source);
      return parsed;
    }
    async function relations(owner) {
      const root = await xml(relsPath(owner));
      return children(root, 'Relationship', R);
    }
    const presentation = await xml('ppt/presentation.xml');
    const presentationRels = await relations('ppt/presentation.xml');
    const order = new Map();
    const slideIds = children(first(presentation, 'sldIdLst', 'http://schemas.openxmlformats.org/presentationml/2006/main'), 'sldId', 'http://schemas.openxmlformats.org/presentationml/2006/main');
    if (slideIds.length > 1000) return records;
    for (let i = 0; i < slideIds.length; i++) {
      const id = Object.values(slideIds[i].$ || {}).find(a => a.local === 'id' && a.uri === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')?.value;
      const matching = presentationRels.filter(r => attr(r, 'Id') === id && attr(r, 'Type') === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide' && !attr(r, 'TargetMode'));
      if (matching.length !== 1) continue;
      const part = resolve('ppt/presentation.xml', attr(matching[0], 'Target'));
      if (part) order.set(part, order.has(part) ? null : i + 1);
    }
    const visible = typeof renderedHtml === 'string' ? new Set(Array.from(renderedHtml.matchAll(/<div class="slide"[^>]*\bdata-slide="(\d+)"/g), m => Number(m[1]))) : null;
    // Diagnostics supply both the owner part and rId. Do not scan arbitrary media,
    // notes, external workbooks or package filenames to guess chart identity.
    const selected = Array.isArray(fallbacks) ? fallbacks.slice(0, LIMITS.charts) : null;
    const slides = selected ? [...new Set(selected.map(f => f.part))].filter(p => /^ppt\/slides\/[^/]+\.xml$/.test(p)) :
      Object.keys(zip.files).filter(p => /^ppt\/slides\/[^/]+\.xml$/.test(p)).slice(0, 1000);
    for (const owner of slides) {
      const slideNumber = order.get(owner);
      if (!slideNumber || (visible && !visible.has(slideNumber))) continue;
      let rels;
      try { rels = await relations(owner); } catch { continue; }
      for (const rel of rels) {
        if (records.length >= LIMITS.charts) return records;
        if (attr(rel, 'Type') !== 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart' ||
            String(attr(rel, 'TargetMode') || '').toLowerCase() === 'external') continue;
        const relationshipId = attr(rel, 'Id'), target = resolve(owner, attr(rel, 'Target'));
        const matches = selected?.filter(f => f.part === owner && f.relationshipId === relationshipId);
        if (selected && matches.length !== 1) continue;
        if (!target || !relationshipId || relationshipId.length > 256 || rels.filter(r => attr(r, 'Id') === relationshipId).length !== 1) continue;
        let result;
        try { result = chartRows(await xml(target)); } catch { result = {status: 'invalid', columns: [], rows: []}; }
        const record = {...(matches?.[0] || {}), part: owner, relationshipId, slideNumber, ...result};
        const length = JSON.stringify(record).length * 3;
        if ((metadata += length) > LIMITS.metadataBytes) return records;
        records.push(record);
      }
    }
  } catch { /* Optional data recovery must not fail conversion. */ }
  return records;
}

module.exports = {inspect, LIMITS};
