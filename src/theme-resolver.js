'use strict';
// Resolve only unambiguous multi-master theme references. Edits are source spans:
// XML child order, unknown nodes, comments, and all untouched parts stay intact.
const {Parser} = require('xml2js');
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const COLOR_KEYS = ['accent1','accent2','accent3','accent4','accent5','accent6','bg1','bg2','folHlink','hlink','tx1','tx2'];
const COLORS = new Set(['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink']);
const MAX_XML = 4 * 1024 * 1024;
const MAX_INPUT = 64 * 1024 * 1024;
const MAX_METADATA_BYTES = 16 * 1024 * 1024;
const MAX_NODES = 50000;
function limitError() {const error=new Error('Theme metadata limit'); error.code='THEME_LIMIT'; return error;}

function parse(source, budget = {nodes:0}) {
  if (source.length > MAX_XML || /<!DOCTYPE/i.test(source)) throw new Error('Unsupported theme XML');
  // Reuse the namespace-aware SAX parser already shipped with xml2js 0.6.2.
  const parser = new Parser({xmlns: true, strict: true}).saxParser;
  const nodes = [], stack = [];
  parser.onopentag = tag => {
    if (++budget.nodes > MAX_NODES) throw limitError();
    const node = {name: tag.name, local: tag.local, uri: tag.uri, attrs: tag.attributes, namespaces: Object.assign({}, ...(() => {const list=[]; for (const prefix in tag.ns) list.push({[prefix]:tag.ns[prefix]}); return list;})()),
      start: parser.startTagPosition - 1, openEnd: parser.position, children: [], parent: stack[stack.length-1]};
    if (node.parent) node.parent.children.push(node);
    nodes.push(node); stack.push(node);
  };
  parser.onclosetag = () => {const node = stack.pop(); node.end = parser.position;};
  parser.ontext = parser.oncdata = parser.onend = () => {};
  parser.onerror = error => {throw error;};
  parser.write(source).close();
  return {source, nodes};
}
function attr(node, name, uri = '') {
  return Object.values(node.attrs).find(value => value.local === name && value.uri === uri)?.value;
}
function descendants(doc, uri, local) {return doc.nodes.filter(node => node.uri === uri && node.local === local);}
function child(node, uri, local) {return node?.children.find(value => value.uri === uri && value.local === local);}
function resolvePath(owner, target) {
  if (!target || /[?#\\]/.test(target) || /^[A-Za-z][A-Za-z0-9+.-]*:/.test(target) || target.startsWith('//')) return null;
  const parts = target.startsWith('/') ? [] : owner.split('/').slice(0,-1);
  for (const part of target.split('/')) {
    if (!part || part === '.') continue;
    if (part === '..') {if (!parts.length) return null; parts.pop();} else parts.push(part);
  }
  return parts.join('/');
}
function escapeXml(value) {return value.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;');}
function attributeEdit(doc, node, name, value) {
  const open = doc.source.slice(node.start,node.openEnd);
  const re = /([^\s=<>/]+)\s*=\s*(['"])([\s\S]*?)\2/g;
  for (const match of open.matchAll(re)) {
    if (match[1] !== name) continue;
    const valueOffset = match.index + match[0].indexOf(match[2]) + 1;
    return {start: node.start + valueOffset, end: node.start + valueOffset + match[3].length, value: escapeXml(value)};
  }
  throw new Error('Missing theme attribute span');
}
function apply(source, edits) {
  let output = source, lastStart = source.length;
  for (const edit of edits.sort((a,b) => b.start-a.start)) {
    if (edit.end > lastStart || edit.start < 0) throw new Error('Overlapping theme edits');
    output = output.slice(0,edit.start) + edit.value + output.slice(edit.end);
    lastStart = edit.start;
  }
  return output;
}
function themeValues(doc) {
  const elements = descendants(doc,A,'themeElements')[0];
  const scheme = child(elements,A,'clrScheme'), fontScheme = child(elements,A,'fontScheme');
  const colors = {}, fonts = {};
  for (const entry of scheme?.children || []) {
    if (entry.uri !== A || !COLORS.has(entry.local)) continue;
    const rgb = child(entry,A,'srgbClr'), value = rgb && attr(rgb,'val');
    if (rgb && !rgb.children.length && /^[0-9A-Fa-f]{6}$/.test(value || '')) colors[entry.local] = value;
  }
  for (const [kind, prefix] of [['majorFont','mj'],['minorFont','mn']]) {
    const list = child(fontScheme,A,kind);
    for (const [name,suffix] of [['latin','lt'],['ea','ea'],['cs','cs']]) {
      const value = attr(child(list,A,name) || {attrs:{}},'typeface');
      // Empty EA defaults require script/language-aware selection; do not guess.
      if (value && value.length <= 200 && !/^[+]/.test(value)) fonts['+'+prefix+'-'+suffix] = value;
    }
  }
  return {colors,fonts};
}
function editsFor(doc, theme, colorMap, state, skipMasterStyles = false) {
  const edits = [];
  for (const node of doc.nodes) {
    if (node.uri !== A) continue;
    let ancestor=node.parent, inMasterStyle=false;
    while (ancestor) {if (ancestor.uri===P && ancestor.local==='txStyles') inMasterStyle=true; ancestor=ancestor.parent;}
    if (skipMasterStyles && inMasterStyle) continue;
    if (node.local === 'schemeClr') {
      const token = attr(node,'val'), key = colorMap[token] || token, value = theme.colors[key];
      if (!value) {state.partial = true; continue;}
      edits.push(attributeEdit(doc,node,'val',value));
      const prefix = node.name.includes(':') ? node.name.split(':')[0]+':' : '';
      edits.push({start:node.start+1,end:node.start+1+node.name.length,value:prefix+'srgbClr'});
      const open = doc.source.slice(node.start,node.openEnd);
      if (!/\/\s*>$/.test(open)) {
        const closeAt = doc.source.lastIndexOf('</',node.end-1);
        if (closeAt < node.openEnd) throw new Error('Missing scheme color closing span');
        edits.push({start:closeAt+2,end:closeAt+2+node.name.length,value:prefix+'srgbClr'});
      }
    } else if (['latin','ea','cs'].includes(node.local)) {
      const token = attr(node,'typeface');
      if (!token || !/^\+(?:mj|mn)-(?:lt|ea|cs)$/.test(token)) continue;
      if (!theme.fonts[token]) {state.partial = true; continue;}
      edits.push(attributeEdit(doc,node,'typeface',theme.fonts[token]));
    }
  }
  return edits;
}

function placeholder(shape) {return child(child(child(shape,P,'nvSpPr'),P,'nvPr'),P,'ph');}
function placeholderShapes(doc) {return descendants(doc,P,'sp').filter(shape=>placeholder(shape));}
function matchingPlaceholder(doc, ph) {
  const idx=attr(ph,'idx'), type=attr(ph,'type');
  return placeholderShapes(doc).filter(shape=>{
    const candidate=placeholder(shape);
    return idx !== undefined ? attr(candidate,'idx')===idx : attr(candidate,'type')===type;
  });
}
function hasLocalTextStyle(shape) {
  if (child(shape,P,'style')?.children.length) return true;
  const body=child(shape,P,'txBody');
  if (!body) return false;
  const list=child(body,A,'lstStyle');
  if (list?.children.length) return true;
  for (const para of body.children.filter(node=>node.uri===A && node.local==='p')) {
    const props=child(para,A,'pPr');
    if (props && (props.children.length || Object.values(props.attrs).some(value=>value.uri==='' && value.local!=='lvl'))) return true;
    for (const node of para.children.filter(value=>value.uri===A && ['r','endParaRPr'].includes(value.local))) {
      const props=node.local==='r' ? child(node,A,'rPr') : node;
      if (props && (props.children.length || Object.values(props.attrs).some(value=>value.uri==='' && !['lang','dirty'].includes(value.local)))) return true;
    }
  }
  return false;
}
function inheritedStyleEdits(slideDoc,layoutDoc,masterDoc,theme,colorMap,state,budget) {
  const edits=[];
  for (const shape of placeholderShapes(slideDoc)) {
    const ph=placeholder(shape), type=attr(ph,'type');
    const body=child(shape,P,'txBody'), list=child(body,A,'lstStyle');
    if (!body) continue;
    const bodyProps=child(body,A,'bodyPr');
    // CT_TextBody permits an omitted lstStyle between bodyPr and its paragraphs.
    // Only insert into that known sequence; do not repair unknown/invalid bodies.
    const canInsert=!list && body.children.length>1 && body.children[0]===bodyProps &&
      body.children.slice(1).every(node=>node.uri===A && node.local==='p');
    const listAttributes=list ? Object.values(list.attrs).filter(value=>value.uri!=='http://www.w3.org/2000/xmlns/') : [];
    const layouts=matchingPlaceholder(layoutDoc,ph), masters=matchingPlaceholder(masterDoc,ph);
    if (child(shape,P,'style')?.children.length || !['body','title','ctrTitle'].includes(type) || (!list && !canInsert) || list?.children.length || listAttributes.length ||
        layouts.length>1 || masters.length>1 || [...layouts,...masters].some(hasLocalTextStyle)) {state.partial=true;continue;}
    const styles=descendants(masterDoc,P,'txStyles')[0];
    const sourceStyle=child(styles,P,type==='body' ? 'bodyStyle' : 'titleStyle');
    if (!sourceStyle || !sourceStyle.children.length) {state.partial=true;continue;}
    // Copy the complete paragraph defaults: this engine replaces defRPr as a unit,
    // so copying only color/font would discard inherited size and other defaults.
    const closeAt=masterDoc.source.lastIndexOf('</',sourceStyle.end-1);
    let prefix='ptctheme';
    while (sourceStyle.namespaces[prefix]) prefix+='x';
    const declarations=Object.entries(sourceStyle.namespaces).filter(([name])=>name!=='xml' && name!=='xmlns').map(([name,uri])=>
      ' xmlns'+(name ? ':'+name : '')+'="'+escapeXml(uri)+'"').join('');
    const snippet='<'+prefix+':lstStyle xmlns:'+prefix+'="'+A+'"'+declarations+'>'+masterDoc.source.slice(sourceStyle.openEnd,closeAt)+'</'+prefix+':lstStyle>';
    // Include generated defaults as well as source XML in the shared metadata
    // budget. Three bytes per UTF-16 code unit conservatively bounds UTF-8.
    budget.bytes+=snippet.length*3;
    if (budget.bytes>MAX_METADATA_BYTES) {state.limit=true;throw limitError();}
    let copied;
    try {copied=parse(snippet,budget);} catch(error) {if(error.code==='THEME_LIMIT') state.limit=true;throw error;}
    const copyState={partial:false}, copiedEdits=editsFor(copied,theme,colorMap,copyState);
    if (copyState.partial) {state.partial=true;continue;}
    const resolved=apply(snippet,copiedEdits);
    edits.push({start:list ? list.start : bodyProps.end,end:list ? list.end : bodyProps.end,value:resolved});
  }
  return edits;
}

async function resolveThemes(bytes, JSZip) {
  const unchanged = () => ({bytes,diagnostics:[]});
  if (!(bytes instanceof Uint8Array) || !bytes.length || bytes.length > MAX_INPUT) throw new Error('Invalid theme input size');
  if (!JSZip || typeof JSZip.loadAsync!=='function') throw new Error('JSZip must be supplied');
  const zip = await JSZip.loadAsync(bytes);
  if (Object.keys(zip.files).filter(name=>/^ppt\/theme\/[^/]+\.xml$/.test(name)).length<2) return unchanged();
  const cache = new Map(), state = {partial:false}, budget={nodes:0,bytes:0};
  function unresolved() {return {bytes,diagnostics:[{code:'THEME_INHERITANCE_UNRESOLVED',message:'복수 테마의 일부 연결·색상·글꼴 상속은 확실하게 해석하지 못해 원본을 유지했습니다.'}]};}
  async function read(name) {
    if (cache.has(name)) return cache.get(name);
    const entry = zip.file(name);
    if (!entry) throw new Error('Unavailable theme part');
    if (!Number.isSafeInteger(entry._data?.uncompressedSize) || entry._data.uncompressedSize>MAX_XML) {state.limit=true;throw limitError();}
    budget.bytes+=entry._data.uncompressedSize;
    if (budget.bytes>MAX_METADATA_BYTES) {state.limit=true;throw limitError();}
    const source=await entry.async('string');
    let doc;
    try {doc=parse(source,budget);} catch(error) {if(error.code==='THEME_LIMIT') state.limit=true;throw error;}
    cache.set(name,doc); return doc;
  }
  async function relationships(owner) {
    const pieces=owner.split('/'), filename=pieces.pop();
    const name = pieces.concat('_rels',filename+'.rels').join('/');
    const doc = await read(name);
    return descendants(doc,REL,'Relationship').map(node => ({id:attr(node,'Id'),type:attr(node,'Type'),
      target: attr(node,'TargetMode')?.toLowerCase() === 'external' ? null : resolvePath(owner,attr(node,'Target'))}));
  }
  function one(rels, suffix) {
    const values=rels.filter(rel => rel.type?.endsWith('/'+suffix));
    if (values.length !== 1 || !values[0].target) throw new Error('Ambiguous theme relationship');
    return values[0].target;
  }
  const chains=[], failedParts=new Set();
  try {
    const pres = await read('ppt/presentation.xml'), presRels = await relationships('ppt/presentation.xml');
    const slidePaths = descendants(pres,P,'sldId').map(node => {
      const id=attr(node,'id',R); const matches=presRels.filter(rel=>rel.id===id && rel.type?.endsWith('/slide'));
      return matches.length===1 ? matches[0].target : null;
    });
    if (slidePaths.length > 1000) throw new Error('Theme slide limit');
    for (const slidePath of slidePaths) {
      if (!slidePath) {state.partial=true;continue;}
      const currentParts=[slidePath];
      try {
        const slideRels=await relationships(slidePath), layoutPath=one(slideRels,'slideLayout');
        currentParts.push(layoutPath);
        const layoutRels=await relationships(layoutPath), masterPath=one(layoutRels,'slideMaster');
        currentParts.push(masterPath);
        const masterRels=await relationships(masterPath), themePath=one(masterRels,'theme');
        const parts=[slidePath,layoutPath,masterPath], docs=[];
        for (const part of parts) docs.push(await read(part));
        const mapNode=descendants(docs[2],P,'clrMap')[0], colorMap={};
        let safe=!!mapNode;
        for (const key of COLOR_KEYS) {
          colorMap[key]=mapNode && attr(mapNode,key);
          if (!COLORS.has(colorMap[key])) safe=false;
        }
        // Shared ancestors cannot be flattened correctly for per-slide overrides.
        if (docs.some(doc=>descendants(doc,A,'overrideClrMapping').length) ||
            [slideRels,layoutRels,masterRels].some(list=>list.some(rel=>rel.type?.endsWith('/themeOverride')))) safe=false;
        if ([slideRels,layoutRels,masterRels].some(list=>list.some(rel=>/\/(?:chart|diagramData|diagramDrawing)$/.test(rel.type || '')))) state.partial=true;
        chains.push({parts,themePath,colorMap,safe});
      } catch {state.partial=true;currentParts.forEach(part=>failedParts.add(part));}
    }
  } catch {state.partial=true;}
  if (state.limit) return unresolved();
  // Single-theme packages retain byte identity and existing renderer behavior.
  if (new Set(chains.map(chain=>chain.themePath)).size < 2) {
    if (state.partial && Object.keys(zip.files).filter(name=>/^ppt\/theme\/[^/]+\.xml$/.test(name)).length>1) return {bytes,diagnostics:[{code:'THEME_INHERITANCE_UNRESOLVED',message:'복수 테마의 연결 또는 상속을 확실하게 해석하지 못했습니다. 원본과 비교해 주세요.'}]};
    return unchanged();
  }
  const blocked=new Set(failedParts), contexts=new Map();
  for (const chain of chains) {
    const signature=JSON.stringify([chain.themePath,chain.colorMap]);
    for (const part of chain.parts) {
      if (!chain.safe || (contexts.has(part) && contexts.get(part)!==signature)) blocked.add(part);
      contexts.set(part,signature);
    }
  }
  // A rejected slide may share an ancestor with another slide. Keep its complete
  // connected chain untouched so baking a parent cannot change override semantics.
  let expanded=true;
  while (expanded) {
    expanded=false;
    for (const chain of chains) if (chain.parts.some(part=>blocked.has(part))) {
      for (const part of chain.parts) if (!blocked.has(part)) {blocked.add(part);expanded=true;}
    }
  }
  const changes=new Map(), visited=new Set();
  for (const chain of chains) {
    if (chain.parts.some(part=>blocked.has(part))) {state.partial=true;continue;}
    try {
      const values=themeValues(await read(chain.themePath)), proposed=new Map();
      for (const part of chain.parts) {
        if (visited.has(part)) continue;
        const doc=await read(part), edits=editsFor(doc,values,chain.colorMap,state,part===chain.parts[2]);
        if (part===chain.parts[0]) edits.push(...inheritedStyleEdits(doc,await read(chain.parts[1]),await read(chain.parts[2]),values,chain.colorMap,state,budget));
        if (edits.length) proposed.set(part,apply(doc.source,edits));
      }
      for (const [part,source] of proposed) changes.set(part,source);
      chain.parts.forEach(part=>visited.add(part));
    } catch {state.partial=true;chain.parts.forEach(part=>blocked.add(part));}
  }
  if (state.limit) return unresolved();
  expanded=true;
  while (expanded) {
    expanded=false;
    for (const chain of chains) if (chain.parts.some(part=>blocked.has(part))) {
      for (const part of chain.parts) if (!blocked.has(part)) {blocked.add(part);expanded=true;}
    }
  }
  for (const part of blocked) changes.delete(part);
  if (!changes.size) return {bytes,diagnostics:[{code:'THEME_INHERITANCE_UNRESOLVED',message:'복수 테마의 일부 색상·글꼴 상속을 확실하게 해석하지 못했습니다. 원본과 비교해 주세요.'}]};
  for (const [name,source] of changes) zip.file(name,source,{createFolders:false});
  const output = await zip.generateAsync({type:'uint8array',compression:'DEFLATE',compressionOptions:{level:6}});
  if (output.byteLength > MAX_INPUT) throw new Error('Resolved PPTX exceeds 64 MiB');
  return {bytes:output,diagnostics:state.partial ? [{code:'THEME_INHERITANCE_UNRESOLVED',message:'복수 테마의 일부 색상·글꼴 또는 차트·다이어그램은 원본과 다를 수 있습니다.'}] : []};
}
module.exports = {resolveThemes};
