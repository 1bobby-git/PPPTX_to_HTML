'use strict';
// Deterministic OOXML/CSS checks, not screenshot or installed-font equivalence.
// --raw-theme-probe reproduces the documented unadapted engine defect as a hard failure.
const assert = require('node:assert/strict');
const test = require('node:test');
const fs = require('node:fs');
const path = require('node:path');
const {pathToFileURL} = require('node:url');
const JSZip = require('jszip');
const {JSDOM} = require('jsdom');
const fixtures = require('../scripts/create-fidelity-fixtures.cjs');
const {resolveThemes} = require('../src/theme-resolver.js');
const root = path.resolve(__dirname, '..');
const ready = (async () => {
  await fixtures.main();
  const packageDir = path.join(root, 'node_modules/@briank-dev/pptx-to-html');
  const engine = await import(pathToFileURL(path.join(packageDir, 'pptx2html_wasm.js')).href);
  engine.initSync({module: fs.readFileSync(path.join(packageDir, 'pptx2html_wasm_bg.wasm'))});
  return engine;
})();
async function render(name, resolve = false) {
  await ready;
  const bytes = fs.readFileSync(path.join(root, 'tests/fixtures/fidelity-' + name + '.pptx'));
  return renderBytes(bytes,resolve);
}
async function renderBytes(bytes, resolve = false) {
  const engine = await ready;
  const prepared = resolve ? await resolveThemes(new Uint8Array(bytes),JSZip) : {bytes:new Uint8Array(bytes),diagnostics:[]};
  const converted = engine.convert_with_options_metadata(prepared.bytes, true, false, new Uint32Array(), 1);
  try { return {dom: new JSDOM(converted.html), bytes, prepared}; } finally {converted.free();}
}
function closeTo(value, expected, label) {assert.ok(Math.abs(value - expected) <= 0.1, label + ': expected ' + expected + ', got ' + value);}

test('combined rowSpan and gridSpan preserve the two-by-two anchor and remaining grid cells', async () => {
  const {dom} = await render('merged-table');
  try {
    const table = dom.window.document.querySelector('table');
    assert.ok(table, 'table must render');
    const rows = [...table.rows];
    assert.deepEqual(rows.map(row => row.cells.length), [2, 1, 3]);
    assert.equal(table.querySelectorAll('td').length, 6);
    const anchor = rows[0].cells[0];
    assert.equal(anchor.colSpan, 2);
    assert.equal(anchor.rowSpan, 2);
    assert.equal(anchor.textContent.trim(), 'MERGED_2_BY_2');
    assert.deepEqual(rows.map(row => [...row.cells].map(cell => cell.textContent.trim())), [
      ['MERGED_2_BY_2', 'R1C3'], ['R2C3'], ['R3C1', 'R3C2', 'R3C3']
    ]);
    assert.deepEqual(rows.map(row => row.style.height), ['50px', '50px', '50px']);
  } finally {dom.window.close();}
});
test('nested group child offsets and unequal scales compose to the expected slide rectangle', async () => {
  const {dom} = await render('groups');
  try {
    const shapes = [...dom.window.document.querySelectorAll('.slide .shape')];
    assert.equal(shapes.length, 3, 'two enclosing groups and one leaf rectangle');
    const leaf = shapes[2];
    assert.match(leaf.textContent, /NESTED_GROUP/);
    // Engine emits already-scaled positions relative to nested absolute containers.
    let x = 0, y = 0, ancestors = 0;
    for (let node = leaf; node; node = node.parentElement) {
      if (!node.classList.contains('shape')) continue;
      x += parseFloat(node.style.left); y += parseFloat(node.style.top); ancestors++;
    }
    assert.equal(ancestors, 3);
    closeTo(x, 180, 'absolute left'); closeTo(y, 107, 'absolute top');
    closeTo(parseFloat(leaf.style.width), 40, 'leaf width');
    closeTo(parseFloat(leaf.style.height), 18, 'leaf height');
    assert.equal(leaf.style.backgroundColor, 'rgb(17, 170, 85)');
    // Rotation and reflection are intentionally outside this measured case.
  } finally {dom.window.close();}
});
test('first master inherits placeholder geometry, theme color, point size and CJK font family', async () => {
  const {dom} = await render('themes');
  try {
    const slide = dom.window.document.querySelector('.slide');
    const shape = slide.querySelector('.shape'), run = slide.querySelector('.run');
    closeTo(parseFloat(shape.style.left), 60, 'inherited left');
    closeTo(parseFloat(shape.style.top), 70, 'inherited top');
    closeTo(parseFloat(shape.style.width), 500, 'inherited width');
    assert.equal(run.style.fontSize, '24pt');
    assert.equal(run.style.color, 'rgb(17, 34, 170)');
    const segments = [...run.querySelectorAll('.run-segment')];
    assert.equal(segments[0].textContent, '상속 한국어 ');
    assert.match(segments[0].style.fontFamily, /Malgun Gothic/);
    assert.match(segments[1].style.fontFamily, /Arial/);
  } finally {dom.window.close();}
});
if (process.argv.includes('--raw-theme-probe')) test('unadapted engine must use the second master theme (known failure probe)', async () => {
  const {dom, bytes} = await render('themes');
  try {
    // Verify that the source really has independent themes and master relationships.
    const zip = await JSZip.loadAsync(bytes);
    const theme2 = await zip.file('ppt/theme/theme2.xml').async('string');
    const masterRels2 = await zip.file('ppt/slideMasters/_rels/slideMaster2.xml.rels').async('string');
    assert.match(theme2, /<a:accent1><a:srgbClr val="BB3311"/);
    assert.match(theme2, /<a:ea typeface="Microsoft YaHei"/);
    assert.match(masterRels2, /Target="\.\.\/theme\/theme2\.xml"/);
    const second = dom.window.document.querySelectorAll('.slide')[1];
    const run = second.querySelector('.run'), segment = run.querySelector('.run-segment');
    assert.deepEqual({color: run.style.color, font: segment.style.fontFamily.replace(/['"]/g, '')},
      {color: 'rgb(187, 51, 17)', font: 'Microsoft YaHei'});
  } finally {dom.window.close();}
});


test('narrow multi-theme resolution fixes the second theme and preserves complete 24pt defaults', async () => {
  const {dom,prepared} = await render('themes',true);
  try {
    assert.deepEqual(prepared.diagnostics,[]);
    const slides=[...dom.window.document.querySelectorAll('.slide')];
    assert.equal(slides.length,2);
    const actual=slides.map(slide=>{
      const run=slide.querySelector('.run'), segments=[...run.querySelectorAll('.run-segment')];
      return {size:run.style.fontSize,color:run.style.color,ea:segments[0].style.fontFamily.replace(/['"]/g,''),latin:segments[1].style.fontFamily.replace(/['"]/g,'')};
    });
    assert.deepEqual(actual,[
      {size:'24pt',color:'rgb(17, 34, 170)',ea:'Malgun Gothic',latin:'Arial'},
      {size:'24pt',color:'rgb(187, 51, 17)',ea:'Microsoft YaHei',latin:'Arial'}
    ]);
  } finally {dom.window.close();}
});
async function modifiedThemes(change) {
  await ready;
  const zip=await JSZip.loadAsync(fs.readFileSync(path.join(root,'tests/fixtures/fidelity-themes.pptx')));
  await change(zip);
  const bytes=await zip.generateAsync({type:'uint8array'});
  const result=await resolveThemes(bytes,JSZip);
  return {original:zip,bytes,result,output:await JSZip.loadAsync(result.bytes)};
}
async function omitOptionalListStyles(zip) {
  for (const id of [1,2]) {
    const name='ppt/slides/slide'+id+'.xml';
    let source=(await zip.file(name).async('string')).replace('<a:lstStyle/>','<!--KEEP_PARAGRAPH_BOUNDARY_'+id+'-->');
    if (id===2) source=source.replace('<a:bodyPr/>','<a:bodyPr></a:bodyPr>').replace(/a:/g,'draw:').replace('xmlns:a=','xmlns:draw=');
    zip.file(name,source);
  }
}
test('omitted optional list styles retain complete inherited defaults and untouched paragraph source', async () => {
  const {original,result,output}=await modifiedThemes(omitOptionalListStyles);
  const {dom}=await renderBytes(result.bytes);
  try {
    const actual=[...dom.window.document.querySelectorAll('.slide')].map(slide=>{
      const run=slide.querySelector('.run'), segments=[...run.querySelectorAll('.run-segment')];
      return {text:run.textContent,size:run.style.fontSize,color:run.style.color,ea:segments[0].style.fontFamily.replace(/['"]/g,''),latin:segments[1].style.fontFamily.replace(/['"]/g,'')};
    });
    assert.deepEqual(actual,[
      {text:'상속 한국어 THEME_ONE',size:'24pt',color:'rgb(17, 34, 170)',ea:'Malgun Gothic',latin:'Arial'},
      {text:'继承中文 THEME_TWO',size:'24pt',color:'rgb(187, 51, 17)',ea:'Microsoft YaHei',latin:'Arial'}
    ]);
    assert.deepEqual(result.diagnostics,[]);
    for (const id of [1,2]) {
      const name='ppt/slides/slide'+id+'.xml', before=await original.file(name).async('string'), after=await output.file(name).async('string');
      const inserted=after.match(/<ptctheme:lstStyle\b[\s\S]*?<\/ptctheme:lstStyle>/g);
      assert.equal(inserted?.length,1,'one list style must be added');
      assert.equal(after.replace(inserted[0],''),before,'all original XML, including comments and text, must retain its exact source order');
      const bodyEnd=id===1 ? '<a:bodyPr/>' : '<draw:bodyPr></draw:bodyPr>';
      assert.ok(after.includes(bodyEnd+inserted[0]+'<!--KEEP_PARAGRAPH_BOUNDARY_'+id+'-->'),'list style must follow bodyPr and precede the existing paragraph boundary');
    }
  } finally {dom.window.close();}
});
test('absent list styles do not trigger repair of invalid or unknown text-body children', async () => {
  const cases=[
    ['missing body properties',source=>source.replace('<a:bodyPr/>','')],
    ['body properties after paragraphs',source=>source.replace('<a:bodyPr/>','').replace('</p:txBody>','<a:bodyPr/></p:txBody>')],
    ['duplicate body properties',source=>source.replace('<a:bodyPr/>','<a:bodyPr/><a:bodyPr/>')],
    ['missing paragraphs',source=>source.replace(/<a:p>[\s\S]*?<\/a:p>/,'')],
    ['foreign child',source=>source.replace('<a:bodyPr/>','<a:bodyPr/><x:p xmlns:x="urn:fidelity-unknown"/>')]
  ];
  for (const [label,change] of cases) {
    const name='ppt/slides/slide2.xml';
    const {original,result,output}=await modifiedThemes(async zip=>{
      zip.file(name,change((await zip.file(name).async('string')).replace('<a:lstStyle/>','')));
    });
    assert.ok(result.diagnostics.some(item=>item.code==='THEME_INHERITANCE_UNRESOLVED'),label);
    assert.equal(await output.file(name).async('string'),await original.file(name).async('string'),label+' must retain its original source');
  }
});
test('single-theme packages preserve byte identity with no re-ZIP', async () => {
  await ready;
  const bytes=new Uint8Array(fs.readFileSync(path.join(root,'tests/fixtures/fidelity-groups.pptx')));
  const result=await resolveThemes(bytes,JSZip);
  assert.strictEqual(result.bytes,bytes);
  assert.deepEqual(result.diagnostics,[]);
});
test('layout or master placeholder local formatting prevents inherited defaults materialization', async () => {
  for (const omitList of [false,true]) {
    for (const part of ['ppt/slideLayouts/slideLayout2.xml','ppt/slideMasters/slideMaster2.xml']) {
      const {original,result,output}=await modifiedThemes(async zip=>{
        if (omitList) await omitOptionalListStyles(zip);
        const source=await zip.file(part).async('string');
        zip.file(part,source.replace('<a:lstStyle/>','<a:lstStyle><a:lvl1pPr><a:defRPr sz="1700"/></a:lvl1pPr></a:lstStyle>'));
      });
      assert.ok(result.diagnostics.some(item=>item.code==='THEME_INHERITANCE_UNRESOLVED'));
      assert.equal(await output.file('ppt/slides/slide2.xml').async('string'),await original.file('ppt/slides/slide2.xml').async('string'));
    }
  }
});
test('override color mappings keep the connected master chain untouched and report the boundary', async () => {
  const {original,result,output}=await modifiedThemes(async zip=>{
    const name='ppt/slides/slide2.xml';
    zip.file(name,(await zip.file(name).async('string')).replace('<a:masterClrMapping/>','<a:overrideClrMapping accent1="accent2"/>'));
  });
  assert.ok(result.diagnostics.some(item=>item.code==='THEME_INHERITANCE_UNRESOLVED'));
  for (const part of ['ppt/slides/slide2.xml','ppt/slideLayouts/slideLayout2.xml','ppt/slideMasters/slideMaster2.xml']) {
    assert.equal(await output.file(part).async('string'),await original.file(part).async('string'));
  }
});
test('source-span edits preserve mixed drawing order, foreign tags, attributes and color transforms', async () => {
  const markers='<p:sp><p:nvSpPr><p:cNvPr id="70" name="ORDER_SP"/></p:nvSpPr><p:spPr><a:solidFill><a:schemeClr val="accent1"><a:tint val="20000"/><a:alpha val="80000"/></a:schemeClr></a:solidFill></p:spPr></p:sp><!--untouched--><p:pic><p:nvPicPr><p:cNvPr id="71" name="ORDER_PIC"/></p:nvPicPr></p:pic><p:grpSp><p:nvGrpSpPr><p:cNvPr id="72" name="ORDER_GROUP"/></p:nvGrpSpPr><p:grpSpPr/></p:grpSp><p:extLst><x:schemeClr xmlns:x="urn:foreign" val="accent1" secret="PRIVATE_UNKNOWN"> literal &amp; untouched </x:schemeClr></p:extLst>';
  const {original,result,output}=await modifiedThemes(async zip=>{
    const name='ppt/slides/slide2.xml';
    zip.file(name,(await zip.file(name).async('string')).replace('</p:spTree>',markers+'</p:spTree>').replace(/a:/g,'draw:').replace('xmlns:a=','xmlns:draw='));
    zip.file('custom/untouched.xml','<unknown><!--PRIVATE_UNKNOWN--><item value="a &amp; b"/></unknown>');
  });
  assert.deepEqual(result.diagnostics,[]);
  const after=await output.file('ppt/slides/slide2.xml').async('string');
  const before=await original.file('ppt/slides/slide2.xml').async('string');
  assert.deepEqual([...after.matchAll(/name="ORDER_[A-Z]+"/g)].map(match=>match[0]),[...before.matchAll(/name="ORDER_[A-Z]+"/g)].map(match=>match[0]));
  assert.ok(after.includes('<!--untouched-->'));
  assert.ok(after.includes('<draw:srgbClr val="BB3311"><draw:tint val="20000"/><draw:alpha val="80000"/></draw:srgbClr>'));
  const foreign='<x:schemeClr xmlns:x="urn:foreign" val="accent1" secret="PRIVATE_UNKNOWN"> literal &amp; untouched </x:schemeClr>';
  assert.ok(after.includes(foreign));
  assert.equal(await output.file('custom/untouched.xml').async('string'),await original.file('custom/untouched.xml').async('string'));
  for (const name of Object.keys(original.files)) {
    if (original.files[name].dir || name==='ppt/slides/slide1.xml' || name==='ppt/slides/slide2.xml') continue;
    assert.equal(await output.file(name).async('string'),await original.file(name).async('string'),name+' must remain byte-identical XML/media content');
  }
});


test('theme metadata node and XML size limits preserve original bytes with a bounded warning', async () => {
  for (const extra of ['<x:n xmlns:x="urn:budget"/>'.repeat(50001), '<!--'+'x'.repeat(4*1024*1024)+'-->']) {
    const {bytes,result}=await modifiedThemes(async zip=>{
      const name='ppt/slides/slide2.xml';
      zip.file(name,(await zip.file(name).async('string')).replace('</p:spTree>',extra+'</p:spTree>'));
    });
    assert.strictEqual(result.bytes,bytes);
    assert.equal(result.diagnostics.length,1);
    assert.equal(result.diagnostics[0].code,'THEME_INHERITANCE_UNRESOLVED');
    assert.ok(result.diagnostics[0].message.length<200);
  }
});


test('generated master-default copies share the metadata and node budget and roll back atomically', async () => {
  for (const omitList of [false,true]) {
    const {bytes,result}=await modifiedThemes(async zip=>{
      if (omitList) await omitOptionalListStyles(zip);
      for (const id of [1,2]) {
        const name='ppt/slideMasters/slideMaster'+id+'.xml';
        const extensions='<a:extLst><a:ext uri="urn:fidelity-budget"><x:root xmlns:x="urn:fidelity-budget">'+'<x:n/>'.repeat(20000)+'</x:root></a:ext></a:extLst>';
        zip.file(name,(await zip.file(name).async('string')).replace('</a:defRPr>',extensions+'</a:defRPr>'));
      }
    });
    assert.strictEqual(result.bytes,bytes,'generated-copy node overflow must roll back every proposed chain');
    assert.equal(result.diagnostics.length,1);
    assert.equal(result.diagnostics[0].code,'THEME_INHERITANCE_UNRESOLVED');
  }
});


test('native worker pipeline resolves both themes with empty or omitted list styles before sanitizing', async () => {
  await ready;
  const {Worker}=require('node:worker_threads');
  const originalBytes=fs.readFileSync(path.join(root,'tests/fixtures/fidelity-themes.pptx'));
  for (const omitList of [false,true]) {
    let bytes=originalBytes;
    if (omitList) {
      const zip=await JSZip.loadAsync(bytes);
      await omitOptionalListStyles(zip);
      bytes=await zip.generateAsync({type:'uint8array'});
    }
    const result=await new Promise((resolve,reject)=>{
      const worker=new Worker(path.join(root,'src/native-cli-worker.cjs'),{workerData:{
        bytes:new Uint8Array(bytes),title:'Theme fidelity',slides:'',includeHidden:false
      }});
      let settled=false;
      const timer=setTimeout(()=>{worker.terminate();reject(new Error('Native fidelity worker timed out'));},30000);
      worker.on('message',message=>{
        if (message.type==='result' || message.type==='error') {
          settled=true;clearTimeout(timer);
          if (message.type==='error') reject(new Error(message.code+': '+message.message)); else resolve(message);
        }
      });
      worker.on('error',error=>{clearTimeout(timer);reject(error);});
      worker.on('exit',code=>{clearTimeout(timer);if(!settled) reject(new Error('Native fidelity worker exited without result: '+code));});
    });
    assert.equal(result.slideCount,2);
    assert.deepEqual(result.sourceSlides,[1,2]);
    assert.ok(!result.warnings.some(item=>item.code==='THEME_INHERITANCE_UNRESOLVED'));
    const dom=new JSDOM(result.html);
    try {
      const actual=[...dom.window.document.querySelectorAll('.slide')].map(slide=>{
        const run=slide.querySelector('.run'), segment=run.querySelector('.run-segment');
        return {size:run.style.fontSize,color:run.style.color,font:segment.style.fontFamily.replace(/['"]/g,'')};
      });
      assert.deepEqual(actual,[
        {size:'24pt',color:'rgb(17, 34, 170)',font:'Malgun Gothic'},
        {size:'24pt',color:'rgb(187, 51, 17)',font:'Microsoft YaHei'}
      ]);
      assert.match(result.html,/Content-Security-Policy/);
    } finally {dom.window.close();}
  }
});

if (process.argv.includes('--browser')) test('file browser UI resolves per-master theme colors, CJK typefaces and 24pt size', async () => {
  await ready;
  const {chromium,firefox,webkit}=require('playwright');
  const name=process.env.PPTX_TEST_BROWSER || 'chromium';
  const type={chromium,firefox,webkit}[name];
  assert.ok(type,'Supported PPTX_TEST_BROWSER is required');
  const explicit=process.env[name==='chromium' ? 'CHROME_BIN' : name==='firefox' ? 'FIREFOX_BIN' : 'WEBKIT_BIN'] || process.env['PLAYWRIGHT_'+name.toUpperCase()+'_EXECUTABLE_PATH'];
  const executablePath=[explicit,type.executablePath(),...(name==='chromium' ? [
    '/tmp/pptx-browsers/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium-1234/chrome-linux64/chrome',
    '/tmp/pptx-browsers/chromium_headless_shell-1234/chrome-headless-shell-linux64/chrome-headless-shell',
    '/usr/bin/chromium','/usr/bin/google-chrome'
  ] : [])].find(candidate=>candidate && fs.existsSync(candidate));
  assert.ok(executablePath,'Install the requested Playwright browser or supply its executable path');
  const browser=await type.launch({executablePath,headless:true,args:name==='chromium' ? ['--no-sandbox'] : []});
  try {
    const context=await browser.newContext(),requests=[];
    await context.route(/^https?:\/\//,route=>{requests.push(route.request().url());return route.abort('internetdisconnected');});
    const page=await context.newPage(),errors=[];
    page.on('pageerror',error=>errors.push(error.message));
    await page.goto(pathToFileURL(path.join(root,'index.html')).href);
    await page.locator('#file-input').setInputFiles(path.join(root,'tests/fixtures/fidelity-themes.pptx'));
    await page.locator('#convert-button').click();
    await page.waitForFunction(()=>!document.querySelector('#download-button').disabled || document.querySelector('#status-line').dataset.tone==='error',null,{timeout:30000});
    assert.equal(await page.locator('#download-button').isEnabled(),true,'Conversion failed: '+await page.locator('#status-text').innerText()+'; '+errors.join('; '));
    const frame=page.frameLocator('#preview-frame');
    await frame.locator('.slide').first().waitFor();
    const actual=await frame.locator('.slide').evaluateAll(slides=>slides.map(slide=>{
      const run=slide.querySelector('.run'),segment=run.querySelector('.run-segment');
      return {size:run.style.fontSize,color:run.style.color,font:segment.style.fontFamily.replace(/['"]/g,'')};
    }));
    assert.deepEqual(actual,[
      {size:'24pt',color:'rgb(17, 34, 170)',font:'Malgun Gothic'},
      {size:'24pt',color:'rgb(187, 51, 17)',font:'Microsoft YaHei'}
    ]);
    assert.deepEqual(requests,[],'Theme fixture must not request external resources');
    assert.deepEqual(errors,[]);
    await context.close();
  } finally {await browser.close();}
});
