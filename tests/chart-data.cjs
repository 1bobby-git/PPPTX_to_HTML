'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const {pathToFileURL} = require('node:url');
const JSZip = require('jszip');
const {JSDOM} = require('jsdom');
const DOMPurify = require('dompurify');
const {inspect, LIMITS} = require('../src/chart-data');
const core = require('../src/conversion-core');
const root = path.resolve(__dirname, '..');
const fixture = n => fs.readFileSync(path.join(root, 'tests/fixtures', n));
const base = fixture('complex.pptx');
const norm = vm.createContext({JSZip, Uint8Array, Error, Number, Object, String});
vm.runInContext(fs.readFileSync(path.join(root, 'src/normalize-package.js'), 'utf8') + ';this.normalize=normalizePptxPackage', norm);
async function change(fn) { const z = await JSZip.loadAsync(base); await fn(z); return z.generateAsync({type:'uint8array'}); }
async function chartChange(fn) { return change(async z => z.file('ppt/charts/chart1.xml', fn(await z.file('ppt/charts/chart1.xml').async('string')))); }
function exported(rendered, chartData) {
  const dom = new JSDOM('<html><body></body></html>', {runScripts:'outside-only'});
  dom.window.DOMPurify = DOMPurify(dom.window);
  let source = fs.readFileSync(path.join(root,'src/exporter.js'),'utf8');
  for (const [marker,value] of [['/* VIEWER_CSS */ null',fs.readFileSync(path.join(root,'src/viewer.css'),'utf8')],['/* VIEWER_SCRIPT */ null',fs.readFileSync(path.join(root,'src/viewer.js'),'utf8')],['/* VIEWER_HASH */ null',require('node:crypto').createHash('sha256').update(fs.readFileSync(path.join(root,'src/viewer.js'))).digest('base64')]]) source=source.replace(marker,()=>JSON.stringify(value));
  dom.window.eval(source);
  const result = dom.window.PptxExport.prepare({...rendered, chartData});
  const document = new JSDOM(result.html).window.document;
  dom.window.close(); return {result,document};
}
async function browserIntegration() {
  const {chromium}=require('playwright');
  const browser=await chromium.launch({headless:true,executablePath:process.env.CHROME_BIN||'/tmp/pptx-browsers/chrome-linux64/chrome',args:['--no-sandbox']});
  try {
    const context=await browser.newContext({acceptDownloads:true});const requests=[];
    await context.route(/^https?:\/\//,route=>{requests.push(route.request().url());return route.abort('internetdisconnected');});
    const page=await context.newPage();await page.goto(pathToFileURL(path.join(root,'index.html')).href);
    async function convert(name) {
      await page.locator('#file-input').setInputFiles(path.join(root,'tests/fixtures',name));
      await page.locator('#convert-button').click();
      await page.waitForFunction(()=>!document.querySelector('#download-button').disabled||document.querySelector('#status-line').dataset.tone==='error',null,{timeout:60000});
      assert(await page.locator('#download-button').isEnabled(),await page.locator('#status-text').innerText());
    }
    await convert('complex.pptx');
    const frame=page.frameLocator('#preview-frame');
    await frame.locator('.chart-cache-data table').waitFor();
    assert.match(await frame.locator('.chart-cache-data').innerText(),/월/);
    assert.equal(await frame.locator('.chart-placeholder').getAttribute('role'),'region');
    assert((await page.locator('#warning-list').textContent()).includes('차트'));
    const downloadPromise=page.waitForEvent('download');await page.locator('#download-button').click();
    const download=await downloadPromise;const standalone=await context.newPage();
    const downloadedHtml=path.join(require('node:os').tmpdir(),'pptx-chart-'+Date.now()+'.html');
    await download.saveAs(downloadedHtml);
    await standalone.goto(pathToFileURL(downloadedHtml).href);
    fs.unlinkSync(downloadedHtml);
    await standalone.locator('.chart-cache-data table').waitFor({state:'attached'});
    await standalone.locator('#ptv-select').selectOption('2');
    assert.match(await standalone.locator('.chart-cache-data').innerText(),/캐시의 원시 값/);
    assert.equal(await standalone.locator('.chart-cache-data tbody tr').count(),3);
    await convert('basic-chart.pptx');
    await frame.locator('.chart-direct svg').waitFor();
    assert.equal(await frame.locator('.chart-cache-data').count(),0);
    assert.deepEqual(requests,[]);
    await standalone.close();await context.close();
  } finally {await browser.close();}
  console.log('PASS chart cache browser UI, downloaded standalone HTML, direct SVG preservation, zero network');
}
async function main() {
  const engine = await import(pathToFileURL(path.join(root,'node_modules/@briank-dev/pptx-to-html/pptx2html_wasm.js')));
  engine.initSync({module:fs.readFileSync(path.join(root,'node_modules/@briank-dev/pptx-to-html/pptx2html_wasm_bg.wasm'))});
  async function render(bytes,options={}) { const data=await norm.normalize(bytes); const r=core.render(engine,data,options); return {r, data:await inspect(data,JSZip,r.chartFallbacks,r.html)}; }
  const normal=await render(base);
  assert.equal(normal.data.length,1); assert.equal(normal.data[0].slideNumber,3);
  assert.deepEqual(normal.data[0].rows,[['월','2'],['화','4'],['수','3']]);
  assert(!JSON.stringify(normal.r.chartFallbacks).includes('raw_reference'));
  assert(!JSON.stringify(normal.data).includes('PRIVATE_SPEAKER'));
  const first=exported(normal.r,normal.data);
  assert.equal(first.document.querySelectorAll('.chart-cache-data table').length,1);
  assert.match(first.document.querySelector('.chart-placeholder').textContent,/원본 차트 대신/);
  assert(first.result.diagnostics.some(d=>/CHART/.test(d.code)));
  assert.deepEqual(Array.from(first.result.sourceSlides),[1,2,3]);
  const selected=await render(base,{slideRange:'3'});
  assert.deepEqual(Array.from(exported(selected.r,selected.data).result.sourceSlides),[3]);
  assert.equal(exported(selected.r,selected.data).document.querySelectorAll('.chart-cache-data table').length,1);
  const noChart=await render(base,{slideRange:'1'}); assert.equal(noChart.data.length,0);
  const reordered=await change(async z=>{let s=await z.file('ppt/presentation.xml').async('string');const ids=[...s.matchAll(/<p:sldId\b[^>]*\/>/g)].map(m=>m[0]);s=s.replace(ids[0],'SWAP').replace(ids[2],ids[0]).replace('SWAP',ids[2]);z.file('ppt/presentation.xml',s);});
  const reorderedResult=await render(reordered,{slideRange:'1'});assert.equal(reorderedResult.data[0].slideNumber,1);
  assert.equal(exported(reorderedResult.r,reorderedResult.data).document.querySelectorAll('.chart-cache-data table').length,1);
  const hidden=await change(async z=>z.file('ppt/slides/slide3.xml',(await z.file('ppt/slides/slide3.xml').async('string')).replace('<p:sld ','<p:sld show="0" ')));
  assert.equal((await render(hidden)).data.length,0);
  assert.equal((await render(hidden,{includeHidden:true})).data.length,1);
  const sparse=await chartChange(s=>s.replace(/<c:numCache>[\s\S]*?<\/c:numCache>/, '<c:numCache><c:ptCount val="3"/><c:pt idx="2"><c:v>30</c:v></c:pt><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache>'));
  const sparseData=await inspect(sparse,JSZip);assert.deepEqual(sparseData[0].rows.map(r=>r[1]),['10','','30']);
  const trailing=await chartChange(s=>s.replace(/(<c:numCache>[\s\S]*?<c:ptCount) val="3"/,'$1 val="5"'));
  assert.equal((await inspect(trailing,JSZip))[0].rows.length,5);
  // Every idx is a row coordinate; preserve explicit trailing blanks as well.
  const prefixed=await chartChange(s=>s.replace(/c:/g,'q:').replace('xmlns:c=','xmlns:q='));
  assert.deepEqual((await inspect(prefixed,JSZip))[0].rows,[['월','2'],['화','4'],['수','3']]);
  const manySeries=await chartChange(s=>{const ser=s.match(/<c:ser>[\s\S]*?<\/c:ser>/)[0];return s.replace(ser,ser.repeat(LIMITS.series+1));});
  assert.equal((await inspect(manySeries,JSZip))[0].status,'limit');
  let zipRead=false;assert.deepEqual(await inspect(base,{loadAsync(){zipRead=true;throw new Error('must not run');}},[]),[]);assert.equal(zipRead,false);
  const hostile=await chartChange(s=>s.replace('<c:v>월</c:v>','<c:v>&lt;img src=x onerror=alert(1)&gt; &amp; &quot;quote&quot;</c:v>'));
  const hostileResult=await render(hostile);const hostileDoc=exported(hostileResult.r,hostileResult.data).document;
  assert.match(hostileDoc.querySelector('.chart-cache-data').textContent,/<img src=x onerror=alert\(1\)> & "quote"/);
  assert.equal(hostileDoc.querySelectorAll('.chart-cache-data img, .chart-cache-data script, .chart-cache-data [onerror]').length,0);
  const empty=await chartChange(s=>s.replace(/<c:pt idx="\d+">[\s\S]*?<\/c:pt>/g,''));assert.equal((await inspect(empty,JSZip))[0].status,'empty');
  const oversized=await chartChange(s=>s.replace('<c:ptCount val="3"/>','<c:ptCount val="1000000000"/>'));assert.equal((await inspect(oversized,JSZip))[0].status,'limit');
  const duplicate=await chartChange(s=>s.replace('<c:pt idx="1"><c:v>4</c:v>','<c:pt idx="0"><c:v>4</c:v>'));assert.equal((await inspect(duplicate,JSZip))[0].status,'invalid');
  const corrupt=await chartChange(()=>'<c:chartSpace>');assert.equal((await inspect(corrupt,JSZip))[0].status,'invalid');
  const bigXml=await chartChange(s=>s+' '.repeat(LIMITS.xmlBytes));assert.equal((await inspect(bigXml,JSZip))[0].status,'invalid');
  const external=await change(async z=>z.file('ppt/slides/_rels/slide3.xml.rels',(await z.file('ppt/slides/_rels/slide3.xml.rels').async('string')).replace(/Target="\/ppt\/charts\/chart1.xml"/,'Target="https://example.invalid/chart.xml" TargetMode="External"')));
  assert.equal((await inspect(external,JSZip)).length,0);
  assert.equal(exported(normal.r,[normal.data[0],normal.data[0]]).document.querySelectorAll('.chart-cache-data').length,0);
  const wrong=structuredClone(normal.data);wrong[0].rect.x+=9525*10;assert.equal(exported(normal.r,wrong).document.querySelectorAll('.chart-cache-data').length,0);
  const ambiguous={...normal.r,html:normal.r.html.replace(/(<div class="shape" style="left: 384.0px; top: 115.2px; width: 700.8px; height: 432.0px">)/,'$1<div class="chart-placeholder">Chart</div>')};
  assert.equal(exported(ambiguous,normal.data).document.querySelectorAll('.chart-cache-data').length,0);
  const pie=await render(fixture('basic-chart.pptx'));assert.equal(pie.data.length,0);assert.equal(exported(pie.r,pie.data).document.querySelectorAll('.chart-direct svg').length,1);
  const native=await new Promise((resolve,reject)=>{
    const worker=new (require('node:worker_threads').Worker)(path.join(root,'src/native-cli-worker.cjs'),{workerData:{bytes:Uint8Array.from(base).buffer,title:'Chart cache test'}});
    const timeout=setTimeout(()=>{worker.terminate();reject(new Error('Native chart worker timed out'));},60000);
    worker.on('error',reject);worker.on('message',message=>{if(message.type==='result'||message.type==='error'){clearTimeout(timeout);worker.terminate();message.type==='error'?reject(new Error(message.message)):resolve(message);}});
  });
  const nativeDoc=new JSDOM(native.html).window.document;
  assert.equal(nativeDoc.querySelectorAll('.chart-cache-data table').length,1);
  assert.match(nativeDoc.querySelector('.chart-cache-data').textContent,/월/);
  assert(native.warnings.some(w=>/CHART/.test(w.code)));
  if(process.argv.includes('--browser')) await browserIntegration();
  console.log('PASS chart cache recovery: indices, text safety, limits, mapping, hidden/range/order, direct SVG, native integration');
}
main().catch(e=>{console.error(e);process.exitCode=1;});
