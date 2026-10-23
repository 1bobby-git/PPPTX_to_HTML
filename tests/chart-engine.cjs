'use strict';

// Real WASM regression for equivalent XML spellings and nested chart styles.
// Test fixtures are synthesized in memory; input conversion never repairs axes.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const crypto = require('node:crypto');
const {pathToFileURL} = require('node:url');
const JSZip = require('jszip');
const {JSDOM} = require('jsdom');
const root = path.resolve(__dirname, '..');
const engineDirectory = process.env.PPTX_ENGINE_DIR || path.join(root, 'node_modules/@briank-dev/pptx-to-html');
const fixture = name => Uint8Array.from(fs.readFileSync(path.join(root, 'tests/fixtures', name)));
const context = vm.createContext({JSZip, Uint8Array, Error, Number, Object, String});
vm.runInContext(fs.readFileSync(path.join(root, 'src/normalize-package.js'), 'utf8') + ';this.normalize=normalizePptxPackage', context);
const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function chart(style, extraAxis = '') {
  return `<c:chartSpace xmlns:c="${C}" xmlns:a="${A}"><c:chart><c:plotArea>
    <c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
      <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>판매량</c:v></c:tx>${style}
        <c:cat><c:strLit><c:ptCount val="3"/><c:pt idx="0"><c:v>월</c:v></c:pt><c:pt idx="1"><c:v>화</c:v></c:pt><c:pt idx="2"><c:v>수</c:v></c:pt></c:strLit></c:cat>
        <c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numLit></c:val>
      </c:ser><c:axId val="10"/><c:axId val="20"/>${extraAxis}
    </c:barChart>
    <c:catAx><c:axId val="10"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/><c:crossAx val="20"/><c:crosses val="autoZero"/></c:catAx>
    <c:valAx><c:axId val="20"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/><c:crossAx val="10"/><c:crosses val="autoZero"/></c:valAx>
  </c:plotArea></c:chart></c:chartSpace>`;
}
async function withChart(xml) {
  const zip = await JSZip.loadAsync(fixture('complex.pptx'));
  zip.file('ppt/charts/chart1.xml', xml);
  return zip.generateAsync({type: 'uint8array'});
}

async function main() {
  const buildInfo = JSON.parse(fs.readFileSync(path.join(engineDirectory, 'BUILD-INFO.json'), 'utf8'));
  const installed = JSON.parse(fs.readFileSync(path.join(engineDirectory, 'package.json'), 'utf8'));
  assert.equal(installed.version, buildInfo.package.version);
  for (const [name, record] of Object.entries(buildInfo.artifacts)) {
    assert.equal(crypto.createHash('sha256').update(fs.readFileSync(path.join(engineDirectory, name))).digest('hex'), record.patchedSha256, name);
  }
  const engine = await import(pathToFileURL(path.join(engineDirectory, 'pptx2html_wasm.js')));
  engine.initSync({module: fs.readFileSync(path.join(engineDirectory, 'pptx2html_wasm_bg.wasm'))});
  async function render(bytes) {
    const converted = engine.convert_with_options_metadata(await context.normalize(bytes), true, false, new Uint32Array(), 1);
    try { return {dom: new JSDOM(converted.html), diagnostics: JSON.parse(converted.diagnosticsJson)}; }
    finally {converted.free();}
  }
  const styles = [
    ['self-closing noFill', '<c:spPr><a:noFill/></c:spPr>'],
    ['paired noFill', '<c:spPr><a:noFill></a:noFill></c:spPr>'],
    ['nested solidFill', '<c:spPr><a:solidFill><a:srgbClr val="C0504D"/></a:solidFill></c:spPr>']
  ];
  let equivalent;
  for (const [name, style] of styles) {
    const {dom, diagnostics} = await render(await withChart(chart(style)));
    try {
      const direct = dom.window.document.querySelector('.chart-direct');
      assert.ok(direct?.querySelector('svg'), name);
      assert.equal(dom.window.document.querySelectorAll('.chart-placeholder').length, 0, name);
      assert.equal(diagnostics.filter(d => /CHART.*(?:FALLBACK|UNSUPPORTED)/.test(d.code)).length, 0, name);
      const heights = [...direct.querySelectorAll('.chart-bar')].map(bar => Number(bar.getAttribute('height')));
      assert.equal(heights.length, 3, name);
      assert.ok(heights.every(height => Number.isFinite(height) && height > 0), name);
      assert.ok(Math.abs(heights[1] / heights[0] - 2) < 0.001, name + ': values 2 and 4');
      assert.ok(Math.abs(heights[2] / heights[0] - 1.5) < 0.001, name + ': values 2 and 3');
      for (const label of ['판매량', '월', '화', '수']) assert.ok(direct.textContent.includes(label), name + ': ' + label);
      if (name === 'self-closing noFill') equivalent = direct.outerHTML;
      if (name === 'paired noFill') assert.equal(direct.outerHTML, equivalent, 'equivalent XML must render identically');
    } finally {dom.window.close();}
  }
  const aliased = chart(styles[2][1]).replace(/c:/g, 'chart:').replace('xmlns:c=', 'xmlns:chart=').replace(/a:/g, 'draw:').replace('xmlns:a=', 'xmlns:draw=');
  const aliasResult = await render(await withChart(aliased));
  try {assert.equal(aliasResult.dom.window.document.querySelectorAll('.chart-direct svg').length, 1);}
  finally {aliasResult.dom.window.close();}
  for (const bytes of [fixture('complex.pptx'), await withChart(chart(styles[2][1], '<c:axId val="30"/>'))]) {
    const {dom, diagnostics} = await render(bytes);
    try {
      assert.equal(dom.window.document.querySelectorAll('.chart-direct').length, 0, 'undefined axes must retain fallback');
      assert.equal(diagnostics.filter(d => d.code === 'CHART_STRUCTURE_UNSUPPORTED').length, 1);
    } finally {dom.window.close();}
  }
  const pie = await render(fixture('basic-chart.pptx'));
  try {assert.equal(pie.dom.window.document.querySelectorAll('.chart-direct svg').length, 1);}
  finally {pie.dom.window.close();}
  console.log('PASS patched chart engine: package hashes, equivalent XML, nested styles, SVG values/labels, namespace aliases, invalid axes and pie');
}
main().catch(error => {console.error(error); process.exitCode = 1;});
