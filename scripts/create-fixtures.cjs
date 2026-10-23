#!/usr/bin/env node
'use strict';

// Reproducible PPTX inputs for the file:// browser test. No checked-in binary is
// needed: pptxgenjs builds the OOXML, then JSZip changes properties that
// pptxgenjs does not expose (presentation order and slide visibility).
const fs = require('node:fs');
const path = require('node:path');
const pptxgen = require('pptxgenjs');
const JSZip = require('jszip');

const outDir = path.resolve(__dirname, '../tests/fixtures');
// Valid 1x1 RGBA red pixel; all chunk CRCs and the zlib scanline are verified.
const pngData = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/iZk9HQAAAABJRU5ErkJggg==';

async function makeDeck() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.author = 'Offline browser regression test';
  pptx.subject = 'Korean text, master, image, table, shapes, chart, order, hidden slide';
  pptx.title = '오프라인 변환 테스트';
  pptx.lang = 'ko-KR';
  pptx.defineSlideMaster({
    title: 'TEST_MASTER',
    background: { color: 'E9F2FF' },
    objects: [{ text: { text: '마스터 문구', options: { x: 0.35, y: 7.06, w: 3, h: 0.2, fontSize: 8, color: '334455' } } }],
    slideNumber: { x: 12.35, y: 7.05, color: '334455' },
  });

  // File slide1.xml intentionally becomes presentation position 2.
  const intro = pptx.addSlide('TEST_MASTER');
  intro.addText('안녕하세요 오프라인 프레젠테이션', { x: 0.7, y: 0.7, w: 10.8, h: 0.8, fontSize: 28, bold: true, color: '123D66' });
  intro.addText('INTRO_FILE_ONE', { x: 0.8, y: 2, w: 7, h: 0.5, fontSize: 18 });
  intro.addNotes('PRIVATE_SPEAKER_NOTE_MUST_NOT_EXPORT');
  intro.addShape(pptx.ShapeType.rect, { x: 0.8, y: 3.1, w: 3.5, h: 1.3, fill: { color: 'FFB347' }, line: { color: '7F3F00' } });

  // File slide2.xml intentionally becomes presentation position 1.
  const table = pptx.addSlide('TEST_MASTER');
  table.addText('TABLE_FILE_TWO', { x: 0.6, y: 0.4, w: 7, h: 0.6, fontSize: 25 });
  table.addTable([
    [{ text: '항목' }, { text: '수량' }],
    [{ text: '사과' }, { text: '3' }],
    [{ text: '배' }, { text: '5' }],
  ], { x: 0.7, y: 1.4, w: 5.7, h: 2, border: { type: 'solid', color: '1C334D', pt: 1 }, color: '1C334D', fill: 'F9FCFF' });
  table.addShape(pptx.ShapeType.ellipse, { x: 7.5, y: 1.7, w: 2.1, h: 1.3, fill: { color: '55CC88' }, line: { color: '147A40' } });

  const visual = pptx.addSlide('TEST_MASTER');
  visual.addText('IMAGE_CHART_FILE_THREE', { x: 0.6, y: 0.4, w: 9, h: 0.6, fontSize: 24 });
  visual.addImage({ data: pngData, x: 0.9, y: 1.5, w: 2.3, h: 2.3 });
  visual.addChart(pptx.ChartType.bar, [{ name: '판매량', labels: ['월', '화', '수'], values: [2, 4, 3] }], {
    x: 4, y: 1.2, w: 7.3, h: 4.5, showLegend: false, showTitle: true, title: '주간 차트',
  });

  const hidden = pptx.addSlide('TEST_MASTER');
  hidden.addText('HIDDEN_FILE_FOUR 비공개', { x: 0.7, y: 0.8, w: 10, h: 0.8, fontSize: 26 });

  const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'nodebuffer' }));
  let presentation = await zip.file('ppt/presentation.xml').async('string');
  const ids = [...presentation.matchAll(/<p:sldId\b[^>]*\/>/g)].map(m => m[0]);
  if (ids.length !== 4) throw new Error(`Expected four slide IDs, found ${ids.length}`);
  // Swap XML sequence while keeping r:id -> target relationships unchanged.
  presentation = presentation.replace(ids[0], '__SWAP_SLIDE_ID__').replace(ids[1], ids[0]).replace('__SWAP_SLIDE_ID__', ids[1]);
  zip.file('ppt/presentation.xml', presentation);
  let fourth = await zip.file('ppt/slides/slide4.xml').async('string');
  fourth = fourth.replace(/<p:sld\b/, '<p:sld show="0"');
  zip.file('ppt/slides/slide4.xml', fourth);
  const target = path.join(outDir, 'complex.pptx');
  fs.writeFileSync(target, await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
  return target;
}

async function makeHostileDeck() {
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_WIDE';
  const slide = pptx.addSlide();
  slide.addText('안전한 텍스트 <script>window.__pptxExecuted=1</script>', { x: 0.5, y: 0.7, w: 11, h: 0.7 });
  slide.addText('EVENT_TEXT <img src=x onerror="window.__pptxExecuted=2">', { x: 0.5, y: 1.5, w: 11, h: 0.7 });
  const zip = await JSZip.loadAsync(await pptx.write({ outputType: 'nodebuffer' }));
  let rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('string');
  rels = rels.replace('</Relationships>', '<Relationship Id="rIdExternalImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://example.invalid/pptx-leak.png" TargetMode="External"/></Relationships>');
  zip.file('ppt/slides/_rels/slide1.xml.rels', rels);
  fs.writeFileSync(path.join(outDir, 'hostile.pptx'), await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }));
}

async function makeBasicChart() {
  // A pie chart exercises the WASM converter's direct SVG chart renderer.
  // The complex fixture's PptxGenJS bar chart deliberately retains its
  // incompatible third axis reference so fallback handling is tested too.
  const pptx = new pptxgen();
  pptx.layout = 'LAYOUT_WIDE';
  const slide = pptx.addSlide();
  slide.addText('PIE_CHART_VALUES_2_4_3', { x: 0.5, y: 0.3, w: 10, h: 0.6 });
  slide.addChart(pptx.ChartType.pie, [
    { name: 'Sales', labels: ['Mon', 'Tue', 'Wed'], values: [2, 4, 3] },
  ], { x: 1, y: 1, w: 8, h: 4, showLegend: true });
  fs.writeFileSync(path.join(outDir, 'basic-chart.pptx'), await pptx.write({ outputType: 'nodebuffer' }));
}

async function main() {
  fs.mkdirSync(outDir, { recursive: true });
  const deck = await makeDeck();
  await makeHostileDeck();
  await makeBasicChart();
  fs.writeFileSync(path.join(outDir, 'invalid.pptx'), Buffer.from('This is deliberately not an OOXML ZIP file.'));
  console.log(`Created ${deck} and companion fixtures`);
}
main().catch(error => { console.error(error); process.exitCode = 1; });
