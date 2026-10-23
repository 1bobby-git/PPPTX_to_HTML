#!/usr/bin/env node
'use strict';
// Synthetic OOXML only: no Office runtime, external inputs, or network access.
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');
const PptxGenJS = require('pptxgenjs');
const OUT = path.resolve(__dirname, '../tests/fixtures');
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const px = n => Math.round(n * 9525);
const groupBase = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>';
const doc = (tag, body, attrs = '') => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:' + tag + ' xmlns:a="' + A + '" xmlns:r="' + R + '" xmlns:p="' + P + '" ' + attrs + '>' + body + '</p:' + tag + '>';
const slide = body => doc('sld', '<p:cSld><p:spTree>' + groupBase + body + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>');
const rels = entries => '<Relationships xmlns="' + REL + '">' + entries.map(([id, type, target]) => '<Relationship Id="' + id + '" Type="' + R + '/' + type + '" Target="' + target + '"/>').join('') + '</Relationships>';
function xfrm(x, y, w, h, child) {
  return '<a:xfrm><a:off x="' + px(x) + '" y="' + px(y) + '"/><a:ext cx="' + px(w) + '" cy="' + px(h) + '"/>' + (child ? '<a:chOff x="' + px(child[0]) + '" y="' + px(child[1]) + '"/><a:chExt cx="' + px(child[2]) + '" cy="' + px(child[3]) + '"/>' : '') + '</a:xfrm>';
}
function rect(id, transform, text = '') {
  return '<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="Fidelity rectangle ' + id + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' + transform + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="11AA55"/></a:solidFill></p:spPr>' + (text ? '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1000"/><a:t>' + text + '</a:t></a:r></a:p></p:txBody>' : '') + '</p:sp>';
}
function group(id, transform, children) {
  return '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="' + id + '" name="Fidelity group ' + id + '"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>' + transform + '</p:grpSpPr>' + children + '</p:grpSp>';
}
async function base(slides) {
  const pptx = new PptxGenJS(); pptx.layout = 'LAYOUT_WIDE'; pptx.author = 'Synthetic fidelity regression';
  slides.forEach(() => pptx.addSlide());
  const zip = await JSZip.loadAsync(await pptx.write({outputType: 'nodebuffer'}));
  slides.forEach((body, i) => zip.file('ppt/slides/slide' + (i + 1) + '.xml', slide(body)));
  return zip;
}
async function save(zip, name) {
  fs.mkdirSync(OUT, {recursive: true});
  await fs.promises.writeFile(path.join(OUT, name), await zip.generateAsync({type: 'nodebuffer', compression: 'DEFLATE'}));
}
async function makeGroups() {
  // Child (15,4,10,3) -> inner (50,29,20,6) -> outer (180,107,40,18).
  const nested = group(2, xfrm(100,80,400,180,[10,20,200,60]),
    group(3, xfrm(30,25,80,20,[5,2,40,10]), rect(4, xfrm(15,4,10,3), 'NESTED_GROUP')));
  await save(await base([nested]), 'fidelity-groups.pptx');
}
function tc(text, attrs = '') {
  return '<a:tc ' + attrs + '><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>' + text + '</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
}
async function makeMergedTable() {
  const rows = [
    tc('MERGED_2_BY_2', 'gridSpan="2" rowSpan="2"') + tc('', 'hMerge="1" rowSpan="2"') + tc('R1C3'),
    tc('', 'vMerge="1" gridSpan="2"') + tc('', 'hMerge="1" vMerge="1"') + tc('R2C3'),
    tc('R3C1') + tc('R3C2') + tc('R3C3')
  ];
  const table = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Combined span table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="' + px(50) + '" y="' + px(60) + '"/><a:ext cx="' + px(300) + '" cy="' + px(150) + '"/></p:xfrm><a:graphic><a:graphicData uri="' + A + '/table"><a:tbl><a:tblPr/><a:tblGrid>' + [0,1,2].map(() => '<a:gridCol w="' + px(100) + '"/>').join('') + '</a:tblGrid>' + rows.map(row => '<a:tr h="' + px(50) + '">' + row + '</a:tr>').join('') + '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>';
  await save(await base([table]), 'fidelity-merged-table.pptx');
}
function placeholder(text, transform = '') {
  return '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Inherited body"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr>' + transform + '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="ko-KR"/><a:t>' + text + '</a:t></a:r></a:p></p:txBody></p:sp>';
}
async function makeThemes() {
  const zip = await base([placeholder('상속 한국어 THEME_ONE'), placeholder('继承中文 THEME_TWO')]);
  const originalTheme = await zip.file('ppt/theme/theme1.xml').async('string');
  const colors = ['1122AA', 'BB3311'], fonts = ['Malgun Gothic', 'Microsoft YaHei'];
  let types = await zip.file('[Content_Types].xml').async('string');
  for (let i = 1; i <= 2; i++) {
    const theme = originalTheme.replace(/<a:accent1>[\s\S]*?<\/a:accent1>/, '<a:accent1><a:srgbClr val="' + colors[i-1] + '"/></a:accent1>').replace(/<a:minorFont>[\s\S]*?<\/a:minorFont>/, '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface="' + fonts[i-1] + '"/><a:cs typeface=""/></a:minorFont>');
    zip.file('ppt/theme/theme' + i + '.xml', theme);
    const style = '<p:txStyles><p:titleStyle/><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/></a:defRPr></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>';
    zip.file('ppt/slideMasters/slideMaster' + i + '.xml', doc('sldMaster', '<p:cSld><p:spTree>' + groupBase + placeholder('',xfrm(60,70,500,100)) + '</p:spTree></p:cSld><p:clrMap accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" bg1="lt1" bg2="lt2" folHlink="folHlink" hlink="hlink" tx1="dk1" tx2="dk2"/><p:sldLayoutIdLst><p:sldLayoutId id="' + (2147483648+i) + '" r:id="rId1"/></p:sldLayoutIdLst>' + style));
    zip.file('ppt/slideMasters/_rels/slideMaster' + i + '.xml.rels', rels([['rId1','slideLayout','../slideLayouts/slideLayout' + i + '.xml'],['rId2','theme','../theme/theme' + i + '.xml']]));
    zip.file('ppt/slideLayouts/slideLayout' + i + '.xml', doc('sldLayout','<p:cSld><p:spTree>' + groupBase + placeholder('',xfrm(60,70,500,100)) + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>', 'type="obj" preserve="1"'));
    zip.file('ppt/slideLayouts/_rels/slideLayout' + i + '.xml.rels', rels([['rId1','slideMaster','../slideMasters/slideMaster' + i + '.xml']]));
    zip.file('ppt/slides/_rels/slide' + i + '.xml.rels', rels([['rId1','slideLayout','../slideLayouts/slideLayout' + i + '.xml']]));
    for (const [part, contentType] of [['slideMasters/slideMaster'+i+'.xml','application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml'],['slideLayouts/slideLayout'+i+'.xml','application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml'],['theme/theme'+i+'.xml','application/vnd.openxmlformats-officedocument.theme+xml']]) {
      if (!types.includes('PartName="/ppt/' + part + '"')) types=types.replace('</Types>','<Override PartName="/ppt/' + part + '" ContentType="'+contentType+'"/></Types>');
    }
  }
  zip.file('[Content_Types].xml',types);
  let pres = await zip.file('ppt/presentation.xml').async('string');
  pres=pres.replace(/<p:sldMasterIdLst>[\s\S]*?<\/p:sldMasterIdLst>/,'<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdFidelityMaster1"/><p:sldMasterId id="2147483649" r:id="rIdFidelityMaster2"/></p:sldMasterIdLst>');
  zip.file('ppt/presentation.xml',pres);
  let pr = await zip.file('ppt/_rels/presentation.xml.rels').async('string');
  pr=pr.replace(/<Relationship\b[^>]*Type="[^"]*\/slideMaster"[^>]*\/>/g,'');
  pr=pr.replace('</Relationships>',[1,2].map(i=>'<Relationship Id="rIdFidelityMaster'+i+'" Type="'+R+'/slideMaster" Target="slideMasters/slideMaster'+i+'.xml"/>').join('')+'</Relationships>');
  zip.file('ppt/_rels/presentation.xml.rels',pr);
  await save(zip,'fidelity-themes.pptx');
}
async function main() { await makeGroups(); await makeMergedTable(); await makeThemes(); console.log('Created three synthetic fidelity fixtures.'); }
if (require.main === module) main().catch(error => {console.error(error);process.exitCode=1;});
module.exports = {main};
