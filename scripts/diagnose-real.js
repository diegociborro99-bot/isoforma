#!/usr/bin/env node
/**
 * Diagnóstico de Fase 7: pasa el engine actual sobre los pares reales y
 * reporta qué se conserva, qué se pierde y qué se clasifica mal.
 *
 *  node scripts/diagnose-real.js
 *
 * Reporta:
 *   - tamaño in / out
 *   - conteo de imágenes en input vs output
 *   - conteo de párrafos por estilo (input vs output)
 *   - presencia de header/footer en output
 *   - warnings / errores del engine
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { createRequire } = require('node:module');

const ROOT = path.resolve(__dirname, '..');
const FIX = path.join(ROOT, 'tests/fixtures');
const requireFromRoot = createRequire(path.join(ROOT, 'package.json'));
const Engine = requireFromRoot(path.join(ROOT, 'isoforma-engine.js'));
const JSZip = requireFromRoot('jszip');

async function inspectDocx(buf) {
  const zip = await JSZip.loadAsync(buf);
  const files = Object.keys(zip.files);
  const docXml = await zip.file('word/document.xml')?.async('string') || '';
  const styleCounts = {};
  const re = /<w:pStyle w:val="([^"]+)"\/>/g;
  let m;
  while ((m = re.exec(docXml)) !== null) {
    styleCounts[m[1]] = (styleCounts[m[1]] || 0) + 1;
  }
  const paraCount = (docXml.match(/<w:p[\s>]/g) || []).length;
  const tblCount = (docXml.match(/<w:tbl[\s>]/g) || []).length;
  const drawingCount = (docXml.match(/<w:drawing/g) || []).length;
  const pictCount = (docXml.match(/<w:pict/g) || []).length;
  const mediaFiles = files.filter(f => f.startsWith('word/media/'));
  const headers = files.filter(f => /^word\/header\d*\.xml$/.test(f));
  const footers = files.filter(f => /^word\/footer\d*\.xml$/.test(f));
  const numIdRefs = (docXml.match(/<w:numId w:val="(\d+)"/g) || []).length;
  return {
    files: files.length,
    paraCount, tblCount, drawingCount, pictCount,
    mediaFiles: mediaFiles.length, mediaList: mediaFiles,
    headers: headers.length, footers: footers.length,
    headerList: headers, footerList: footers,
    styleCounts, numIdRefs,
    docXmlSize: docXml.length
  };
}

async function diagnose(refName, contentName, label) {
  console.log(`\n========================================`);
  console.log(`CASE: ${label}`);
  console.log(`  ref:     ${refName}`);
  console.log(`  content: ${contentName}`);
  console.log(`========================================`);

  const refBuf = fs.readFileSync(path.join(FIX, refName));
  const contentBuf = fs.readFileSync(path.join(FIX, contentName));

  console.log('\n--- INPUT: ref ---');
  const ref = await inspectDocx(refBuf);
  console.log(JSON.stringify({
    files: ref.files, paraCount: ref.paraCount, tblCount: ref.tblCount,
    drawingCount: ref.drawingCount, pictCount: ref.pictCount,
    mediaFiles: ref.mediaFiles, headers: ref.headers, footers: ref.footers,
    numIdRefs: ref.numIdRefs, docXmlSize: ref.docXmlSize,
    topStyles: Object.entries(ref.styleCounts).sort((a,b)=>b[1]-a[1]).slice(0, 8)
  }, null, 2));

  console.log('\n--- INPUT: content ---');
  const content = await inspectDocx(contentBuf);
  console.log(JSON.stringify({
    files: content.files, paraCount: content.paraCount, tblCount: content.tblCount,
    drawingCount: content.drawingCount, pictCount: content.pictCount,
    mediaFiles: content.mediaFiles, headers: content.headers, footers: content.footers,
    numIdRefs: content.numIdRefs, docXmlSize: content.docXmlSize,
    topStyles: Object.entries(content.styleCounts).sort((a,b)=>b[1]-a[1]).slice(0, 8)
  }, null, 2));

  // Run engine
  console.log('\n--- ENGINE RUN ---');
  let result;
  try {
    result = await Engine.process({
      refFile: refBuf,
      contentFile: contentBuf,
      outputType: 'nodebuffer'
    });
  } catch (err) {
    console.log('ERROR:', err && err.code, '-', err && err.message);
    if (err && err.context) console.log('  context:', err.context);
    return;
  }
  console.log(`  blob size: ${(result.blob.length/1024).toFixed(1)} KB`);
  console.log(`  warnings:  ${(result.warnings||[]).length}`);
  for (const w of (result.warnings || []).slice(0, 20)) {
    console.log(`    [${w.code}] ${w.message || ''}`);
  }

  console.log('\n--- OUTPUT ---');
  const out = await inspectDocx(result.blob);
  console.log(JSON.stringify({
    files: out.files, paraCount: out.paraCount, tblCount: out.tblCount,
    drawingCount: out.drawingCount, pictCount: out.pictCount,
    mediaFiles: out.mediaFiles, headers: out.headers, footers: out.footers,
    numIdRefs: out.numIdRefs, docXmlSize: out.docXmlSize,
    topStyles: Object.entries(out.styleCounts).sort((a,b)=>b[1]-a[1]).slice(0, 12)
  }, null, 2));

  // Diff summary
  console.log('\n--- DIFF (content -> output) ---');
  console.log(`  paragraphs:  ${content.paraCount} -> ${out.paraCount}  (diff ${out.paraCount - content.paraCount})`);
  console.log(`  tables:      ${content.tblCount} -> ${out.tblCount}`);
  console.log(`  drawings:    ${content.drawingCount} -> ${out.drawingCount}  ${out.drawingCount < content.drawingCount ? '⚠️ LOST' : ''}`);
  console.log(`  media files: ${content.mediaFiles} -> ${out.mediaFiles}  ${out.mediaFiles < content.mediaFiles ? '⚠️ LOST' : ''}`);
  console.log(`  headers:     ${ref.headers} (ref) / ${content.headers} (content) -> ${out.headers}  ${out.headers === 0 ? '⚠️ NO HEADER' : ''}`);
  console.log(`  footers:     ${ref.footers} (ref) / ${content.footers} (content) -> ${out.footers}  ${out.footers === 0 ? '⚠️ NO FOOTER' : ''}`);
}

async function main() {
  // Caso 1: el doc canónico como ref + el "pnt claude" como content (stress)
  await diagnose('ref-gestion-info.docx', 'content-pnt-claude.docx', 'Stress: gestion-info vs pnt-claude');

  // Caso 2: la teoría como ref + pnt-claude como content
  await diagnose('ref-recomendaciones.docx', 'content-pnt-claude.docx', 'Theory ref vs pnt-claude');

  // Caso 3: ref vs ref (autoconsistencia — debería salir limpio)
  await diagnose('ref-gestion-info.docx', 'ref-recomendaciones.docx', 'Self-test: ref vs ref');
}

main().catch(err => {
  console.error('Diagnose falló:', err && err.stack || err);
  process.exit(1);
});
