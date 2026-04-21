/**
 * Fase 8 — tests del auto-fix normativo.
 *
 * Bloque A: applyNormativaFixesDom corrige underline, fuentes no-Arial,
 *           ALL-CAPS de cuerpo y listas vacías antes del validador.
 *
 * El engine tiene autoFix=false por defecto (backward compat). Las tests de
 * Fase 7 siguen verdes porque el validador ve el documento sin corregir.
 * Aquí probamos el modo autoFix=true explícito.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

import {
  buildReferentDocx,
  buildContentDocx
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function runFixed(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', autoFix: true, ...opts });
}
async function runPlain(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', ...opts });
}

async function unpack(buf) {
  const zip = await JSZip.loadAsync(buf);
  const out = {};
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (path.endsWith('.xml')) out[path] = await file.async('string');
  }
  return { zip, files: out };
}

// -----------------------------------------------------------------------------
// Backward compatibility: autoFix=false por defecto
// -----------------------------------------------------------------------------

describe('Fase 8 — autoFix opt-in (backward compat)', () => {
  it('sin autoFix: stats.fixes.* son 0 y el body se queda como estaba', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Texto subrayado</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats, warnings } = await runPlain({ refFile, contentFile });
    expect(stats.fixes.underline).toBe(0);
    expect(stats.fixes.font).toBe(0);
    expect(stats.fixes.allCaps).toBe(0);
    expect(stats.fixes.emptyList).toBe(0);
    expect(stats.autoFixApplied).toBe(false);
    // El validador sí se queja, porque nadie arregló nada.
    expect(warnings.map(w => w.code)).toContain('NORMATIVA_UNDERLINE');
  });
});

// -----------------------------------------------------------------------------
// Bloque A: auto-fix corrige cada tipo de infracción
// -----------------------------------------------------------------------------

describe('Fase 8 — autoFix: underline', () => {
  it('elimina <w:u> del body y hace que el warning NO se emita', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Subrayado uno</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="double"/></w:rPr><w:t xml:space="preserve">Subrayado dos</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob, stats, warnings } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBeGreaterThanOrEqual(2);
    expect(stats.autoFixApplied).toBe(true);
    expect(warnings.map(w => w.code)).not.toContain('NORMATIVA_UNDERLINE');
    // El XML resultante no tiene <w:u> con val != none en el body.
    const { files } = await unpack(blob);
    const doc = files['word/document.xml'];
    expect(doc).not.toMatch(/<w:u\s+w:val="(single|double|thick)"/);
  });

  it('<w:u w:val="none"/> no cuenta ni se toca', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t xml:space="preserve">No subrayado explícito</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBe(0);
  });
});

describe('Fase 8 — autoFix: fuentes no-Arial', () => {
  it('reescribe rFonts Times New Roman a Arial y silencia el warning', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t xml:space="preserve">Texto Times.</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t xml:space="preserve">Texto Calibri.</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob, stats, warnings } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.font).toBeGreaterThanOrEqual(2);
    expect(warnings.map(w => w.code)).not.toContain('NORMATIVA_FONT_NON_ARIAL');
    const { files } = await unpack(blob);
    const doc = files['word/document.xml'];
    expect(doc).not.toMatch(/Times New Roman/);
    expect(doc).not.toMatch(/w:ascii="Calibri"/);
  });

  it('rFonts con Arial sólo en ascii no lo altera (ya es Arial)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t xml:space="preserve">Arial ya correcto.</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.font).toBe(0);
  });
});

describe('Fase 8 — autoFix: ALL-CAPS body', () => {
  it('descapitaliza párrafo FHJPrrafo en mayúsculas y silencia el warning', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJPrrafo"/></w:pPr><w:r><w:t xml:space="preserve">ESTE TEXTO ESTÁ ESCRITO ENTERO EN MAYÚSCULAS Y NO DEBERÍA ESTARLO PORQUE LA NORMATIVA DEL HOSPITAL DICE QUE NO.</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob, stats, warnings } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.allCaps).toBeGreaterThanOrEqual(1);
    expect(warnings.map(w => w.code)).not.toContain('NORMATIVA_ALL_CAPS_BODY');
    const { files } = await unpack(blob);
    const doc = files['word/document.xml'];
    // El texto debe empezar con mayúscula pero estar mayormente en minúsculas.
    expect(doc).toMatch(/Este texto está escrito entero/);
    expect(doc).not.toMatch(/ESTE TEXTO ESTÁ ESCRITO/);
  });

  it('NO descapitaliza títulos (FHJTtulo*)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJTtulo1"/></w:pPr><w:r><w:t xml:space="preserve">1.- ESTE ES UN TÍTULO EN MAYÚSCULAS Y DEBE QUEDARSE ASÍ PORQUE ES TÍTULO.</w:t></w:r></w:p>' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob, stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.allCaps).toBe(0);
    const { files } = await unpack(blob);
    // El título sigue en mayúsculas.
    expect(files['word/document.xml']).toMatch(/ESTE ES UN TÍTULO EN MAYÚSCULAS/);
  });
});

describe('Fase 8 — autoFix: listas vacías', () => {
  it('elimina párrafos FHJVieta* vacíos y silencia el warning', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJVietaNivel1"/></w:pPr></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJVietaNivel1"/></w:pPr><w:r><w:t xml:space="preserve">Bullet válido</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJVietaNivel2"/></w:pPr></w:p>' }
      ]
    });
    const { blob, stats, warnings } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.emptyList).toBeGreaterThanOrEqual(2);
    expect(warnings.map(w => w.code)).not.toContain('NORMATIVA_EMPTY_LIST');
    const { files } = await unpack(blob);
    // El bullet con contenido se preserva.
    expect(files['word/document.xml']).toMatch(/Bullet válido/);
  });

  it('párrafo FHJVietaNivel1 sin texto pero con drawing NO se elimina', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJVietaNivel1"/></w:pPr><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/></wp:inline></w:drawing></w:r></w:p>' }
      ]
    });
    const { blob, stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.emptyList).toBe(0);
    const { files } = await unpack(blob);
    // El <w:drawing> sigue presente.
    expect(files['word/document.xml']).toMatch(/<w:drawing/);
  });
});

// -----------------------------------------------------------------------------
// Combo: varias infracciones a la vez
// -----------------------------------------------------------------------------

describe('Fase 8 — autoFix combinado', () => {
  it('un documento con underline + fuentes + all-caps + lista vacía queda limpio tras autoFix', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Subrayado.</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr><w:t xml:space="preserve">Calibri.</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJPrrafo"/></w:pPr><w:r><w:t xml:space="preserve">UN PÁRRAFO LARGO EN MAYÚSCULAS PARA TESTEAR LA NORMATIVA DE FORMATO QUE NO LAS PERMITE EN EL CUERPO.</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJVietaNivel1"/></w:pPr></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats, warnings } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.font).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.allCaps).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.emptyList).toBeGreaterThanOrEqual(1);
    const codes = warnings.map(w => w.code);
    expect(codes).not.toContain('NORMATIVA_UNDERLINE');
    expect(codes).not.toContain('NORMATIVA_FONT_NON_ARIAL');
    expect(codes).not.toContain('NORMATIVA_ALL_CAPS_BODY');
    expect(codes).not.toContain('NORMATIVA_EMPTY_LIST');
  });

  it('documento ya limpio: todos los contadores de fixes son 0', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();  // limpio por defecto
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBe(0);
    expect(stats.fixes.font).toBe(0);
    expect(stats.fixes.allCaps).toBe(0);
    expect(stats.fixes.emptyList).toBe(0);
  });
});
