/**
 * Fase 9 — tests para extracción de metadatos y samples de auto-fix.
 *
 * Bloque A: extractMetadata alias + patrones de versión extendidos
 *           ("Versión N", "Edición N", "Rev. N").
 * Bloque C: stats.fixes.samples.* contiene snippets de texto para que la UI
 *           pueda mostrar evidencia concreta de qué corrigió el auto-fix.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

import {
  buildReferentDocx,
  buildContentDocx
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');
const { extractMetadata, inspectContent } = IsoformaEngine;

async function runFixed(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', autoFix: true, ...opts });
}

// -----------------------------------------------------------------------------
// Bloque A: extractMetadata alias + patrones de versión
// -----------------------------------------------------------------------------

describe('Fase 9 — extractMetadata alias', () => {
  it('expone extractMetadata que devuelve sólo { code, version, title }', async () => {
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    const meta = await extractMetadata(contentFile);
    expect(meta).toHaveProperty('code');
    expect(meta).toHaveProperty('version');
    expect(meta).toHaveProperty('title');
    expect(meta).not.toHaveProperty('hasFhjHeader');
    expect(meta.code).toBe('P.01.00.001');
    expect(meta.version).toBe('V.0.1');
  });

  it('extractMetadata con input null devuelve nulls sin lanzar', async () => {
    const meta = await extractMetadata(null);
    expect(meta).toEqual({ code: null, version: null, title: null });
  });
});

describe('Fase 9 — patrones de versión extendidos', () => {
  it('detecta "Versión 1.0"', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'PNT P.05.02.015 · Versión 1.0', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' }
      ]
    });
    const meta = await extractMetadata(contentFile);
    expect(meta.version).toMatch(/Versi[oó]n\s+1\.0/i);
  });

  it('detecta "Edición 2"', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Procedimiento P.05.02.016 — Edición 2', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' }
      ]
    });
    const meta = await extractMetadata(contentFile);
    expect(meta.version).toMatch(/Edici[oó]n\s+2/i);
  });

  it('detecta "Rev. 3"', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Código P.05.02.017 Rev. 3', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' }
      ]
    });
    const meta = await extractMetadata(contentFile);
    expect(meta.version).toMatch(/Rev\.?\s+3/i);
  });

  it('sigue detectando el formato clásico "V.1.2"', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'P.05.02.018 V.1.2', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' }
      ]
    });
    const meta = await extractMetadata(contentFile);
    expect(meta.version).toBe('V.1.2');
  });

  it('no confunde código con versión', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'P.05.02.015', kind: 'plain' },
        { text: 'Versión 2.1', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' }
      ]
    });
    const meta = await extractMetadata(contentFile);
    expect(meta.code).toBe('P.05.02.015');
    expect(meta.version).toMatch(/Versi[oó]n\s+2\.1/i);
  });
});

// -----------------------------------------------------------------------------
// Bloque C: samples en auto-fix
// -----------------------------------------------------------------------------

describe('Fase 9 — auto-fix samples', () => {
  it('samples.underline contiene snippet del párrafo del run subrayado', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Texto crítico subrayado aquí</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.samples.underline.length).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.samples.underline[0].text).toContain('Texto crítico');
  });

  it('samples.font incluye el nombre de la fuente original', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t xml:space="preserve">Párrafo en Times</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.font).toBeGreaterThanOrEqual(1);
    const sample = stats.fixes.samples.font[0];
    expect(sample).toBeDefined();
    expect(sample.font).toMatch(/Times/i);
    expect(sample.text).toContain('Times');
  });

  it('samples.allCaps captura el texto ANTES de descapitalizar', async () => {
    const refFile = await buildReferentDocx();
    // >120 chars para que el classifier no lo marque como FHJTtulo1 (all-caps-short)
    const allCapsText = 'ESTE PÁRRAFO DEL CUERPO ESTÁ COMPLETAMENTE ESCRITO EN MAYÚSCULAS Y NO DEBERÍA ESTARLO NUNCA BAJO NINGÚN CONCEPTO NORMATIVO APLICABLE AL PNT';
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: allCapsText, kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.allCaps).toBeGreaterThanOrEqual(1);
    const sample = stats.fixes.samples.allCaps[0];
    expect(sample).toBeDefined();
    // El sample se captura antes de modificar, por lo que aún está en MAYÚSCULAS
    expect(sample.text).toMatch(/[A-ZÁÉÍÓÚÑ]{10,}/);
    expect(sample.text).toContain('MAYÚSCULAS');
  });

  it('samples se limitan a 3 por tipo (no crecen sin control)', async () => {
    const refFile = await buildReferentDocx();
    // Creamos 5 párrafos con subrayado
    const underlineP = '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">Subrayado</w:t></w:r></w:p>';
    const paragraphs = [
      { text: '1.- OBJETO', kind: 'plain' },
      { kind: 'raw', xml: underlineP },
      { kind: 'raw', xml: underlineP },
      { kind: 'raw', xml: underlineP },
      { kind: 'raw', xml: underlineP },
      { kind: 'raw', xml: underlineP },
      { text: '2.- ALCANCE', kind: 'plain' },
      { text: 'Alc.', kind: 'plain' },
      { text: '3.- DESARROLLO', kind: 'plain' }
    ];
    const contentFile = await buildContentDocx({ paragraphs });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.underline).toBe(5);
    expect(stats.fixes.samples.underline.length).toBeLessThanOrEqual(3);
  });

  it('snippet se trunca a 80 chars con ellipsis', async () => {
    const refFile = await buildReferentDocx();
    // Texto largo >> 80 chars
    const longText = 'Este es un texto muy largo que supera con creces los ochenta caracteres permitidos en un snippet de muestra, para forzar truncado';
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t xml:space="preserve">' + longText + '</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    const sample = stats.fixes.samples.underline[0];
    expect(sample).toBeDefined();
    expect(sample.text.length).toBeLessThanOrEqual(80);
    expect(sample.text.endsWith('…')).toBe(true);
  });

  it('sin auto-fix: samples siguen siendo arrays vacíos (nunca undefined)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();
    const { stats } = await IsoformaEngine.process({
      outputType: 'nodebuffer', refFile, contentFile
    });
    expect(stats.fixes.samples).toBeDefined();
    expect(stats.fixes.samples.underline).toEqual([]);
    expect(stats.fixes.samples.font).toEqual([]);
    expect(stats.fixes.samples.allCaps).toEqual([]);
    expect(stats.fixes.samples.emptyList).toEqual([]);
  });
});

// -----------------------------------------------------------------------------
// Bloque A bis: inspectContent retrocompat (no se rompe la API existente)
// -----------------------------------------------------------------------------

describe('Fase 9 — backward compat', () => {
  it('inspectContent sigue devolviendo hasFhjHeader + detected', async () => {
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    const res = await inspectContent(contentFile);
    expect(res.hasFhjHeader).toBe(true);
    expect(res.detected.code).toBe('P.01.00.001');
  });
});
