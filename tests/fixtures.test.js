/**
 * Tests contra docx REALES del hospital.
 * Se saltan automáticamente si los archivos no están en tests/fixtures/.
 * Ver tests/fixtures/README.md.
 */

import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

import {
  unpackDocx,
  assertStructuralIntegrity,
  extractText
} from './helpers/docx.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const FIXTURES_DIR = path.join(__dirname, 'fixtures');

const fixturePath = (name) => path.join(FIXTURES_DIR, name);
const fixtureExists = (name) => fs.existsSync(fixturePath(name));
const readFixture = (name) => fs.readFileSync(fixturePath(name));

describe('Fixtures reales — caso A (P.02.03.001 Gestión de información documentada)', () => {
  const refName = 'referente.docx';
  const contentName = 'caso-A.docx';
  const available = fixtureExists(refName) && fixtureExists(contentName);

  if (!available) {
    it.skip('saltado: falta tests/fixtures/' + refName + ' o tests/fixtures/' + contentName, () => {});
    return;
  }

  it('detecta cabecera FHJ propia y la conserva', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    expect(stats.preservedHeaders).toBe(true);
  });

  it('aplica estilos a los párrafos esperados (según changelog: 5 títulos + 111 párrafos)', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    expect(stats.title1).toBeGreaterThanOrEqual(4);
    expect(stats.paragraph).toBeGreaterThanOrEqual(80);
  });

  it('produce un docx con document.xml válido', async () => {
    const { blob } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    const { files } = await unpackDocx(blob);
    const integrity = await assertStructuralIntegrity(files);
    expect(integrity.errors).toEqual([]);
  });
});

describe('Fixtures reales — caso B (Procedimiento analítico Cobas c303)', () => {
  const refName = 'referente.docx';
  const contentName = 'caso-B.docx';
  const available = fixtureExists(refName) && fixtureExists(contentName);

  if (!available) {
    it.skip('saltado: falta tests/fixtures/' + refName + ' o tests/fixtures/' + contentName, () => {});
    return;
  }

  it('inyecta cabeceras del referente (preservedHeaders false)', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      metadata: {
        code: 'P.00.00.000',
        version: 'V.0.1',
        title: 'Procedimiento analítico Cobas c303 [test]'
      },
      outputType: 'nodebuffer'
    });
    expect(stats.preservedHeaders).toBe(false);
  });

  it('personaliza el header con los metadatos aportados', async () => {
    const { blob } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      metadata: {
        code: 'P.00.00.000',
        version: 'V.0.1',
        title: 'Procedimiento analítico Cobas c303 [test]'
      },
      outputType: 'nodebuffer'
    });
    const { files } = await unpackDocx(blob);
    const header2Text = extractText(files['word/header2.xml']);
    expect(header2Text).toContain('P.00.00.000');
    expect(header2Text).toContain('V.0.1');
    expect(header2Text).toContain('Cobas c303');
  });

  it('encuentra cantidades coherentes con el changelog (15+15+465+11 / 5 tablas / 47 figuras)', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      metadata: {
        code: 'P.00.00.000',
        version: 'V.0.1',
        title: 'Cobas c303'
      },
      outputType: 'nodebuffer'
    });
    expect(stats.title1).toBeGreaterThanOrEqual(10);
    expect(stats.titPar).toBeGreaterThanOrEqual(10);
    expect(stats.paragraph).toBeGreaterThanOrEqual(400);
    expect(stats.vignette).toBeGreaterThanOrEqual(5);
    expect(stats.tables).toBeGreaterThanOrEqual(3);
    expect(stats.figures).toBeGreaterThanOrEqual(30);
  });

  it('produce un docx con document.xml válido', async () => {
    const { blob } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      metadata: {
        code: 'P.00.00.000',
        version: 'V.0.1',
        title: 'Cobas c303'
      },
      outputType: 'nodebuffer'
    });
    const { files } = await unpackDocx(blob);
    const integrity = await assertStructuralIntegrity(files);
    expect(integrity.errors).toEqual([]);
  });
});

// ----------------------------------------------------------------------------
// Fase 7 — fixtures reales (gestion-info, recomendaciones, pnt-claude)
// ----------------------------------------------------------------------------

describe('Fase 7 — stress real: ref-gestion-info + content-pnt-claude', () => {
  const refName = 'ref-gestion-info.docx';
  const contentName = 'content-pnt-claude.docx';
  const available = fixtureExists(refName) && fixtureExists(contentName);
  if (!available) {
    it.skip('saltado: faltan tests/fixtures/' + refName + ' o ' + contentName, () => {});
    return;
  }

  it('classifier no se dispara: FHJTtulo1 ≤ 30 (antes de Fase 7 daba ~147)', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    // Pre-Fase 7 reportaba 147 títulos por la regex débil. Tras Fase 7 esperamos ≤ 30.
    expect(stats.title1).toBeLessThanOrEqual(30);
  });

  it('preserva las 47 imágenes del content', async () => {
    const { blob } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    const { files } = await unpackDocx(blob);
    const drawings = (files['word/document.xml'].match(/<w:drawing/g) || []).length;
    expect(drawings).toBe(47);
  });

  it('emite warnings normativos sobre el contenido stress', async () => {
    const { warnings } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(contentName),
      outputType: 'nodebuffer'
    });
    const codes = warnings.map(w => w.code);
    // El stress doc tiene runs subrayados y fuentes no-Arial → ambos warnings esperados.
    expect(codes).toContain('NORMATIVA_UNDERLINE');
    expect(codes).toContain('NORMATIVA_FONT_NON_ARIAL');
  });
});

describe('Fase 7 — autoconsistencia: ref vs ref', () => {
  const refName = 'ref-gestion-info.docx';
  const otherName = 'ref-recomendaciones.docx';
  const available = fixtureExists(refName) && fixtureExists(otherName);
  if (!available) {
    it.skip('saltado: faltan tests/fixtures/' + refName + ' o ' + otherName, () => {});
    return;
  }

  it('procesa sin error y preserva mayoría de estilos FHJ pre-existentes', async () => {
    const { stats } = await IsoformaEngine.process({
      refFile: readFixture(refName),
      contentFile: readFixture(otherName),
      outputType: 'nodebuffer'
    });
    // El ref-recomendaciones ya viene perfectamente formateado, casi todo
    // el body trae pStyle FHJ* → preservedStyles debe ser elevado.
    expect(stats.preservedStyles).toBeGreaterThanOrEqual(100);
  });
});
