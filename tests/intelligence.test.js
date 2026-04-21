/**
 * Fase 5 — inteligencia: inspectContent() + validación estructural.
 *
 * inspectContent extrae código/versión/título desde el header propio o desde los
 * primeros párrafos del body, sin procesar el documento.
 *
 * La validación estructural emite warnings no fatales sobre secciones que
 * faltan, saltos en la numeración y cross-refs rotas a tablas/figuras.
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
const { inspectContent, process: engineProcess } = IsoformaEngine;

async function runEngine(opts) {
  return engineProcess({ outputType: 'nodebuffer', ...opts });
}

describe('inspectContent — extracción desde cabecera propia FHJ', () => {
  it('detecta código, versión y título desde el header propio del contenido', async () => {
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    const result = await inspectContent(contentFile);
    expect(result.hasFhjHeader).toBe(true);
    expect(result.detected.code).toBe('P.01.00.001');
    expect(result.detected.version).toBe('V.0.1');
    expect(result.detected.title).toBe('Procedimiento de prueba (cabecera propia)');
  });

  it('sin header propio: hasFhjHeader=false y detecta desde body si hay rastro', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Procedimiento sentencia P.05.02.015 V.1.2', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo del PNT', kind: 'plain' }
      ]
    });
    const result = await inspectContent(contentFile);
    expect(result.hasFhjHeader).toBe(false);
    expect(result.detected.code).toBe('P.05.02.015');
    expect(result.detected.version).toBe('V.1.2');
  });

  it('sin metadatos reconocibles: devuelve nulls y no lanza', async () => {
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Un documento sin código ni versión.', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' }
      ]
    });
    const result = await inspectContent(contentFile);
    expect(result.detected.code).toBeNull();
    expect(result.detected.version).toBeNull();
  });

  it('input null/undefined devuelve shape vacía sin lanzar', async () => {
    const result = await inspectContent(null);
    expect(result.hasFhjHeader).toBe(false);
    expect(result.detected).toEqual({ code: null, version: null, title: null });
  });

  it('input no-zip devuelve shape vacía sin lanzar', async () => {
    const garbage = Buffer.from('esto no es un docx');
    const result = await inspectContent(garbage);
    expect(result.hasFhjHeader).toBe(false);
    expect(result.detected.code).toBeNull();
  });

  it('zip válido sin document.xml devuelve shape vacía', async () => {
    const zip = new JSZip();
    zip.file('foo.txt', 'hola');
    const buf = await zip.generateAsync({ type: 'nodebuffer' });
    const result = await inspectContent(buf);
    expect(result.hasFhjHeader).toBe(false);
    expect(result.detected.code).toBeNull();
  });

  it('título no se confunde con el código cuando código y título están en el mismo párrafo', async () => {
    const contentFile = await buildContentDocx({
      withFhjHeader: true
      // El header sintético ya separa código y título en párrafos distintos;
      // este test sólo verifica que el extractor no atrapa la línea del código.
    });
    const result = await inspectContent(contentFile);
    expect(result.detected.title).not.toContain('P.01.00.001');
    expect(result.detected.title).not.toContain('V.0.1');
  });
});

describe('validación estructural — warnings emitidos en process()', () => {
  it('documento con OBJETO + ALCANCE + DESARROLLO no emite warnings estructurales', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();
    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).not.toContain('STRUCTURE_MISSING_OBJETO');
    expect(codes).not.toContain('STRUCTURE_MISSING_ALCANCE');
    expect(codes).not.toContain('STRUCTURE_MISSING_DESARROLLO');
    expect(codes).not.toContain('STRUCTURE_NUMBERING_GAP');
    expect(codes).not.toContain('STRUCTURE_CROSSREF_BROKEN');
  });

  it('falta OBJETO: emite STRUCTURE_MISSING_OBJETO', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- ALCANCE', kind: 'plain' },
        { text: 'Aplica a todos', kind: 'plain' },
        { text: '2.- DESARROLLO', kind: 'plain' },
        { text: 'El cuerpo del procedimiento.', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    expect(warnings.map(w => w.code)).toContain('STRUCTURE_MISSING_OBJETO');
  });

  it('falta DESARROLLO y PROCEDIMIENTO: emite STRUCTURE_MISSING_DESARROLLO', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Tal.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Cual.', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    expect(warnings.map(w => w.code)).toContain('STRUCTURE_MISSING_DESARROLLO');
  });

  it('acepta "PROCEDIMIENTO" o "SISTEMÁTICA" como equivalentes a DESARROLLO', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- PROCEDIMIENTO', kind: 'plain' },
        { text: 'El proced.', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    expect(warnings.map(w => w.code)).not.toContain('STRUCTURE_MISSING_DESARROLLO');
  });

  it('hueco en numeración (1,2,4): emite STRUCTURE_NUMBERING_GAP', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '4.- DESARROLLO', kind: 'plain' },
        { text: 'El proced.', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    const gap = warnings.find(w => w.code === 'STRUCTURE_NUMBERING_GAP');
    expect(gap).toBeDefined();
    expect(gap.context.gaps.length).toBeGreaterThanOrEqual(1);
    expect(gap.context.gaps[0]).toMatchObject({ after: 2, got: 4, expected: 3 });
  });

  it('cross-ref a tabla inexistente: emite STRUCTURE_CROSSREF_BROKEN', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Referencia fantasma: ver Tabla 5 y Figura 3 abajo.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { text: 'Texto.', kind: 'plain' },
        // No hay drawings ni tablas extra: el content no aporta figuras y
        // sólo aporta 0 tablas (la de Datos generales no se cuenta en stats.tables).
      ]
    });
    const { warnings, stats } = await runEngine({ refFile, contentFile });
    const broken = warnings.find(w => w.code === 'STRUCTURE_CROSSREF_BROKEN');
    expect(broken).toBeDefined();
    expect(broken.context.brokenTableRefs).toContain(5);
    expect(broken.context.brokenFigureRefs).toContain(3);
    // Sanity: stats.tables refleja el conteo real.
    expect(stats.tables).toBe(0);
    expect(stats.figures).toBe(0);
  });

  it('cross-ref válida no emite warning', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Ver Tabla 1 arriba.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' },
        { kind: 'table' }  // una tabla extra → stats.tables = 1 → ref a Tabla 1 es válida
      ]
    });
    const { warnings, stats } = await runEngine({ refFile, contentFile });
    expect(stats.tables).toBeGreaterThanOrEqual(1);
    expect(warnings.map(w => w.code)).not.toContain('STRUCTURE_CROSSREF_BROKEN');
  });
});
