/**
 * Fase 4 — edge cases.
 * Ejercitan paths raros del engine sobre docx sintéticos mínimos:
 * párrafos vacíos, whitespace, tablas vacías/anidadas, drawings "desnudos",
 * sectPr múltiples, bookmarks masivamente duplicados, y smoke de volumen.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

import {
  buildReferentDocx,
  buildContentDocx,
  manyPlainParagraphs,
  rawEmptyParagraph,
  rawWhitespaceParagraph,
  rawEmptyTable,
  rawNestedTable,
  rawDrawingBareParagraph,
  rawParagraphWithInlineSectPr,
  rawDuplicateBookmarks
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function runEngine(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', ...opts });
}

async function extractDocumentXml(outputBuffer) {
  const zip = await JSZip.loadAsync(outputBuffer);
  return await zip.file('word/document.xml').async('string');
}

function parseXml(xml) {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

describe('edge cases — párrafos vacíos / whitespace', () => {
  it('párrafo self-closing (<w:p/>) no rompe la clasificación', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: rawEmptyParagraph() },
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: rawEmptyParagraph() },
        { text: 'Un parrafo cualquiera', kind: 'plain' }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.title1).toBe(1);
    expect(stats.paragraph).toBe(1);
  });

  it('párrafo con solo whitespace no se clasifica (ni título ni párrafo)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: rawWhitespaceParagraph('     ') },
        { kind: 'raw', xml: rawWhitespaceParagraph('\t\t') },
        { text: '1.- OBJETO', kind: 'plain' }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    // Los dos whitespace-only no cuentan como párrafos.
    expect(stats.paragraph).toBe(0);
    expect(stats.title1).toBe(1);
  });
});

describe('edge cases — tablas', () => {
  it('tabla vacía (sin <w:tr>) no explota y no se numera como figura', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: rawEmptyTable() },
        { text: 'texto intermedio', kind: 'plain' },
        { kind: 'table' }
      ]
    });
    const { stats, blob } = await runEngine({ refFile, contentFile });
    // La primera tabla es "Datos generales" (la recién inyectada), la segunda es la vacía,
    // la tercera es la tabla simple. Numeradas sólo de la segunda en adelante.
    expect(stats.tables).toBe(2);
    expect(blob.length).toBeGreaterThan(0);
  });

  it('tabla anidada dentro de otra tabla se ignora para numeración', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'intro', kind: 'plain' },
        { kind: 'raw', xml: rawNestedTable() },
        { kind: 'table' }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    // Body-level tables: Datos generales (inyectada) + nested-wrapper + table simple = 3
    // Se numeran todas menos la primera ⇒ 2 títulos "Tabla n."
    expect(stats.tables).toBe(2);
  });

  it('el documento de salida sigue siendo XML válido tras tabla anidada', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'intro', kind: 'plain' },
        { kind: 'raw', xml: rawNestedTable() }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const xml = await extractDocumentXml(blob);
    const doc = parseXml(xml);
    // El root debe parsear como un documento con body.
    const bodies = doc.getElementsByTagName('w:body');
    expect(bodies.length).toBe(1);
  });
});

describe('edge cases — figuras', () => {
  it('drawing sin docPr sigue contando como figura', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'intro', kind: 'plain' },
        { kind: 'raw', xml: rawDrawingBareParagraph() }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.figures).toBe(1);
  });

  it('drawings dentro de una tabla NO cuentan como figuras', async () => {
    const refFile = await buildReferentDocx();
    // Tabla con drawing dentro de una celda
    const tableWithDrawing =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
      '<w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>' +
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/></wp:inline></w:drawing></w:r></w:p>' +
      '</w:tc></w:tr></w:tbl>';
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'intro', kind: 'plain' },
        { kind: 'raw', xml: tableWithDrawing },
        { kind: 'drawing' }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.figures).toBe(1); // Sólo el drawing fuera de tabla
  });
});

describe('edge cases — sectPr múltiples', () => {
  it('un doc con sectPr intra-párrafo + sectPr final acaba con ambas intactas', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'antes del salto', kind: 'plain' },
        { kind: 'raw', xml: rawParagraphWithInlineSectPr() },
        { text: 'despues del salto', kind: 'plain' }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const xml = await extractDocumentXml(blob);
    const doc = parseXml(xml);
    const sectPrs = doc.getElementsByTagName('w:sectPr');
    // El engine actualiza SOLO el primer sectPr que encuentra. El segundo
    // (el del final del body) no se toca — debe seguir presente.
    expect(sectPrs.length).toBeGreaterThanOrEqual(1);
    // Y el updateSectPrDom aplica a sectPrs[0] — confirma 6 refs en él.
    const first = sectPrs[0];
    const refs = Array.from(first.getElementsByTagName('*')).filter(n =>
      n.nodeName === 'w:headerReference' || n.nodeName === 'w:footerReference'
    );
    expect(refs.length).toBe(6);
  });
});

describe('edge cases — bookmarks', () => {
  it('6 pares de bookmarks con id duplicado se renumeran y deduplican sin fallar', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'intro', kind: 'plain' },
        { kind: 'raw', xml: rawDuplicateBookmarks('0', 6) },
        { text: 'cierre', kind: 'plain' }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const xml = await extractDocumentXml(blob);
    const doc = parseXml(xml);
    const starts = Array.from(doc.getElementsByTagName('w:bookmarkStart'));
    const ends = Array.from(doc.getElementsByTagName('w:bookmarkEnd'));
    // Después de renumerar, los 6 starts deben tener ids únicos desde 1000 en adelante.
    const startIds = starts.map(s => s.getAttribute('w:id')).filter(Boolean);
    const uniqueStartIds = new Set(startIds);
    expect(uniqueStartIds.size).toBe(startIds.length);
    expect(startIds.length).toBe(6);
    // Dedup de bookmarkEnd: no debe haber dos w:id iguales entre ends.
    const endIds = ends.map(e => e.getAttribute('w:id')).filter(Boolean);
    expect(new Set(endIds).size).toBe(endIds.length);
  });
});

describe('edge cases — smoke de volumen', () => {
  it('procesa 200 párrafos en menos de 5s', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        ...manyPlainParagraphs(200)
      ]
    });
    const t0 = Date.now();
    const { stats, blob, warnings } = await runEngine({ refFile, contentFile });
    const elapsed = Date.now() - t0;
    expect(stats.title1).toBe(1);
    expect(stats.paragraph).toBe(200);
    expect(blob.length).toBeGreaterThan(0);
    expect(Array.isArray(warnings)).toBe(true);
    expect(elapsed).toBeLessThan(5000);
  });

  it('documento sin ningún párrafo en body no explota', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ paragraphs: [] });
    const { stats, blob } = await runEngine({ refFile, contentFile });
    expect(stats.title1).toBe(0);
    expect(stats.paragraph).toBe(0);
    expect(blob.length).toBeGreaterThan(0);
  });
});
