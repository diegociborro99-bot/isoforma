/**
 * Golden tests del Isoforma engine con docx sintéticos.
 * No dependen de documentos reales — cubren forma y conteo de estilos.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

import { buildReferentDocx, buildContentDocx } from './helpers/synthetic.js';
import {
  unpackDocx,
  countPStyle,
  isValidXml,
  extractText,
  assertStructuralIntegrity
} from './helpers/docx.js';

// Engine sigue siendo CJS (compatible con el browser). Lo cargamos con createRequire.
const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function runEngine(opts) {
  return IsoformaEngine.process({
    outputType: 'nodebuffer',
    ...opts
  });
}

describe('Isoforma engine — caso B (contenido SIN cabecera FHJ)', () => {
  it('inyecta cabeceras/pies del referente y stats.preservedHeaders es false', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob, stats } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    expect(stats.preservedHeaders).toBe(false);
    expect(files['word/header1.xml']).toBeDefined();
    expect(files['word/header2.xml']).toBeDefined();
    expect(files['word/header3.xml']).toBeDefined();
    expect(files['word/footer1.xml']).toBeDefined();
    expect(files['word/footer2.xml']).toBeDefined();
    expect(files['word/footer3.xml']).toBeDefined();
  });

  it('aplica estilos FHJ al cuerpo con los contadores esperados', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { stats } = await runEngine({ refFile, contentFile });

    expect(stats.title1).toBe(4);
    expect(stats.titPar).toBe(2);
    // 3 párrafos descriptivos del cuerpo + 2 de celdas de la tabla
    // (el engine aplica estilos a párrafos DENTRO de tablas — ver known issues).
    expect(stats.paragraph).toBe(5);
    expect(stats.vignette).toBe(2);
  });

  it('numera una tabla y una figura', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.tables).toBe(1);
    expect(stats.figures).toBe(1);
  });

  it('inyecta la tabla de Datos generales (estilos específicos)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);
    const docXml = files['word/document.xml'];

    expect(countPStyle(docXml, 'FHJTitulodatosgenerales')).toBe(7);
    expect(countPStyle(docXml, 'FHJContenidodatosgenerales')).toBe(7);
    expect(docXml).toContain('Datos generales');
  });

  it('personaliza el header2 cuando se aportan metadatos completos', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const metadata = {
      code: 'P.99.99.999',
      version: 'V.1.2',
      title: 'Procedimiento de prueba sintético'
    };
    const { blob } = await runEngine({ refFile, contentFile, metadata });
    const { files } = await unpackDocx(blob);

    const header2Text = extractText(files['word/header2.xml']);
    expect(header2Text).toContain('P.99.99.999');
    expect(header2Text).toContain('V.1.2');
    expect(header2Text).toContain('Procedimiento de prueba sintético');
  });

  it('no personaliza el header2 cuando los metadatos están incompletos', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const metadata = { code: 'P.01.00.001', version: '', title: '' };
    const { blob } = await runEngine({ refFile, contentFile, metadata });
    const { files } = await unpackDocx(blob);

    expect(files['word/header2.xml']).toContain('[CODIGO / VERSION]');
    expect(files['word/header2.xml']).toContain('[TITULO DEL PROCEDIMIENTO]');
  });

  it('renombra el logo con prefijo fhj_ y actualiza header2.xml.rels', async () => {
    // Fase 7 (Bloque C): el logo se importa como fhj_<basename> para evitar
    // colisión con cualquier media/image1.emf que el contenido ya traiga.
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob } = await runEngine({ refFile, contentFile });
    const { zip, files } = await unpackDocx(blob);

    expect(zip.file('word/media/fhj_image1.emf')).not.toBeNull();
    expect(files['word/_rels/header2.xml.rels']).toContain('fhj_image1.emf');
  });

  it('añade las 6 relaciones header/footer en document.xml.rels', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    const relsXml = files['word/_rels/document.xml.rels'];
    const headerRels = (relsXml.match(/Type="[^"]+\/header"/g) || []).length;
    const footerRels = (relsXml.match(/Type="[^"]+\/footer"/g) || []).length;
    expect(headerRels).toBe(3);
    expect(footerRels).toBe(3);
  });

  it('el document.xml resultante es XML válido y tiene un único <w:body>', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    const integrity = await assertStructuralIntegrity(files);
    expect(integrity.errors).toEqual([]);
    expect(integrity.ok).toBe(true);
  });

  it('registra los content types de header/footer en [Content_Types].xml', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    const ct = files['[Content_Types].xml'];
    for (let i = 1; i <= 3; i++) {
      expect(ct).toContain('PartName="/word/header' + i + '.xml"');
      expect(ct).toContain('PartName="/word/footer' + i + '.xml"');
    }
    expect(ct).toContain('Extension="emf"');
  });
});

describe('Isoforma engine — caso A (contenido CON cabecera FHJ propia)', () => {
  it('detecta la cabecera FHJ y reporta preservedHeaders true', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.preservedHeaders).toBe(true);
  });

  it('conserva el drawing original de la cabecera del contenido', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    expect(files['word/header1.xml']).toContain('<w:drawing>');
    expect(files['word/header1.xml']).toContain('cabecera propia');
    expect(files['word/header1.xml']).not.toContain('[CODIGO / VERSION]');
  });

  it('aplica los estilos al body igual que en caso B', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.title1).toBe(4);
    expect(stats.titPar).toBe(2);
    expect(stats.paragraph).toBe(5); // incluye 2 párrafos de celdas de tabla
    expect(stats.vignette).toBe(2);
  });

  it('también inyecta Datos generales y numera tabla/figura', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { blob, stats } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);
    expect(files['word/document.xml']).toContain('Datos generales');
    expect(stats.tables).toBe(1);
    expect(stats.figures).toBe(1);
  });

  it('produce document.xml válido', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);
    const integrity = await assertStructuralIntegrity(files);
    expect(integrity.errors).toEqual([]);
  });
});

describe('Isoforma engine — clasificación de párrafos (variantes)', () => {
  it('clasifica correctamente variantes de numeración de título', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: '1.- CON GUIÓN', kind: 'plain' },          // FHJTtulo1
      { text: '2. SIN GUIÓN', kind: 'plain' },           // FHJTtulo1
      { text: 'ANEXO II', kind: 'plain' },               // FHJTtulo1
      { text: '1.1. Con punto', kind: 'plain' },         // FHJTtuloprrafo
      { text: '1.1.- Con punto-guión', kind: 'plain' },  // FHJTtuloprrafo
      { text: '1.1.1.- Tres niveles', kind: 'plain' },   // FHJTtuloprrafo
      { text: '• Viñeta con bullet', kind: 'plain' },    // FHJVietaNivel1
      { text: '· Viñeta con medio', kind: 'plain' },     // FHJVietaNivel1
      { text: '* Viñeta asterisco', kind: 'plain' },     // FHJVietaNivel1
      { text: 'Texto normal de párrafo', kind: 'plain' } // FHJPrrafo
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });

    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.title1).toBe(3);
    expect(stats.titPar).toBe(3);
    expect(stats.vignette).toBe(3);
    expect(stats.paragraph).toBe(1);
  });

  it('no aplica estilo a párrafos vacíos', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: '', kind: 'plain' },
      { text: '   ', kind: 'plain' },
      { text: 'Texto real', kind: 'plain' }
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });

    const { stats } = await runEngine({ refFile, contentFile });
    const total = stats.title1 + stats.titPar + stats.paragraph + stats.vignette;
    expect(total).toBe(1);
    expect(stats.paragraph).toBe(1);
  });

  // Antes era it.fails (Fase 1). Fase 2 lo arregla: el clasificador permite
  // punto final opcional en subtítulos multinivel ("3.1 Foo", "3.1. Foo", "3.1.- Foo").
  it('acepta "3.1 Subtítulo" sin punto final como FHJTtuloprrafo', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: '3.1 Subtítulo sin punto final', kind: 'plain' },
      { text: '3.1- Subtítulo con solo guión', kind: 'plain' }
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.titPar).toBe(2);
  });
});

describe('Isoforma engine — robustez estructural (Fase 2)', () => {
  // Antes era it.fails. Al pasar a DOM, el engine ya no depende de cómo Word
  // serialice <w:tbl>; el nodo se detecta igual con o sin atributos.
  it('cuenta tablas con atributos ("<w:tbl w:rsidR=...>") igual que las literales', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: '1.- OBJETO', kind: 'plain' },
      { text: 'Párrafo', kind: 'plain' },
      { kind: 'table' }, // sin atributos
      { kind: 'table', attrs: 'w:rsidR="00112233" w:rsidTr="00445566"' },
      { kind: 'drawing' }
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });
    const { stats } = await runEngine({ refFile, contentFile });
    // La primera w:tbl del documento tras Datos generales es la que se salta;
    // la segunda (con atributos) y la tercera se numeran → 2 tablas numeradas.
    expect(stats.tables).toBe(2);
    expect(stats.figures).toBe(1);
  });

  it('parse-once DOM: document.xml resultante es XML válido tras todas las mutaciones', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: '1.- OBJETO', kind: 'plain' },
      { text: '1.1 Sin punto final', kind: 'plain' },
      { text: '1.1.- Con punto-guión', kind: 'plain' },
      { kind: 'table', attrs: 'w:rsidR="DEADBEEF"' },
      { kind: 'drawing' },
      { kind: 'drawing' }
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });
    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    const integrity = await assertStructuralIntegrity(files);
    expect(integrity.errors).toEqual([]);
    expect(files['word/document.xml']).toBeDefined();
    const isValid = isValidXml(files['word/document.xml']);
    expect(isValid.valid).toBe(true);
  });

  it('sectPr actualizado conserva exactamente 6 references (3 header + 3 footer)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });
    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);
    const docXml = files['word/document.xml'];

    const headerRefs = (docXml.match(/<w:headerReference\b/g) || []).length;
    const footerRefs = (docXml.match(/<w:footerReference\b/g) || []).length;
    expect(headerRefs).toBe(3);
    expect(footerRefs).toBe(3);
    // Parámetros de página fijos del PNT
    expect(docXml).toMatch(/w:w="11906"/);
    expect(docXml).toMatch(/w:h="16838"/);
  });
});

describe('Isoforma engine — idempotencia / Datos generales', () => {
  it('si el documento ya contiene "Datos generales" al inicio, no la duplica', async () => {
    const refFile = await buildReferentDocx();
    const paragraphs = [
      { text: 'Datos generales', kind: 'plain' },
      { text: '1.- OBJETO', kind: 'plain' }
    ];
    const contentFile = await buildContentDocx({ withFhjHeader: false, paragraphs });

    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpackDocx(blob);

    const occurrences = (files['word/document.xml'].match(/Datos generales/g) || []).length;
    expect(occurrences).toBe(1);
  });
});
