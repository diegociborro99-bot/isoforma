/**
 * Fase 7 — tests de las mejoras del clasificador, transferencia de
 * headers/footers y validador normativo.
 *
 *   Bloque A: classifier feature-based + listas + trust path
 *   Bloque C: transferencia de headers/footers desde el ref
 *   Bloque D: warnings normativos (ALL-CAPS, subrayado, fuentes)
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

async function runEngine(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', ...opts });
}

async function unpack(buf) {
  const zip = await JSZip.loadAsync(buf);
  const out = {};
  for (const [path, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (path.endsWith('.xml') || path.endsWith('.rels')) {
      out[path] = await file.async('string');
    }
  }
  return { zip, files: out };
}

function getStyleCounts(docXml) {
  const counts = {};
  const re = /<w:pStyle w:val="([^"]+)"\/>/g;
  let m;
  while ((m = re.exec(docXml)) !== null) {
    counts[m[1]] = (counts[m[1]] || 0) + 1;
  }
  return counts;
}

// -----------------------------------------------------------------------------
// Bloque A: classifier feature-based
// -----------------------------------------------------------------------------

describe('Fase 7 — Bloque A: classifier feature-based', () => {
  it('"1. Lavarse las manos" NO se clasifica como Título 1 (es un paso, no un título)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- PROCEDIMIENTO', kind: 'plain' },
        { text: '1. Lavarse las manos según protocolo, durante al menos 40 segundos con jabón.', kind: 'plain' },
        { text: '2. Aplicar la solución hidroalcohólica en las dos caras de las manos.', kind: 'plain' },
        { text: '3. Secarse con papel de un solo uso y desechar.', kind: 'plain' }
      ]
    });
    const { blob, stats } = await runEngine({ refFile, contentFile });
    const { files } = await unpack(blob);
    const styleCounts = getStyleCounts(files['word/document.xml']);
    // Sólo "1.- OBJETO", "2.- ALCANCE", "3.- PROCEDIMIENTO" como Título 1.
    expect(stats.title1).toBe(3);
    // Los pasos "1. Lavarse...", "2. Aplicar...", "3. Secarse..." se clasifican
    // como FHJPrrafo (no son ALL-CAPS y son largos).
    expect(styleCounts.FHJPrrafo).toBeGreaterThanOrEqual(3);
  });

  it('"3.1.- Subetapa" sigue siendo FHJTtuloprrafo (numbered-level-2plus, multinivel)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { text: '3.1.- Primera subetapa de prueba', kind: 'plain' }
      ]
    });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.titPar).toBeGreaterThanOrEqual(1);
  });

  it('respeta el pStyle FHJ pre-existente del contenido (trust path)', async () => {
    const refFile = await buildReferentDocx();
    // Contenido con un párrafo que ya viene marcado como FHJTtuloprrafo
    // por el redactor — el classifier no debe sobreescribirlo.
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJTtuloprrafo"/></w:pPr><w:r><w:t xml:space="preserve">Encabezado custom del autor</w:t></w:r></w:p>' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob, stats } = await runEngine({ refFile, contentFile });
    expect(stats.preservedStyles).toBeGreaterThanOrEqual(1);
    // Sigue habiendo al menos 1 párrafo FHJTtuloprrafo (el preservado).
    const { files } = await unpack(blob);
    const counts = getStyleCounts(files['word/document.xml']);
    expect(counts.FHJTtuloprrafo).toBeGreaterThanOrEqual(1);
  });

  it('párrafo con <w:numPr> bullet → FHJVietaNivel(N) según ilvl', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Texto.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Texto.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Bullet nivel 1</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Bullet nivel 2</w:t></w:r></w:p>' }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpack(blob);
    const counts = getStyleCounts(files['word/document.xml']);
    // Sin numbering.xml con definiciones, default es bullet → ambos van a FHJVieta*.
    expect((counts.FHJVietaNivel1 || 0) + (counts.FHJVietaNivel2 || 0)).toBeGreaterThanOrEqual(2);
  });

  it('párrafo con pStyle "Normal" → se reclasifica como FHJPrrafo', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="Normal"/></w:pPr><w:r><w:t xml:space="preserve">Texto en estilo Normal del usuario.</w:t></w:r></w:p>' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpack(blob);
    const counts = getStyleCounts(files['word/document.xml']);
    // El "Normal" no está en FHJ* → entra al classifier → fall-through FHJPrrafo.
    expect(counts.FHJPrrafo).toBeGreaterThanOrEqual(1);
  });

  it('párrafo con pStyle "ListParagraph" sin numPr → FHJVietaNivel1', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr><w:r><w:t xml:space="preserve">Item de lista huérfano</w:t></w:r></w:p>' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { blob } = await runEngine({ refFile, contentFile });
    const { files } = await unpack(blob);
    const counts = getStyleCounts(files['word/document.xml']);
    expect(counts.FHJVietaNivel1).toBeGreaterThanOrEqual(1);
  });
});

// -----------------------------------------------------------------------------
// Bloque C: transferencia robusta de headers/footers desde el ref
// -----------------------------------------------------------------------------

describe('Fase 7 — Bloque C: transferencia de headers/footers', () => {
  it('content sin cabecera FHJ → output trae los 3 headers + 3 footers del ref', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });
    const { blob, stats } = await runEngine({ refFile, contentFile });
    const { zip } = await unpack(blob);
    expect(stats.preservedHeaders).toBe(false);
    for (let i = 1; i <= 3; i++) {
      expect(zip.file('word/header' + i + '.xml')).not.toBeNull();
      expect(zip.file('word/footer' + i + '.xml')).not.toBeNull();
    }
  });

  it('media del ref se importa con prefijo fhj_ (no colisiona con media del content)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: false });
    const { blob } = await runEngine({ refFile, contentFile });
    const { zip, files } = await unpack(blob);
    // El ref aporta media/image1.emf — debe llegar como fhj_image1.emf.
    expect(zip.file('word/media/fhj_image1.emf')).not.toBeNull();
    // header2.xml.rels apunta al renombrado.
    expect(files['word/_rels/header2.xml.rels']).toContain('fhj_image1.emf');
  });

  it('content con cabecera con drawing pero SIN marcador FHJ → ref gana', async () => {
    // Construimos un content con header propio que contiene drawing + image rel
    // pero cuyo texto NO menciona FHJ ni código PNT — el detector estricto
    // de Fase 7 NO lo cuenta como cabecera FHJ → ref se transfiere.
    const refFile = await buildReferentDocx();
    const baseContent = await buildContentDocx({ withFhjHeader: false });
    const zip = await JSZip.loadAsync(baseContent);
    zip.file('word/header1.xml',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/></wp:inline></w:drawing></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">Logo de otro centro sanitario — sin marca corporativa</w:t></w:r></w:p>' +
      '</w:hdr>'
    );
    zip.file('word/_rels/header1.xml.rels',
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>' +
      '</Relationships>'
    );
    zip.file('word/media/image1.emf', new Uint8Array([0xAB, 0xCD]));
    const sneakyContent = await zip.generateAsync({ type: 'nodebuffer' });

    const { blob, stats } = await runEngine({ refFile, contentFile: sneakyContent });
    expect(stats.preservedHeaders).toBe(false);
    const { zip: outZip } = await unpack(blob);
    expect(outZip.file('word/media/fhj_image1.emf')).not.toBeNull();
  });

  it('content con cabecera FHJ propia (texto que matchea P.dd.dd.ddd) → se preserva', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    const { stats } = await runEngine({ refFile, contentFile });
    expect(stats.preservedHeaders).toBe(true);
  });
});

// -----------------------------------------------------------------------------
// Bloque D: validador normativo
// -----------------------------------------------------------------------------

describe('Fase 7 — Bloque D: validador normativo', () => {
  it('párrafo de cuerpo en ALL-CAPS largo (FHJPrrafo) → emite NORMATIVA_ALL_CAPS_BODY', async () => {
    const refFile = await buildReferentDocx();
    // Forzamos pStyle FHJPrrafo para evitar que el classifier lo detecte como
    // título por la regla all-caps-short. El validador debe quejarse igual.
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:pStyle w:val="FHJPrrafo"/></w:pPr><w:r><w:t xml:space="preserve">ESTE TEXTO ESTÁ ESCRITO ENTERO EN MAYÚSCULAS Y NO DEBERÍA ESTARLO PORQUE LA NORMATIVA DEL HOSPITAL DICE QUE NO SE USEN MAYÚSCULAS EN EL CUERPO.</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).toContain('NORMATIVA_ALL_CAPS_BODY');
  });

  it('runs con <w:u/> en el body → emite NORMATIVA_UNDERLINE', async () => {
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
    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).toContain('NORMATIVA_UNDERLINE');
  });

  it('runs con fuente no-Arial → emite NORMATIVA_FONT_NON_ARIAL', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/></w:rPr><w:t xml:space="preserve">En Times New Roman.</w:t></w:r></w:p>' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    // Fase 12 B2 enforceTypography ya reescribe la fuente automáticamente; aquí
    // queremos validar que el *validador* detecta la desviación, así que lo
    // desactivamos para forzar que la fuente Times sobreviva hasta validate().
    const { warnings } = await runEngine({ refFile, contentFile, enforceTypography: false });
    const codes = warnings.map(w => w.code);
    expect(codes).toContain('NORMATIVA_FONT_NON_ARIAL');
  });

  it('engine auto-inyecta Datos Generales → no emite NORMATIVA_MISSING_DATOS_GENERALES', async () => {
    // El engine inyecta la tabla Datos Generales si no la encuentra. Por tanto,
    // el output siempre debe pasar este check sin warning.
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Obj.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).not.toContain('NORMATIVA_MISSING_DATOS_GENERALES');
  });

  it('documento limpio (sin desviaciones) NO emite warnings normativos', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();   // body por defecto
    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).not.toContain('NORMATIVA_ALL_CAPS_BODY');
    expect(codes).not.toContain('NORMATIVA_UNDERLINE');
    expect(codes).not.toContain('NORMATIVA_FONT_NON_ARIAL');
    expect(codes).not.toContain('NORMATIVA_EMPTY_LIST');
  });
});
