/**
 * Fase 6 — detección robusta de "Datos generales" en el contenido.
 *
 * El engine inyecta la tabla FHJ "Datos generales" al principio del body
 * si no la detecta. La heurística antigua se fijaba sólo en el texto
 * "Datos generales". La nueva, además, reconoce tablas cuyo primera
 * columna contiene las etiquetas canónicas (Código, Versión, Elaborado
 * por, …) aunque el título literal haya sido eliminado.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

import {
  buildReferentDocx,
  buildContentDocx,
  replaceInDocx
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function runEngine(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', ...opts });
}

async function readDocumentXml(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  return await zip.file('word/document.xml').async('string');
}

function buildDocumentWithFhjTable() {
  // Una tabla 2 columnas con las etiquetas canónicas — SIN título "Datos generales".
  const rows = [
    ['Código:', 'P.01.00.001'],
    ['Versión:', 'V.1.0'],
    ['Elaborado por:', 'Lab'],
    ['Aprobado por:', 'Dirección']
  ];
  const tableRows = rows.map(([l, v]) =>
    '<w:tr>' +
      '<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>' +
        '<w:p><w:r><w:t xml:space="preserve">' + l + '</w:t></w:r></w:p>' +
      '</w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="6500" w:type="dxa"/></w:tcPr>' +
        '<w:p><w:r><w:t xml:space="preserve">' + v + '</w:t></w:r></w:p>' +
      '</w:tc>' +
    '</w:tr>'
  ).join('');

  const tableXml =
    '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="6500"/></w:tblGrid>' +
    tableRows +
    '</w:tbl>';

  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<w:body>' +
    tableXml +
    '<w:p><w:r><w:t xml:space="preserve">1.- OBJETO</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">cuerpo</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">2.- DESARROLLO</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">bla</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
    '</w:body></w:document>';
}

describe('Fase 6 — detección de Datos generales por contenido de celdas', () => {
  it('NO inyecta la tabla si el body ya tiene una tabla con las etiquetas canónicas (sin título)', async () => {
    const refFile = await buildReferentDocx();
    const baseContent = await buildContentDocx();
    const contentFile = await replaceInDocx(baseContent, 'word/document.xml', buildDocumentWithFhjTable());

    const { blob } = await runEngine({ refFile, contentFile });
    const docXml = await readDocumentXml(blob);

    // El título literal "Datos generales" NO debería existir porque no lo
    // inyectamos y el contenido no lo trae.
    expect(docXml).not.toContain('Datos generales');
    // Y debería haber exactamente UNA tabla (la del contenido), no dos.
    const tbls = docXml.match(/<w:tbl[\s>]/g) || [];
    expect(tbls.length).toBe(1);
  });

  it('SÍ inyecta la tabla cuando el body no trae ni título ni tabla canónica', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo sin tabla previa.', kind: 'plain' },
        { text: '2.- DESARROLLO', kind: 'plain' },
        { text: 'Bla.', kind: 'plain' }
      ]
    });

    const { blob } = await runEngine({ refFile, contentFile });
    const docXml = await readDocumentXml(blob);

    // Se ha inyectado la tabla con el título canónico.
    expect(docXml).toContain('Datos generales');
    expect(docXml).toContain('FHJTitulodatosgenerales');
  });

  it('con título literal "Datos generales" en el body: tampoco duplica (camino antiguo)', async () => {
    const refFile = await buildReferentDocx();
    // Por defecto el content no lleva "Datos generales" literal → lo metemos manualmente.
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Datos generales', kind: 'plain' },
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo.', kind: 'plain' },
        { text: '2.- DESARROLLO', kind: 'plain' },
        { text: 'Bla.', kind: 'plain' }
      ]
    });

    const { blob } = await runEngine({ refFile, contentFile });
    const docXml = await readDocumentXml(blob);

    // "Datos generales" debería aparecer exactamente UNA vez — la del content,
    // no una segunda inyectada por nosotros.
    const occurrences = (docXml.match(/Datos generales/g) || []).length;
    expect(occurrences).toBe(1);
  });

  it('tabla con sólo 2 etiquetas canónicas (por debajo del umbral de 3): SÍ inyecta', async () => {
    const refFile = await buildReferentDocx();
    const baseContent = await buildContentDocx();

    // Tabla con 2 filas → sólo 2 labels → no se reconoce como FHJ datos generales.
    const smallTableXml =
      '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="6500"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">Código:</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p><w:r><w:t xml:space="preserve">X</w:t></w:r></w:p></w:tc></w:tr>' +
      '<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">Versión:</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p><w:r><w:t xml:space="preserve">Y</w:t></w:r></w:p></w:tc></w:tr>' +
      '</w:tbl>';

    const docXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<w:body>' +
      smallTableXml +
      '<w:p><w:r><w:t xml:space="preserve">1.- OBJETO</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">cuerpo</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">2.- DESARROLLO</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">bla</w:t></w:r></w:p>' +
      '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
      '</w:body></w:document>';
    const contentFile = await replaceInDocx(baseContent, 'word/document.xml', docXml);

    const { blob } = await runEngine({ refFile, contentFile });
    const out = await readDocumentXml(blob);

    // Se inyectó la tabla FHJ porque la del contenido no supera el umbral.
    expect(out).toContain('Datos generales');
  });
});
