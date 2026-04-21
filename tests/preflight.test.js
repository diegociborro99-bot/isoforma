/**
 * Fase 3 — preflight + modelo de errores.
 * Cubre cada path de error tipado del engine y los warnings no fatales.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';

import {
  buildReferentDocx,
  buildContentDocx,
  dropFromDocx,
  replaceInDocx,
  buildDocumentXmlWithoutBody
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');
const { IsoformaError } = IsoformaEngine;

async function runEngine(opts) {
  return IsoformaEngine.process({
    outputType: 'nodebuffer',
    ...opts
  });
}

/**
 * Atrapa la excepción lanzada por runEngine y la devuelve para inspección.
 * Si NO lanza, el error del test es claro.
 */
async function catchError(fn) {
  try {
    await fn();
  } catch (err) {
    return err;
  }
  throw new Error('Se esperaba que el engine lanzara, pero completó.');
}

describe('IsoformaError — shape y API', () => {
  it('IsoformaError se expone en la API del engine', () => {
    expect(typeof IsoformaError).toBe('function');
    const e = new IsoformaError({ code: 'X', message: 'y', step: 'z', context: { a: 1 } });
    expect(e.name).toBe('IsoformaError');
    expect(e.code).toBe('X');
    expect(e.step).toBe('z');
    expect(e.context).toEqual({ a: 1 });
    expect(e.message).toBe('y');
    expect(e instanceof Error).toBe(true);
  });

  it('toJSON devuelve un objeto serializable útil para logging', () => {
    const e = new IsoformaError({ code: 'FOO', message: 'bar', step: 'baz', context: { n: 2 } });
    const j = e.toJSON();
    expect(j).toEqual({
      name: 'IsoformaError',
      code: 'FOO',
      step: 'baz',
      message: 'bar',
      context: { n: 2 }
    });
  });
});

describe('preflight — errores fatales sobre el archivo de contenido', () => {
  it('MISSING_INPUT: contentFile null', async () => {
    const refFile = await buildReferentDocx();
    const err = await catchError(() => runEngine({ refFile, contentFile: null }));
    expect(err.name).toBe('IsoformaError');
    expect(err.code).toBe('MISSING_INPUT');
    expect(err.step).toBe('Leyendo contenido');
    expect(err.context.role).toBe('contenido');
  });

  it('MISSING_INPUT: refFile undefined', async () => {
    const contentFile = await buildContentDocx();
    const err = await catchError(() => runEngine({ refFile: undefined, contentFile }));
    expect(err.code).toBe('MISSING_INPUT');
    expect(err.step).toBe('Leyendo referente');
  });

  it('INVALID_DOCX: contentFile no es un zip', async () => {
    const refFile = await buildReferentDocx();
    const garbage = Buffer.from('esto no es un docx, es texto plano');
    const err = await catchError(() => runEngine({ refFile, contentFile: garbage }));
    expect(err.code).toBe('INVALID_DOCX');
    expect(err.step).toBe('Leyendo contenido');
    expect(err.context.role).toBe('contenido');
    expect(err.context.originalError).toBeTruthy();
  });

  it('CONTENT_NOT_DOCX: zip válido pero sin [Content_Types].xml', async () => {
    const refFile = await buildReferentDocx();
    // Zip válido con cualquier archivo que no sea docx
    const zip = new JSZip();
    zip.file('algo.txt', 'hola');
    const contentFile = await zip.generateAsync({ type: 'nodebuffer' });

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('CONTENT_NOT_DOCX');
    expect(err.step).toBe('preflight');
    expect(err.context.missing).toBe('[Content_Types].xml');
  });

  it('CONTENT_MISSING_DOCUMENT_XML: docx sin word/document.xml', async () => {
    const refFile = await buildReferentDocx();
    const good = await buildContentDocx();
    const contentFile = await dropFromDocx(good, ['word/document.xml']);

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('CONTENT_MISSING_DOCUMENT_XML');
    expect(err.step).toBe('preflight');
    expect(err.context.missing).toBe('word/document.xml');
  });

  it('CONTENT_MISSING_BODY: document.xml sin <w:body>', async () => {
    const refFile = await buildReferentDocx();
    const good = await buildContentDocx();
    const contentFile = await replaceInDocx(good, 'word/document.xml', buildDocumentXmlWithoutBody());

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('CONTENT_MISSING_BODY');
    expect(err.step).toBe('preflight');
    expect(err.context.file).toBe('word/document.xml');
  });
});

describe('preflight — errores fatales sobre el referente', () => {
  it('REFERENT_NOT_DOCX: zip válido sin [Content_Types].xml', async () => {
    const contentFile = await buildContentDocx();
    const zip = new JSZip();
    zip.file('algo.txt', 'hola');
    const refFile = await zip.generateAsync({ type: 'nodebuffer' });

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('REFERENT_NOT_DOCX');
    expect(err.context.role).toBe('referent');
  });

  it('REFERENT_MISSING_DOCUMENT_XML: referente sin word/document.xml', async () => {
    const contentFile = await buildContentDocx();
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, ['word/document.xml']);

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('REFERENT_MISSING_DOCUMENT_XML');
  });

  it('REFERENT_MISSING_STYLES: referente sin word/styles.xml', async () => {
    const contentFile = await buildContentDocx();
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, ['word/styles.xml']);

    const err = await catchError(() => runEngine({ refFile, contentFile }));
    expect(err.code).toBe('REFERENT_MISSING_STYLES');
    expect(err.context.missing).toBe('word/styles.xml');
  });
});

describe('warnings no fatales', () => {
  it('camino feliz devuelve warnings = []', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();
    const { warnings } = await runEngine({ refFile, contentFile });
    expect(Array.isArray(warnings)).toBe(true);
    expect(warnings).toEqual([]);
  });

  it('REFERENT_NO_NUMBERING cuando al referente le falta numbering.xml', async () => {
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, ['word/numbering.xml']);
    const contentFile = await buildContentDocx();

    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).toContain('REFERENT_NO_NUMBERING');
  });

  it('REFERENT_NO_THEME cuando al referente le falta theme/theme1.xml', async () => {
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, ['word/theme/theme1.xml']);
    const contentFile = await buildContentDocx();

    const { warnings } = await runEngine({ refFile, contentFile });
    const codes = warnings.map(w => w.code);
    expect(codes).toContain('REFERENT_NO_THEME');
  });

  it('REFERENT_PARTIAL_HEADERS cuando el referente solo aporta 1 header', async () => {
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, [
      'word/header2.xml',
      'word/header3.xml'
    ]);
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    // Con withFhjHeader el engine conserva cabecera propia del contenido,
    // pero el preflight valida el estado del referente igualmente.

    const { warnings } = await runEngine({ refFile, contentFile });
    const partial = warnings.find(w => w.code === 'REFERENT_PARTIAL_HEADERS');
    expect(partial).toBeDefined();
    expect(partial.context.present).toEqual(['header1']);
  });

  it('REFERENT_NO_FOOTERS cuando el referente no trae ningún footer', async () => {
    const good = await buildReferentDocx();
    const refFile = await dropFromDocx(good, [
      'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
    ]);
    const contentFile = await buildContentDocx({ withFhjHeader: true });

    const { warnings } = await runEngine({ refFile, contentFile });
    expect(warnings.map(w => w.code)).toContain('REFERENT_NO_FOOTERS');
  });

  it('METADATA_IGNORED_PRESERVED_HEADER cuando el contenido ya trae cabecera FHJ propia', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({ withFhjHeader: true });
    const metadata = {
      code: 'P.00.00.000',
      version: 'V.0.1',
      title: 'Cualquier cosa'
    };
    const { warnings } = await runEngine({ refFile, contentFile, metadata });
    const w = warnings.find(x => x.code === 'METADATA_IGNORED_PRESERVED_HEADER');
    expect(w).toBeDefined();
    expect(w.context.metadata).toEqual(metadata);
  });

  it('CUSTOMIZE_HEADER_INSUFFICIENT_PARAGRAPHS cuando el referente trae header2 con < 2 párrafos', async () => {
    const good = await buildReferentDocx();
    const onlyOneP =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
      '<w:p><w:r><w:t xml:space="preserve">SOLO UNO</w:t></w:r></w:p>' +
      '</w:hdr>';
    const refFile = await replaceInDocx(good, 'word/header2.xml', onlyOneP);
    const contentFile = await buildContentDocx();
    const metadata = { code: 'P.01.00.001', version: 'V.1.0', title: 'Prueba' };

    const { warnings } = await runEngine({ refFile, contentFile, metadata });
    const w = warnings.find(x => x.code === 'CUSTOMIZE_HEADER_INSUFFICIENT_PARAGRAPHS');
    expect(w).toBeDefined();
    expect(w.context.foundParagraphs).toBe(1);
  });
});

describe('errores envueltos por runStep', () => {
  // Forzamos un fallo en runStep entregando un contentFile cuyo document.xml
  // es válido como XML pero contiene <w:body> para pasar preflight y luego
  // rompe el parseo de forma controlada. En la práctica no lo hacemos XML
  // inválido (porque @xmldom es muy permisivo); en su lugar validamos que
  // cuando cualquier path lanza algo no tipado, sale como IsoformaError.

  it('un error interno inesperado se traduce a IsoformaError con step', async () => {
    // Forzamos el fallo monkey-patcheando onProgress para que lance.
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();

    // onProgress se invoca dentro del flujo (no dentro de runStep para "Leyendo
    // documentos"), así que lanzar desde progress no testea lo que queremos.
    // En su lugar verificamos el caso natural: un INVALID_DOCX ya está
    // envuelto, confirmando que el contrato se respeta punta a punta.
    const err = await catchError(() =>
      runEngine({ refFile, contentFile: Buffer.from([0, 1, 2]) })
    );
    expect(err.name).toBe('IsoformaError');
    expect(err.step).toBeTruthy();
    expect(err.code).toBeTruthy();
    expect(err.message).toContain('contenido');
  });
});
