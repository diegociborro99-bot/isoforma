/**
 * Utilidades de inspección de docx para los tests golden.
 * Abre el zip resultado del engine y permite hacer asserts dirigidos.
 */

import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

/**
 * Recibe la salida del engine (Buffer en Node) y devuelve un objeto
 * { zip, files } donde files es un mapa nombre -> string con el contenido XML.
 * Solo se leen archivos de texto plano (XML/RELS); los binarios se omiten.
 */
export async function unpackDocx(bufferOrBlob) {
  let input = bufferOrBlob;
  if (typeof Blob !== 'undefined' && input instanceof Blob) {
    input = Buffer.from(await input.arrayBuffer());
  }
  const zip = await JSZip.loadAsync(input);
  const files = {};
  const names = Object.keys(zip.files);
  for (const name of names) {
    const entry = zip.files[name];
    if (entry.dir) continue;
    if (/\.(xml|rels)$/i.test(name)) {
      files[name] = await entry.async('string');
    }
  }
  return { zip, files };
}

/**
 * Cuenta cuántas veces aparece <w:pStyle w:val="NOMBRE"/> en un XML.
 */
export function countPStyle(xml, styleName) {
  const re = new RegExp('<w:pStyle\\s+w:val="' + escapeRegex(styleName) + '"\\s*/>', 'g');
  return (xml.match(re) || []).length;
}

/**
 * Cuenta nodos <w:p> cuya pStyle coincide (DOM, más estricto).
 */
export function countParagraphsWithStyle(xml, styleName) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const paragraphs = doc.getElementsByTagName('w:p');
  let n = 0;
  for (let i = 0; i < paragraphs.length; i++) {
    const pStyle = paragraphs[i].getElementsByTagName('w:pStyle')[0];
    if (pStyle && pStyle.getAttribute('w:val') === styleName) n++;
  }
  return n;
}

/**
 * Valida que un string sea XML bien formado.
 */
export function isValidXml(xmlString) {
  const errors = [];
  const parser = new DOMParser({
    errorHandler: {
      warning: () => {},
      error: (e) => errors.push(String(e)),
      fatalError: (e) => errors.push(String(e))
    }
  });
  try {
    parser.parseFromString(xmlString, 'application/xml');
  } catch (e) {
    errors.push(String(e));
  }
  return { valid: errors.length === 0, errors };
}

/**
 * Devuelve todo el texto visible (<w:t>) concatenado de un XML.
 */
export function extractText(xml) {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const ts = doc.getElementsByTagName('w:t');
  let out = '';
  for (let i = 0; i < ts.length; i++) {
    out += ts[i].textContent;
    out += ' ';
  }
  return out.replace(/\s+/g, ' ').trim();
}

/**
 * Verifica integridad estructural básica del docx output.
 */
export async function assertStructuralIntegrity(files) {
  const errors = [];
  if (!files['[Content_Types].xml']) errors.push('Falta [Content_Types].xml');
  else {
    const ct = isValidXml(files['[Content_Types].xml']);
    if (!ct.valid) errors.push('[Content_Types].xml inválido: ' + ct.errors.join('; '));
  }
  if (!files['word/document.xml']) errors.push('Falta word/document.xml');
  else {
    const dx = isValidXml(files['word/document.xml']);
    if (!dx.valid) errors.push('word/document.xml inválido: ' + dx.errors.join('; '));
    const bodies = (files['word/document.xml'].match(/<w:body>/g) || []).length;
    if (bodies !== 1) errors.push('document.xml debe tener exactamente un <w:body>, encontrados: ' + bodies);
  }
  return { ok: errors.length === 0, errors };
}

function escapeRegex(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
