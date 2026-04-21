/**
 * Factorías de docx sintéticos mínimos para tests golden.
 * Producen zips válidos que ejercen el engine sin necesitar Word real.
 *
 * buildReferentDocx(opts) → Buffer
 * buildContentDocx(opts)  → Buffer  (caso A o caso B según withFhjHeader)
 */

import JSZip from 'jszip';

// ----- bloques XML reusables -----

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const W_NS_ATTR = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const R_NS_ATTR = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const ROOT_RELS =
  XML_HEAD +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  '</Relationships>';

const STYLES_XML =
  XML_HEAD +
  '<w:styles ' + W_NS_ATTR + '>' +
  fhjStyle('FHJTtulo1', 'FHJ Titulo 1') +
  fhjStyle('FHJTtuloprrafo', 'FHJ Titulo Parrafo') +
  fhjStyle('FHJPrrafo', 'FHJ Parrafo') +
  fhjStyle('FHJVietaNivel1', 'FHJ Vineta Nivel 1') +
  fhjStyle('FHJTitulodatosgenerales', 'FHJ Datos Titulo') +
  fhjStyle('FHJContenidodatosgenerales', 'FHJ Datos Contenido') +
  '</w:styles>';

const THEME_XML = XML_HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements/></a:theme>';
const SETTINGS_XML = XML_HEAD + '<w:settings ' + W_NS_ATTR + '/>';
const NUMBERING_XML = XML_HEAD + '<w:numbering ' + W_NS_ATTR + '/>';

const FOOTER_XML =
  XML_HEAD +
  '<w:ftr ' + W_NS_ATTR + ' ' + R_NS_ATTR + '>' +
  '<w:p><w:r><w:t xml:space="preserve">Footer FHJ</w:t></w:r></w:p>' +
  '</w:ftr>';

function fhjStyle(id, name) {
  return '<w:style w:type="paragraph" w:styleId="' + id + '">' +
    '<w:name w:val="' + name + '"/>' +
    '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr>' +
    '</w:style>';
}

// Un header tipo "referente" con dos párrafos (código/versión + título) — lo que
// customizeHeader espera encontrar para personalizar.
function referentHeader2() {
  return XML_HEAD +
    '<w:hdr ' + W_NS_ATTR + ' ' + R_NS_ATTR + '>' +
    '<w:p><w:r><w:t xml:space="preserve">[CODIGO / VERSION]</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">[TITULO DEL PROCEDIMIENTO]</w:t></w:r></w:p>' +
    '</w:hdr>';
}

// Header con drawing + rel a imagen (lo que el engine detecta como "FHJ header propio")
function contentHeaderWithDrawing() {
  return XML_HEAD +
    '<w:hdr ' + W_NS_ATTR + ' ' + R_NS_ATTR + '>' +
    '<w:p><w:r>' +
    '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/></wp:inline></w:drawing>' +
    '</w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">P.01.00.001 / V.0.1</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t xml:space="preserve">Procedimiento de prueba (cabecera propia)</w:t></w:r></w:p>' +
    '</w:hdr>';
}

function headerImageRels() {
  return XML_HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>' +
    '</Relationships>';
}

function referentHeader2Rels() {
  // Apunta a media/image1.emf — transferHeadersFootersAndLogo renombrará a image1_fhj_logo.emf
  return headerImageRels();
}

function simpleHeader(label) {
  return XML_HEAD +
    '<w:hdr ' + W_NS_ATTR + '>' +
    '<w:p><w:r><w:t xml:space="preserve">' + escapeXml(label) + '</w:t></w:r></w:p>' +
    '</w:hdr>';
}

// Default body: mezcla de títulos, subtítulos, párrafos, viñetas, una tabla y un drawing.
function defaultBodyParagraphs() {
  return [
    { text: '1.- OBJETO', kind: 'plain' },
    { text: 'Este procedimiento establece la sistemática de prueba.', kind: 'plain' },
    { text: '2.- ALCANCE', kind: 'plain' },
    { text: 'Aplica al personal del laboratorio & servicios anexos.', kind: 'plain' },
    { text: '3.- DESARROLLO', kind: 'plain' },
    { text: '3.1.- Primera subetapa', kind: 'plain' },
    { text: 'Descripción detallada de la primera subetapa.', kind: 'plain' },
    { text: '3.1.1.- Sub-subetapa', kind: 'plain' },
    { text: '- Primer elemento de lista', kind: 'plain' },
    { text: '- Segundo elemento de lista', kind: 'plain' },
    { text: 'ANEXO I', kind: 'plain' },
    { text: '', kind: 'table' },   // una tabla simple
    { text: '', kind: 'drawing' }  // un drawing suelto
  ];
}

function buildBodyXml(paragraphs) {
  return paragraphs.map(p => {
    if (p.kind === 'raw') {
      // Escapatoria para tests de edge cases: XML literal que se mete tal cual.
      return p.xml || '';
    }
    if (p.kind === 'table') {
      // Permite simular tablas reales generadas por Word con metadata (w:rsidR, etc.)
      // pasando { kind: 'table', attrs: 'w:rsidR="00112233"' }.
      const attrs = p.attrs ? ' ' + p.attrs : '';
      return '<w:tbl' + attrs + '><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>' +
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">A</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">B</w:t></w:r></w:p></w:tc></w:tr>' +
        '</w:tbl>';
    }
    if (p.kind === 'drawing') {
      return '<w:p><w:r>' +
        '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="1" cy="1"/></wp:inline></w:drawing>' +
        '</w:r></w:p>';
    }
    return '<w:p><w:r><w:t xml:space="preserve">' + escapeXml(p.text) + '</w:t></w:r></w:p>';
  }).join('');
}

function buildDocumentXml(paragraphs) {
  const body = buildBodyXml(paragraphs);
  return XML_HEAD +
    '<w:document ' + W_NS_ATTR + ' ' + R_NS_ATTR + '>' +
    '<w:body>' + body +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>' +
    '</w:body></w:document>';
}

function buildDocumentRels(extraRels = []) {
  const core = [
    { id: 'rId1', type: 'styles', target: 'styles.xml' },
    { id: 'rId2', type: 'settings', target: 'settings.xml' },
    { id: 'rId3', type: 'numbering', target: 'numbering.xml' },
    { id: 'rId4', type: 'theme', target: 'theme/theme1.xml' }
  ];
  const all = core.concat(extraRels);
  return XML_HEAD +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    all.map(r => {
      return '<Relationship Id="' + r.id +
        '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + r.type +
        '" Target="' + r.target + '"/>';
    }).join('') +
    '</Relationships>';
}

function buildContentTypes({ withHeadersFooters }) {
  let xml = XML_HEAD +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>';
  if (withHeadersFooters) {
    for (let i = 1; i <= 3; i++) {
      xml += '<Override PartName="/word/header' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>';
      xml += '<Override PartName="/word/footer' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>';
    }
  }
  xml += '</Types>';
  return xml;
}

function escapeXml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ----- builders públicos -----

export async function buildReferentDocx({ withLogo = true, withImage = true } = {}) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', buildContentTypes({ withHeadersFooters: true }));
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/document.xml', buildDocumentXml([{ text: 'REFERENTE', kind: 'plain' }]));
  zip.file('word/_rels/document.xml.rels', buildDocumentRels());
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/theme/theme1.xml', THEME_XML);
  zip.file('word/settings.xml', SETTINGS_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/header1.xml', simpleHeader('Header 1 referente'));
  zip.file('word/header2.xml', referentHeader2());
  zip.file('word/header3.xml', simpleHeader('Header 3 referente'));
  zip.file('word/footer1.xml', FOOTER_XML);
  zip.file('word/footer2.xml', FOOTER_XML);
  zip.file('word/footer3.xml', FOOTER_XML);
  if (withLogo) {
    zip.file('word/_rels/header2.xml.rels', referentHeader2Rels());
  }
  if (withImage) {
    // bytes bogus pero válidos como archivo; el engine no los abre
    zip.file('word/media/image1.emf', new Uint8Array([0x01, 0x02, 0x03, 0x04]));
  }
  return await zip.generateAsync({ type: 'nodebuffer' });
}

export async function buildContentDocx({
  withFhjHeader = false,
  paragraphs = defaultBodyParagraphs()
} = {}) {
  const zip = new JSZip();
  zip.file('[Content_Types].xml', buildContentTypes({ withHeadersFooters: withFhjHeader }));
  zip.file('_rels/.rels', ROOT_RELS);
  zip.file('word/document.xml', buildDocumentXml(paragraphs));
  zip.file('word/_rels/document.xml.rels', buildDocumentRels(
    withFhjHeader
      ? [{ id: 'rId10', type: 'header', target: 'header1.xml' }]
      : []
  ));
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/theme/theme1.xml', THEME_XML);
  zip.file('word/settings.xml', SETTINGS_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  if (withFhjHeader) {
    // Caso A: el doc ya trae su propio header FHJ con drawing + image rel
    zip.file('word/header1.xml', contentHeaderWithDrawing());
    zip.file('word/_rels/header1.xml.rels', headerImageRels());
    zip.file('word/media/image1.emf', new Uint8Array([0x99, 0x98, 0x97]));
  }
  return await zip.generateAsync({ type: 'nodebuffer' });
}

// ----- helpers para tests de error: mutar docx existentes -----

/**
 * Devuelve un nuevo docx (Buffer) sin las entradas indicadas.
 * Uso: const broken = await dropFromDocx(buf, ['word/styles.xml']).
 */
export async function dropFromDocx(buffer, paths) {
  const zip = await JSZip.loadAsync(buffer);
  for (const p of paths) zip.remove(p);
  return await zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Devuelve un nuevo docx (Buffer) reemplazando el contenido del archivo dado.
 * Uso: replaceInDocx(buf, 'word/document.xml', '<?xml ...?><w:document/>')
 */
export async function replaceInDocx(buffer, path, content) {
  const zip = await JSZip.loadAsync(buffer);
  zip.file(path, content);
  return await zip.generateAsync({ type: 'nodebuffer' });
}

/**
 * Documento cuyo word/document.xml es válido como XML pero NO contiene <w:body>.
 * Útil para forzar el path CONTENT_MISSING_BODY.
 */
export function buildDocumentXmlWithoutBody() {
  return XML_HEAD +
    '<w:document ' + W_NS_ATTR + ' ' + R_NS_ATTR + '>' +
    '<w:metadata/>' +
    '</w:document>';
}

export { defaultBodyParagraphs };

// ----- builders de edge cases (Fase 4) -----

/** Párrafo vacío — self-closing, sin runs ni pPr. */
export function rawEmptyParagraph() {
  return '<w:p/>';
}

/** Párrafo cuyo único run es un w:t con solo espacios. */
export function rawWhitespaceParagraph(spaces = '     ') {
  return '<w:p><w:r><w:t xml:space="preserve">' + spaces + '</w:t></w:r></w:p>';
}

/** Tabla con tblPr/tblGrid pero cero filas. */
export function rawEmptyTable() {
  return '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    '</w:tbl>';
}

/** Tabla con una celda que a su vez contiene otra tabla. */
export function rawNestedTable() {
  const inner = '<w:tbl><w:tblPr><w:tblW w:w="2500" w:type="pct"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:t xml:space="preserve">INNER</w:t></w:r></w:p></w:tc></w:tr>' +
    '</w:tbl>';
  return '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>' +
    '<w:p><w:r><w:t xml:space="preserve">OUTER</w:t></w:r></w:p>' +
    inner +
    '</w:tc></w:tr>' +
    '</w:tbl>';
}

/** Drawing sin wp:docPr ni wp:inline children relevantes — la figura debe seguir contando. */
export function rawDrawingBareParagraph() {
  return '<w:p><w:r><w:drawing><wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"/></w:drawing></w:r></w:p>';
}

/**
 * Párrafo con un sectPr anidado (ruptura de sección intra-documento) —
 * Word permite varios sectPr. El engine solo usa el primero.
 */
export function rawParagraphWithInlineSectPr() {
  return '<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906"/></w:sectPr></w:pPr>' +
    '<w:r><w:t xml:space="preserve">salto de seccion</w:t></w:r></w:p>';
}

/**
 * Genera `count` parejas bookmarkStart/bookmarkEnd con el mismo id — caso extremo
 * de bookmarks duplicados que el engine debe deduplicar sin fallar.
 */
export function rawDuplicateBookmarks(id = '0', count = 6) {
  let xml = '';
  for (let i = 0; i < count; i++) {
    xml += '<w:p>' +
      '<w:bookmarkStart w:id="' + id + '" w:name="dup' + i + '"/>' +
      '<w:r><w:t xml:space="preserve">marca ' + i + '</w:t></w:r>' +
      '<w:bookmarkEnd w:id="' + id + '"/>' +
      '</w:p>';
  }
  return xml;
}

/** Genera `n` párrafos normales con números crecientes (stress/perf smoke test). */
export function manyPlainParagraphs(n) {
  const out = [];
  for (let i = 1; i <= n; i++) {
    out.push({ text: 'Linea ' + i + ' del documento sintético de stress.', kind: 'plain' });
  }
  return out;
}
