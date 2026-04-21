/**
 * Fase 12 — tests de cumplimiento normativo §4.
 *
 * Cubre los 5 bloques del motor:
 *   B1 · enforceFHJSpacing       — espaciado e interlineado §4.3
 *   B2 · enforceArialTypography  — Arial 10/9 §4.2.1–§4.2.3.2
 *   B3 · unwrapNarrativeParagraphs — reconstrucción de prosa fragmentada
 *   B4 · enforceListIndent + normalizeListSymbols — listas §4.2.4
 *   B6 · enforceSemanticTypography — latinismos/extranjerismos/alertas §4.2.2
 *
 * Los flags enforceSpacing / enforceTypography / unwrapNarrative /
 * normalizeLists / semanticTypography están a `true` por defecto — los
 * tests de backward-compat los desactivan explícitamente para verificar
 * que se respetan los opt-outs.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';

import {
  buildReferentDocx,
  buildContentDocx,
  replaceInDocx
} from './helpers/synthetic.js';

/**
 * numbering.xml sintético con 3 niveles bullet (ilvl 0/1/2) usando símbolos
 * NO canónicos — fuerza que normalizeListSymbols reescriba lvlText a ●/–/▪.
 */
const RICH_NUMBERING_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:abstractNum w:abstractNumId="0">' +
      '<w:lvl w:ilvl="0">' +
        '<w:numFmt w:val="bullet"/>' +
        '<w:lvlText w:val="&#183;"/>' + // · — NO canónico
        '<w:pPr><w:ind w:left="99" w:hanging="99"/></w:pPr>' +
      '</w:lvl>' +
      '<w:lvl w:ilvl="1">' +
        '<w:numFmt w:val="bullet"/>' +
        '<w:lvlText w:val="o"/>' + // o — NO canónico
        '<w:pPr><w:ind w:left="99" w:hanging="99"/></w:pPr>' +
      '</w:lvl>' +
      '<w:lvl w:ilvl="2">' +
        '<w:numFmt w:val="bullet"/>' +
        '<w:lvlText w:val="&#167;"/>' + // § — NO canónico
        '<w:pPr><w:ind w:left="99" w:hanging="99"/></w:pPr>' +
      '</w:lvl>' +
    '</w:abstractNum>' +
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>' +
  '</w:numbering>';

async function buildReferentWithBullets() {
  const ref = await buildReferentDocx();
  return await replaceInDocx(ref, 'word/numbering.xml', RICH_NUMBERING_XML);
}

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function run(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', ...opts });
}

async function openXml(blob, file) {
  const zip = await JSZip.loadAsync(blob);
  const f = zip.file(file);
  if (!f) return null;
  return f.async('string');
}

function parse(xml) {
  return new DOMParser().parseFromString(xml, 'application/xml');
}

function firstBodyParagraphMatching(doc, predicate) {
  const body = doc.getElementsByTagName('w:body')[0];
  const ps = body.getElementsByTagName('w:p');
  for (let i = 0; i < ps.length; i++) {
    let inTbl = false;
    let n = ps[i].parentNode;
    while (n) { if (n.nodeName === 'w:tbl') { inTbl = true; break; } n = n.parentNode; }
    if (inTbl) continue;
    const ts = ps[i].getElementsByTagName('w:t');
    let txt = '';
    for (let k = 0; k < ts.length; k++) txt += ts[k].textContent || '';
    if (predicate(txt, ps[i])) return ps[i];
  }
  return null;
}

// -----------------------------------------------------------------------------
// B1 — Espaciado §4.3
// -----------------------------------------------------------------------------

describe('Fase 12 · B1 — enforce espaciado §4.3', () => {
  it('FHJPrrafo → before=0, after=120, line=360, lineRule=auto', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Párrafo de prueba con espaciado heredado raro.', kind: 'plain' }
      ]
    });
    const { stats } = await run({ refFile, contentFile });
    expect(stats.normativa.spacing.applied).toBe(true);
    expect(stats.normativa.spacing.touched).toBeGreaterThan(0);

    // Verificar el XML real — FHJPrrafo debe tener line=360.
    const { blob } = await run({ refFile, contentFile });
    const docXml = await openXml(blob, 'word/document.xml');
    const doc = parse(docXml);
    const p = firstBodyParagraphMatching(doc, t => /espaciado heredado/.test(t));
    expect(p).toBeTruthy();
    const sp = p.getElementsByTagName('w:spacing')[0];
    expect(sp).toBeTruthy();
    expect(sp.getAttribute('w:line')).toBe('360');
    expect(sp.getAttribute('w:after')).toBe('120');
    expect(sp.getAttribute('w:before')).toBe('0');
    expect(sp.getAttribute('w:lineRule')).toBe('auto');
  });

  it('FHJTtulo1 → before=240, after=120, line=360', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' }
      ]
    });
    const { blob } = await run({ refFile, contentFile });
    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /OBJETO/.test(t));
    expect(p).toBeTruthy();
    const sp = p.getElementsByTagName('w:spacing')[0];
    expect(sp.getAttribute('w:before')).toBe('240');
    expect(sp.getAttribute('w:after')).toBe('120');
    expect(sp.getAttribute('w:line')).toBe('360');
  });

  it('desactivar enforceSpacing respeta el spacing original', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:t xml:space="preserve">Texto con line=276</w:t></w:r></w:p>' }
      ]
    });
    const { stats } = await run({ refFile, contentFile, enforceSpacing: false });
    expect(stats.normativa.spacing.applied).toBe(false);
    expect(stats.normativa.spacing.touched).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// B2 — Tipografía Arial 10
// -----------------------------------------------------------------------------

describe('Fase 12 · B2 — enforce Arial 10 §4.2.1', () => {
  it('fuerza Arial + sz=20 en runs del body', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">Calibri 11 heredado</w:t></w:r></w:p>' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.typography.applied).toBe(true);
    expect(stats.normativa.typography.runs).toBeGreaterThan(0);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /Calibri 11/.test(t));
    const r = p.getElementsByTagName('w:r')[0];
    const rFonts = r.getElementsByTagName('w:rFonts')[0];
    const sz = r.getElementsByTagName('w:sz')[0];
    expect(rFonts.getAttribute('w:ascii')).toBe('Arial');
    expect(rFonts.getAttribute('w:hAnsi')).toBe('Arial');
    expect(sz.getAttribute('w:val')).toBe('20');
  });

  it('"Fuente:" bajo tabla → Arial 9 cursiva', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '', kind: 'table' },
        { text: 'Fuente: Registro hospitalario 2024.', kind: 'plain' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.typography.tableSource).toBeGreaterThan(0);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /^Fuente:/.test(t));
    expect(p).toBeTruthy();
    const runs = p.getElementsByTagName('w:r');
    const r0 = runs[0];
    const sz = r0.getElementsByTagName('w:sz')[0];
    const hasI = r0.getElementsByTagName('w:i').length > 0;
    expect(sz.getAttribute('w:val')).toBe('18'); // 9pt
    expect(hasI).toBe(true);
  });
});

// -----------------------------------------------------------------------------
// B3 — Unwrap prosa fragmentada
// -----------------------------------------------------------------------------

describe('Fase 12 · B3 — unwrap prosa fragmentada', () => {
  it('fusiona fragmentos consecutivos sin puntuación fuerte', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Este es un párrafo largo', kind: 'plain' },
        { text: 'que se ha roto en varios', kind: 'plain' },
        { text: 'fragmentos consecutivos por culpa del portapapeles.', kind: 'plain' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.unwrap.applied).toBe(true);
    expect(stats.normativa.unwrap.merged).toBeGreaterThanOrEqual(2);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /portapapeles/.test(t));
    expect(p).toBeTruthy();
    const ts = p.getElementsByTagName('w:t');
    let txt = ''; for (let i = 0; i < ts.length; i++) txt += ts[i].textContent;
    // Debe contener las 3 partes fusionadas
    expect(txt).toContain('Este es un párrafo largo');
    expect(txt).toContain('que se ha roto');
    expect(txt).toContain('portapapeles.');
  });

  it('NO fusiona cuando un fragmento acaba en punto', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Primera frase completa.', kind: 'plain' },
        { text: 'Segunda frase completa.', kind: 'plain' }
      ]
    });
    const { stats } = await run({ refFile, contentFile });
    expect(stats.normativa.unwrap.merged).toBe(0);
  });

  it('desactivar unwrapNarrative deja los fragmentos como párrafos separados', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Fragmento uno', kind: 'plain' },
        { text: 'continuación del anterior sin puntuación.', kind: 'plain' }
      ]
    });
    const { stats } = await run({ refFile, contentFile, unwrapNarrative: false });
    expect(stats.normativa.unwrap.applied).toBe(false);
    expect(stats.normativa.unwrap.merged).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// B4 — Listas §4.2.4
// -----------------------------------------------------------------------------

describe('Fase 12 · B4 — normalizar listas §4.2.4', () => {
  it('fuerza indent canónico (0/357/720 left, 357/363/720 hanging) en párrafos con numPr', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { kind: 'raw', xml: '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:t xml:space="preserve">Lista nivel 1</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Lista nivel 2</w:t></w:r></w:p>' },
        { kind: 'raw', xml: '<w:p><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t xml:space="preserve">Lista nivel 3</w:t></w:r></w:p>' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.lists.applied).toBe(true);
    expect(stats.normativa.lists.paragraphs).toBe(3);
    expect(stats.normativa.lists.byLevel[0]).toBe(1);
    expect(stats.normativa.lists.byLevel[1]).toBe(1);
    expect(stats.normativa.lists.byLevel[2]).toBe(1);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p0 = firstBodyParagraphMatching(doc, t => /Lista nivel 1/.test(t));
    const ind0 = p0.getElementsByTagName('w:ind')[0];
    expect(ind0.getAttribute('w:left')).toBe('0');
    expect(ind0.getAttribute('w:hanging')).toBe('357');

    const p1 = firstBodyParagraphMatching(doc, t => /Lista nivel 2/.test(t));
    const ind1 = p1.getElementsByTagName('w:ind')[0];
    expect(ind1.getAttribute('w:left')).toBe('357');
    expect(ind1.getAttribute('w:hanging')).toBe('363');

    const p2 = firstBodyParagraphMatching(doc, t => /Lista nivel 3/.test(t));
    const ind2 = p2.getElementsByTagName('w:ind')[0];
    expect(ind2.getAttribute('w:left')).toBe('720');
    expect(ind2.getAttribute('w:hanging')).toBe('720');
  });

  it('reescribe lvlText bullet → ●/–/▪ por nivel en numbering.xml', async () => {
    const refFile = await buildReferentWithBullets();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.lists.bulletsRewritten).toBeGreaterThan(0);

    const numXml = await openXml(blob, 'word/numbering.xml');
    expect(numXml).toBeTruthy();
    const doc = parse(numXml);
    const abstractNums = doc.getElementsByTagName('w:abstractNum');
    let foundBullet0 = false;
    for (let i = 0; i < abstractNums.length; i++) {
      const lvls = abstractNums[i].getElementsByTagName('w:lvl');
      for (let j = 0; j < lvls.length; j++) {
        const ilvl = lvls[j].getAttribute('w:ilvl');
        const numFmt = lvls[j].getElementsByTagName('w:numFmt')[0];
        const lvlText = lvls[j].getElementsByTagName('w:lvlText')[0];
        if (numFmt && numFmt.getAttribute('w:val') === 'bullet' && ilvl === '0') {
          expect(lvlText.getAttribute('w:val')).toBe('●');
          foundBullet0 = true;
        }
        if (numFmt && numFmt.getAttribute('w:val') === 'bullet' && ilvl === '1') {
          expect(lvlText.getAttribute('w:val')).toBe('–');
        }
        if (numFmt && numFmt.getAttribute('w:val') === 'bullet' && ilvl === '2') {
          expect(lvlText.getAttribute('w:val')).toBe('▪');
        }
      }
    }
    expect(foundBullet0).toBe(true);
  });
});

// -----------------------------------------------------------------------------
// B6 — Tipografía semántica §4.2.2
// -----------------------------------------------------------------------------

describe('Fase 12 · B6 — tipografía semántica §4.2.2', () => {
  it('aplica cursiva a términos latinos (in situ, ad hoc, et al.)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'Este procedimiento se aplica in situ y ad hoc.', kind: 'plain' },
        { text: 'Según Pérez et al. conviene revisar.', kind: 'plain' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.semantic.applied).toBe(true);
    expect(stats.normativa.semantic.italicTerms).toBeGreaterThanOrEqual(3);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /in situ/.test(t));
    const runs = p.getElementsByTagName('w:r');
    let foundItalicTerm = false;
    for (let i = 0; i < runs.length; i++) {
      const ts = runs[i].getElementsByTagName('w:t');
      let txt = ''; for (let k = 0; k < ts.length; k++) txt += ts[k].textContent;
      const rPr = runs[i].getElementsByTagName('w:rPr')[0];
      const hasI = rPr && rPr.getElementsByTagName('w:i').length > 0;
      if (hasI && (txt === 'in situ' || txt === 'ad hoc')) foundItalicTerm = true;
    }
    expect(foundItalicTerm).toBe(true);
  });

  it('aplica negrita a palabras-alerta (ADVERTENCIA, ATENCIÓN)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'ADVERTENCIA: paciente alérgico a penicilina.', kind: 'plain' }
      ]
    });
    const { stats, blob } = await run({ refFile, contentFile });
    expect(stats.normativa.semantic.boldAlerts).toBeGreaterThanOrEqual(1);

    const doc = parse(await openXml(blob, 'word/document.xml'));
    const p = firstBodyParagraphMatching(doc, t => /ADVERTENCIA/.test(t));
    const runs = p.getElementsByTagName('w:r');
    let foundBoldAlert = false;
    for (let i = 0; i < runs.length; i++) {
      const ts = runs[i].getElementsByTagName('w:t');
      let txt = ''; for (let k = 0; k < ts.length; k++) txt += ts[k].textContent;
      const rPr = runs[i].getElementsByTagName('w:rPr')[0];
      const hasB = rPr && rPr.getElementsByTagName('w:b').length > 0;
      if (hasB && txt === 'ADVERTENCIA') foundBoldAlert = true;
    }
    expect(foundBoldAlert).toBe(true);
  });

  it('desactivar semanticTypography omite cursivas y negritas', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: 'ADVERTENCIA: se aplica in situ.', kind: 'plain' }
      ]
    });
    const { stats } = await run({ refFile, contentFile, semanticTypography: false });
    expect(stats.normativa.semantic.applied).toBe(false);
    expect(stats.normativa.semantic.italicTerms).toBe(0);
    expect(stats.normativa.semantic.boldAlerts).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// Integración — stats.normativa completo
// -----------------------------------------------------------------------------

describe('Fase 12 · integración — stats.normativa', () => {
  it('devuelve stats.normativa con las 5 subclaves', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [{ text: '1.- OBJETO', kind: 'plain' }]
    });
    const { stats } = await run({ refFile, contentFile });
    expect(stats.normativa).toBeTruthy();
    expect(stats.normativa.unwrap).toBeTruthy();
    expect(stats.normativa.spacing).toBeTruthy();
    expect(stats.normativa.typography).toBeTruthy();
    expect(stats.normativa.lists).toBeTruthy();
    expect(stats.normativa.semantic).toBeTruthy();
  });

  it('todos los flags a false → todo "applied: false"', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [{ text: '1.- OBJETO', kind: 'plain' }]
    });
    const { stats } = await run({
      refFile, contentFile,
      enforceSpacing: false,
      enforceTypography: false,
      unwrapNarrative: false,
      normalizeLists: false,
      semanticTypography: false
    });
    expect(stats.normativa.unwrap.applied).toBe(false);
    expect(stats.normativa.spacing.applied).toBe(false);
    expect(stats.normativa.typography.applied).toBe(false);
    expect(stats.normativa.lists.applied).toBe(false);
    expect(stats.normativa.semantic.applied).toBe(false);
  });
});
