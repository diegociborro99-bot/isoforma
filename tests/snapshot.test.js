/**
 * Fase 4 — snapshot test.
 * Congela una huella compacta y determinista del engine sobre un input sintético
 * fijo. Cualquier cambio no-intencional en clasificación, bookmarks, sectPr o
 * numeración de tablas/figuras rompe el snapshot y obliga a revisión.
 *
 * El snapshot se guarda en tests/__snapshots__/snapshot.test.js.snap y debe
 * commitearse al repo. Para actualizar intencionadamente: `npx vitest -u`.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';
import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

import {
  buildReferentDocx,
  buildContentDocx
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

/**
 * Extrae una huella determinista del word/document.xml post-procesado:
 * recuentos, secuencia de pStyle por párrafo del body, resumen de sectPr y
 * una vista normalizada de bookmarks.
 */
async function fingerprint(outputBuffer) {
  const zip = await JSZip.loadAsync(outputBuffer);
  const xml = await zip.file('word/document.xml').async('string');
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const body = doc.getElementsByTagName('w:body')[0];

  // Secuencia de pStyle por párrafo del body (null si no hay estilo).
  const pStyles = [];
  for (let n = body.firstChild; n; n = n.nextSibling) {
    if (n.nodeType !== 1 || n.nodeName !== 'w:p') continue;
    const pPr = Array.from(n.childNodes).find(c => c.nodeName === 'w:pPr');
    if (!pPr) { pStyles.push(null); continue; }
    const pStyle = Array.from(pPr.childNodes).find(c => c.nodeName === 'w:pStyle');
    pStyles.push(pStyle ? pStyle.getAttribute('w:val') : null);
  }

  // Resumen de bookmarks: ids ordenados.
  const bookmarkStartIds = Array.from(doc.getElementsByTagName('w:bookmarkStart'))
    .map(b => b.getAttribute('w:id'))
    .filter(Boolean)
    .sort();
  const bookmarkEndIds = Array.from(doc.getElementsByTagName('w:bookmarkEnd'))
    .map(b => b.getAttribute('w:id'))
    .filter(Boolean)
    .sort();

  // sectPr primario: serializado tal cual para congelar headerRefs/pgSz/pgMar/cols/docGrid.
  const sectPrs = doc.getElementsByTagName('w:sectPr');
  const firstSectPrXml = sectPrs.length
    ? new XMLSerializer().serializeToString(sectPrs[0])
    : null;

  return {
    pStyleSequence: pStyles,
    bodyParagraphCount: pStyles.length,
    bodyTableCount: (function () {
      // Sólo tablas del body, no anidadas.
      const tables = Array.from(body.getElementsByTagName('w:tbl'));
      return tables.filter(t => {
        let p = t.parentNode;
        while (p) {
          if (p === body) return true;
          if (p.nodeName === 'w:tbl') return false;
          p = p.parentNode;
        }
        return false;
      }).length;
    })(),
    bookmarkStartIds,
    bookmarkEndIds,
    firstSectPr: firstSectPrXml
  };
}

describe('snapshot — document.xml post-procesado', () => {
  it('produce una huella determinista sobre un input sintético fijo', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx(); // usa defaultBodyParagraphs
    const { blob, stats, warnings } = await IsoformaEngine.process({
      refFile,
      contentFile,
      metadata: {
        code: 'P.99.00.999',
        version: 'V.1.0',
        title: 'Snapshot fijo'
      },
      outputType: 'nodebuffer'
    });

    const fp = await fingerprint(blob);

    expect({
      stats,
      warnings,
      fingerprint: fp
    }).toMatchSnapshot();
  });
});
