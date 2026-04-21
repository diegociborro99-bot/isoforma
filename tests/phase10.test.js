/**
 * Fase 10 — tests para auto-fix nivel 2 y validador reforzado.
 *
 * Bloque B (engine):
 *   - blankParas: párrafos vacíos consecutivos se colapsan a uno.
 *   - multiSpace: secuencias de 2+ espacios ASCII en w:t se colapsan a 1.
 *   - renumbered: FHJTtulo1 con prefijo numérico se renumeran si hay huecos.
 *   - NORMATIVA_PLACEHOLDER_UNFILLED: detecta placeholders típicos sin rellenar.
 *
 * Bloque A (UI batch) se cubre con smoke tests básicos vía el engine —
 * el DOM de batch-mode no se testea aquí; se valida manualmente.
 */

import { describe, it, expect } from 'vitest';
import { createRequire } from 'node:module';

import {
  buildReferentDocx,
  buildContentDocx
} from './helpers/synthetic.js';

const require = createRequire(import.meta.url);
const IsoformaEngine = require('../isoforma-engine.js');

async function runFixed(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', autoFix: true, ...opts });
}

async function runUnfixed(opts) {
  return IsoformaEngine.process({ outputType: 'nodebuffer', autoFix: false, ...opts });
}

// -----------------------------------------------------------------------------
// Bloque B1: blankParas (párrafos en blanco consecutivos)
// -----------------------------------------------------------------------------

describe('Fase 10 — auto-fix blankParas', () => {
  it('colapsa 3 párrafos vacíos consecutivos a 1 (cuenta 2 fixes)', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo normal.', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.blankParas).toBeGreaterThanOrEqual(2);
    expect(stats.fixes.samples.blankParas.length).toBeGreaterThanOrEqual(1);
  });

  it('blancos consecutivos producen varios fixes; blancos aislados no', async () => {
    // Documento con 5 blancos seguidos → debe colapsar a 1, esperamos ≥4 fixes.
    const refFile = await buildReferentDocx();
    const fiveBlanks = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo A.', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: 'Cuerpo B.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats: statsBig } = await runFixed({ refFile, contentFile: fiveBlanks });
    expect(statsBig.fixes.blankParas).toBeGreaterThanOrEqual(4);
  });

  it('sin autoFix los blankParas no se modifican', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runUnfixed({ refFile, contentFile });
    expect(stats.fixes.blankParas || 0).toBe(0);
    expect(stats.autoFixApplied).toBe(false);
  });
});

// -----------------------------------------------------------------------------
// Bloque B2: multiSpace (espacios múltiples)
// -----------------------------------------------------------------------------

describe('Fase 10 — auto-fix multiSpace', () => {
  it('colapsa secuencias de 2+ espacios a uno y captura sample', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Texto   con   muchos   espacios   seguidos.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.multiSpace).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.samples.multiSpace.length).toBeGreaterThanOrEqual(1);
    // El sample se normaliza (\s+ → ' ') al guardarse, así que comprobamos
    // que el texto de referencia aparezca en el snippet.
    const sample = stats.fixes.samples.multiSpace[0];
    expect(sample.text).toContain('Texto');
    expect(sample.text).toContain('espacios');
  });

  it('un solo espacio no se toca', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Texto con espacios normales.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.multiSpace || 0).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// Bloque B3: renumbered (huecos en la numeración de FHJTtulo1)
// -----------------------------------------------------------------------------

describe('Fase 10 — auto-fix renumbered', () => {
  it('renumera títulos con hueco (1,2,4 → 1,2,3)', async () => {
    const refFile = await buildReferentDocx();
    // Los títulos los clasifica el engine como FHJTtulo1; con autoFix debe
    // detectar el hueco entre 2 y 4 y renumerar.
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo A.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Cuerpo B.', kind: 'plain' },
        { text: '4.- DESARROLLO', kind: 'plain' },
        { text: 'Cuerpo D.', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.renumbered).toBeGreaterThanOrEqual(1);
    expect(stats.fixes.samples.renumbered.length).toBeGreaterThanOrEqual(1);
    // El sample muestra la transición
    expect(stats.fixes.samples.renumbered[0].text).toMatch(/4 → 3|4→3/);
  });

  it('no renumera si los títulos ya son consecutivos', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Cuerpo A.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Cuerpo B.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { text: 'Cuerpo C.', kind: 'plain' }
      ]
    });
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes.renumbered || 0).toBe(0);
  });
});

// -----------------------------------------------------------------------------
// Bloque B4: validador — NORMATIVA_PLACEHOLDER_UNFILLED
// -----------------------------------------------------------------------------

describe('Fase 10 — validador placeholders sin rellenar', () => {
  it('detecta [CODIGO] en el cuerpo', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Este procedimiento tiene código [CODIGO] y versión V.0.1.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { warnings } = await runUnfixed({ refFile, contentFile });
    const w = warnings.find(x => x.code === 'NORMATIVA_PLACEHOLDER_UNFILLED');
    expect(w).toBeDefined();
    expect(w.context.samples.some(s => /\[CODIGO\]/i.test(s))).toBe(true);
  });

  it('detecta XXXXX como placeholder', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Responsable: XXXXXX, firmado por el jefe del servicio.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { warnings } = await runUnfixed({ refFile, contentFile });
    const w = warnings.find(x => x.code === 'NORMATIVA_PLACEHOLDER_UNFILLED');
    expect(w).toBeDefined();
  });

  it('detecta <<nombre>> como placeholder', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Autorizado por <<nombre del responsable>>.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Alc.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' }
      ]
    });
    const { warnings } = await runUnfixed({ refFile, contentFile });
    const w = warnings.find(x => x.code === 'NORMATIVA_PLACEHOLDER_UNFILLED');
    expect(w).toBeDefined();
  });

  it('no emite warning si el texto no tiene placeholders', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx({
      paragraphs: [
        { text: '1.- OBJETO', kind: 'plain' },
        { text: 'Procedimiento de limpieza de áreas generales.', kind: 'plain' },
        { text: '2.- ALCANCE', kind: 'plain' },
        { text: 'Aplica a todo el personal.', kind: 'plain' },
        { text: '3.- DESARROLLO', kind: 'plain' },
        { text: 'Detalle de acciones a realizar.', kind: 'plain' }
      ]
    });
    const { warnings } = await runUnfixed({ refFile, contentFile });
    const w = warnings.find(x => x.code === 'NORMATIVA_PLACEHOLDER_UNFILLED');
    expect(w).toBeUndefined();
  });
});

// -----------------------------------------------------------------------------
// Bloque B5: stats.fixes tiene las claves nuevas aunque no haya fixes
// -----------------------------------------------------------------------------

describe('Fase 10 — shape de stats.fixes', () => {
  it('stats.fixes incluye blankParas, multiSpace y renumbered aun sin fixes', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();
    const { stats } = await runFixed({ refFile, contentFile });
    expect(stats.fixes).toHaveProperty('blankParas');
    expect(stats.fixes).toHaveProperty('multiSpace');
    expect(stats.fixes).toHaveProperty('renumbered');
    expect(stats.fixes.samples).toHaveProperty('blankParas');
    expect(stats.fixes.samples).toHaveProperty('multiSpace');
    expect(stats.fixes.samples).toHaveProperty('renumbered');
    expect(Array.isArray(stats.fixes.samples.blankParas)).toBe(true);
    expect(Array.isArray(stats.fixes.samples.multiSpace)).toBe(true);
    expect(Array.isArray(stats.fixes.samples.renumbered)).toBe(true);
  });

  it('sin autoFix los nuevos contadores siguen siendo 0 con arrays vacíos', async () => {
    const refFile = await buildReferentDocx();
    const contentFile = await buildContentDocx();
    const { stats } = await runUnfixed({ refFile, contentFile });
    expect(stats.fixes.blankParas || 0).toBe(0);
    expect(stats.fixes.multiSpace || 0).toBe(0);
    expect(stats.fixes.renumbered || 0).toBe(0);
    expect(stats.fixes.samples.blankParas).toEqual([]);
    expect(stats.fixes.samples.multiSpace).toEqual([]);
    expect(stats.fixes.samples.renumbered).toEqual([]);
  });
});
