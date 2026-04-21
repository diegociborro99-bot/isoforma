#!/usr/bin/env node
/**
 * Smoke test del build minificado:
 *   - Carga dist/isoforma-engine.js
 *   - Ejecuta process() con fixtures sintéticos
 *   - Verifica que sale un blob válido (nodebuffer) y sin errores
 *
 * Se lanza desde la raíz del proyecto con:
 *   node scripts/smoke.js
 */

'use strict';

const path = require('node:path');
const { createRequire } = require('node:module');

const ROOT = path.resolve(__dirname, '..');
const DIST_ENGINE = path.join(ROOT, 'dist', 'isoforma-engine.js');

// createRequire con un path absoluto válido (en vez de __filename de eval).
const requireFromRoot = createRequire(path.join(ROOT, 'package.json'));

async function main() {
  // helpers sintéticos están en ESM; los importamos con import().
  const synthetic = await import(
    require('url').pathToFileURL(path.join(ROOT, 'tests/helpers/synthetic.js')).href
  );

  const Engine = requireFromRoot(DIST_ENGINE);

  // Sanity de la API pública.
  for (const fn of ['process', 'inspectContent', 'IsoformaError']) {
    if (typeof Engine[fn] !== 'function') {
      throw new Error(`Export ausente: ${fn}`);
    }
  }

  // Fixtures sintéticos mínimos.
  const refFile = await synthetic.buildReferentDocx();
  const contentFile = await synthetic.buildContentDocx();

  const result = await Engine.process({
    refFile,
    contentFile,
    outputType: 'nodebuffer'
  });

  if (!result || !result.blob) {
    throw new Error('process() no devolvió blob');
  }
  if (!(result.blob instanceof Buffer)) {
    throw new Error('blob no es Buffer (outputType=nodebuffer)');
  }
  if (result.blob.length < 1000) {
    throw new Error('blob demasiado pequeño (' + result.blob.length + ' bytes)');
  }

  // Cabecera ZIP: 'PK\x03\x04'
  const sig = result.blob.slice(0, 4).toString('hex');
  if (sig !== '504b0304') {
    throw new Error('blob no parece ZIP (sig=' + sig + ')');
  }

  // inspectContent también debe funcionar.
  const meta = await Engine.inspectContent(contentFile);
  if (!meta || typeof meta !== 'object') {
    throw new Error('inspectContent() no devolvió objeto');
  }

  // Report.
  console.log('Smoke OK');
  console.log('  blob size:  ' + (result.blob.length / 1024).toFixed(1) + ' KB');
  console.log('  warnings:   ' + (result.warnings ? result.warnings.length : 0));
  console.log('  inspectContent keys: ' + Object.keys(meta).join(', '));
}

main().catch((err) => {
  console.error('Smoke FALLÓ:', err && err.stack || err);
  process.exit(1);
});
