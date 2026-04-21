#!/usr/bin/env node
/**
 * Build para producción:
 *   - Minifica isoforma-engine.js con terser.
 *   - Minifica index.html con html-minifier-terser (minifica también inline
 *     <script> y <style>).
 *   - Copia README.md como referencia (opcional; no se sirve).
 *   - Añade un 404.html que redirige a / para que los deep links de
 *     GitHub Pages no rompan.
 *
 * Salida: ./dist/
 *
 * El workflow de GitHub Actions llama a `npm run build` y sube ./dist/ como
 * artefacto de Pages.
 */

'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { minify: minifyJs } = require('terser');
const { minify: minifyHtml } = require('html-minifier-terser');

const ROOT = path.resolve(__dirname, '..');
const DIST = path.join(ROOT, 'dist');

function readUtf8(file) {
  return fs.readFileSync(path.join(ROOT, file), 'utf8');
}

function writeFile(relPath, content) {
  const dest = path.join(DIST, relPath);
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, content);
}

function sizeKb(s) {
  const bytes = Buffer.byteLength(s, 'utf8');
  return (bytes / 1024).toFixed(1) + ' KB';
}

async function buildEngine() {
  const src = readUtf8('isoforma-engine.js');
  const out = await minifyJs(src, {
    compress: {
      passes: 2,
      drop_console: false, // conservamos console.error/warn para debugging en prod
      pure_funcs: []
    },
    mangle: {
      // Protege nombres exportados por la API pública UMD.
      reserved: ['IsoformaEngine', 'IsoformaError', 'process', 'inspectContent', 'toJSON']
    },
    format: {
      comments: /^\s*Isoforma Engine/, // conserva el banner de versión
      ecma: 2017
    },
    sourceMap: false
  });
  if (!out.code) throw new Error('terser no devolvió código');
  writeFile('isoforma-engine.js', out.code);
  console.log(`  engine:  ${sizeKb(src)}  →  ${sizeKb(out.code)}`);
}

async function buildHtml() {
  const src = readUtf8('index.html');
  const out = await minifyHtml(src, {
    collapseWhitespace: true,
    conservativeCollapse: false,
    removeComments: true,
    removeRedundantAttributes: true,
    removeScriptTypeAttributes: true,
    removeStyleLinkTypeAttributes: true,
    useShortDoctype: true,
    minifyCSS: true,
    minifyJS: {
      compress: { passes: 2 },
      mangle: true,
      format: { ecma: 2017 }
    },
    // Evita romper los atributos del viewer SVG inline / data URIs.
    keepClosingSlash: true,
    preserveLineBreaks: false
  });
  writeFile('index.html', out);
  console.log(`  html:    ${sizeKb(src)}  →  ${sizeKb(out)}`);
}

function writeExtras() {
  // 404.html = copia de index.html — así los deep links #procesador funcionan
  // aunque GitHub Pages reciba una URL "limpia" inexistente.
  const indexDist = path.join(DIST, 'index.html');
  fs.copyFileSync(indexDist, path.join(DIST, '404.html'));
  // .nojekyll para que GitHub Pages no procese la carpeta con Jekyll.
  fs.writeFileSync(path.join(DIST, '.nojekyll'), '');
  console.log('  extras:  404.html + .nojekyll');
}

async function main() {
  // Limpia dist/.
  if (fs.existsSync(DIST)) fs.rmSync(DIST, { recursive: true, force: true });
  fs.mkdirSync(DIST, { recursive: true });

  console.log('Construyendo dist/ ...');
  await buildEngine();
  await buildHtml();
  writeExtras();
  console.log('Listo. dist/ generado en', DIST);
}

main().catch((err) => {
  console.error('Build falló:', err);
  process.exit(1);
});
