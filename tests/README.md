# Tests — Fase 1: harness golden

Suite de tests del motor de Isoforma que ejecuta en Node sin necesidad de navegador.

## Cómo correr

```bash
npm install
npm test
```

Los tests golden sintéticos **siempre corren** y no requieren fixtures reales. Los tests contra documentos reales del hospital (`tests/fixtures.test.js`) se saltan automáticamente si los .docx no están presentes en `tests/fixtures/`.

## Estructura

```
tests/
├── engine.test.js          → 20+ asserts sobre docx sintéticos
├── fixtures.test.js        → asserts sobre caso A y caso B reales
├── fixtures/
│   └── README.md           → nombres de archivo esperados
└── helpers/
    ├── docx.js             → unpackDocx, countPStyle, isValidXml, ...
    └── synthetic.js        → buildReferentDocx, buildContentDocx
```

## Qué cubre la suite sintética

- **Caso B**: inyección de 3 headers + 3 footers, 6 relaciones en `document.xml.rels`, renombrado del logo, personalización del header2 con metadatos, registro de content types (header/footer/emf).
- **Caso A**: detección de la cabecera FHJ propia, conservación del drawing original, aplicación de estilos al body igualmente.
- **Clasificación de párrafos**: variantes de numeración de títulos (`1.-`, `1.`, `1.1`, `1.1.1`, `ANEXO`), viñetas con distintos bullets (`-`, `•`, `·`, `*`), exclusión de párrafos vacíos.
- **Idempotencia**: no duplica "Datos generales" si ya existe al inicio del documento.
- **Integridad estructural**: el `document.xml` resultante es XML válido y tiene un único `<w:body>`.

## Qué NO cubre todavía (Fase 2+)

- Tablas con atributos en la apertura (`<w:tbl w:rsidR="...">`). El motor las ignora — bug conocido, se arreglará al pasar a DOM.
- Documentos con múltiples `<w:sectPr>` (múltiples secciones).
- Referente sin logo EMF (logo en PNG u otro formato).
- Entidades HTML escapadas en títulos (`&amp;`).
- Bookmarks anidados o desequilibrados preexistentes.

Cada edge case no cubierto se corresponde con un ítem de la lista de fragilidad del review. Conforme se vayan cubriendo con tests + fix, se mueven a la sección "Qué cubre".
