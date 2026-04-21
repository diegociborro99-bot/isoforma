# Isoforma

**Un documento, mil formas. Un solo formato.**

Herramienta de finalización de documentos PNT del Hospital de Jove. Aplica el formato corporativo a cualquier Word, respetando lo que ya esté bien puesto.

Procesamiento 100% local en el navegador. Ningún archivo se envía a ningún servidor.

---

## Qué hace

Isoforma **detecta qué trae el documento** y actúa en consecuencia:

### Caso A — el documento ya tiene cabecera FHJ propia

Este es el caso habitual: alguien partió de la plantilla oficial del hospital, rellenó la cabecera (código, versión, título) y el pie en Word, y escribió el contenido.

- **Cabecera y pie se respetan tal cual.** Nada que hayas escrito en Word se modifica.
- Se aplican estilos FHJ al texto del cuerpo (títulos, subtítulos, párrafos, viñetas).
- Se inyecta la tabla "Datos generales" al inicio si no existe ya.
- Se numeran las tablas y figuras automáticamente con highlight amarillo para revisar.

### Caso B — el documento no tiene cabecera FHJ

Un procedimiento heredado, un texto pelado sin envolver, un documento que alguien trajo de fuera del hospital.

- Se inyectan los 3 headers y 3 footers del referente, con logo FHJ incluido.
- Si aportaste código / versión / título en el formulario, se rellenan en la cabecera.
- Si no los aportaste, la cabecera queda con los placeholders del referente y los editas directamente en Word.
- Todo lo demás (estilos, datos generales, numeración) igual que en el caso A.

## Cómo usar

1. **Abre `index.html`** en cualquier navegador moderno.
2. **Sube el PNT de referencia** — un documento ya formateado que marca el estándar.
3. **Sube el documento nuevo** — tu Word recién escrito o heredado.
4. **Opcionalmente rellena los datos** (código, versión, título). Solo hacen falta si el documento no trae cabecera propia.
5. **Pulsa "Formatear documento"**. En pocos segundos obtienes un `.docx` listo.
6. **Revisa los highlights amarillos** antes de entregar: son los nombres de tablas y figuras que tienes que completar manualmente.

## Detección automática

Isoforma decide entre caso A y caso B mirando si existe `word/header1.xml`, `header2.xml` o `header3.xml` con:

- Un drawing embebido (imagen).
- Un fichero `_rels` asociado con una relación de tipo imagen.

Si encuentra ambos, asume que hay cabecera FHJ propia y no la toca. Si no, la inyecta del referente.

## Estilos FHJ aplicados

El motor clasifica cada párrafo del cuerpo aplicando estos patrones sobre el texto:

- **`FHJTtulo1`** — `^\d+\.- `, `^\d+\. ` seguido de mayúscula, o `^ANEXO ` (ej: `"1.- OBJETO"`, `"4. DESCRIPCIÓN"`, `"ANEXO I"`).
- **`FHJTtuloprrafo`** — `^\d+\.\d+\.?-? ` o `^\d+\.\d+\.\d+\.?-? ` (ej: `"3.1"`, `"4.2.1"`).
- **`FHJVietaNivel1`** — empieza por `- `, `• `, `· ` o `* `.
- **`FHJPrrafo`** — todo lo demás.

Si aparecen documentos con otras convenciones, se amplían las reglas editando las funciones `classifyParagraph` e `isListItem` en `isoforma-engine.js` (líneas ~220-245).

## Privacidad

El procesamiento ocurre íntegramente dentro del navegador usando JavaScript. Los archivos nunca abandonan el ordenador. No hay servidor intermedio, ni base de datos, ni registros. Se puede usar en la intranet del hospital sin acceso exterior una vez cargada la app.

## Qué se conserva siempre

- Todo el texto del cuerpo y su orden.
- Imágenes, gráficos, diagramas (probado con documentos de 47+ imágenes).
- Tablas y sus contenidos.
- Hyperlinks.

## Requisitos

- Cualquier navegador moderno (últimos 2 años).
- Conexión a internet **solo** la primera vez (para cargar fuentes y librerías desde CDN). Después funciona offline.
- Documentos en formato `.docx` (Word moderno).

## Estructura

```
isoforma/
├── index.html            → Interfaz y lógica de UI
├── isoforma-engine.js    → Motor de transformación docx
└── README.md             → Este documento
```

Self-contained. Sin build step, sin dependencias locales. Librerías JSZip y FileSaver cargadas por CDN.

## Versión

v0.11.0 · Engine v1.9 — abril 2026
Fundación Hospital de Jove · Servicio de Laboratorio

## Changelog

- **v0.11.0 · Engine v1.9 (Fase 12)** — Cumplimiento normativo §4 integral (*Recomendaciones para elaboración de documentos*):
  - **B1 · Enforce espaciado e interlineado §4.3** — nueva pasada `enforceFHJSpacing(doc)` que reescribe `<w:spacing>` por cada estilo FHJ (antes/después/interlineado canónico): `FHJTtulo1` 240/120/1.5, `FHJTtuloprrafo` 240/60/1.5, `FHJPrrafo` 0/120/1.5, `FHJVietaNivel*` 0/60/1.0, `FHJListaNivel*` 0/60/1.0. Forzando `lineRule=auto` y eliminando autospacings huérfanos. Activa por defecto; flag `enforceSpacing: false`.
  - **B2 · Enforce Arial §4.2.1 / §4.2.3.2** — nueva pasada `enforceArialTypography(doc)` que fuerza `rFonts=Arial` y `sz=20` (10 pt) en todos los runs del cuerpo; detecta párrafos "Fuente: …" bajo tabla y les aplica Arial 9 cursiva (`sz=18`, `<w:i/>`) por §4.2.3.2. Activa por defecto; flag `enforceTypography: false`.
  - **B3 · Unwrap prosa fragmentada (Don Quijote)** — `unwrapNarrativeParagraphs(doc)` fusiona párrafos `FHJPrrafo` contiguos sin puntuación fuerte al final cuando el siguiente empieza con minúscula/dígito/coma. Safety-rails: nunca dentro de tabla, con `numPr`, con drawings/picts, con `w:br type=page/column`; máximo 30 fusiones por bloque. **Preserva bookmarks, hyperlinks y comment-refs** al migrar runs (fix colateral de un bug silente).
  - **B4 · Normalizar listas §4.2.4** — `enforceListIndent(doc)` fuerza `<w:ind>` canónico por `ilvl` (0: 0/357 | 1: 357/363 | 2: 720/720 twips). `normalizeListSymbols(outputZip)` reescribe `numbering.xml`: `lvlText` de bullets → `●` / `–` / `▪` por nivel, `numFmt` numerado → `decimal` / `lowerLetter` / `lowerRoman` (respetando `upperLetter`/`upperRoman`/`decimalZero`). Activa por defecto; flag `normalizeLists: false`.
  - **B6 · Tipografía semántica §4.2.2** — `enforceSemanticTypography(doc)` detecta latinismos/extranjerismos ("in situ", "ad hoc", "et al.", "in vitro", "ex profeso", "motu proprio"…) y palabras-alerta ("ADVERTENCIA", "ATENCIÓN", "PRECAUCIÓN", "IMPORTANTE", "NOTA", "CUIDADO") en `<w:t>` y **splitea runs** en los bordes exactos, aplicando `<w:i/>` a los primeros y `<w:b/>` a los segundos (empate → negrita). Preserva `rPr` clonando por run. Activa por defecto; flag `semanticTypography: false`.
  - **Stats enriquecidos**: `stats.normativa = { unwrap, spacing, typography, lists, semantic }` con `applied` + contadores por sub-pasada, alimentando la nueva checklist de UI.
  - **UI · Checklist de cumplimiento normativo** — nuevo grupo de 5 checkboxes bajo "Cumplimiento normativo §4" con badge "Nuevo", y un panel de resultado `result-normativa` que renderiza los 5 chips con iconos `check` / `circle`.
  - Tests: 15 nuevos en `tests/phase12.test.js` (151 total verdes — 2 skipped). Suite completa re-verde tras ajustes a `phase7`/`phase8`/`phase9`/`edge-cases`/`fixtures` y re-snapshot de `snapshot.test.js` por la nueva sub-clave `stats.normativa`.
- **v0.10.0 · Engine v1.8 (Fase 10)** — Modo lote + auto-fix nivel 2 + validador reforzado:
  - **Modo lote**: nuevo toggle en la UI que permite seleccionar (o soltar) N documentos `.docx` contra el mismo referente de una sola vez. El resultado es un `.zip` que contiene todos los documentos formateados + un `_resumen_lote.csv` con estado, fixes y warnings por archivo. En lote los metadatos se autodetectan de cada documento (no se rellenan manualmente).
  - **Auto-fix nivel 2** — tres nuevas correcciones automáticas (opt-in, activas con `autoFix: true`):
    - `blankParas`: colapsa cualquier cluster de párrafos vacíos consecutivos a uno solo.
    - `multiSpace`: dentro de cada `w:t`, sustituye secuencias de 2+ espacios ASCII por uno.
    - `renumbered`: detecta huecos en la numeración de `FHJTtulo1` (p. ej. 1, 2, 4) y reescribe el prefijo numérico al valor correcto. No toca `ANEXO` ni títulos sin prefijo numérico.
  - **Validador reforzado** — nuevo warning `NORMATIVA_PLACEHOLDER_UNFILLED` que detecta placeholders típicos sin rellenar en el cuerpo: `[CODIGO]`, `[TITULO]`, `[VERSION]`, `[FECHA]`, `XXXXX`, `[XX.XX.XX]`, `<<...>>`, `[xxx]`. Evidencia los casos concretos en `context.samples` para la UI.
  - Contadores y `samples` para cada nuevo tipo en `stats.fixes` (`blankParas`, `multiSpace`, `renumbered`) con la misma mecánica que Fase 9 (MAX 3 por tipo, 80 chars, elipsis).
  - Tests: 13 nuevos en `tests/phase10.test.js` (136 total verdes).
- **v0.9.0 · Engine v1.7 (Fase 9)** — Transparencia y extracción de metadatos:
  - `extractMetadata(file)` como alias descriptivo de `inspectContent` (devuelve solo `{ code, version, title }`).
  - Patrones de versión extendidos: además del clásico `V.1.2`, ahora detecta `Versión 1.0`, `Edición 2`, `Rev. 3`, `Revisión 4`.
  - `stats.fixes.samples` — el auto-fix ahora devuelve hasta 3 snippets de texto (truncados a 80 chars) por cada tipo de corrección, para que la UI pueda mostrar evidencia concreta de qué cambió.
  - `samples.font` incluye el nombre de la fuente original (`{ text, font }`).
  - UI: bloque "Correcciones aplicadas automáticamente" con chips visuales por tipo + sección expandible "Ver ejemplos concretos" que muestra los snippets capturados.
  - UI: etiquetas bumped a v0.9.
  - Tests: 14 nuevos en `tests/phase9.test.js` (123 total verdes).
- **v0.8.0 · Engine v1.6 (Fase 8)** — Auto-fix normativo opt-in:
  - Nueva opción `autoFix: true` en `process()` que repara en DOM las 4 infracciones normativas más comunes antes del validador:
    - Subrayado en cuerpo → elimina `<w:u>`.
    - Fuentes no-Arial → reescribe `rFonts` (ascii/hAnsi/cs) a Arial.
    - ALL-CAPS en párrafos de cuerpo → descapitaliza (primera mayúscula, resto minúsculas); respeta títulos `FHJTtulo*`.
    - Ítems de lista vacíos → elimina párrafos `FHJVieta*`/`FHJLista*` sin texto y sin imágenes.
  - Contadores por tipo en `stats.fixes`; bandera `stats.autoFixApplied`.
  - UI: checkbox "Corregir automáticamente antes de exportar" activado por defecto; la meta del resultado muestra cuántas correcciones se aplicaron.
  - Backward compat: `autoFix: false` por defecto en el engine — los tests de Fase 7 siguen viendo warnings normativos como esperan.
  - Tests: 11 nuevos en `tests/phase8.test.js` (109 total verdes).
- **v0.7.0 · Engine v1.5 (Fase 7)** — Mejoras masivas de inteligencia:
  - Clasificador feature-based con scoring multi-señal: el texto tipo `"1. Lavarse las manos"` ya no se mal-clasifica como Título 1 (antes 147/147 falsos positivos → ahora ~15 reales).
  - Path de confianza: si el párrafo ya trae `pStyle` FHJ* se preserva; cuenta como `preservedStyles` en las estadísticas.
  - Detección de listas via `numPr` → `numFmt` del `abstractNum`: bullets (`●/–/▪`) se mapean a `FHJVieta*`, numeradas (`1./a./i.`) a `FHJLista*`.
  - Transferencia real de headers/footers del referente: los 6 XML + sus `_rels` + toda la media con prefijo `fhj_` para evitar colisiones.
  - Detección estricta de cabecera FHJ propia (antes se quedaba con headers ajenos solo por tener una imagen): requiere marcadores textuales (`Hospital de Jove`, `FHJ`, código `P.XX.XX.XXX`).
  - Normal / ListParagraph heredados → remapeados a FHJ equivalente.
  - Validador normativo post-proceso: emite `NORMATIVA_ALL_CAPS_BODY`, `NORMATIVA_UNDERLINE`, `NORMATIVA_EMPTY_LIST`, `NORMATIVA_MISSING_DATOS_GENERALES`, `NORMATIVA_FONT_NON_ARIAL`.
  - Tests: 94 verdes + 4 fixtures reales de stress incluidos.
- **v0.6.0 · Engine v1.2 (Fase 6)** — IsoformaError tipificado (code, step, context, cause, toJSON), runStep wrapper, warnings estructurados, UMD-lite.
- **v0.5 · Engine v1.1** — Lógica condicional: respeta la cabecera propia del documento cuando existe, inyecta la del referente cuando no. Metadatos opcionales. Detección robusta en los 3 headers posibles.
- **v0.4 · Engine v1.0** — Motor completo: estilos FHJ, tabla Datos generales, numeración de tablas/figuras con highlight.
- **v0.3** — Rediseño visual masivo: navbar fijo, cursor custom, orbes animados, secciones how/features/faq.
- **v0.2** — Logo animado, iconografía custom, sistema de animaciones.
- **v0.1** — MVP: transferencia del envoltorio corporativo.

## Validación

Probado contra dos documentos reales del Hospital de Jove:

- **P.02.03.001 Gestión de información documentada** (caso A) → cabecera preservada, 5 títulos + 111 párrafos estilados, validación `All PASSED`.
- **Procedimiento analítico Cobas c303** (caso B) → cabecera inyectada, 15 títulos + 15 subtítulos + 465 párrafos + 11 viñetas + 5 tablas + 47 figuras, validación `All PASSED`.
