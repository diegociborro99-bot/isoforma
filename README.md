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

v0.7.0 · Engine v1.5 — abril 2026
Fundación Hospital de Jove · Servicio de Laboratorio

## Changelog

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
