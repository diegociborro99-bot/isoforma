# Fase 7 — Resultado

**Versión:** v0.7.0 · Engine v1.5
**Fecha:** abril 2026
**Estado:** ✅ todos los bloques verdes (98 tests passed, 2 skipped por ausencia de fixtures opcionales)

## Objetivo

Cerrar las 3 grietas detectadas en `PHASE7_DIAGNOSTIC.md`:

1. **Clasificador débil** — la regex `^\d+\.- ` capturaba pasos de procedimiento (`"1. Lavarse las manos"`) como `FHJTtulo1` (147 falsos positivos en `content-pnt-claude.docx`).
2. **Headers/footers no transferidos** — `content-pnt-claude.docx` salía con su header genérico en vez del membrete FHJ.
3. **Estilos heredados sin remapear** — 199 párrafos `Normal` y 9 `ListParagraph` ignorados por el motor.

## Antes / Después (caso stress: ref-gestion-info + content-pnt-claude)

| Métrica                              | Antes (Engine v1.4) | Después (Engine v1.5) |
|--------------------------------------|--------------------:|----------------------:|
| FHJTtulo1 mal-clasificados           |                 147 |                 ≤ 30  |
| Drawings preservados                 |               47/47 |                 47/47 |
| Headers FHJ inyectados               |                  no |                    sí |
| Footers FHJ inyectados               |                  no |                    sí |
| Logo + media renombrados con `fhj_`  |                  no |                    sí |
| Párrafos `Normal` remapeados         |                   0 |    todos → FHJPrrafo  |
| Párrafos `ListParagraph` remapeados  |                   0 |  todos → FHJVietaNiv1 |
| Detección de listas vía `numPr`      |                  no |                    sí |
| Validador normativo                  |                  no |       5 reglas activas |

## Bloques implementados

### Bloque A — Clasificador feature-based

- Trust path: si el párrafo ya trae `pStyle` que empieza por `FHJ`, se preserva (cuenta como `preservedStyles`).
- Detección de listas via `numPr` → `numFmt` del `abstractNum`:
  - `bullet` → `FHJVietaNivel{1|2|3}` según `ilvl`.
  - `decimal` / `lowerLetter` / `lowerRoman` → `FHJListaNivel{1|2|3}` según `ilvl`.
- Regex de Título 1 endurecida: requiere `length ≤ 90` Y ratio de mayúsculas ≥ 0.6.
- Fall-through al clasificador feature-based con scoring multi-señal (ALL-CAPS, longitud, prefijo numerado, marcadores ANEXO).

### Bloque B — Preservar imágenes del content

Tests sintéticos + reales validan que `<w:drawing>` y `<w:pict>` se conservan 1:1.

### Bloque C — Transferencia real de headers/footers

- Copia los 6 XML (`header[1-3].xml`, `footer[1-3].xml`) del referente.
- Copia todos los `_rels` asociados.
- Renombra TODA la media del referente con prefijo `fhj_` para evitar colisiones con la media del content (probado con 47 imágenes en content + logo en ref).
- `detectFhjHeader` ahora exige marcadores textuales (`Hospital de Jove`, `FHJ`, código `P.XX.XX.XXX`) además de la imagen embebida — antes una cabecera ajena con cualquier imagen pasaba por FHJ.

### Bloque D — Validador normativo post-proceso

5 reglas, todas no fatales (warnings):

- `NORMATIVA_ALL_CAPS_BODY` — párrafos de cuerpo en mayúsculas.
- `NORMATIVA_UNDERLINE` — texto subrayado fuera de hiperenlaces.
- `NORMATIVA_EMPTY_LIST` — ítems de lista vacíos.
- `NORMATIVA_MISSING_DATOS_GENERALES` — falta tabla de datos generales (raramente dispara: el motor la inyecta automáticamente).
- `NORMATIVA_FONT_NON_ARIAL` — runs con fuente distinta de Arial.

### Bloque E — UI panel de inteligencia

- `index.html` muestra `preservedStyles` y `list` en la meta del resultado.
- `WARNING_LABELS` cubre todos los códigos `NORMATIVA_*`.
- Versión visible: `v0.7 beta` en hero y footer.

## Tests

```
Test Files  9 passed (9)
     Tests  98 passed | 2 skipped (100)
```

- `tests/phase7.test.js` — 15 tests cubren los 4 bloques (A clasificador, B media, C headers, D normativa).
- `tests/fixtures.test.js` — 4 tests Fase 7 sobre fixtures reales del hospital:
  - FHJTtulo1 ≤ 30 sobre stress doc (antes 147).
  - 47 drawings preservados.
  - Warnings normativos disparados (UNDERLINE + FONT_NON_ARIAL).
  - Autoconsistencia ref vs ref: ≥ 100 estilos FHJ preservados.

## Build

- `dist/isoforma-engine.js`: 33.4 KB minificado (de 76 KB fuente).
- `dist/index.html`: 54.8 KB minificado (de 61 KB fuente).
- `node scripts/smoke.js`: ✅ Smoke OK.
