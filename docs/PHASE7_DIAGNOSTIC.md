# Fase 7 — Diagnóstico del engine v0.6 sobre fixtures reales

## Setup

Tres docx reales en `tests/fixtures/`:

| Archivo | Rol | paraCount | tblCount | drawings | media | headers | footers | Estilos top |
|---------|-----|-----------|----------|----------|-------|---------|---------|-------------|
| `ref-gestion-info.docx` | ref canónico | 119 | 1 | 0 | 1 | 1 | 1 | FHJVietaNivel1 (61), FHJTtuloprrafo (28), FHJPrrafo (18) |
| `ref-recomendaciones.docx` | ref normativa | 180 | 5 | 0 | 1 | 3 | 3 | FHJPrrafo (145), FHJTtuloprrafo (16), FHJVietaNivel1 (7) |
| `content-pnt-claude.docx` | content stress | 777 | 6 | 47 | 48 | 3 | 3 | FHJPrrafo (516), Normal (199), FHJTtulo1 (15), FHJTtuloprrafo (15) |

## Casos probados

Se corrieron tres casos via `scripts/diagnose-real.js`:

1. **Stress**: `ref-gestion-info` + `content-pnt-claude` — ref canónico contra contenido denso con 47 imágenes.
2. **Theory**: `ref-recomendaciones` + `content-pnt-claude` — ref con normativa contra contenido denso.
3. **Self-test**: `ref-gestion-info` + `ref-recomendaciones` — dos PNT buenos entre sí para ver auto-consistencia.

## Hallazgos

### H1. ✅ Imágenes: preservadas correctamente

En los tres casos el conteo de `<w:drawing>` y media files se preserva 1:1 desde content hasta output. El "dolor" percibido de imágenes rotas NO se reproduce con el engine actual sobre estos fixtures.

**Conclusión**: Bloque B (preservar imágenes) **no es prioritario**. Cuando el usuario vea imágenes rotas, necesitamos un caso concreto para reproducir.

### H2. ❌ Headers/footers: NO se transfieren desde el ref

En todos los casos, el output mantiene los headers/footers del **content**, no del ref:

- Caso Stress: ref tiene 1+1 headers/footers (membrete FHJ), content tiene 3+3 (distintos) → output: 3+3 del content. El membrete del Hospital de Jove NO aparece.
- Caso Self-test: ref tiene 1+1, content tiene 3+3 → output: 3+3. Mismo fallo.

**Impacto**: el output no lleva el membrete correcto del FHJ — uno de los dolores principales del usuario.

**Prioridad**: ALTA.

### H3. ❌ Clasificación: FHJTtulo1 se dispara (x10)

En el caso Stress:
- Content tiene **15** párrafos como FHJTtulo1
- Output tiene **147** párrafos como FHJTtulo1

Es decir, el classifier reclasifica ~130 párrafos como Título 1 que no lo son. El conteo de FHJPrrafo bajó de 516 a 448 (68 perdidos), compensando parcialmente.

**Causa probable**: la regex de Título 1 mete como título cualquier párrafo que empiece con un número y un punto, pero `pnt-claude.docx` tiene pasos de procedimientos tipo "1. Lavarse las manos", "2. Aplicar..." que son listas, no títulos.

**Impacto**: el output visualmente queda arruinado — todos los pasos de procedimiento se muestran con tipografía y estilo de Título 1.

**Prioridad**: MÁXIMA.

### H4. ⚠️ 199 párrafos estilo `Normal` no se tocan

El content tiene 199 `<w:pStyle w:val="Normal"/>` que el engine deja pasar intactos.

**Decisión pendiente**: ¿convertirlos a FHJPrrafo (normalización total) o respetarlos (preservación)?

La normativa dice "Arial 10, interlineado 1.5, alineado a izquierda" — si el estilo `Normal` del content no cumple esto, queda out-of-brand. La solución correcta es **detectar `Normal` como sinónimo de FHJPrrafo** y reclasificar.

**Prioridad**: MEDIA.

### H5. ⚠️ STYLES_UNKNOWN_ID: 1 estilo huérfano

En los tres casos aparece la warning de que algún estilo referenciado por el body no existe en styles.xml tras el merge. Probable culpable: `ListParagraph` (9 apariciones en content).

**Causa**: `ListParagraph` no es FHJ ni está en styles.xml del ref → tras el merge, sobrevive la referencia pero no la definición.

**Solución**: al detectar `ListParagraph` en content, convertir a `FHJVietaNivel1` / `FHJListaNivel2` según corresponda, o preservar el estilo (copiarlo del content al merge).

**Prioridad**: MEDIA.

### H6. ⚠️ Numbering: 41 numIds remapeados

Todos los casos generan `NUMBERING_MERGED_REMAPPED` con 41 numIds renombrados. Es el path esperado — ya funciona. El warning es informativo.

## Plan de Fase 7 revisado

En orden de impacto:

1. **Bloque A1: Clasificador — regex de Título 1 más estricta** (H3)
   - Anclar al comienzo de línea + espacio
   - Requiere que el resto del texto sea corto y ALL-CAPS o capitalized
   - Si el párrafo tiene `<w:numPr>` (numbered list) → nunca es título, es lista

2. **Bloque A2: Path "trust existing FHJ pStyle"** (H3, H4)
   - Si el content ya tiene `FHJ*` → mantener
   - Si tiene `Normal` → convertir a `FHJPrrafo`
   - Si tiene `ListParagraph` → convertir según contexto (lista numerada → FHJListaNivel, bullets → FHJVietaNivel)

3. **Bloque A3: Detección de listas por w:numPr** (H3)
   - Si un párrafo tiene `<w:numPr>` con numId → mapear a FHJVieta* o FHJLista*
   - Consultar el `<w:lvl>` del abstractNum para determinar nivel (0, 1, 2) y tipo (bullet, decimal, lowerLetter, lowerRoman)

4. **Bloque C: Transferir headers/footers del ref** (H2)
   - Copiar `word/header*.xml`, `word/footer*.xml` del ref
   - Update `word/_rels/document.xml.rels`
   - Update `[Content_Types].xml`
   - Actualizar `w:headerReference`/`w:footerReference` en `<w:sectPr>` del body

5. **Bloque D: Validador normativo** (bonus, a partir de la normativa)
   - ALL-CAPS en body → warning
   - Subrayado → warning
   - Listas vacías → warning
   - Datos generales ausente → warning

6. **Bloque B: Imágenes** (H1)
   - Descartado como prioridad salvo reporte concreto de fallo.

## Next steps inmediatos

- [x] Guardar fixtures en `tests/fixtures/`
- [x] Correr diagnóstico y documentar hallazgos
- [ ] Implementar Bloque A (classifier) — máximo impacto visual
- [ ] Implementar Bloque C (headers/footers) — resuelve membrete
- [ ] Añadir validador normativo
- [ ] Tests con fixtures reales (no snapshot exacto — rangos)
- [ ] Bump a 0.7.0

---

*Generado desde `scripts/diagnose-real.js` — engine v0.6.0 sobre fixtures reales.*
