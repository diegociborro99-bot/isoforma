# Fixtures reales

Esta carpeta es donde van los .docx reales del hospital. **No los subas a git si contienen información sensible** — añádelos al `.gitignore` si procede.

## Archivos esperados por el test runner

Los tests en `tests/fixtures.test.js` buscan estos nombres exactos:

| Archivo                    | Rol      | Descripción                                                                          |
|----------------------------|----------|--------------------------------------------------------------------------------------|
| `ref-gestion-info.docx`    | ref      | PNT canónico bien formateado del hospital (Gestión de Información Documentada).      |
| `ref-recomendaciones.docx` | ref      | PNT normativo sobre recomendaciones de elaboración de documentos.                    |
| `content-pnt-claude.docx`  | content  | Documento con 47 imágenes y múltiples secciones — el "stress case" para Fase 7.      |
| `referente.docx`           | ref      | (Legacy) PNT ya formateado — usado por tests de Fase 5.                              |
| `caso-A.docx`              | content  | (Legacy) Doc con cabecera FHJ propia.                                                |
| `caso-B.docx`              | content  | (Legacy) Doc sin cabecera FHJ.                                                       |

Si alguno falta, los tests correspondientes se **saltan automáticamente** con un mensaje claro — el resto de la suite sigue corriendo.

## Qué valida cada uno

- **caso-A**: que el engine detecta la cabecera existente (`preservedHeaders === true`), la conserva, aplica estilos al cuerpo y produce un `document.xml` válido.
- **caso-B**: que el engine inyecta las cabeceras del referente (`preservedHeaders === false`), personaliza el header2 con los metadatos que le pasamos, numera tablas/figuras y produce un `document.xml` válido.

Las asserts usan **rangos en vez de cantidades exactas** (ej: `toBeGreaterThanOrEqual(80)` en vez de `toBe(111)`) para que no se rompan por cambios menores del documento original. Si quieres fijar cantidades exactas, cámbialas tras la primera ejecución verde.

## Añadir más casos

Para añadir un nuevo edge case (ej: documento sin `sectPr`, con múltiples secciones, con bookmarks duplicados, etc.):

1. Deja el .docx aquí con un nombre descriptivo (`edge-sin-sectpr.docx`, etc.).
2. Añade un `describe` block en `tests/fixtures.test.js` siguiendo el mismo patrón de skip automático.

## Privacidad

Si estos archivos contienen datos reales de procedimientos, pacientes o personal:

- No los commitees públicamente.
- Considera usar versiones anonimizadas (cambia código/versión/título pero mantén estructura).
- Añade esta carpeta (excepto este README) al `.gitignore`:

```
tests/fixtures/*
!tests/fixtures/README.md
```
