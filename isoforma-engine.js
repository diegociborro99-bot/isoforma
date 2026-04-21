/**
 * Isoforma Engine v1.8
 * Motor de transformación de documentos PNT del Hospital de Jove.
 *
 * Comportamiento:
 *   - Si el documento nuevo YA tiene header/footer FHJ → se respetan tal cual.
 *     (Se aplican estilos al cuerpo, tabla Datos generales y numeración).
 *   - Si el documento nuevo NO tiene header/footer FHJ → se inyectan desde el referente.
 *     (Si se aportan metadatos de código/versión/título, se personalizan).
 *
 * Entry point: IsoformaEngine.process({ refFile, contentFile, metadata, onProgress })
 *   metadata es opcional; si no se aporta o viene vacío, no se personaliza el header.
 *
 * Returns: Promise<{ blob, stats, warnings }>
 *   - warnings: array de { code, message, context } no fatales
 *   - en error: lanza IsoformaError con { code, step, context, message }
 *
 * UMD-lite: exports as window.IsoformaEngine en browser, module.exports en Node.
 * En Node requiere 'jszip' y '@xmldom/xmldom' instalados como devDeps.
 *
 * Fase 2: todas las mutaciones sobre word/document.xml se hacen por DOM
 * (un solo parse, un solo serialize) en lugar de regex sobre string.
 *
 * Fase 3: preflight validation + errores tipados con contexto.
 *   Cualquier excepción genérica del pipeline se traduce a IsoformaError con
 *   `step` indicando el paso afectado. Los caminos blandos (referente sin
 *   numbering.xml, customizeHeader sin header2 válido, etc.) emiten warnings
 *   no fatales en lugar de fallar silenciosamente.
 */

(function () {
  'use strict';

  // Resolve deps: require() en Node, globals en browser.
  var JSZip, DOMParser, XMLSerializer;
  var isNode = (typeof module !== 'undefined' && module.exports && typeof require === 'function');

  if (isNode) {
    JSZip = require('jszip');
    var xmldom = require('@xmldom/xmldom');
    DOMParser = xmldom.DOMParser;
    XMLSerializer = xmldom.XMLSerializer;
  } else {
    var g = (typeof window !== 'undefined') ? window : globalThis;
    JSZip = g.JSZip;
    DOMParser = g.DOMParser;
    XMLSerializer = g.XMLSerializer;
  }

  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  // ============================================================
  // Fase 3: error model
  // ============================================================

  class IsoformaError extends Error {
    constructor({ code, message, step, context, cause }) {
      super(message);
      this.name = 'IsoformaError';
      this.code = code || 'UNKNOWN';
      this.step = step || null;
      this.context = context || {};
      if (cause) this.cause = cause;
    }
    toJSON() {
      return { name: this.name, code: this.code, step: this.step, message: this.message, context: this.context };
    }
  }

  /**
   * Envoltura para cada paso del pipeline. Si fn() lanza algo que no es ya
   * IsoformaError, lo traducimos a un IsoformaError(STEP_FAILED) anotado con
   * el nombre del paso para que el caller sepa exactamente dónde se rompió.
   */
  async function runStep(stepName, fn) {
    try {
      return await fn();
    } catch (err) {
      if (err && err.name === 'IsoformaError') {
        if (!err.step) err.step = stepName;
        throw err;
      }
      throw new IsoformaError({
        code: 'STEP_FAILED',
        step: stepName,
        message: 'Falló durante "' + stepName + '": ' + (err && err.message ? err.message : String(err)),
        context: { originalError: err && err.message ? err.message : String(err) },
        cause: err
      });
    }
  }

  /**
   * Carga un buffer/blob como zip y traduce errores de JSZip a IsoformaError
   * (INVALID_DOCX) con contexto de qué archivo falló.
   */
  async function loadDocxZip(buffer, role) {
    if (buffer == null) {
      // step se rellena en runStep para evitar duplicar la fuente de verdad.
      throw new IsoformaError({
        code: 'MISSING_INPUT',
        message: 'No se ha aportado el archivo "' + role + '".',
        context: { role }
      });
    }
    try {
      return await JSZip.loadAsync(buffer);
    } catch (err) {
      throw new IsoformaError({
        code: 'INVALID_DOCX',
        message: 'El archivo "' + role + '" no es un .docx válido (no se pudo abrir como zip).',
        context: { role, originalError: err && err.message ? err.message : String(err) },
        cause: err
      });
    }
  }

  /**
   * Validación temprana de la estructura del refZip y del contentZip.
   * Detecta los errores fatales (sin los cuales no podemos continuar) y
   * acumula warnings para condiciones no fatales.
   */
  async function preflight({ refZip, contentZip }) {
    const warnings = [];

    // --- contentZip (obligatorio) ---
    if (!contentZip.file('[Content_Types].xml')) {
      throw new IsoformaError({
        code: 'CONTENT_NOT_DOCX',
        step: 'preflight',
        message: 'El archivo de contenido no parece un documento Word: falta [Content_Types].xml.',
        context: { role: 'content', missing: '[Content_Types].xml' }
      });
    }
    const contentDocFile = contentZip.file('word/document.xml');
    if (!contentDocFile) {
      throw new IsoformaError({
        code: 'CONTENT_MISSING_DOCUMENT_XML',
        step: 'preflight',
        message: 'El archivo de contenido no es un .docx válido: no contiene word/document.xml.',
        context: { role: 'content', missing: 'word/document.xml' }
      });
    }
    const contentDocXml = await contentDocFile.async('string');
    if (!/<w:body[\s>]/.test(contentDocXml)) {
      throw new IsoformaError({
        code: 'CONTENT_MISSING_BODY',
        step: 'preflight',
        message: 'El word/document.xml del contenido no contiene <w:body>. ¿Está corrupto?',
        context: { role: 'content', file: 'word/document.xml' }
      });
    }

    // --- refZip (obligatorio) ---
    if (!refZip.file('[Content_Types].xml')) {
      throw new IsoformaError({
        code: 'REFERENT_NOT_DOCX',
        step: 'preflight',
        message: 'El referente no parece un documento Word: falta [Content_Types].xml.',
        context: { role: 'referent', missing: '[Content_Types].xml' }
      });
    }
    if (!refZip.file('word/document.xml')) {
      throw new IsoformaError({
        code: 'REFERENT_MISSING_DOCUMENT_XML',
        step: 'preflight',
        message: 'El referente no es un .docx válido: no contiene word/document.xml.',
        context: { role: 'referent', missing: 'word/document.xml' }
      });
    }
    if (!refZip.file('word/styles.xml')) {
      throw new IsoformaError({
        code: 'REFERENT_MISSING_STYLES',
        step: 'preflight',
        message: 'El referente no tiene word/styles.xml. Sin estilos no podemos aplicar el formato FHJ al contenido.',
        context: { role: 'referent', missing: 'word/styles.xml' }
      });
    }

    // --- referente: condiciones no fatales (warnings) ---
    const refHeaders = ['header1', 'header2', 'header3'].filter(n => refZip.file('word/' + n + '.xml'));
    if (refHeaders.length === 0) {
      warnings.push({
        code: 'REFERENT_NO_HEADERS',
        message: 'El referente no contiene headers (header1/2/3.xml). Si el contenido tampoco trae cabecera FHJ propia, el resultado quedará sin cabecera.',
        context: { role: 'referent' }
      });
    } else if (refHeaders.length < 3) {
      warnings.push({
        code: 'REFERENT_PARTIAL_HEADERS',
        message: 'El referente solo tiene ' + refHeaders.length + ' de 3 headers (' + refHeaders.join(', ') + '). El resultado puede no incluir header en todas las páginas.',
        context: { role: 'referent', present: refHeaders }
      });
    }
    const refFooters = ['footer1', 'footer2', 'footer3'].filter(n => refZip.file('word/' + n + '.xml'));
    if (refFooters.length === 0) {
      warnings.push({
        code: 'REFERENT_NO_FOOTERS',
        message: 'El referente no contiene footers (footer1/2/3.xml).',
        context: { role: 'referent' }
      });
    }
    if (!refZip.file('word/numbering.xml')) {
      warnings.push({
        code: 'REFERENT_NO_NUMBERING',
        message: 'El referente no tiene word/numbering.xml. Las viñetas y listas numeradas pueden no respetar el formato FHJ.',
        context: { role: 'referent' }
      });
    }
    if (!refZip.file('word/theme/theme1.xml')) {
      warnings.push({
        code: 'REFERENT_NO_THEME',
        message: 'El referente no tiene word/theme/theme1.xml. Los colores y fuentes del tema pueden no transferirse.',
        context: { role: 'referent' }
      });
    }

    return { warnings };
  }

  async function process({ refFile, contentFile, metadata, onProgress, outputType, autoFix }) {
    const progress = onProgress || (() => {});
    metadata = metadata || {};
    // outputType: 'blob' (browser default) | 'nodebuffer' | 'uint8array' | 'arraybuffer'
    const blobType = outputType || (isNode ? 'nodebuffer' : 'blob');
    // autoFix: false por defecto (backward-compat). Pasar { autoFix: true }
    // para aplicar correcciones normativas antes del validador. La UI lo
    // activa por defecto mediante un checkbox.
    const doAutoFix = autoFix === true;
    const warnings = [];
    let fixes = {
      underline: 0, font: 0, allCaps: 0, emptyList: 0,
      blankParas: 0, multiSpace: 0, renumbered: 0,
      samples: {
        underline: [], font: [], allCaps: [], emptyList: [],
        blankParas: [], multiSpace: [], renumbered: []
      }
    };

    progress('Leyendo documentos');
    const refZip = await runStep('Leyendo referente', () => loadDocxZip(refFile, 'referente'));
    const outputZip = await runStep('Leyendo contenido', () => loadDocxZip(contentFile, 'contenido'));

    progress('Validando estructura');
    const pre = await runStep('preflight', () => preflight({ refZip, contentZip: outputZip }));
    warnings.push(...pre.warnings);

    progress('Analizando documento');
    const alreadyHasFhjHeader = await runStep('Detectando cabecera FHJ', () => detectFhjHeader(outputZip));

    progress('Transfiriendo estilos y tema del referente');
    const styleMerge = await runStep('Fusionando estilos del referente', () => mergeStylesAndTransferCore(refZip, outputZip));
    warnings.push(...styleMerge.warnings);

    let relIds = null;

    if (alreadyHasFhjHeader) {
      progress('Header y pie de página existentes conservados');
    } else {
      progress('Instalando cabeceras y pies del referente');
      await runStep('Transfiriendo cabeceras/pies del referente', () => transferHeadersFootersAndLogo(refZip, outputZip));

      if (hasMetadata(metadata)) {
        progress('Personalizando cabecera con datos aportados');
        const customResult = await runStep('Personalizando cabecera', () => customizeHeader(outputZip, metadata));
        if (customResult && customResult.warning) warnings.push(customResult.warning);
      }

      progress('Actualizando relaciones');
      relIds = await runStep('Actualizando relaciones del documento', () => updateDocumentRels(outputZip));
    }

    if (alreadyHasFhjHeader && hasMetadata(metadata)) {
      // Si el contenido trae su propia cabecera, los metadatos no se pueden aplicar.
      warnings.push({
        code: 'METADATA_IGNORED_PRESERVED_HEADER',
        message: 'Se han aportado metadatos pero el contenido ya trae su propia cabecera FHJ. La cabecera del contenido se ha conservado tal cual y los metadatos NO se han aplicado.',
        context: { metadata }
      });
    }

    // ---- document.xml: un único parse → mutaciones DOM → un único serialize ----
    progress('Parseando document.xml');
    const docDoc = await runStep('Parseando word/document.xml', async () => {
      const documentXmlIn = await outputZip.file('word/document.xml').async('string');
      return new DOMParser().parseFromString(documentXmlIn, 'application/xml');
    });

    progress('Indexando formatos de lista');
    // El classifier necesita saber si cada numId es bullet o numérico para
    // distinguir FHJVieta* vs FHJLista* en párrafos con <w:numPr>.
    const listAbstractIndex = await runStep('Indexando numbering.xml', () =>
      buildListAbstractIndex(outputZip)
    );

    progress('Aplicando estilos FHJ a párrafos');
    const restyleStats = await runStep('Aplicando estilos FHJ', () =>
      applyFhjStylesDom(docDoc, { listAbstractIndex })
    );
    if (restyleStats.lowConfidence && restyleStats.lowConfidence.length > 0) {
      warnings.push({
        code: 'CLASSIFIER_LOW_CONFIDENCE',
        message: 'Algunos párrafos se clasificaron con baja confianza. Revísalos por si el estilo aplicado no es el correcto.',
        context: { count: restyleStats.lowConfidence.length, samples: restyleStats.lowConfidence.slice(0, 8) }
      });
    }

    progress('Comprobando tabla de datos generales');
    await runStep('Inyectando tabla Datos generales', () => injectDatosGeneralesTableDom(docDoc));

    progress('Numerando tablas y figuras');
    const numberingStats = await runStep('Numerando tablas y figuras', () => addTableAndFigureTitlesDom(docDoc));

    if (relIds) {
      progress('Configurando página y secciones');
      await runStep('Actualizando sectPr', () => updateSectPrDom(docDoc, relIds));
    }

    await runStep('Renumerando bookmarks', () => renumberBookmarksDom(docDoc));
    await runStep('Deduplicando bookmarks', () => dedupeBookmarksDom(docDoc));

    // Fase 6: merge de numbering.xml (requiere docDoc ya parseado para remapear body).
    progress('Fusionando numeración / listas');
    const numMerge = await runStep('Fusionando numbering.xml', () =>
      mergeNumberingIntoDoc(refZip, outputZip, docDoc)
    );
    warnings.push(...numMerge.warnings);

    // Fase 5: validación estructural post-transformación (soft-fail).
    progress('Validando estructura del PNT');
    const structuralWarnings = await runStep('Validación estructural', () =>
      validateStructureDom(docDoc, { tables: numberingStats.tables, figures: numberingStats.figures })
    );
    for (const w of structuralWarnings) warnings.push(w);

    // Fase 8: auto-fix normativo (antes del validador) si está habilitado.
    if (doAutoFix) {
      progress('Corrigiendo desviaciones normativas');
      fixes = await runStep('Auto-fix normativo', () => applyNormativaFixesDom(docDoc));
    }

    // Fase 7 (Bloque D): validador normativo PNT.
    progress('Validando normativa de formato');
    const normativaWarnings = await runStep('Validación normativa', () =>
      validateNormativaDom(docDoc)
    );
    for (const w of normativaWarnings) warnings.push(w);

    // Fase 6: validar que los pStyle refs del body estén en styles.xml.
    progress('Validando referencias a estilos');
    const pStyleWarning = await runStep('Validando pStyle IDs', () =>
      validatePStyleRefs(outputZip, docDoc)
    );
    if (pStyleWarning) warnings.push(pStyleWarning);

    outputZip.file('word/document.xml', new XMLSerializer().serializeToString(docDoc));

    progress('Registrando tipos de contenido');
    await runStep('Actualizando [Content_Types].xml', async () => {
      let contentTypes = await outputZip.file('[Content_Types].xml').async('string');
      contentTypes = updateContentTypes(contentTypes, alreadyHasFhjHeader);
      outputZip.file('[Content_Types].xml', contentTypes);
    });

    progress('Generando archivo final');
    const blob = await runStep('Generando archivo final', () => outputZip.generateAsync({
      type: blobType,
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 }
    }));

    return {
      blob,
      stats: {
        preservedHeaders: alreadyHasFhjHeader,
        title1: restyleStats.title1,
        titPar: restyleStats.titPar,
        paragraph: restyleStats.paragraph,
        vignette: restyleStats.vignette,
        list: restyleStats.list || 0,
        preservedStyles: restyleStats.preserved || 0,
        tables: numberingStats.tables,
        figures: numberingStats.figures,
        fixes: fixes,
        autoFixApplied: doAutoFix
      },
      warnings
    };
  }

  function hasMetadata(m) {
    return m && (m.code || '').trim() && (m.version || '').trim() && (m.title || '').trim();
  }

  // Marcadores de texto que identifican un header como "del FHJ". Cualquier
  // documento con un header propio que contenga uno de estos se considera
  // cabecera FHJ y se respeta tal cual.
  const FHJ_HEADER_MARKERS = [
    /Hospital\s+de\s+Jove/i,
    /Fundaci[oó]n\s+Hospital\s+de\s+Jove/i,
    /FHJ\b/i,
    /\bP\.\d{2}\.\d{2}\.\d{3,}\b/   // código PNT → header claramente oficial
  ];

  // Detectar si el documento ya trae un header con formato FHJ.
  //
  // Fase 7: el check anterior era "tiene drawing + image rel" → machista,
  // dejaba pasar cualquier documento con cualquier imagen en el header
  // (tablas de membrete corporativas ajenas, logotipos de otros centros, etc.).
  // Ahora exigimos **imagen + texto que lo identifique como FHJ**.
  async function detectFhjHeader(zip) {
    const candidates = ['header1', 'header2', 'header3'];
    for (const name of candidates) {
      const headerFile = zip.file(`word/${name}.xml`);
      if (!headerFile) continue;
      const headerXml = await headerFile.async('string');
      const hasDrawing = headerXml.includes('<w:drawing>') || headerXml.includes('<w:pict>');
      if (!hasDrawing) continue;

      const relsFile = zip.file(`word/_rels/${name}.xml.rels`);
      if (!relsFile) continue;
      const relsText = await relsFile.async('string');
      const hasImageRel = /Type="[^"]*\/image"/.test(relsText) || /Target="[^"]*\/media\//.test(relsText);
      if (!hasImageRel) continue;

      // Texto plano del header (stripped de etiquetas) para matching de marcadores.
      const plainText = headerXml.replace(/<[^>]+>/g, ' ');
      const isFhj = FHJ_HEADER_MARKERS.some(re => re.test(plainText));
      if (isFhj) return true;
    }
    return false;
  }

  /**
   * Fase 6: en vez de machacar word/styles.xml con la copia del referente,
   * hacemos UNIÓN:
   *   - base = referente (garantiza que todos los estilos FHJ estén presentes)
   *   - se añaden los estilos del contenido cuyo styleId no existe en ref
   *     (así no perdemos estilos custom/numeración/tracked-changes del origen)
   *   - ref gana en caso de colisión por styleId
   *
   * theme1.xml y settings.xml se siguen sobreescribiendo desde el referente —
   * son presentación "global" que el referente controla por diseño.
   *
   * numbering.xml NO se toca aquí: se mergeará más tarde con acceso al DOM
   * de document.xml para poder remapear numIds si hay colisión.
   *
   * Devuelve { mergedStyles, styleWarnings } para que el caller pueda propagar
   * avisos. Nunca lanza: si el merge no es posible, cae a "copiar del ref"
   * y emite warning.
   */
  async function mergeStylesAndTransferCore(refZip, outputZip) {
    const warnings = [];

    // theme + settings: comportamiento previo — overwrite desde ref.
    for (const path of ['word/theme/theme1.xml', 'word/settings.xml']) {
      const file = refZip.file(path);
      if (file) outputZip.file(path, await file.async('uint8array'));
    }

    // styles.xml: merge
    const refStylesFile = refZip.file('word/styles.xml');
    if (!refStylesFile) {
      // preflight ya ha bloqueado este caso, pero por si acaso:
      return { warnings };
    }
    const refStylesXml = await refStylesFile.async('string');

    const contentStylesFile = outputZip.file('word/styles.xml');
    if (!contentStylesFile) {
      // No hay nada que mergear: copiamos tal cual ref.
      outputZip.file('word/styles.xml', refStylesXml);
      return { warnings };
    }

    const contentStylesXml = await contentStylesFile.async('string');

    try {
      const { mergedXml, addedIds, conflictIds } = mergeStylesXml(refStylesXml, contentStylesXml);
      outputZip.file('word/styles.xml', mergedXml);
      // Sólo avisamos cuando realmente preservamos estilos del content. Los
      // conflictos (ref gana) son el camino esperado para los estilos FHJ.
      if (addedIds.length > 0) {
        warnings.push({
          code: 'STYLES_MERGED',
          message: 'Se fusionaron los estilos del referente con los del contenido. ' +
                   addedIds.length + ' estilos del contenido se preservaron (no existían en el referente).',
          context: {
            addedFromContent: addedIds.slice(0, 20),
            conflictsWonByRef: conflictIds.slice(0, 20)
          }
        });
      }
    } catch (err) {
      // Merge falló: fallback al comportamiento viejo (copiar del ref).
      outputZip.file('word/styles.xml', refStylesXml);
      warnings.push({
        code: 'STYLES_MERGE_FAILED',
        message: 'No se pudo fusionar styles.xml; se ha copiado el del referente. ' +
                 'Estilos custom del contenido que no existan en el referente se pueden perder.',
        context: { error: err && err.message ? err.message : String(err) }
      });
    }

    return { warnings };
  }

  /**
   * Merge puro de dos strings XML de word/styles.xml.
   * - Toma ref como DOM base.
   * - Recorre `<w:style>` hijos de content. Si su @w:styleId NO está en ref,
   *   lo importa y lo apendea al `<w:styles>` del ref.
   * - Si ya está en ref, gana ref (no se toca).
   * - No se tocan `<w:docDefaults>` ni `<w:latentStyles>` del ref.
   */
  function mergeStylesXml(refStylesXml, contentStylesXml) {
    const refDoc = new DOMParser().parseFromString(refStylesXml, 'application/xml');
    const contentDoc = new DOMParser().parseFromString(contentStylesXml, 'application/xml');

    const refRoot = refDoc.getElementsByTagName('w:styles')[0] || refDoc.documentElement;
    if (!refRoot) throw new Error('ref styles.xml sin <w:styles>');

    const refStyleNodes = refDoc.getElementsByTagName('w:style');
    const refIds = new Set();
    for (let i = 0; i < refStyleNodes.length; i++) {
      const id = refStyleNodes[i].getAttributeNS(W_NS, 'styleId') || refStyleNodes[i].getAttribute('w:styleId');
      if (id) refIds.add(id);
    }

    const contentStyleNodes = contentDoc.getElementsByTagName('w:style');
    const addedIds = [];
    const conflictIds = [];
    for (let i = 0; i < contentStyleNodes.length; i++) {
      const node = contentStyleNodes[i];
      const id = node.getAttributeNS(W_NS, 'styleId') || node.getAttribute('w:styleId');
      if (!id) continue;
      if (refIds.has(id)) {
        conflictIds.push(id);
        continue;
      }
      const imported = refDoc.importNode(node, true);
      refRoot.appendChild(imported);
      addedIds.push(id);
      refIds.add(id);
    }

    return {
      mergedXml: new XMLSerializer().serializeToString(refDoc),
      addedIds,
      conflictIds
    };
  }

  /**
   * Fase 6: merge de word/numbering.xml + remap de numIds en document.xml.
   *
   * Flujo:
   *   - Si ref no tiene numbering.xml: si content sí lo tiene, se queda el del
   *     content. (Warning REFERENT_NO_NUMBERING ya emitido por preflight.)
   *   - Si content no tiene numbering.xml: se copia el del ref (comportamiento
   *     previo) — no hay nada que remapear en body.
   *   - Si ambos tienen: UNIÓN con ref como base.
   *     - abstractNumIds del content que colisionen con ref → renumerar al
   *       max(ref.abstract)+1, +2, ...
   *     - numIds del content que colisionen con ref → renumerar al
   *       max(ref.nums)+1, +2, ...
   *     - El `<w:abstractNumId>` dentro de cada `<w:num>` del content también
   *       se remapea si apuntaba a un abstract remapeado.
   *     - Todas las referencias `<w:numId w:val="X"/>` dentro del body de
   *       document.xml se remapean según el mapa oldNumId → newNumId.
   *
   * Nunca lanza: en error emite warning NUMBERING_MERGE_FAILED y cae al
   * comportamiento antiguo (copiar ref).
   */
  async function mergeNumberingIntoDoc(refZip, outputZip, docDoc) {
    const warnings = [];

    const refNumFile = refZip.file('word/numbering.xml');
    const contentNumFile = outputZip.file('word/numbering.xml');

    if (!refNumFile && !contentNumFile) return { warnings };

    if (!refNumFile && contentNumFile) {
      // Sólo el content lo tiene — ya está en outputZip. Nada que hacer.
      return { warnings };
    }

    const refNumXml = await refNumFile.async('string');

    if (!contentNumFile) {
      // Content no tenía numbering.xml — copiamos del ref tal cual.
      outputZip.file('word/numbering.xml', refNumXml);
      return { warnings };
    }

    const contentNumXml = await contentNumFile.async('string');

    try {
      const { mergedXml, numIdRemap, remappedAbstracts, remappedNums } =
        mergeNumberingXml(refNumXml, contentNumXml);
      outputZip.file('word/numbering.xml', mergedXml);

      const remapCount = Object.keys(numIdRemap).length;
      if (remapCount > 0) {
        applyNumIdRemapToBody(docDoc, numIdRemap);
        warnings.push({
          code: 'NUMBERING_MERGED_REMAPPED',
          message: 'Se fusionaron las definiciones de viñetas/numeración. ' +
                   remapCount + ' numIds del contenido se renumeraron para evitar colisión con los del referente.',
          context: {
            numIdsRemapped: remapCount,
            abstractIdsRemapped: remappedAbstracts,
            sampleMap: Object.entries(numIdRemap).slice(0, 10).map(([k, v]) => ({ from: Number(k), to: v }))
          }
        });
      }
      // Si no hubo remaps, el merge es silencioso: es comportamiento esperado.
    } catch (err) {
      // Fallback: comportamiento viejo — copiar ref y dejar el body como esté.
      outputZip.file('word/numbering.xml', refNumXml);
      warnings.push({
        code: 'NUMBERING_MERGE_FAILED',
        message: 'No se pudo fusionar numbering.xml; se ha copiado el del referente. ' +
                 'Las viñetas o listas numeradas del contenido pueden perderse.',
        context: { error: err && err.message ? err.message : String(err) }
      });
    }

    return { warnings };
  }

  /**
   * Merge puro de numbering.xml. Devuelve el DOM serializado + el mapa de
   * remapeo de numId para aplicar al body.
   */
  function mergeNumberingXml(refNumXml, contentNumXml) {
    const refDoc = new DOMParser().parseFromString(refNumXml, 'application/xml');
    const contentDoc = new DOMParser().parseFromString(contentNumXml, 'application/xml');

    const refRoot = refDoc.getElementsByTagName('w:numbering')[0] || refDoc.documentElement;
    if (!refRoot) throw new Error('ref numbering.xml sin <w:numbering>');

    // Recolectar IDs ya usados en ref.
    const refAbstractIds = new Set();
    const refAbsNodes = refDoc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < refAbsNodes.length; i++) {
      const v = refAbsNodes[i].getAttributeNS(W_NS, 'abstractNumId') ||
                refAbsNodes[i].getAttribute('w:abstractNumId');
      if (v != null && v !== '') refAbstractIds.add(Number(v));
    }
    const refNumIds = new Set();
    const refNumNodes = refDoc.getElementsByTagName('w:num');
    for (let i = 0; i < refNumNodes.length; i++) {
      const v = refNumNodes[i].getAttributeNS(W_NS, 'numId') ||
                refNumNodes[i].getAttribute('w:numId');
      if (v != null && v !== '') refNumIds.add(Number(v));
    }

    // Siguiente ID disponible.
    let nextAbstract = refAbstractIds.size ? Math.max(...refAbstractIds) + 1 : 0;
    let nextNum = refNumIds.size ? Math.max(...refNumIds) + 1 : 0;

    // Paso 1: importar abstracts del content, remapeando los que colisionen.
    const abstractRemap = {}; // oldAbstractId → newAbstractId
    let remappedAbstracts = 0;
    const contentAbsNodes = contentDoc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < contentAbsNodes.length; i++) {
      const node = contentAbsNodes[i];
      const oldV = node.getAttributeNS(W_NS, 'abstractNumId') ||
                   node.getAttribute('w:abstractNumId');
      if (oldV == null || oldV === '') continue;
      const oldId = Number(oldV);
      let newId = oldId;
      if (refAbstractIds.has(oldId)) {
        newId = nextAbstract++;
        abstractRemap[oldId] = newId;
        remappedAbstracts++;
      } else {
        refAbstractIds.add(oldId);
      }
      const imported = refDoc.importNode(node, true);
      imported.setAttributeNS(W_NS, 'w:abstractNumId', String(newId));
      // Insertar antes del primer <w:num> si existe, si no al final.
      const firstNum = refNumNodes[0];
      if (firstNum && firstNum.parentNode === refRoot) {
        refRoot.insertBefore(imported, firstNum);
      } else {
        refRoot.appendChild(imported);
      }
    }

    // Paso 2: importar <w:num>, remapeando numId si colisiona y su <w:abstractNumId>
    // si el abstract al que apunta fue remapeado.
    const numIdRemap = {}; // oldNumId → newNumId
    let remappedNums = 0;
    const contentNumNodes = contentDoc.getElementsByTagName('w:num');
    for (let i = 0; i < contentNumNodes.length; i++) {
      const node = contentNumNodes[i];
      const oldV = node.getAttributeNS(W_NS, 'numId') || node.getAttribute('w:numId');
      if (oldV == null || oldV === '') continue;
      const oldId = Number(oldV);
      let newId = oldId;
      if (refNumIds.has(oldId)) {
        newId = nextNum++;
        numIdRemap[oldId] = newId;
        remappedNums++;
      } else {
        refNumIds.add(oldId);
      }
      const imported = refDoc.importNode(node, true);
      imported.setAttributeNS(W_NS, 'w:numId', String(newId));
      // Remapear el <w:abstractNumId> interno si apuntaba a un abstract renombrado.
      const absRef = imported.getElementsByTagName('w:abstractNumId')[0];
      if (absRef) {
        const pointedOld = absRef.getAttributeNS(W_NS, 'val') || absRef.getAttribute('w:val');
        if (pointedOld != null && pointedOld !== '') {
          const oldAbs = Number(pointedOld);
          if (Object.prototype.hasOwnProperty.call(abstractRemap, oldAbs)) {
            absRef.setAttributeNS(W_NS, 'w:val', String(abstractRemap[oldAbs]));
          }
        }
      }
      refRoot.appendChild(imported);
    }

    return {
      mergedXml: new XMLSerializer().serializeToString(refDoc),
      numIdRemap,
      remappedAbstracts,
      remappedNums
    };
  }

  /**
   * Aplica el mapa { oldNumId → newNumId } a todas las `<w:numId w:val=X/>`
   * del documento, tanto dentro de `<w:numPr>` (body) como fuera (por si hay
   * referencias raras).
   */
  function applyNumIdRemapToBody(docDoc, numIdRemap) {
    if (!docDoc) return 0;
    const numIdEls = docDoc.getElementsByTagName('w:numId');
    let remapped = 0;
    for (let i = 0; i < numIdEls.length; i++) {
      const el = numIdEls[i];
      const v = el.getAttributeNS(W_NS, 'val') || el.getAttribute('w:val');
      if (v == null || v === '') continue;
      const oldId = Number(v);
      if (Object.prototype.hasOwnProperty.call(numIdRemap, oldId)) {
        el.setAttributeNS(W_NS, 'w:val', String(numIdRemap[oldId]));
        remapped++;
      }
    }
    return remapped;
  }

  /**
   * Fase 6: valida que todos los `<w:pStyle w:val="X"/>` del body apunten
   * a un styleId presente en styles.xml tras el merge. Emite un único warning
   * con la lista de IDs desconocidos (si hay).
   */
  async function validatePStyleRefs(outputZip, docDoc) {
    const stylesFile = outputZip.file('word/styles.xml');
    if (!stylesFile) return null;
    const stylesXml = await stylesFile.async('string');
    const stylesDoc = new DOMParser().parseFromString(stylesXml, 'application/xml');
    const knownIds = new Set();
    const styleNodes = stylesDoc.getElementsByTagName('w:style');
    for (let i = 0; i < styleNodes.length; i++) {
      const id = styleNodes[i].getAttributeNS(W_NS, 'styleId') ||
                 styleNodes[i].getAttribute('w:styleId');
      if (id) knownIds.add(id);
    }

    const pStyleEls = docDoc.getElementsByTagName('w:pStyle');
    const unknown = new Set();
    for (let i = 0; i < pStyleEls.length; i++) {
      const v = pStyleEls[i].getAttributeNS(W_NS, 'val') || pStyleEls[i].getAttribute('w:val');
      if (!v) continue;
      if (!knownIds.has(v)) unknown.add(v);
    }

    if (unknown.size === 0) return null;
    return {
      code: 'STYLES_UNKNOWN_ID',
      message: 'El documento referencia ' + unknown.size + ' estilos que no existen en styles.xml. ' +
               'Word mostrará esos párrafos con el estilo "Normal".',
      context: { unknownStyleIds: Array.from(unknown).slice(0, 20) }
    };
  }

  /**
   * Fase 7 (Bloque C): transfiere desde el ref todas las partes que componen
   * el membrete: header1-3, footer1-3, cada .rels asociado, y el media referenciado
   * por esos .rels (logo, sellos, firmas…). Para evitar colisión con media que
   * el contenido ya tenga en word/media/, renombramos los targets del ref a
   * "fhj_<nombreOriginal>" y reescribimos los .rels apuntando al nuevo nombre.
   *
   * Devuelve el listado de targets inyectados para debug.
   */
  async function transferHeadersFootersAndLogo(refZip, outputZip) {
    const parts = ['header1', 'header2', 'header3', 'footer1', 'footer2', 'footer3'];
    const injectedMedia = new Set();

    for (const name of parts) {
      const partFile = refZip.file('word/' + name + '.xml');
      if (!partFile) continue;

      // Copiar el xml del header/footer.
      outputZip.file('word/' + name + '.xml', await partFile.async('uint8array'));

      // Copiar + reescribir su .rels.
      const relsFile = refZip.file('word/_rels/' + name + '.xml.rels');
      if (!relsFile) continue;

      let relsText = await relsFile.async('string');

      // Encontrar todas las Relationships que apuntan a media y prefijarlas.
      // media/image1.emf → media/fhj_image1.emf, etc.
      const mediaRefRe = /Target="(media\/[^"]+)"/g;
      const substitutions = [];
      let m;
      while ((m = mediaRefRe.exec(relsText)) !== null) {
        const original = m[1];          // "media/image1.emf"
        const basename = original.replace(/^media\//, '');
        const fhjTarget = 'media/fhj_' + basename;
        substitutions.push({ original, fhjTarget, basename });
      }
      for (const s of substitutions) {
        const esc = s.original.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
        relsText = relsText.replace(new RegExp('Target="' + esc + '"', 'g'), 'Target="' + s.fhjTarget + '"');

        // Copiar el archivo media si no está ya inyectado.
        if (!injectedMedia.has(s.fhjTarget)) {
          const mediaFile = refZip.file('word/' + s.original);
          if (mediaFile) {
            outputZip.file('word/' + s.fhjTarget, await mediaFile.async('uint8array'));
            injectedMedia.add(s.fhjTarget);
          }
        }
      }

      outputZip.file('word/_rels/' + name + '.xml.rels', relsText);
    }

    return { injectedMedia: Array.from(injectedMedia) };
  }

  async function customizeHeader(outputZip, metadata) {
    const headerFile = outputZip.file('word/header2.xml');
    if (!headerFile) {
      return {
        warning: {
          code: 'CUSTOMIZE_HEADER_SKIPPED_NO_HEADER2',
          message: 'No se pudo personalizar la cabecera: el referente no aportó word/header2.xml. Los metadatos no se han aplicado.',
          context: { metadata }
        }
      };
    }
    const headerXml = await headerFile.async('string');
    const doc = new DOMParser().parseFromString(headerXml, 'application/xml');
    const paragraphs = doc.getElementsByTagName('w:p');
    if (paragraphs.length < 2) {
      return {
        warning: {
          code: 'CUSTOMIZE_HEADER_INSUFFICIENT_PARAGRAPHS',
          message: 'No se pudo personalizar la cabecera: header2.xml tiene menos de 2 párrafos (encontrados: ' + paragraphs.length + '). Los metadatos no se han aplicado.',
          context: { metadata, foundParagraphs: paragraphs.length }
        }
      };
    }
    clearParagraphText(paragraphs[0]);
    addTextToParagraph(doc, paragraphs[0], metadata.code + ' / ' + metadata.version, false);
    clearParagraphText(paragraphs[1]);
    addTextToParagraph(doc, paragraphs[1], metadata.title, true);
    outputZip.file('word/header2.xml', new XMLSerializer().serializeToString(doc));
    return null;
  }

  function clearParagraphText(paragraph) {
    Array.from(paragraph.getElementsByTagName('w:r')).forEach(r => r.parentNode.removeChild(r));
    Array.from(paragraph.getElementsByTagName('w:proofErr')).forEach(p => p.parentNode.removeChild(p));
  }

  function addTextToParagraph(doc, paragraph, text, bold) {
    const r = doc.createElementNS(W_NS, 'w:r');
    const rPr = doc.createElementNS(W_NS, 'w:rPr');
    const rFonts = doc.createElementNS(W_NS, 'w:rFonts');
    rFonts.setAttribute('w:cs', 'Arial');
    rPr.appendChild(rFonts);
    if (bold) {
      rPr.appendChild(doc.createElementNS(W_NS, 'w:b'));
      rPr.appendChild(doc.createElementNS(W_NS, 'w:bCs'));
    }
    r.appendChild(rPr);
    const t = doc.createElementNS(W_NS, 'w:t');
    t.setAttribute('xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    paragraph.appendChild(r);
  }

  async function updateDocumentRels(outputZip) {
    const relsXml = await outputZip.file('word/_rels/document.xml.rels').async('string');
    const doc = new DOMParser().parseFromString(relsXml, 'application/xml');
    const rels = Array.from(doc.getElementsByTagName('Relationship'));
    let maxId = 0;
    const toRemove = [];
    for (const r of rels) {
      const id = r.getAttribute('Id');
      const n = parseInt(id.replace(/\D/g, ''), 10);
      if (n > maxId) maxId = n;
      const type = r.getAttribute('Type');
      if (type.endsWith('/header') || type.endsWith('/footer')) toRemove.push(r);
    }
    toRemove.forEach(r => r.parentNode.removeChild(r));
    const baseType = R_NS + '/';
    const newRels = [
      { type: 'header', target: 'header1.xml' },
      { type: 'header', target: 'header2.xml' },
      { type: 'header', target: 'header3.xml' },
      { type: 'footer', target: 'footer1.xml' },
      { type: 'footer', target: 'footer2.xml' },
      { type: 'footer', target: 'footer3.xml' }
    ];
    const ids = {};
    let nextId = maxId + 1;
    const root = doc.documentElement;
    for (const rel of newRels) {
      const id = 'rId' + (nextId++);
      const el = doc.createElement('Relationship');
      el.setAttribute('Id', id);
      el.setAttribute('Type', baseType + rel.type);
      el.setAttribute('Target', rel.target);
      root.appendChild(el);
      ids[rel.target.replace('.xml', '')] = id;
    }
    outputZip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(doc));
    return ids;
  }

  // ============================================================
  // Fase 2: funciones DOM-based sobre document.xml
  // ============================================================

  /**
   * Fase 7 (Bloque A): clasificador con respeto a estilos pre-existentes.
   *
   * Orden de decisión por párrafo:
   *  1. Si ya tiene un pStyle FHJ* válido → respetarlo (trust path).
   *  2. Si tiene <w:numPr> con numId → es lista. Mapear a FHJVieta(1|2|3) o
   *     FHJLista(1|2|3) según el numFmt del nivel (bullet vs numérico) y el ilvl.
   *  3. Si su pStyle es "Normal" → tratarlo como párrafo no clasificado y
   *     pasarlo por el classifier (acabará como FHJPrrafo si nada más coincide).
   *  4. Si su pStyle es "ListParagraph" sin numPr → bullet por defecto.
   *  5. Si arranca con bullet ASCII (-, •, ▪…) → FHJVietaNivel1.
   *  6. classifyParagraphDetailed() — fall-through normativo.
   *
   * El listAbstractIndex permite, para un numId del body, recuperar el lvl 0/1/2
   * y su numFmt para distinguir bullet vs numérico.
   */
  function applyFhjStylesDom(doc, opts) {
    opts = opts || {};
    const listAbstractIndex = opts.listAbstractIndex || null; // { numId → {ilvl0,1,2: 'bullet'|'decimal'|'lowerLetter'|'lowerRoman'|'upperLetter'} }

    let title1 = 0, titPar = 0, paragraph = 0, vignette = 0, list = 0, preserved = 0;
    const lowConfidence = [];
    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));
    for (const p of paragraphs) {
      const rawText = getParagraphTextDom(p);
      const text = normalizeParagraphText(rawText);

      // Inspeccionar el pStyle pre-existente y la presencia de numPr.
      const existingStyle = getExistingPStyle(p);
      const numPrInfo = readNumPr(p);

      // 1) Trust path: el contenido ya viene marcado con FHJ* — lo respetamos.
      if (existingStyle && /^FHJ/.test(existingStyle)) {
        preserved++;
        if (existingStyle === 'FHJTtulo1') title1++;
        else if (existingStyle === 'FHJTtuloprrafo') titPar++;
        else if (existingStyle === 'FHJPrrafo') paragraph++;
        else if (/^FHJVieta/.test(existingStyle) || /^FHJLista/.test(existingStyle)) {
          vignette++;
          list++;
        }
        continue;
      }

      // Sin texto → puede ser un párrafo "ancla" para una imagen o tabla.
      // Si no tiene FHJ ni numPr, lo dejamos como está para no romper layout.
      if (!text && !numPrInfo) continue;

      let style = null, confidence = 'high', reason = '';

      // 2) Lista por numPr → nunca puede ser título.
      if (numPrInfo) {
        const fmt = listAbstractIndex && listAbstractIndex[numPrInfo.numId]
          ? (listAbstractIndex[numPrInfo.numId][numPrInfo.ilvl] || null)
          : null;
        const isBullet = !fmt || fmt === 'bullet' || fmt === 'none';
        const lvl = numPrInfo.ilvl + 1; // 0-based → 1-based
        const family = isBullet ? 'FHJVietaNivel' : 'FHJListaNivel';
        const clamped = Math.max(1, Math.min(3, lvl));
        style = family + clamped;
        reason = 'numPr-' + (isBullet ? 'bullet' : 'numbered') + '-lvl' + lvl;
      } else if (existingStyle === 'ListParagraph') {
        // 4) Sin numPr pero estilo "ListParagraph" — Word lo usa para bullets sueltos.
        style = 'FHJVietaNivel1';
        reason = 'list-paragraph';
      } else if (text && isListItem(text)) {
        // 5) Bullet ASCII al inicio del texto.
        style = 'FHJVietaNivel1';
        reason = 'list-item';
      } else if (text) {
        // 3 + 6) Normal o sin estilo → classifier normativo.
        const result = classifyParagraphDetailed(text);
        style = result.style;
        confidence = result.confidence;
        reason = result.reason;
      }

      if (!style) continue;

      if (confidence === 'medium') {
        lowConfidence.push({ style, reason, text: (text || '').slice(0, 80) });
      }
      if (style === 'FHJTtulo1') title1++;
      else if (style === 'FHJTtuloprrafo') titPar++;
      else if (style === 'FHJPrrafo') paragraph++;
      else if (/^FHJVieta/.test(style)) { vignette++; list++; }
      else if (/^FHJLista/.test(style)) { list++; }

      setParagraphStyleDom(doc, p, style);
    }
    return { title1, titPar, paragraph, vignette, list, preserved, lowConfidence };
  }

  /** Devuelve el styleId del <w:pStyle> dentro del <w:pPr> de un párrafo, o null. */
  function getExistingPStyle(p) {
    const pPr = firstChildByName(p, 'w:pPr');
    if (!pPr) return null;
    const pStyle = firstChildByName(pPr, 'w:pStyle');
    if (!pStyle) return null;
    return pStyle.getAttributeNS(W_NS, 'val') || pStyle.getAttribute('w:val') || null;
  }

  /** Lee <w:numPr> de un párrafo y devuelve { numId, ilvl } o null. */
  function readNumPr(p) {
    const pPr = firstChildByName(p, 'w:pPr');
    if (!pPr) return null;
    const numPr = firstChildByName(pPr, 'w:numPr');
    if (!numPr) return null;
    const numIdEl = firstChildByName(numPr, 'w:numId');
    const ilvlEl = firstChildByName(numPr, 'w:ilvl');
    const numId = numIdEl ? Number(numIdEl.getAttributeNS(W_NS, 'val') || numIdEl.getAttribute('w:val') || 0) : 0;
    const ilvl = ilvlEl ? Number(ilvlEl.getAttributeNS(W_NS, 'val') || ilvlEl.getAttribute('w:val') || 0) : 0;
    if (!numId || Number.isNaN(numId)) return null;
    return { numId, ilvl: Math.max(0, ilvl) };
  }

  /**
   * Construye un índice { numId → { 0: 'bullet'|'decimal'|..., 1: ..., 2: ... } }
   * leyendo word/numbering.xml del outputZip. Tolera ausencia (devuelve {}).
   */
  async function buildListAbstractIndex(outputZip) {
    const idx = {};
    const numFile = outputZip.file('word/numbering.xml');
    if (!numFile) return idx;
    const xml = await numFile.async('string');
    let doc;
    try { doc = new DOMParser().parseFromString(xml, 'application/xml'); } catch (e) { return idx; }

    // Primero: abstractNumId → { ilvl → numFmt }.
    const abstractFmts = {};
    const abstractNodes = doc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < abstractNodes.length; i++) {
      const aNode = abstractNodes[i];
      const aId = aNode.getAttributeNS(W_NS, 'abstractNumId') || aNode.getAttribute('w:abstractNumId');
      if (aId == null) continue;
      const lvls = aNode.getElementsByTagName('w:lvl');
      const map = {};
      for (let j = 0; j < lvls.length; j++) {
        const lvl = lvls[j];
        const ilvl = Number(lvl.getAttributeNS(W_NS, 'ilvl') || lvl.getAttribute('w:ilvl') || 0);
        const fmtEl = firstChildByName(lvl, 'w:numFmt');
        const fmt = fmtEl ? (fmtEl.getAttributeNS(W_NS, 'val') || fmtEl.getAttribute('w:val') || 'bullet') : 'bullet';
        map[ilvl] = fmt;
      }
      abstractFmts[Number(aId)] = map;
    }

    // Después: numId → abstractNumId.
    const numNodes = doc.getElementsByTagName('w:num');
    for (let i = 0; i < numNodes.length; i++) {
      const nNode = numNodes[i];
      const nId = nNode.getAttributeNS(W_NS, 'numId') || nNode.getAttribute('w:numId');
      if (nId == null) continue;
      const absRef = firstChildByName(nNode, 'w:abstractNumId');
      if (!absRef) continue;
      const aId = absRef.getAttributeNS(W_NS, 'val') || absRef.getAttribute('w:val');
      if (aId == null) continue;
      const fmts = abstractFmts[Number(aId)];
      if (fmts) idx[Number(nId)] = fmts;
    }
    return idx;
  }

  /**
   * Normaliza el texto de un párrafo antes de clasificarlo:
   *   - NFKC (combina acentos sueltos, normaliza formas compuestas)
   *   - NBSP (\u00A0) → espacio
   *   - dashes unicode (‒ – — ― ‐ ‑) → guion ASCII
   *   - tabs → espacio
   *   - colapsa runs de whitespace
   *   - trim
   * Fase 6: resuelve la fragilidad del classifier ante espacios dobles,
   * guiones tipográficos y caracteres raros pegados desde Word.
   */
  function normalizeParagraphText(text) {
    if (!text) return '';
    let t = String(text);
    if (typeof t.normalize === 'function') t = t.normalize('NFKC');
    t = t.replace(/\u00A0/g, ' ');
    t = t.replace(/[\u2010\u2011\u2012\u2013\u2014\u2015]/g, '-');
    t = t.replace(/\t/g, ' ');
    t = t.replace(/\s+/g, ' ');
    return t.trim();
  }

  function getParagraphTextDom(p) {
    const ts = p.getElementsByTagName('w:t');
    let out = '';
    for (let i = 0; i < ts.length; i++) {
      out += ts[i].textContent || '';
    }
    return out;
  }

  function setParagraphStyleDom(doc, p, styleName) {
    let pPr = firstChildByName(p, 'w:pPr');
    if (!pPr) {
      pPr = doc.createElementNS(W_NS, 'w:pPr');
      p.insertBefore(pPr, p.firstChild);
    }
    let pStyle = firstChildByName(pPr, 'w:pStyle');
    if (pStyle) {
      pStyle.setAttribute('w:val', styleName);
    } else {
      pStyle = doc.createElementNS(W_NS, 'w:pStyle');
      pStyle.setAttribute('w:val', styleName);
      pPr.insertBefore(pStyle, pPr.firstChild);
    }
  }

  // Palabras ancla que — solas o casi — marcan sección de primer nivel.
  // Se usan en classifyParagraphDetailed para coger títulos sin numeración.
  const FHJ_SECTION_ANCHORS = [
    'OBJETO', 'OBJETIVO',
    'ALCANCE', 'APLICACIÓN', 'APLICACION',
    'DEFINICIONES',
    'RESPONSABILIDADES', 'RESPONSABLES',
    'DESARROLLO', 'PROCEDIMIENTO', 'SISTEMÁTICA', 'SISTEMATICA',
    'METODOLOGÍA', 'METODOLOGIA',
    'REFERENCIAS', 'BIBLIOGRAFÍA', 'BIBLIOGRAFIA',
    'ANEXOS',
    'INTRODUCCIÓN', 'INTRODUCCION',
    'DOCUMENTOS RELACIONADOS', 'DIAGRAMA DE FLUJO',
    'MATERIAL', 'MATERIALES',
    'REGISTROS', 'DISTRIBUCIÓN', 'DISTRIBUCION'
  ];

  /**
   * Clasificador rico: devuelve { style, confidence, reason }.
   * - style: 'FHJTtulo1' | 'FHJTtuloprrafo' | 'FHJPrrafo' | null
   * - confidence: 'high' | 'medium'
   * - reason: identificador del criterio que disparó la clasificación
   *
   * Recibe texto YA normalizado. La envoltura classifyParagraph() conserva
   * compat con tests antiguos devolviendo sólo el style string.
   */
  function classifyParagraphDetailed(text) {
    if (!text) return { style: null, confidence: 'high', reason: 'empty' };

    const upper = text.toUpperCase();

    // ANEXO I | ANEXO 1 | ANEXO II.A | ANEXO A:
    if (/^ANEXO\s+([IVX]+|\d+|[A-Z])(\b|[\.\-:])/i.test(text)) {
      return { style: 'FHJTtulo1', confidence: 'high', reason: 'annex' };
    }

    // FHJTtuloprrafo — numeración multi-nivel: "1.1", "1.1.", "1.1.-", "1.1)", "1.1.1.- ", etc.
    // Admite cualquier combinación de ". - ) espacio" entre el número y el texto.
    if (/^\d+(\.\d+)+[\.\-\)\s]+\S/.test(text)) {
      return { style: 'FHJTtuloprrafo', confidence: 'high', reason: 'numbered-level-2plus' };
    }

    // FHJTtulo1 — numeración de primer nivel ("1.-", "1.", "1)", "1 ") seguido de
    // mayúscula Y con resto del texto que parece título (corto, ALL-CAPS o
    // capitalizado, sin verbos imperativos típicos de pasos de procedimiento).
    //
    // Fase 7: el classifier viejo metía como Título 1 cualquier "1. Lavarse..."
    // que es un paso, no un título. Endurecemos las condiciones:
    //   - texto total ≤ 90 caracteres
    //   - todo el texto tras el prefijo numérico es ALL-CAPS o tiene ratio alto
    //     de mayúsculas (≥ 0.6 sobre las letras alfabéticas)
    const leadM = text.match(/^(\d+)[\.\-\)\s]+([A-ZÁÉÍÓÚÑ].*)$/);
    if (leadM) {
      const rest = leadM[2];
      if (text.length <= 90) {
        const letters = rest.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g, '');
        const upperLetters = rest.replace(/[^A-ZÁÉÍÓÚÑ]/g, '').length;
        const ratio = letters.length ? upperLetters / letters.length : 0;
        if (ratio >= 0.6) {
          return { style: 'FHJTtulo1', confidence: 'high', reason: 'numbered-level-1' };
        }
      }
      // Si no cumple, es un paso de procedimiento — lo trataremos como párrafo.
    }

    // FHJTtulo1 por ancla — palabra clave sola, o casi (≤ 50 chars).
    if (text.length <= 50) {
      for (const anchor of FHJ_SECTION_ANCHORS) {
        if (upper === anchor || upper === anchor + '.' || upper === anchor + ':' ||
            upper === anchor + ';' || upper.startsWith(anchor + ' ')) {
          return { style: 'FHJTtulo1', confidence: 'high', reason: 'anchor-' + anchor.toLowerCase() };
        }
      }
    }

    // PROCEDIMIENTO ANALÍTICO… (título habitual en PNTs de laboratorio)
    if (upper.includes('PROCEDIMIENTO ANALÍTICO') && text.length < 100) {
      return { style: 'FHJTtulo1', confidence: 'medium', reason: 'analytical-procedure' };
    }

    // ALL-CAPS corto y denso en letras → probable título (confianza media).
    if (text.length > 0 && text.length <= 120) {
      const letters = text.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g, '');
      if (letters.length >= 4) {
        const upperLetters = text.replace(/[^A-ZÁÉÍÓÚÑ]/g, '').length;
        const ratio = upperLetters / letters.length;
        if (ratio >= 0.8) {
          return { style: 'FHJTtulo1', confidence: 'medium', reason: 'all-caps-short' };
        }
      }
    }

    return { style: 'FHJPrrafo', confidence: 'high', reason: 'default' };
  }

  // Retro-compat: callers externos (tests antiguos, otros módulos) esperan string.
  function classifyParagraph(text) {
    return classifyParagraphDetailed(text).style;
  }

  // Cubre más caracteres de viñeta + numeración de lista entre paréntesis suave.
  // No hace match con "1) Título" (eso sigue siendo título, no lista).
  function isListItem(text) {
    if (!text) return false;
    return /^[\-\u2022\u00B7\*\u25E6\u25AA\u25AB\u25A0\u25A1\u25C6]\s+/.test(text);
  }

  // Fase 6: etiquetas canónicas que esperamos en la primera columna de la
  // tabla de Datos generales FHJ. Si una tabla tiene ≥ 3 de estas etiquetas
  // en su primera columna, la reconocemos como "Datos generales" aunque no
  // lleve ese título explícito.
  const FHJ_DATOS_LABELS = [
    /c[óo]digo/i,
    /versi[óo]n/i,
    /elaborado\s+por/i,
    /revisado\s+por/i,
    /aprobado\s+por/i,
    /fecha\s+de\s+entrada/i,
    /responsables?/i,
    /t[íi]tulo/i
  ];

  function isDatosGeneralesTable(tbl) {
    if (!tbl || !tbl.getElementsByTagName) return false;
    const rows = tbl.getElementsByTagName('w:tr');
    if (!rows || rows.length < 3) return false;
    let matched = 0;
    const maxScan = Math.min(rows.length, 12);
    for (let i = 0; i < maxScan; i++) {
      const cells = rows[i].getElementsByTagName('w:tc');
      if (!cells || cells.length === 0) continue;
      // Recogemos el texto de la primera celda.
      let cellText = '';
      const ts = cells[0].getElementsByTagName('w:t');
      for (let j = 0; j < ts.length; j++) cellText += (ts[j].textContent || '');
      cellText = cellText.trim();
      if (!cellText) continue;
      for (const re of FHJ_DATOS_LABELS) {
        if (re.test(cellText)) { matched++; break; }
      }
      if (matched >= 3) return true;
    }
    return false;
  }

  function injectDatosGeneralesTableDom(doc) {
    const body = doc.getElementsByTagName('w:body')[0];
    if (!body) return;

    // Heurística A (rápida): texto "Datos generales" en los primeros ~10 elementos.
    let scannedText = '';
    let count = 0;
    for (let n = body.firstChild; n && count < 10; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      count++;
      const ts = n.getElementsByTagName ? n.getElementsByTagName('w:t') : null;
      if (ts) {
        for (let i = 0; i < ts.length; i++) {
          scannedText += (ts[i].textContent || '') + ' ';
        }
      }
    }
    if (/datos\s+generales/i.test(scannedText)) return;

    // Heurística B (Fase 6): detectar tablas con forma de "Datos generales"
    // por contenido de celdas (Código/Versión/Elaborado por/…) aunque no
    // lleven el título explícito. Evita duplicar la tabla en documentos
    // donde el autor eliminó la cabecera pero conservó el cuerpo tabular.
    count = 0;
    for (let n = body.firstChild; n && count < 10; n = n.nextSibling) {
      if (n.nodeType !== 1) continue;
      count++;
      if (n.nodeName === 'w:tbl' && isDatosGeneralesTable(n)) return;
    }

    const frag = parseFragment(doc, buildDatosGeneralesTable());
    const anchor = body.firstChild;
    for (const node of frag) {
      body.insertBefore(node, anchor);
    }
  }

  /**
   * Parsea un fragmento de XML con prefijo w: y lo importa a `doc`.
   * Devuelve el array de nodos importados (en el orden original).
   */
  function parseFragment(doc, xmlString) {
    const wrapped =
      '<w:wrap xmlns:w="' + W_NS + '" xmlns:r="' + R_NS + '">' +
      xmlString +
      '</w:wrap>';
    const fragDoc = new DOMParser().parseFromString(wrapped, 'application/xml');
    const children = [];
    for (let n = fragDoc.documentElement.firstChild; n; n = n.nextSibling) {
      children.push(n);
    }
    return children.map(n => doc.importNode(n, true));
  }

  function buildDatosGeneralesTable() {
    const rows = [
      ['Código:', ''], ['Versión:', ''], ['Elaborado por:', ''], ['Revisado por:', ''],
      ['Aprobado por:', ''], ['Fecha de entrada en vigor:', ''], ['Responsables:', '']
    ];
    let xml = `<w:p><w:pPr><w:pStyle w:val="FHJTtulo1"/></w:pPr><w:r><w:t xml:space="preserve">Datos generales</w:t></w:r></w:p>`;
    xml += '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>';
    xml += '<w:tblBorders>';
    xml += '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>';
    xml += '</w:tblBorders><w:tblLook w:val="04A0"/></w:tblPr>';
    xml += '<w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="6500"/></w:tblGrid>';
    for (const [label, value] of rows) {
      xml += '<w:tr>';
      xml += '<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr>';
      xml += `<w:p><w:pPr><w:pStyle w:val="FHJTitulodatosgenerales"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(label)}</w:t></w:r></w:p>`;
      xml += '</w:tc>';
      xml += '<w:tc><w:tcPr><w:tcW w:w="6500" w:type="dxa"/></w:tcPr>';
      xml += `<w:p><w:pPr><w:pStyle w:val="FHJContenidodatosgenerales"/></w:pPr><w:r><w:t xml:space="preserve">${escapeXml(value)}</w:t></w:r></w:p>`;
      xml += '</w:tc></w:tr>';
    }
    xml += '</w:tbl><w:p><w:pPr><w:pStyle w:val="FHJPrrafo"/></w:pPr></w:p>';
    return xml;
  }

  function escapeXml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function addTableAndFigureTitlesDom(doc) {
    // Tablas: numeramos todas las w:tbl DEL BODY excepto la primera (que es Datos generales).
    // Ignoramos tablas anidadas (dentro de otras tablas).
    const allTables = Array.from(doc.getElementsByTagName('w:tbl'));
    const bodyTables = allTables.filter(t => !isInsideName(t, 'w:tbl', /* skipSelf */ true));
    let tableCount = 0;
    for (let i = 1; i < bodyTables.length; i++) {
      tableCount++;
      const titleP = makeHighlightedTitleParagraphDom(
        doc,
        `Tabla ${tableCount}.`,
        '[Nombre pendiente de revisión]'
      );
      bodyTables[i].parentNode.insertBefore(titleP, bodyTables[i]);
    }

    // Figuras: cada w:drawing fuera de una w:tbl cuenta como figura.
    // El título se inserta INMEDIATAMENTE DESPUÉS del párrafo contenedor.
    // Varios drawings en el mismo párrafo producen varios títulos consecutivos.
    const drawings = Array.from(doc.getElementsByTagName('w:drawing'));
    let figureCount = 0;
    const paraToNumbers = [];

    for (const d of drawings) {
      if (isInsideName(d, 'w:tbl')) continue;
      const containingP = findAncestorByName(d, 'w:p');
      if (!containingP) continue;
      figureCount++;
      let entry = paraToNumbers.find(e => e.p === containingP);
      if (!entry) {
        entry = { p: containingP, nums: [] };
        paraToNumbers.push(entry);
      }
      entry.nums.push(figureCount);
    }

    for (const { p, nums } of paraToNumbers) {
      const anchor = p.nextSibling;
      for (const n of nums) {
        const titleP = makeHighlightedTitleParagraphDom(
          doc,
          `Figura ${n}.`,
          '[Nombre pendiente de revisión]'
        );
        // insertBefore con anchor=null equivale a appendChild, con anchor lo pone antes de él.
        p.parentNode.insertBefore(titleP, anchor);
      }
    }

    return { tables: tableCount, figures: figureCount };
  }

  function makeHighlightedTitleParagraphDom(doc, labelBold, nameText) {
    const xml =
      '<w:p><w:pPr><w:pStyle w:val="FHJPrrafo"/><w:spacing w:before="240" w:after="80"/><w:ind w:firstLine="0"/><w:jc w:val="left"/></w:pPr>' +
      `<w:r><w:rPr><w:b/><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">${escapeXml(labelBold)}</w:t></w:r>` +
      `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve"> ${escapeXml(nameText)}</w:t></w:r></w:p>`;
    const imported = parseFragment(doc, xml);
    return imported[0];
  }

  function updateSectPrDom(doc, ids) {
    const sectPrs = doc.getElementsByTagName('w:sectPr');
    if (sectPrs.length === 0) return;
    const sectPr = sectPrs[0];
    while (sectPr.firstChild) sectPr.removeChild(sectPr.firstChild);

    const ref = (tag, type, rid) => {
      const el = doc.createElementNS(W_NS, tag);
      el.setAttribute('w:type', type);
      el.setAttribute('r:id', rid);
      sectPr.appendChild(el);
    };
    ref('w:headerReference', 'even', ids.header1);
    ref('w:headerReference', 'default', ids.header2);
    ref('w:footerReference', 'even', ids.footer1);
    ref('w:footerReference', 'default', ids.footer2);
    ref('w:headerReference', 'first', ids.header3);
    ref('w:footerReference', 'first', ids.footer3);

    const pgSz = doc.createElementNS(W_NS, 'w:pgSz');
    pgSz.setAttribute('w:w', '11906');
    pgSz.setAttribute('w:h', '16838');
    pgSz.setAttribute('w:code', '9');
    sectPr.appendChild(pgSz);

    const pgMar = doc.createElementNS(W_NS, 'w:pgMar');
    pgMar.setAttribute('w:top', '1418');
    pgMar.setAttribute('w:right', '1418');
    pgMar.setAttribute('w:bottom', '1418');
    pgMar.setAttribute('w:left', '1701');
    pgMar.setAttribute('w:header', '709');
    pgMar.setAttribute('w:footer', '709');
    pgMar.setAttribute('w:gutter', '0');
    sectPr.appendChild(pgMar);

    const cols = doc.createElementNS(W_NS, 'w:cols');
    cols.setAttribute('w:space', '708');
    sectPr.appendChild(cols);

    const docGrid = doc.createElementNS(W_NS, 'w:docGrid');
    docGrid.setAttribute('w:linePitch', '360');
    sectPr.appendChild(docGrid);
  }

  function renumberBookmarksDom(doc) {
    let counter = 1000;
    const idMap = {};
    const starts = Array.from(doc.getElementsByTagName('w:bookmarkStart'));
    for (const s of starts) {
      const oldId = s.getAttribute('w:id');
      if (oldId === null || oldId === '') continue;
      const newId = String(counter++);
      idMap[oldId] = newId;
      s.setAttribute('w:id', newId);
    }
    const ends = Array.from(doc.getElementsByTagName('w:bookmarkEnd'));
    for (const e of ends) {
      const oldId = e.getAttribute('w:id');
      if (oldId !== null && idMap[oldId]) {
        e.setAttribute('w:id', idMap[oldId]);
      }
    }
  }

  function dedupeBookmarksDom(doc) {
    const seen = new Set();
    const ends = Array.from(doc.getElementsByTagName('w:bookmarkEnd'));
    for (const e of ends) {
      const id = e.getAttribute('w:id');
      if (seen.has(id)) {
        if (e.parentNode) e.parentNode.removeChild(e);
      } else {
        seen.add(id);
      }
    }
  }

  // ---- helpers DOM ----

  function firstChildByName(parent, qname) {
    for (let n = parent.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && nodeNameMatches(n, qname)) return n;
    }
    return null;
  }

  function findAncestorByName(node, qname) {
    for (let n = node.parentNode; n; n = n.parentNode) {
      if (n.nodeType === 1 && nodeNameMatches(n, qname)) return n;
    }
    return null;
  }

  function isInsideName(node, qname, skipSelf) {
    let n = skipSelf ? node.parentNode : node;
    for (; n; n = n.parentNode) {
      if (n.nodeType === 1 && nodeNameMatches(n, qname)) return true;
    }
    return false;
  }

  function nodeNameMatches(el, qname) {
    return el.nodeName === qname || el.tagName === qname;
  }

  function updateContentTypes(contentTypes, preservedHeaders) {
    const imageDefaults = [
      ['emf', 'image/x-emf'], ['wmf', 'image/x-wmf'], ['png', 'image/png'],
      ['jpeg', 'image/jpeg'], ['jpg', 'image/jpeg'], ['gif', 'image/gif']
    ];
    for (const [ext, mime] of imageDefaults) {
      if (!contentTypes.includes(`Extension="${ext}"`)) {
        contentTypes = contentTypes.replace(
          /<Default\s+Extension="xml"/,
          `<Default Extension="${ext}" ContentType="${mime}"/>\n  <Default Extension="xml"`
        );
      }
    }
    if (!preservedHeaders) {
      const needed = [
        ['header1.xml', 'header'], ['header2.xml', 'header'], ['header3.xml', 'header'],
        ['footer1.xml', 'footer'], ['footer2.xml', 'footer'], ['footer3.xml', 'footer']
      ];
      for (const [fname, type] of needed) {
        const partName = '/word/' + fname;
        if (!contentTypes.includes(`PartName="${partName}"`)) {
          const entry = `  <Override PartName="${partName}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${type}+xml"/>\n`;
          contentTypes = contentTypes.replace('</Types>', entry + '</Types>');
        }
      }
    }
    return contentTypes;
  }

  // ============================================================
  // Fase 8: auto-fix normativo — repara en DOM las infracciones
  // que `validateNormativaDom` detecta. Opt-in via `autoFix: true`.
  // ============================================================

  /**
   * Aplica correcciones deterministas sobre el DOM:
   *   - underline: elimina <w:u> de runs del body.
   *   - font: reescribe <w:rFonts> no-Arial a Arial (ascii/hAnsi/cs).
   *   - allCaps: descapitaliza párrafos de cuerpo totalmente en mayúsculas
   *     (primera letra en mayúscula, resto en minúsculas).
   *   - emptyList: elimina párrafos de lista (FHJVieta o FHJLista, cualquier
   *     nivel) sin texto y sin drawings/pict.
   *
   * No toca títulos (FHJTtulo*) ni headers/footers (opera sólo sobre w:body).
   * Devuelve contadores por tipo de corrección.
   */
  function applyNormativaFixesDom(doc) {
    const fixes = {
      underline: 0, font: 0, allCaps: 0, emptyList: 0,
      blankParas: 0, multiSpace: 0, renumbered: 0,
      samples: {
        underline: [], font: [], allCaps: [], emptyList: [],
        blankParas: [], multiSpace: [], renumbered: []
      }
    };
    const MAX_SAMPLES = 3;
    const SNIPPET_LEN = 80;
    const body = doc.getElementsByTagName('w:body')[0];
    if (!body) return fixes;

    const isTitleStyle = (styleId) => /^FHJTtulo/.test(styleId || '');
    const bodyRuns = Array.from(body.getElementsByTagName('w:r'));

    function findParentP(node) {
      let n = node && node.parentNode;
      while (n && n.nodeType === 1) {
        if (n.localName === 'p' || n.nodeName === 'w:p') return n;
        n = n.parentNode;
      }
      return null;
    }

    function pushSample(bucket, text, extra) {
      if (!bucket || bucket.length >= MAX_SAMPLES) return;
      const trimmed = String(text || '').replace(/\s+/g, ' ').trim();
      if (!trimmed) return;
      const snippet = trimmed.length > SNIPPET_LEN
        ? trimmed.slice(0, SNIPPET_LEN - 1) + '…'
        : trimmed;
      const entry = extra ? Object.assign({ text: snippet }, extra) : { text: snippet };
      bucket.push(entry);
    }

    // 1) Underline: borrar <w:u> de los rPr del body.
    for (const r of bodyRuns) {
      const rPr = firstChildByName(r, 'w:rPr');
      if (!rPr) continue;
      const u = firstChildByName(rPr, 'w:u');
      if (!u) continue;
      const val = u.getAttributeNS(W_NS, 'val') || u.getAttribute('w:val') || 'single';
      if (val === 'none') continue;
      const parentP = findParentP(r);
      const paraText = parentP ? normalizeParagraphText(getParagraphTextDom(parentP)) : '';
      rPr.removeChild(u);
      fixes.underline++;
      pushSample(fixes.samples.underline, paraText);
    }

    // 2) Fuentes no-Arial: reescribir ascii/hAnsi/cs a Arial.
    for (const r of bodyRuns) {
      const rPr = firstChildByName(r, 'w:rPr');
      if (!rPr) continue;
      const rFonts = firstChildByName(rPr, 'w:rFonts');
      if (!rFonts) continue;
      let changed = false;
      let originalFont = null;
      const attrs = ['ascii', 'hAnsi', 'cs'];
      for (const attr of attrs) {
        const cur = rFonts.getAttributeNS(W_NS, attr) || rFonts.getAttribute('w:' + attr);
        if (cur && !/^Arial\b/i.test(cur)) {
          if (!originalFont) originalFont = cur;
          rFonts.setAttributeNS(W_NS, 'w:' + attr, 'Arial');
          changed = true;
        }
      }
      if (changed) {
        fixes.font++;
        const parentP = findParentP(r);
        const paraText = parentP ? normalizeParagraphText(getParagraphTextDom(parentP)) : '';
        pushSample(fixes.samples.font, paraText, { font: originalFont });
      }
    }

    // 3) ALL-CAPS body: descapitalizar párrafos no-título en mayúsculas.
    const paragraphs = Array.from(body.getElementsByTagName('w:p'));
    for (const p of paragraphs) {
      const styleId = getExistingPStyle(p);
      if (isTitleStyle(styleId)) continue;
      const text = normalizeParagraphText(getParagraphTextDom(p));
      if (!text || text.length < 40) continue;
      const letters = text.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g, '');
      if (letters.length < 20) continue;
      const upperLetters = text.replace(/[^A-ZÁÉÍÓÚÑ]/g, '').length;
      const ratio = upperLetters / letters.length;
      if (ratio < 0.9) continue;

      pushSample(fixes.samples.allCaps, text);

      const ts = p.getElementsByTagName('w:t');
      let firstDone = false;
      for (let i = 0; i < ts.length; i++) {
        const original = ts[i].textContent || '';
        if (!original) continue;
        let lowered = original.toLowerCase();
        if (!firstDone) {
          // Capitaliza la primera letra encontrada (saltando whitespace inicial).
          const m = lowered.match(/^(\s*)([A-Za-zÁÉÍÓÚÑáéíóúñ])/);
          if (m) {
            lowered = m[1] + m[2].toUpperCase() + lowered.slice(m[0].length);
            firstDone = true;
          }
        }
        ts[i].textContent = lowered;
      }
      fixes.allCaps++;
    }

    // 4) Listas vacías: eliminar párrafos FHJVieta*/FHJLista* sin texto ni dibujos.
    for (const p of paragraphs) {
      const styleId = getExistingPStyle(p);
      if (!styleId) continue;
      if (!/^FHJ(Vieta|Lista)/.test(styleId)) continue;
      const text = normalizeParagraphText(getParagraphTextDom(p));
      if (text) continue;
      if (p.getElementsByTagName('w:drawing').length > 0) continue;
      if (p.getElementsByTagName('w:pict').length > 0) continue;
      if (p.parentNode) {
        p.parentNode.removeChild(p);
        fixes.emptyList++;
        pushSample(fixes.samples.emptyList, '(ítem vacío con estilo ' + styleId + ')');
      }
    }

    // 5) Fase 10B: párrafos en blanco consecutivos → colapsar a 1.
    //    Un párrafo "en blanco" no tiene texto, no tiene drawing/pict, y no es
    //    una lista (ya tratada arriba). Dejamos siempre uno de cada cluster.
    const freshParagraphs = Array.from(body.getElementsByTagName('w:p'));
    let blankRun = 0;
    for (const p of freshParagraphs) {
      const text = normalizeParagraphText(getParagraphTextDom(p));
      const hasDrawing = p.getElementsByTagName('w:drawing').length > 0 ||
                         p.getElementsByTagName('w:pict').length > 0;
      const hasTable = p.getElementsByTagName('w:tbl').length > 0;
      const isBlank = !text && !hasDrawing && !hasTable;
      if (isBlank) {
        blankRun++;
        if (blankRun > 1 && p.parentNode) {
          p.parentNode.removeChild(p);
          fixes.blankParas = (fixes.blankParas || 0) + 1;
          pushSample(fixes.samples.blankParas, '(párrafo en blanco duplicado)');
        }
      } else {
        blankRun = 0;
      }
    }

    // 6) Fase 10B: espacios múltiples dentro de texto → colapsar a 1.
    //    No tocamos NBSP (\u00A0) ni whitespace de preservación explícita
    //    (xml:space="preserve" con espacio único al inicio/fin sigue siendo
    //    válido; solo colapsamos secuencias de 2+ espacios ASCII dentro).
    const allTextNodes = body.getElementsByTagName('w:t');
    for (let i = 0; i < allTextNodes.length; i++) {
      const node = allTextNodes[i];
      const original = node.textContent || '';
      if (!/  +/.test(original)) continue;
      const collapsed = original.replace(/ {2,}/g, ' ');
      if (collapsed !== original) {
        node.textContent = collapsed;
        fixes.multiSpace = (fixes.multiSpace || 0) + 1;
        pushSample(fixes.samples.multiSpace, original.trim());
      }
    }

    // 7) Fase 10B: renumerar FHJTtulo1 con huecos.
    //    Solo tocamos títulos cuyo texto empieza por "N" + separador, donde N
    //    es un entero. No tocamos ANEXO (que usa romanos/letras) ni títulos
    //    sin prefijo numérico.
    const currentParagraphs = Array.from(body.getElementsByTagName('w:p'));
    const titleItems = [];
    for (const p of currentParagraphs) {
      const styleId = getExistingPStyle(p);
      if (styleId !== 'FHJTtulo1') continue;
      const text = normalizeParagraphText(getParagraphTextDom(p));
      if (!text) continue;
      if (/^ANEXO\b/i.test(text)) continue;
      const m = text.match(/^(\d+)([\.\-\)]+\s*|\s+)(.*)$/);
      if (!m) continue;
      titleItems.push({ p, current: parseInt(m[1], 10), sep: m[2], rest: m[3], text });
    }
    if (titleItems.length >= 2) {
      const hasGap = titleItems.some((t, idx) => t.current !== idx + 1);
      if (hasGap) {
        for (let idx = 0; idx < titleItems.length; idx++) {
          const item = titleItems[idx];
          const desired = idx + 1;
          if (item.current === desired) continue;
          const ts = item.p.getElementsByTagName('w:t');
          let rewrote = false;
          for (let j = 0; j < ts.length; j++) {
            const t = ts[j].textContent || '';
            const im = t.match(/^(\s*)(\d+)(.*)$/);
            if (im) {
              ts[j].textContent = im[1] + desired + im[3];
              rewrote = true;
              break;
            }
          }
          if (rewrote) {
            fixes.renumbered = (fixes.renumbered || 0) + 1;
            pushSample(fixes.samples.renumbered,
              `${item.current} → ${desired}: ${item.text.slice(0, 60)}`);
          }
        }
      }
    }

    return fixes;
  }

  // ============================================================
  // Fase 7 (Bloque D): validador normativo — detecta desviaciones
  // respecto a la normativa PNT del Hospital de Jove.
  // ============================================================

  /**
   * Ejecuta un conjunto de checks sobre el documento ya reformateado:
   *   - NORMATIVA_ALL_CAPS_BODY: párrafos de cuerpo en ALL-CAPS largos (se
   *     penaliza sólo cuerpo — títulos pueden ir en mayúsculas).
   *   - NORMATIVA_UNDERLINE: runs con <w:u/> en el body (la normativa pide
   *     prescindir del subrayado salvo en enlaces).
   *   - NORMATIVA_EMPTY_LIST: bullets/listas numeradas sin texto.
   *   - NORMATIVA_MISSING_DATOS_GENERALES: no se detecta la tabla FHJ.
   *   - NORMATIVA_FONT_NON_ARIAL: runs con <w:rFonts> distinto de Arial.
   *
   * No muta el documento. Cada warning incluye un `samples` corto para que
   * el usuario pueda localizar rápidamente.
   */
  function validateNormativaDom(doc) {
    const warnings = [];
    const body = doc.getElementsByTagName('w:body')[0];
    if (!body) return warnings;

    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));

    // Helpers: detectar estilo FHJ títulos para saltarlos en ALL-CAPS check.
    const isTitleStyle = (styleId) => /^FHJTtulo/.test(styleId || '');

    // 1) ALL-CAPS en cuerpo — párrafos no-título, ≥ 40 chars, ratio mayúsculas ≥ 0.9
    const allCapsSamples = [];
    for (const p of paragraphs) {
      const styleId = getExistingPStyle(p);
      if (isTitleStyle(styleId)) continue;
      const text = normalizeParagraphText(getParagraphTextDom(p));
      if (!text || text.length < 40) continue;
      const letters = text.replace(/[^A-Za-zÁÉÍÓÚÑáéíóúñ]/g, '');
      if (letters.length < 20) continue;
      const upperLetters = text.replace(/[^A-ZÁÉÍÓÚÑ]/g, '').length;
      const ratio = upperLetters / letters.length;
      if (ratio >= 0.9) {
        allCapsSamples.push(text.slice(0, 120));
      }
    }
    if (allCapsSamples.length) {
      warnings.push({
        code: 'NORMATIVA_ALL_CAPS_BODY',
        message: 'Se detectaron ' + allCapsSamples.length + ' párrafos de cuerpo escritos enteramente en mayúsculas. La normativa desaconseja el uso de ALL-CAPS en el cuerpo del PNT por ser difícil de leer.',
        context: { count: allCapsSamples.length, samples: allCapsSamples.slice(0, 5) }
      });
    }

    // 2) Subrayado en runs del body (no incluye header/footer).
    const bodyRuns = Array.from(body.getElementsByTagName('w:r'));
    let underlineCount = 0;
    const underlineSamples = [];
    for (const r of bodyRuns) {
      const rPr = firstChildByName(r, 'w:rPr');
      if (!rPr) continue;
      const u = firstChildByName(rPr, 'w:u');
      if (!u) continue;
      const val = u.getAttributeNS(W_NS, 'val') || u.getAttribute('w:val') || 'single';
      if (val === 'none') continue;
      underlineCount++;
      if (underlineSamples.length < 5) {
        const ts = r.getElementsByTagName('w:t');
        let t = '';
        for (let i = 0; i < ts.length; i++) t += ts[i].textContent || '';
        if (t.trim()) underlineSamples.push(t.trim().slice(0, 80));
      }
    }
    if (underlineCount > 0) {
      warnings.push({
        code: 'NORMATIVA_UNDERLINE',
        message: 'Se detectaron ' + underlineCount + ' fragmentos subrayados en el cuerpo. La normativa PNT recomienda reservar el subrayado para hiperenlaces.',
        context: { count: underlineCount, samples: underlineSamples }
      });
    }

    // 3) Listas vacías: párrafos con pStyle FHJVieta(1|2|3) o FHJLista(1|2|3) sin texto.
    let emptyListCount = 0;
    for (const p of paragraphs) {
      const styleId = getExistingPStyle(p);
      if (!styleId) continue;
      if (!/^FHJ(Vieta|Lista)/.test(styleId)) continue;
      const text = normalizeParagraphText(getParagraphTextDom(p));
      if (!text) emptyListCount++;
    }
    if (emptyListCount > 0) {
      warnings.push({
        code: 'NORMATIVA_EMPTY_LIST',
        message: 'Se detectaron ' + emptyListCount + ' ítems de lista sin contenido. Revísalos o elimínalos.',
        context: { count: emptyListCount }
      });
    }

    // 4) Datos generales: ¿hay al menos una tabla que reconozcamos?
    let hasDatosGenerales = false;
    const tables = doc.getElementsByTagName('w:tbl');
    for (let i = 0; i < tables.length; i++) {
      if (isDatosGeneralesTable(tables[i])) { hasDatosGenerales = true; break; }
    }
    if (!hasDatosGenerales) {
      warnings.push({
        code: 'NORMATIVA_MISSING_DATOS_GENERALES',
        message: 'No se detectó la tabla de "Datos generales" (Código, Versión, Elaborado/Revisado/Aprobado, Fecha de entrada). La normativa exige esta tabla al inicio del PNT.',
        context: {}
      });
    }

    // 5) Fuentes no-Arial en runs del body.
    const fontCounts = {};
    let totalFontRuns = 0;
    for (const r of bodyRuns) {
      const rPr = firstChildByName(r, 'w:rPr');
      if (!rPr) continue;
      const rFonts = firstChildByName(rPr, 'w:rFonts');
      if (!rFonts) continue;
      const candidates = [
        rFonts.getAttributeNS(W_NS, 'ascii') || rFonts.getAttribute('w:ascii'),
        rFonts.getAttributeNS(W_NS, 'hAnsi') || rFonts.getAttribute('w:hAnsi'),
        rFonts.getAttributeNS(W_NS, 'cs') || rFonts.getAttribute('w:cs')
      ].filter(Boolean);
      for (const f of candidates) {
        totalFontRuns++;
        fontCounts[f] = (fontCounts[f] || 0) + 1;
      }
    }
    const nonArial = Object.entries(fontCounts).filter(([name]) => !/^Arial\b/i.test(name));
    if (nonArial.length > 0) {
      const totalNonArial = nonArial.reduce((s, [, n]) => s + n, 0);
      warnings.push({
        code: 'NORMATIVA_FONT_NON_ARIAL',
        message: 'Se detectaron fuentes distintas de Arial en el cuerpo (' + totalNonArial + ' runs). La normativa exige Arial 10 para todo el cuerpo.',
        context: { fonts: Object.fromEntries(nonArial.slice(0, 8)) }
      });
    }

    // 6) Fase 10B: placeholders sin rellenar.
    //    Patrones típicos del referente que nadie sustituyó por el valor real:
    //      [CODIGO], [CÓDIGO], [TITULO], [TÍTULO], [VERSION], [VERSIÓN],
    //      [XX.XX.XX], XXXXX, XXXXXX, [xxxx], <<...>>
    const placeholderPatterns = [
      /\[C[OÓ]DIGO\]/i,
      /\[T[IÍ]TULO\]/i,
      /\[VERSI[OÓ]N\]/i,
      /\[FECHA(?:\s+DE\s+ENTRADA)?\]/i,
      /\bX{5,}\b/,
      /\[X{2,}(?:\.X{2,}){1,}\]/i,
      /<<[^<>\n]{2,50}>>/,
      /\[[a-z_]{2,20}\]/i  // e.g., [nombre], [cargo] — genérico
    ];
    const foundPlaceholders = new Set();
    const allBodyTexts = [];
    const allBodyParas = Array.from(body.getElementsByTagName('w:p'));
    for (const p of allBodyParas) {
      const text = (getParagraphTextDom(p) || '').trim();
      if (text) allBodyTexts.push(text);
    }
    const joinedBody = allBodyTexts.join(' \n ');
    for (const re of placeholderPatterns) {
      let m;
      const globalRe = new RegExp(re.source, re.flags.replace('g', '') + 'g');
      while ((m = globalRe.exec(joinedBody)) !== null) {
        foundPlaceholders.add(m[0]);
        if (foundPlaceholders.size >= 20) break;
      }
      if (foundPlaceholders.size >= 20) break;
    }
    if (foundPlaceholders.size > 0) {
      const samples = Array.from(foundPlaceholders).slice(0, 8);
      warnings.push({
        code: 'NORMATIVA_PLACEHOLDER_UNFILLED',
        message: 'Se encontraron placeholders sin rellenar en el cuerpo (ej.: ' +
                 samples.join(', ') + '). Sustituye los marcadores por el valor real.',
        context: { samples }
      });
    }

    return warnings;
  }

  // ============================================================
  // Fase 5: validación estructural — emite warnings no fatales
  // ============================================================

  /**
   * Inspecciona el documento procesado y devuelve un array de warnings sobre
   * estructura y coherencia. No muta el documento.
   *
   * Emite (si procede):
   *   - STRUCTURE_MISSING_OBJETO: no hay FHJTtulo1 que contenga "OBJETO"
   *   - STRUCTURE_MISSING_ALCANCE: no hay FHJTtulo1 que contenga "ALCANCE"
   *   - STRUCTURE_MISSING_DESARROLLO: no hay FHJTtulo1 que contenga "DESARROLLO" o "PROCEDIMIENTO"
   *   - STRUCTURE_NUMBERING_GAP: la secuencia numérica de FHJTtulo1 tiene huecos (1,2,4)
   *   - STRUCTURE_CROSSREF_BROKEN: el texto menciona "Tabla N" o "Figura N" inexistentes
   */
  function validateStructureDom(doc, counts) {
    const warnings = [];
    const body = doc.getElementsByTagName('w:body')[0];
    if (!body) return warnings;

    // Recolectar títulos de primer nivel (FHJTtulo1) con su texto.
    const paragraphs = Array.from(doc.getElementsByTagName('w:p'));
    const title1Texts = [];
    let allBodyText = '';
    for (const p of paragraphs) {
      const text = getParagraphTextDom(p).trim();
      if (!text) continue;
      allBodyText += ' ' + text;
      const pPr = firstChildByName(p, 'w:pPr');
      if (!pPr) continue;
      const pStyle = firstChildByName(pPr, 'w:pStyle');
      if (!pStyle) continue;
      const val = pStyle.getAttribute('w:val');
      if (val === 'FHJTtulo1') title1Texts.push(text);
    }

    const upper = title1Texts.map(t => t.toUpperCase());

    const hasKeyword = (kw) => upper.some(t => t.includes(kw));

    if (!hasKeyword('OBJETO')) {
      warnings.push({
        code: 'STRUCTURE_MISSING_OBJETO',
        message: 'No se detectó un título de primer nivel que contenga "OBJETO". Revisa la sección de objeto del procedimiento.',
        context: { foundTitles: title1Texts }
      });
    }
    if (!hasKeyword('ALCANCE')) {
      warnings.push({
        code: 'STRUCTURE_MISSING_ALCANCE',
        message: 'No se detectó un título de primer nivel que contenga "ALCANCE".',
        context: { foundTitles: title1Texts }
      });
    }
    if (!hasKeyword('DESARROLLO') && !hasKeyword('PROCEDIMIENTO') && !hasKeyword('SISTEMÁTICA') && !hasKeyword('SISTEMATICA')) {
      warnings.push({
        code: 'STRUCTURE_MISSING_DESARROLLO',
        message: 'No se detectó una sección de desarrollo/procedimiento/sistemática entre los títulos de primer nivel.',
        context: { foundTitles: title1Texts }
      });
    }

    // Secuencia numérica de los FHJTtulo1 que empiezan por "N.-" o "N."
    // Ignora ANEXO… y el resto.
    const seq = [];
    for (const t of title1Texts) {
      const m = t.match(/^(\d+)[\.\-]/);
      if (m) seq.push(parseInt(m[1], 10));
    }
    if (seq.length >= 2) {
      const gaps = [];
      for (let i = 1; i < seq.length; i++) {
        const expected = seq[i - 1] + 1;
        if (seq[i] !== expected) {
          gaps.push({ after: seq[i - 1], got: seq[i], expected });
        }
      }
      if (gaps.length > 0) {
        warnings.push({
          code: 'STRUCTURE_NUMBERING_GAP',
          message: 'La numeración de los títulos de primer nivel tiene saltos. Revisa la secuencia.',
          context: { sequence: seq, gaps }
        });
      }
    }

    // Cross-refs: el texto menciona "Tabla N" o "Figura N" con N mayor al que existe.
    const existingTables = counts.tables || 0;
    const existingFigures = counts.figures || 0;
    const tableMentions = [];
    const figureMentions = [];
    const reTable = /\bTabla\s+(\d+)\b/gi;
    const reFigure = /\bFigura\s+(\d+)\b/gi;
    let m;
    while ((m = reTable.exec(allBodyText))) tableMentions.push(parseInt(m[1], 10));
    while ((m = reFigure.exec(allBodyText))) figureMentions.push(parseInt(m[1], 10));
    const brokenTables = tableMentions.filter(n => n > existingTables);
    const brokenFigures = figureMentions.filter(n => n > existingFigures);
    if (brokenTables.length || brokenFigures.length) {
      warnings.push({
        code: 'STRUCTURE_CROSSREF_BROKEN',
        message: 'El texto menciona tablas o figuras con numeración superior a las existentes.',
        context: {
          existingTables,
          existingFigures,
          brokenTableRefs: Array.from(new Set(brokenTables)),
          brokenFigureRefs: Array.from(new Set(brokenFigures))
        }
      });
    }

    return warnings;
  }

  // ============================================================
  // Fase 5: inspectContent — extrae metadatos sin procesar
  // ============================================================

  /**
   * Inspecciona un docx de contenido y devuelve los metadatos detectados
   * (código PNT, versión, título) + flag hasFhjHeader. No modifica el docx.
   * Tolerante a entradas inválidas: si no se puede leer, devuelve nulls.
   */
  async function inspectContent(input) {
    const empty = {
      hasFhjHeader: false,
      detected: { code: null, version: null, title: null }
    };
    if (input == null) return empty;
    let zip;
    try {
      zip = await JSZip.loadAsync(input);
    } catch (e) {
      return empty;
    }
    // Si no hay document.xml es que no es un docx — devolvemos empty.
    if (!zip.file('word/document.xml')) return empty;

    const hasFhjHeader = await detectFhjHeader(zip);

    // Recolectar paragraph-text por cada header del contenido.
    const headerParagraphs = [];
    for (const name of ['header1', 'header2', 'header3']) {
      const f = zip.file('word/' + name + '.xml');
      if (!f) continue;
      const xml = await f.async('string');
      const paras = extractParagraphTexts(xml);
      for (const t of paras) headerParagraphs.push(t);
    }

    // Recolectar paragraph-text de los primeros ~12 paragraphs del body.
    const docXml = await zip.file('word/document.xml').async('string');
    const bodyParagraphs = extractParagraphTexts(docXml).slice(0, 12);

    const allParagraphs = headerParagraphs.concat(bodyParagraphs);
    const joinedText = allParagraphs.join(' \n ');

    // Código PNT: "P.dd.dd.ddd…"
    const codeMatch = joinedText.match(/\bP\.\d{2}\.\d{2}\.\d{3,}\b/);
    const code = codeMatch ? codeMatch[0] : null;

    // Versión: varios formatos frecuentes
    //   V.d, V.d.d, V d, Vd                  → "V.X"
    //   Versión d, Versión d.d               → tal cual
    //   Edición d, Edición d.d               → tal cual
    //   Rev. d, Rev d, Revisión d            → tal cual
    const versionPatterns = [
      /\bV\.\s?\d+(?:\.\d+)?\b/i,
      /\bV\s\d+(?:\.\d+)?\b/i,
      /\bVersi[oó]n\s+\d+(?:\.\d+)?\b/i,
      /\bEdici[oó]n\s+\d+(?:\.\d+)?\b/i,
      /\bRev(?:\.|isi[oó]n)?\s+\d+(?:\.\d+)?\b/i
    ];
    let version = null;
    for (const re of versionPatterns) {
      const m = joinedText.match(re);
      if (m) { version = m[0].trim(); break; }
    }

    // Título: primer párrafo razonable que no sea código/versión, que no empiece
    // por numeración ni por ANEXO, y con longitud 6..200.
    let title = null;
    const source = headerParagraphs.length ? headerParagraphs : bodyParagraphs;
    for (const raw of source) {
      const t = raw.trim();
      if (!t) continue;
      if (code && t.includes(code)) continue;
      if (version && t.includes(version) && t.length < 40) continue;
      if (/^\d+\s*[\.\-]/.test(t)) continue;
      if (/^ANEXO\b/i.test(t)) continue;
      if (t.length < 6 || t.length > 200) continue;
      title = t;
      break;
    }

    return {
      hasFhjHeader,
      detected: { code, version, title }
    };
  }

  function extractParagraphTexts(xml) {
    // Parser-based para resistir orden de atributos, namespaces raros, etc.
    try {
      const doc = new DOMParser().parseFromString(xml, 'application/xml');
      const ps = Array.from(doc.getElementsByTagName('w:p'));
      return ps.map(p => {
        const ts = p.getElementsByTagName('w:t');
        let text = '';
        for (let i = 0; i < ts.length; i++) text += ts[i].textContent || '';
        return text;
      });
    } catch (e) {
      return [];
    }
  }

  // Alias descriptivo: `extractMetadata(file)` devuelve solo `detected`
  // porque el contexto de uso es "rellenar el formulario".
  async function extractMetadata(input) {
    const res = await inspectContent(input);
    return res && res.detected ? res.detected : { code: null, version: null, title: null };
  }

  const api = { process, inspectContent, extractMetadata, IsoformaError };
  if (isNode) {
    module.exports = api;
  } else {
    var gOut = (typeof window !== 'undefined') ? window : globalThis;
    gOut.IsoformaEngine = api;
  }
})();
