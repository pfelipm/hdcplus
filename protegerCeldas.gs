// Envoltorios para protegerFxHoja() - HOJA ACTUAL
function protegerFxHojaAdvertencia() { protegerFxHoja(PROTECCION.modo.advertencia) }
function protegerFxHojaSoloYo() { protegerFxHoja(PROTECCION.modo.soloYo) }

// Envoltorios para protegerFxHoja() - TODA LA HOJA DE CÁLCULO
function protegerFxLibroAdvertencia() { protegerFxMasivo(PROTECCION.modo.advertencia) }
function protegerFxLibroSoloYo() { protegerFxMasivo(PROTECCION.modo.soloYo) }

// Envoltorios para eliminarProteccionesFx() - HOJA ACTUAL
function eliminarProteccionesFxHdCPlus() { eliminarProteccionesFx(true) }
function eliminarProteccionesFxTodas() { eliminarProteccionesFx(false) }

// Envoltorios para eliminarProteccionesFx() - TODA LA HOJA DE CÁLCULO
function eliminarProteccionesFxLibroHdCPlus() { eliminarProteccionesMasivo(true) }
function eliminarProteccionesFxLibroTodas() { eliminarProteccionesMasivo(false) }

/**
 * Ejecuta la protección en varias hojas (toda la hoja de cálculo)
 */
function protegerFxMasivo(modoProteccion) {
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = hdc.getSheets();
  const ui = SpreadsheetApp.getUi();
  
  hdc.toast('Buscando fórmulas en toda la hoja de cálculo...', '👀 HdC+', -1);
  
  let totalFormulas = 0;
  hojas.forEach(hoja => {
    const formulas = hoja.getDataRange().getFormulas();
    formulas.forEach(fila => fila.forEach(celda => { if (celda) totalFormulas++; }));
  });

  if (totalFormulas === 0) {
    hdc.toast('No se han encontrado fórmulas en ninguna pestaña.', '⚠️ HdC+', 10);
    return;
  }

  if (ui.alert(ENCABEZADO_ALERTAS, `🛡️ ¿Deseas proteger las fórmulas en las ${hojas.length} pestañas de la hoja de cálculo?
  
Se han detectado aproximadamente ${totalFormulas} celdas con fórmulas que serán agrupadas y protegidas.`, ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
    
    let intervalosTotales = 0;
    hojas.forEach((hoja, index) => {
      hdc.toast(`Procesando pestaña ${index + 1} de ${hojas.length} [${hoja.getName()}]...`, '🔒 HdC+', -1);
      intervalosTotales += protegerFxHojaInternal_(modoProteccion, hoja);
    });
    
    hdc.toast(`✅ Se han protegido ${intervalosTotales} intervalos en ${hojas.length} pestañas.`, '🟢 HdC+', 10);
  } else {
    hdc.toast('Operación cancelada.', '⚪ HdC+', 5);
  }
}

/**
 * Elimina protecciones en varias hojas (toda la hoja de cálculo)
 */
function eliminarProteccionesMasivo(soloHdCPlus) {
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = hdc.getSheets();
  const ui = SpreadsheetApp.getUi();
  
  hdc.toast('Buscando reglas de protección en toda la hoja de cálculo...', '👀 HdC+', -1);
  
  let totalProtecciones = 0;
  hojas.forEach(hoja => {
    let p = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    if (soloHdCPlus) p = p.filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion);
    totalProtecciones += p.length;
  });

  if (totalProtecciones === 0) {
    hdc.toast(`No se han encontrado protecciones ${soloHdCPlus ? 'de HdC+ ' : ''}en ninguna pestaña.`, '⚠️ HdC+', 10);
    return;
  }

  if (ui.alert(ENCABEZADO_ALERTAS, `🛡️ ¿Deseas eliminar las protecciones en las ${hojas.length} pestañas de la hoja de cálculo?
  
Se van a eliminar un total de ${totalProtecciones} reglas de protección.`, ui.ButtonSet.OK_CANCEL) === ui.Button.OK) {
    
    let desprotegidosTotales = 0;
    hojas.forEach((hoja, index) => {
      hdc.toast(`Desprotegiendo pestaña ${index + 1} de ${hojas.length} [${hoja.getName()}]...`, '🔓 HdC+', -1);
      desprotegidosTotales += eliminarProteccionesFxInternal_(soloHdCPlus, hoja);
    });
    
    hdc.toast(`✅ Se han eliminado ${desprotegidosTotales} reglas de protección en ${hojas.length} pestañas.`, '🟢 HdC+', 10);
  } else {
    hdc.toast('Operación cancelada.', '⚪ HdC+', 5);
  }
}

/**
 * Versión interna de protección que devuelve el número de intervalos protegidos (sin UI)
 */
function protegerFxHojaInternal_(modoProteccion, hoja) {
  const listaCeldasFx = [];
  const intervaloFx = hoja.getDataRange().getFormulas();
  for (let nFil = 0; nFil < intervaloFx.length; nFil++) {
    for (let nCol = 0; nCol < intervaloFx[nFil].length; nCol++) {
      if (intervaloFx[nFil][nCol]) {
        listaCeldasFx.push({ r1: nFil + 1, c1: nCol + 1, r2: nFil + 1, c2: nCol + 1 });
      }
    }
  }
  if (listaCeldasFx.length == 0) return 0;

  listaCeldasFx.sort((a, b) => a.c1 - b.c1 || a.r1 - b.r1);
  const bloquesV = [];
  let actV = { ...listaCeldasFx[0] };
  for (let i = 1; i < listaCeldasFx.length; i++) {
    let sig = listaCeldasFx[i];
    if (sig.c1 === actV.c1 && sig.r1 === actV.r2 + 1) actV.r2 = sig.r2;
    else { bloquesV.push(actV); actV = { ...sig }; }
  }
  bloquesV.push(actV);

  bloquesV.sort((a, b) => a.r1 - b.r1 || a.r2 - b.r2 || a.c1 - b.c1);
  const final = [];
  let actH = { ...bloquesV[0] };
  for (let i = 1; i < bloquesV.length; i++) {
    let sig = bloquesV[i];
    if (sig.r1 === actH.r1 && sig.r2 === actH.r2 && sig.c1 === actH.c2 + 1) actH.c2 = sig.c2;
    else { final.push(actH); actH = { ...sig }; }
  }
  final.push(actH);

  hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion)
    .forEach(regla => regla.remove());

  const opciones = { year: 'numeric', month: 'numeric', day: 'numeric', timeZone: Session.getScriptTimeZone(), hour: 'numeric', minute: 'numeric', hour12: false };
  final.forEach(int => {
    const p = hoja.getRange(int.r1, int.c1, int.r2 - int.r1 + 1, int.c2 - int.c1 + 1).protect();
    const d = `${PROTECCION.descripcion} ${new Date().toLocaleString(Session.getActiveUserLocale(), opciones)}`;
    if (modoProteccion == PROTECCION.modo.advertencia) p.setWarningOnly(true).setDescription(d);
    else p.addEditor(Session.getActiveUser()).setDescription(d);
  });
  return final.length;
}

/**
 * Versión interna de desprotección que devuelve el número de reglas eliminadas (sin UI)
 */
function eliminarProteccionesFxInternal_(soloHdCPlus, hoja) {
  let p = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (soloHdCPlus) p = p.filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion);
  const n = p.length;
  p.forEach(regla => regla.remove());
  return n;
}

/**
 * Protege las celdas que contienen fórmulas, tratando de agruparlas
 * en un número mínimo de intervalos (reglas de protección)
 * 
 * MODIFICADO 10/05/26:
 *  - Optimización crítica de rendimiento: se sustituye el algoritmo de comparación exhaustiva O(N²)
 *    por un algoritmo de agrupación de dos pasadas (Vertical -> Horizontal) con complejidad O(N log N).
 *  - Soporte para hoja específica y modo silencioso para ejecuciones masivas.
 * 
 * @param {string}  modoProteccion  "advertencia" o "soloyo"
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [hoja] Hoja sobre la que actuar (por defecto la activa)
 * @param {boolean} [silencioso] Si es true, no muestra alertas ni toasts individuales
 */
function protegerFxHoja(modoProteccion, hoja = SpreadsheetApp.getActiveSheet(), silencioso = false) {

  // Utilizado para etiquetar las reglas de protección de intervalos
  const opcionesFechaHora = {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    timeZone: Session.getScriptTimeZone(),
    timeZoneName: 'short',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false
  };

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  let numIntervalosProtegidos = 0;
  let iniciado = false;

  if (!silencioso) hdc.toast(`Buscando fórmulas en [${hoja.getName()}]...`, '👀 HdC+', -1);

  // (I) Extraer coordenadas de celdas con fórmulas
  const listaCeldasFx = [];
  const intervaloFx = hoja.getDataRange().getFormulas();
  for (let nFil = 0; nFil < intervaloFx.length; nFil++) {
    for (let nCol = 0; nCol < intervaloFx[nFil].length; nCol++) {
      if (intervaloFx[nFil][nCol]) {
        listaCeldasFx.push({
          r1: nFil + 1,
          c1: nCol + 1,
          r2: nFil + 1,
          c2: nCol + 1
        });
      }
    }
  }

  if (listaCeldasFx.length == 0) {
    if (!silencioso) hdc.toast(`No se han encontrado fórmulas en la hoja [${hoja.getName()}].`, '⚠️ HdC+', 5);
    return;
  }

  try {

    // (II) Algoritmo de agrupación optimizado de dos pasadas (O(N log N))

    // PASA 1: Agrupación Vertical (Columnas)
    // Ordenar por columna (asc) y luego por fila (asc)
    listaCeldasFx.sort((a, b) => a.c1 - b.c1 || a.r1 - b.r1);

    const bloquesVerticales = [];
    if (listaCeldasFx.length > 0) {
      let actual = { ...listaCeldasFx[0] };
      for (let i = 1; i < listaCeldasFx.length; i++) {
        let sig = listaCeldasFx[i];
        // Si es la misma columna y es la fila inmediatamente siguiente, extendemos el bloque
        if (sig.c1 === actual.c1 && sig.r1 === actual.r2 + 1) {
          actual.r2 = sig.r2;
        } else {
          bloquesVerticales.push(actual);
          actual = { ...sig };
        }
      }
      bloquesVerticales.push(actual);
    }

    // PASA 2: Agrupación Horizontal (Bloques idénticos en columnas adyacentes)
    // Ordenar por fila inicio (asc), fila fin (asc) y luego por columna (asc)
    bloquesVerticales.sort((a, b) => a.r1 - b.r1 || a.r2 - b.r2 || a.c1 - b.c1);

    const intervalosFinales = [];
    if (bloquesVerticales.length > 0) {
      let actual = { ...bloquesVerticales[0] };
      for (let i = 1; i < bloquesVerticales.length; i++) {
        let sig = bloquesVerticales[i];
        // Si coinciden exactamente en filas y es la columna inmediatamente siguiente, ensanchamos el bloque
        if (sig.r1 === actual.r1 && sig.r2 === actual.r2 && sig.c1 === actual.c2 + 1) {
          actual.c2 = sig.c2;
        } else {
          intervalosFinales.push(actual);
          actual = { ...sig };
        }
      }
      intervalosFinales.push(actual);
    }

    if (silencioso || ui.alert(ENCABEZADO_ALERTAS, `⚠️ ¿Proteger intervalos con fórmulas en [${hoja.getName()}]?

Se van a proteger ${intervalosFinales.length} intervalos en esta pestaña. Los intervalos protegidos previamente por HdC+ se regenerarán.
      
${modoProteccion == PROTECCION.modo.advertencia
  ? 'Se mostrará una advertencia al tratar de editar las celdas de los intervalos protegidos.'
  : `Solo tú (${Session.getActiveUser().getEmail()}) podrás editar las celdas de los intervalos protegidos.`} `, ui.ButtonSet.OK_CANCEL) == ui.Button.OK) {

      iniciado = true;
      if (!silencioso) hdc.toast(`Protegiendo ${intervalosFinales.length} intervalos...`, '🔒 HdC+', -1);

      // (III) Aplicar protecciones
      hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion)
        .forEach(regla => regla.remove());

      intervalosFinales.forEach(intervalo => {
        const proteccion = hoja.getRange(
          intervalo.r1,
          intervalo.c1,
          intervalo.r2 - intervalo.r1 + 1,
          intervalo.c2 - intervalo.c1 + 1,
        ).protect();
        const descripcion = `${PROTECCION.descripcion} ${new Date().toLocaleString(Session.getActiveUserLocale(), opcionesFechaHora)}`;
        switch (modoProteccion) {
          case PROTECCION.modo.advertencia: proteccion.setWarningOnly(true).setDescription(descripcion); break;
          case PROTECCION.modo.soloYo: proteccion.addEditor(Session.getActiveUser()).setDescription(descripcion); break;      
        }
        numIntervalosProtegidos++;
      });
    }

  } catch (e) {
    if (!silencioso) hdc.toast(`Error en [${hoja.getName()}]: ${e.message}`, '🔴 HdC+', 10);
  }

  if (!silencioso) hdc.toast(
    `Se han protegido ${numIntervalosProtegidos} intervalos en [${hoja.getName()}].`,
    `${numIntervalosProtegidos == 0 || !iniciado ? '🟠' : '🟢'} HdC+`,
    5
  );

}

/**
 * Eliminar todos los intervalos protegidos previamente por HdC+ (o todos los existentes)
 * @param {boolean} [soloHdCPlus]
 * @param {GoogleAppsScript.Spreadsheet.Sheet} [hoja] Hoja sobre la que actuar
 * @param {boolean} [silencioso]
 */
function eliminarProteccionesFx(soloHdCPlus = true, hoja = SpreadsheetApp.getActiveSheet(), silencioso = false) {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  let numIntervalosDesprotegidos = 0;
  let iniciado = false;

  if (!silencioso) hdc.toast(`Buscando protecciones en [${hoja.getName()}]...`, '👀 HdC+', -1);
  
  let protecciones = hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (soloHdCPlus) protecciones = protecciones.filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion);
  
  if (protecciones.length == 0) {
    if (!silencioso) hdc.toast(`No hay protecciones que eliminar en [${hoja.getName()}].`, '⚠️ HdC+', 5);
    return;
  }
  
  if (silencioso || ui.alert(ENCABEZADO_ALERTAS, `⚠️ ¿Eliminar intervalos protegidos en [${hoja.getName()}]?

Se van a eliminar ${protecciones.length} intervalos protegidos en esta pestaña.`,
    ui.ButtonSet.OK_CANCEL) == ui.Button.OK) {
        
        iniciado = true;
        if (!silencioso) hdc.toast(`Desprotegiendo ${protecciones.length} intervalos...`, '🔓 HdC+', -1);
        
        protecciones.forEach(regla => {
          regla.remove();
          numIntervalosDesprotegidos++;
        });
    }

  if (!silencioso) hdc.toast(
    `Se han desprotegido ${numIntervalosDesprotegidos} intervalos en [${hoja.getName()}].`,
    `${numIntervalosDesprotegidos == 0 || !iniciado ? '🟠' : '🟢'} HdC+`,
    5
  );
}

// Funciones para el Panel Lateral de Protección
function abrirPanelProteccion() {
  const html = HtmlService.createTemplateFromFile('panelProteccion')
    .evaluate()
    .setTitle('🛡️ Gestión de Protecciones');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Devuelve la lista de hojas de la hoja de cálculo con metadatos para el panel
 */
function getHojasParaPanel() {
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActiva = hdc.getActiveSheet().getName();
  return hdc.getSheets().map(h => ({
    nombre: h.getName(),
    oculta: h.isSheetHidden(),
    activa: h.getName() === hojaActiva
  }));
}

/**
 * Procesa la acción masiva desde el panel lateral (OBSOLETA por nueva lógica granular)
 */
function procesarAccionPanel(datos) {
  // Se mantiene por compatibilidad si fuera necesario, pero el panel usará ahora la versión individual
  return "Esta función ha sido sustituida por procesarAccionHojaIndividual.";
}

/**
 * Procesa una única hoja desde el panel lateral y devuelve el número de intervalos afectados
 */
function procesarAccionHojaIndividual(datos) {
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = hdc.getSheetByName(datos.nombreHoja);
  
  if (!hoja) throw new Error(`No se encontró la pestaña: ${datos.nombreHoja}`);
  
  let intervalos = 0;
  if (datos.accion === 'proteger') {
    intervalos = protegerFxHojaInternal_(datos.modo, hoja);
  } else if (datos.accion === 'desproteger') {
    intervalos = eliminarProteccionesFxInternal_(datos.soloHdc, hoja);
  }
  
  return {
    nombre: datos.nombreHoja,
    intervalos: intervalos
  };
}