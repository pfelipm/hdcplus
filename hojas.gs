// Envoltorios para funciones invocadas por menú mostrar hojas por color
function mostrarHojasAzul() { conmutarHojasColor(COLORES_HOJAS.azul, true, false); }
function mostrarHojasMorado() { conmutarHojasColor(COLORES_HOJAS.morado, true, false); }
function mostrarHojasVerde() { conmutarHojasColor(COLORES_HOJAS.verde, true, false); }
function mostrarHojasNaranja() { conmutarHojasColor(COLORES_HOJAS.naranja, true, false); }
function mostrarHojasRojo() { conmutarHojasColor(COLORES_HOJAS.rojo, true, false); }

// Envoltorios para funciones invocadas por menú mostrar solo hojas por color
function mostrarSoloHojasAzul() { conmutarHojasColor(COLORES_HOJAS.azul, true, true); }
function mostrarSoloHojasMorado() { conmutarHojasColor(COLORES_HOJAS.morado, true, true); }
function mostrarSoloHojasVerde() { conmutarHojasColor(COLORES_HOJAS.verde, true, true); }
function mostrarSoloHojasNaranja() { conmutarHojasColor(COLORES_HOJAS.naranja, true, true); }
function mostrarSoloHojasRojo() { conmutarHojasColor(COLORES_HOJAS.rojo, true, true); }

// Envoltorios para funciones invocadas por menú ocultar hojas por color
function ocultarHojasAzul() { conmutarHojasColor(COLORES_HOJAS.azul, false, false); }
function ocultarHojasMorado() { conmutarHojasColor(COLORES_HOJAS.morado, false, false); }
function ocultarHojasVerde() { conmutarHojasColor(COLORES_HOJAS.verde, false, false); }
function ocultarHojasNaranja() { conmutarHojasColor(COLORES_HOJAS.naranja, false, false); }
function ocultarHojasRojo() { conmutarHojasColor(COLORES_HOJAS.rojo, false, false); }

// Envoltorios para funciones invocadas por menú eliminar hojas por color
function eliminarHojasAzul() { eliminarHojasColor(COLORES_HOJAS.azul); }
function eliminarHojasMorado() { eliminarHojasColor(COLORES_HOJAS.morado); }
function eliminarHojasVerde() { eliminarHojasColor(COLORES_HOJAS.verde); }
function eliminarHojasNaranja() { eliminarHojasColor(COLORES_HOJAS.naranja); }
function eliminarHojasRojo() { eliminarHojasColor(COLORES_HOJAS.rojo); }

// Envoltorios para funciones invocadas por menú ordenar hojas asc/dsc
function ordenarHojasAsc() { ordenarHojasServicio(true); }
function ordenarHojasDesc() { ordenarHojasServicio(false); }

/**
 * Ordena las hojas alfabéticamente en sentido ascendente o descendente
 * @param {boolean} [ascendente]
 */
function ordenarHojasServicio(ascendente = true) {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  // Precisamos únicamente el código del idioma (idioma_región)
  const collator = new Intl.Collator(hdc.getSpreadsheetLocale().split('_')[0], { numeric: true, sensitivity: 'base' });

  // Vector { hoja, nombre } para evitar múltiples llamadas a getName() en la función del sort().
  const hojasOrdenadas = hdc.getSheets()
    .map(hoja => ({ objeto: hoja, nombre: hoja.getName() }))
    .sort((hoja1, hoja2) => Math.pow(-1, !ascendente) * collator.compare(hoja1.nombre, hoja2.nombre));

  const hojasPorOcultar = [];
  const hojaActual = hdc.getActiveSheet();

  try {

    hojasOrdenadas.forEach((hoja, pos) => {
      if (hoja.objeto.isSheetHidden()) hojasPorOcultar.push(hoja.objeto);
      hdc.setActiveSheet(hoja.objeto);
      hdc.moveActiveSheet(pos + 1);
    });
    // Necesario confirmar cambios antes de procesar hojas a ocultar
    SpreadsheetApp.flush();
    // Sin esta pausa en ocasiones alguna hoja oculta queda visible
    Utilities.sleep(1000);
    if (hojasPorOcultar.length > 0) hojasPorOcultar.forEach(hoja => hoja.hideSheet());
    hdc.setActiveSheet(hojaActual);
    SpreadsheetApp.flush();
    ui.alert(ENCABEZADO_ALERTAS, `Se han ordenado alfabéticamente ${hojasOrdenadas.length} hojas en sentido ${ascendente ? 'ascendente' : 'descendente'}.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ordenar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Ordena las hojas alfabéticamente en sentido ascendente o descendente usando la API de Sheets.
 * Mucho más rapido, pero ***NO USADO*** por el momento, dado que requiere https://www.googleapis.com/auth/spreadsheets.
 * @param {boolean} [ascendente]
 */
function ordenarHojasApi(ascendente = true) {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  // Precisamos únicamente el código del idioma (idioma_región)
  const collator = new Intl.Collator(hdc.getSpreadsheetLocale().split('_')[0], { numeric: true, sensitivity: 'base' });

  try {

    // Obtiene vector de propiedades de hojas (nombre, índice)
    const hojasOrdenadas = Sheets.Spreadsheets.get(
      SpreadsheetApp.getActiveSpreadsheet().getId(), { fields: 'sheets.properties(sheetId,title,index)' }
    ).sheets.sort((hoja1, hoja2) =>
      Math.pow(-1, !ascendente) * collator.compare(hoja1.properties.title, hoja2.properties.title)
    );

    const hojaActual = hdc.getActiveSheet();
    Sheets.Spreadsheets.batchUpdate(
      {
        requests: hojasOrdenadas.map(({ properties: { sheetId } }, pos) => (
          { updateSheetProperties: { properties: { sheetId, index: pos }, fields: 'index' } }
        )),
        includeSpreadsheetInResponse: false
      },
      hdc.getId()
    );
    hdc.setActiveSheet(hojaActual);
    SpreadsheetApp.flush();
    ui.alert(ENCABEZADO_ALERTAS, `Se han ordenado alfabéticamente ${hojasOrdenadas.length} hojas en sentido ${ascendente ? 'ascendente' : 'descendente'}.`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ordenar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Desordena las hojas de un libro de cálculo de forma aleatoria
 * utilizando el servicio integrado de hojas de cálculo.
 * Mantiene el estado de visibilidad de las hojas y restaura la hoja activa.
 *
 * Generado por Gemini 2.5 Pro (experimental) a partir de la función de ordenación ordenarHojasServicio().
 * https://pablofelip.online/velocidad-permisos-ordenando-pestanas-apps-script/
 */
function desordenarHojas() {

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  try {
    const hojas = hdc.getSheets();
    const numHojas = hojas.length;

    // --- Inicio: Algoritmo de Fisher-Yates (Knuth) Shuffle ---
    // Recorre el array desde el final hacia el principio
    for (let i = numHojas - 1; i > 0; i--) {
      // Elige un índice aleatorio j entre 0 e i (inclusive)
      const j = Math.floor(Math.random() * (i + 1));
      // Intercambia el elemento en i con el elemento en j
      [hojas[i], hojas[j]] = [hojas[j], hojas[i]];
    }
    // --- Fin: Algoritmo de Fisher-Yates ---
    // Ahora el array 'hojas' contiene los objetos Sheet en un orden aleatorio.

    const hojasPorOcultar = [];
    const hojaActual = hdc.getActiveSheet(); // Guarda la hoja activa actual

    // Mueve las hojas a su nueva posición aleatoria
    hojas.forEach((hoja, pos) => {
      // Comprueba si la hoja estaba oculta ANTES de moverla
      // (moverla la hará visible temporalmente si estaba oculta)
      if (hoja.isSheetHidden()) {
        hojasPorOcultar.push(hoja); // Añade el objeto hoja a la lista para ocultar después
      }
      hdc.setActiveSheet(hoja); // Activa la hoja para poder moverla
      hdc.moveActiveSheet(pos + 1); // Mueve la hoja a su nueva posición (índice + 1)
    });

    // Necesario confirmar cambios antes de procesar hojas a ocultar
    SpreadsheetApp.flush();
    // Mantenemos la pausa por consistencia con la función original,
    // por si hay problemas de sincronización al ocultar rápidamente.
    Utilities.sleep(1000);

    // Vuelve a ocultar las hojas que estaban ocultas originalmente
    hojasPorOcultar.forEach(hoja => hoja.hideSheet());

    // Restaura la hoja que estaba activa al principio
    hdc.setActiveSheet(hojaActual);
    SpreadsheetApp.flush(); // Asegura que la activación final se aplique

    ui.alert(ENCABEZADO_ALERTAS, `Se han desordenado aleatoriamente ${numHojas} hoja(s).`, ui.ButtonSet.OK);

  } catch (e) {
    // Mismo manejo de errores que la función original
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al desordenar las hojas, inténtalo de nuevo.

      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Hace visible u oculta las hojas de uno de los colores parametrizados, ocultando el resto si se desea
 * @param {Object}  color        Objeto que representa el color como { nombre: 'nombre_color', hex:'#rrggbb' }
 * @param {boolean} visible      true para mostrar, false para ocultar
 * @param {boolean} ocultarResto true para ocultar el resto de hojas (se ignora si visible = false)
 */
function conmutarHojasColor(color, visible, ocultarResto) {

  // Formato canónico de color hex, con canal alfa y en mayúsculas
  const colorHex = color.hex.toUpperCase().slice(1).padStart(9, '#FF');
  const ui = SpreadsheetApp.getUi();
  const hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();

  // Se ignora ocultarResto en operaciones de ocultación de hojas de color
  ocultarResto = ocultarResto && visible;

  let hojasColorProcesar = []; // hojas del color seleccionado cuya visibilidad actual hace que deban procesarse
  let otrasHojasOcultar = []; // resto de hojas de otros colores a procesar (ocultar)
  let numHojasVisibles = 0; // número total de hojas visibles
  let numHojasColor = 0; // número de hojas del color seleccionado, con independencia de su visibilidad actual

  // Identificar hojas a procesar y elementos auxiliares
  hojas.forEach(hoja => {
    if (!hoja.isSheetHidden()) numHojasVisibles++;
    if (hoja.getTabColorObject().getColorType() == SpreadsheetApp.ColorType.RGB &&
      hoja.getTabColorObject().asRgbColor().asHexString().toUpperCase().slice(1).padStart(9, '#FF') == colorHex) {
      numHojasColor++;
      // XOR (^) para invertir resultado booleano de isSheetHidden() 
      if (hoja.isSheetHidden() ^ !visible) hojasColorProcesar.push(hoja);
    } else if (ocultarResto && !hoja.isSheetHidden()) otrasHojasOcultar.push(hoja);
  });

  if (!visible && (hojasColorProcesar.length == numHojasVisibles)) ui.alert(ENCABEZADO_ALERTAS, 'No puedes ocultar todas las hojas visibles.', ui.ButtonSet.OK)
  else if (numHojasColor == 0) ui.alert(ENCABEZADO_ALERTAS, `${color.icono} No hay hojas de color ${color.nombre}.`, ui.ButtonSet.OK);
  else if (hojasColorProcesar.length == 0 && otrasHojasOcultar.length == 0) ui.alert(ENCABEZADO_ALERTAS, `${color.icono} No hay hojas ${visible ? 'ocultas' : 'visibles'} de color ${color.nombre}.`, ui.ButtonSet.OK);
  else try {

    // Vector utilizado para construir el mensaje de estado al finalizar
    let mensajeResultado = [];

    // Ajustar visibilidad de hojas del color designado
    if (hojasColorProcesar.length > 0) {
      hojasColorProcesar.forEach(hoja => visible ? hoja.showSheet() : hoja.hideSheet());
      mensajeResultado.push(`Se han ${visible ? 'hecho visibles' : 'ocultado'} ${hojasColorProcesar.length} hojas de color ${color.nombre}.`);
    }

    // Ocultar en su caso el resto de hojas
    if (otrasHojasOcultar.length > 0) {
      otrasHojasOcultar.forEach(hoja => hoja.hideSheet());
      mensajeResultado.push(`Se han ocultado ${otrasHojasOcultar.length} hojas de color distinto al ${color.nombre}.`);
    }

    SpreadsheetApp.flush();
    ui.alert(ENCABEZADO_ALERTAS, `${color.icono} ${mensajeResultado.join(' ')}`, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
        ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Elimina todas la hojas de un color determinado
 * @param {Object} color Objeto que representa el color como { nombre: 'nombre_color', hex:'#rrggbb' }
 */
function eliminarHojasColor(color) {

  // Formato canónico de color hex, con canal alfa y en mayúsculas
  const colorHex = color.hex.toUpperCase().slice(1).padStart(9, '#FF');

  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let hojasColorEliminar = []; // hojas del color seleccionado a eliminar
  let numHojasVisibles = 0; // número total de hojas visibles
  let numHojasColorVisibles = 0 // número de hojas del color seleccionado a eliminar que son visibles
  let numHojasColorOcultas = 0 // número de hojas del color seleccionado a eliminar que están ocultas

  // Identificar hojas a procesar y elementos auxiliares
  hdc.getSheets().forEach(hoja => {
    if (!hoja.isSheetHidden()) numHojasVisibles++;
    if (hoja.getTabColorObject().getColorType() == SpreadsheetApp.ColorType.RGB &&
      hoja.getTabColorObject().asRgbColor().asHexString().toUpperCase().slice(1).padStart(9, '#FF') == colorHex) {
      if (hoja.isSheetHidden()) numHojasColorOcultas++;
      else numHojasColorVisibles++;
      hojasColorEliminar.push(hoja);
    }
  });

  if (hojasColorEliminar.length == 0) ui.alert(ENCABEZADO_ALERTAS, `${color.icono} No hay hojas de color ${color.nombre} que eliminar.`, ui.ButtonSet.OK);
  else if (numHojasVisibles == numHojasColorVisibles) ui.alert(ENCABEZADO_ALERTAS, `No es posible eliminar todas las hojas visibles.`, ui.ButtonSet.OK);
  else try {

    // Vector utilizado para construir el mensaje de estado al finalizar
    let mensajeResultado = [];
    if (numHojasColorVisibles > 0) mensajeResultado.push(`${numHojasColorVisibles} hojas visibles`);
    if (numHojasColorOcultas > 0) mensajeResultado.push(`${numHojasColorOcultas} hojas ocultas`);

    if (ui.alert(ENCABEZADO_ALERTAS, `⚠️ ¿Eliminar hojas de color ${color.nombre}?`,
      `${color.icono} Se van a eliminar ${mensajeResultado.join(' y ')} de color ${color.nombre}.

      Puedes revertir el proceso utilizando el comando deshacer tantas veces como sea necesario.`,
      ui.ButtonSet.OK_CANCEL) == ui.Button.OK) hojasColorEliminar.forEach(hoja => hdc.deleteSheet(hoja));

  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al eliminar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Elimina todas la hojas ocultas
 */
function eliminarHojasOcultas() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojasEliminar = hdc.getSheets().filter(hoja => hoja.isSheetHidden());

  if (hojasEliminar.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No hay hojas ocultas que eliminar.', ui.ButtonSet.OK);
  else try {
    // Se usa try para fallar graciosamente cuando en un escenario de concurrencia se eliminan
    // algunas de las hojas mientras el usuario visualiza la alerta de confirmación
    if (ui.alert(ENCABEZADO_ALERTAS, '⚠️ ¿Eliminar hojas ocultas?',
      `Se van a eliminar ${hojasEliminar.length} hojas no visibles de la hoja de cálculo.

      Puedes revertir el proceso utilizando el comando deshacer tantas veces como sea necesario.`,
      ui.ButtonSet.OK_CANCEL) == ui.Button.OK) hojasEliminar.forEach(hoja => hdc.deleteSheet(hoja));
  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al eliminar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Elimina todas las hojas de la hdc excepto la activa
 */
function eliminarHojas() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const idHojaActual = hdc.getActiveSheet().getSheetId();
  const hojasEliminar = hdc.getSheets().filter(hoja => hoja.getSheetId() != idHojaActual);

  if (hojasEliminar.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No hay más hojas que puedan eliminarse.', ui.ButtonSet.OK);
  else try {
    if (ui.alert(ENCABEZADO_ALERTAS, '⚠️ ¿Eliminar otras hojas?',
      `Se van a eliminar ${hojasEliminar.length} hojas de la hoja de cálculo.

      Puedes revertir el proceso utilizando el comando deshacer tantas veces como sea necesario.`,
      ui.ButtonSet.OK_CANCEL) == ui.Button.OK) hojasEliminar.forEach(hoja => hdc.deleteSheet(hoja));
  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al eliminar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Oculta todas las hojas de la hdc excepto la activa
 */
function mostrarSoloActiva() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();

  if (hdc.getSheets().length == 1) ui.alert(ENCABEZADO_ALERTAS, 'No hay más hojas que puedan ocultarse.', ui.ButtonSet.OK);
  else {
    const idHojaActual = SpreadsheetApp.getActiveSheet().getSheetId();
    const hojasVisibles = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(hoja => !hoja.isSheetHidden() && hoja.getSheetId() != idHojaActual);
    if (hojasVisibles.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No hay más hojas visibles que ocultar.', ui.ButtonSet.OK);
    else try {
      hojasVisibles.forEach(hoja => hoja.hideSheet());
      SpreadsheetApp.flush();
      ui.alert(ENCABEZADO_ALERTAS, `Se han ocultado ${hojasVisibles.length} hojas.`, ui.ButtonSet.OK);
    } catch (e) {
      ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
    }
  }

}

/** 
 * Hace visibles todas las hojas de la hdc
 */
function mostrarHojas() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojasOcultas = hdc.getSheets().filter(hoja => hoja.isSheetHidden());

  if (hojasOcultas.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No hay hojas ocultas que mostrar.', ui.ButtonSet.OK);
  else try {
    hojasOcultas.forEach(hoja => hoja.showSheet());
    SpreadsheetApp.flush();
    ui.alert(ENCABEZADO_ALERTAS, `Se han hecho visibles ${hojasOcultas.length} hojas.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Hace visibles todas las hojas de la hdc excepto la activa
 */
function mostrarTodasMenosActual() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActual = SpreadsheetApp.getActiveSheet();

  if (hdc.getSheets().length == 1) ui.alert(ENCABEZADO_ALERTAS, 'No es posible ocultar la única hoja existente.', ui.ButtonSet.OK);
  else {
    const hojasOcultas = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(hoja => hoja.isSheetHidden());
    try {
      hojasOcultas.forEach(hoja => hoja.showSheet());
      hojaActual.hideSheet();
      if (hojasOcultas.length > 0) {
        SpreadsheetApp.flush();
        ui.alert(ENCABEZADO_ALERTAS, `Se han hecho visibles ${hojasOcultas.length} hojas.`, ui.ButtonSet.OK);
      }
    } catch (e) {
      ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
        ⚠️ ${e.message}`, ui.ButtonSet.OK);
    }
  }

}

/**
 * Conmuta la visibilidad de todas las hojas existentes
 */
function conmutarVisibilidadHojas() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = hdc.getSheets();

  const hojasOcultas = hojas.filter(hoja => hoja.isSheetHidden());
  const hojasVisibles = hojas.filter(hoja => !hoja.isSheetHidden());

  if (hojasVisibles.length == 1 && hojasOcultas.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No es posible ocultar la única hoja existente.', ui.ButtonSet.OK);
  if (hojasOcultas.length == 0) ui.alert(ENCABEZADO_ALERTAS, 'No hay hojas ocultas que mostrar.', ui.ButtonSet.OK);

  else try {
    hojasOcultas.forEach(hoja => hoja.showSheet());
    hojasVisibles.forEach(hoja => hoja.hideSheet());
    SpreadsheetApp.flush();
    ui.alert(ENCABEZADO_ALERTAS, `Se han ocultado ${hojasVisibles.length} hojas y se han hecho visibles otras ${hojasOcultas.length}.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}