// Funciones invocadas por menú mostrar / ocultar hojas por color
function mostrarHojasNaranja() {conmutarHojasColor(COLORES_HOJAS.naranja,true);}
function mostrarHojasAzul() {conmutarHojasColor(COLORES_HOJAS.azul, true);}
function ocultarHojasNaranja() {conmutarHojasColor(COLORES_HOJAS.naranja, false);}
function ocultarHojasAzul() {conmutarHojasColor(COLORES_HOJAS.azul, false);}

/**
 * Hace visible u oculta las hojas de uno de los colores parametrizados
 * @param {Object}  color   Objeto que representa el color como { nombre: 'nombre_color', hex:'rrggbb' }
 * @param {boolean} visible true para mostrar, false para ocultar
 */
function conmutarHojasColor(color, visible) {

  // Formato canónico de color hex, con canal alfa y en mayúsculas
  const colorHex = color.hex.toUpperCase().slice(1).padStart(9, '#FF');
  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const idHojaActual = SpreadsheetApp.getActiveSheet().getSheetId();
  const numHojas = hdc.getSheets().length;
  const hojasColor = hdc.getSheets()
    .filter(hoja => (hoja.isSheetHidden() ^ !visible) && // XOR para invertir resultado booleano de isSheetHidden()
      hoja.getTabColorObject().getColorType() == SpreadsheetApp.ColorType.RGB  && 
      hoja.getTabColorObject().asRgbColor().asHexString().toUpperCase().slice(1).padStart(9, '#FF') == colorHex
    );
  
  if (hojasColor.length == 0) ui.alert(`No hay hojas ${visible ? 'ocultas' : 'visibles'} de color ${color.nombre}.`, SpreadsheetApp.getUi().ButtonSet.OK);
  else try {
      hojasColor.forEach(hoja => visible ? hoja.showSheet() : hoja.hideSheet());
      ui.alert(`Se han ${visible ? 'hecho visibles' : 'ocultado'} ${hojasColor.length} hoja(s) de color ${color.nombre}.`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
      ui.alert(`Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
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
  
  if (hojasEliminar.length == 0) ui.alert('No hay hojas ocultas que eliminar.', SpreadsheetApp.getUi().ButtonSet.OK);
  else try {
    // Se usa try para fallar graciosamente cuando en un escenario de concurrencia se eliminan
    // algunas de las hojas mientras el usuario visualiza la alerta de confirmación
    if (SpreadsheetApp.getUi().alert('⚠️ ¿Eliminar hojas?',
      `Se van a eliminar ${hojasEliminar.length} hoja(s) de la hoja de cálculo.

      Puedes revertir el proceso utilizando el comando deshacer tantas veces como sea necesario.`,
    ui.ButtonSet.OK_CANCEL) == ui.Button.OK) hojasEliminar.forEach(hoja => hdc.deleteSheet(hoja));
  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al eliminar las hojas, inténtalo de nuevo.
      
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
  const hojasEliminar = hdc.getSheets().filter(hoja => hoja.getSheetId()!= idHojaActual);

  if (hojasEliminar.length == 0) ui.alert('No hay más hojas que puedan eliminarse.', SpreadsheetApp.getUi().ButtonSet.OK);
  else try {
    if (SpreadsheetApp.getUi().alert('⚠️ ¿Eliminar hojas?',
      `Se van a eliminar ${hojasEliminar.length} hoja(s) de la hoja de cálculo.

      Puedes revertir el proceso utilizando el comando deshacer tantas veces como sea necesario.`,
    ui.ButtonSet.OK_CANCEL) == ui.Button.OK) hojasEliminar.forEach(hoja => hdc.deleteSheet(hoja));
  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al eliminar las hojas, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

}

/**
 * Oculta todas las hojas de la hdc excepto la activa
 */
function ocultarHojas() {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();

  if (hdc.getSheets().length == 1) ui.alert('No hay más hojas que puedan ocultarse.', SpreadsheetApp.getUi().ButtonSet.OK);
  else {
    const idHojaActual = SpreadsheetApp.getActiveSheet().getSheetId();
    const hojasVisibles = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(hoja => !hoja.isSheetHidden() && hoja.getSheetId() != idHojaActual);
    if (hojasVisibles.length == 0) ui.alert('No hay más hojas visibles que ocultar.', SpreadsheetApp.getUi().ButtonSet.OK);
    else try {
      hojasVisibles.forEach(hoja => hoja.hideSheet());
      ui.alert(`Se han ocultado ${hojasVisibles.length} hoja(s).`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      ui.alert(`Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
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
  const hojasOcultas = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(hoja => hoja.isSheetHidden());
  
  if (hojasOcultas.length == 0) ui.alert('No hay hojas ocultas que mostrar.', SpreadsheetApp.getUi().ButtonSet.OK);
  else try {
    hojasOcultas.forEach(hoja => hoja.showSheet());
    ui.alert(`Se han hecho visibles ${hojasOcultas.length} hoja(s).`, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
      
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
  
  if (hdc.getSheets().length == 1) ui.alert('No puedes ocultar la única hoja existente.', ui.ButtonSet.OK);
  else {
    const hojasOcultas = SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(hoja => hoja.isSheetHidden());
    try {
      hojasOcultas.forEach(hoja => hoja.showSheet());
      hojaActual.hideSheet();
      if (hojasOcultas.length > 0) ui.alert(`Se han hecho visibles ${hojasOcultas.length} hoja(s).`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      ui.alert(`Se ha producido un error inesperado al ajustar la visibilidad de las hojas, inténtalo de nuevo.
        
        ⚠️ ${e.message}`, ui.ButtonSet.OK);
    }
  }

}