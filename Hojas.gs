// Funciones invocadas por menú mostrar / ocultar hojas por color

function mostrarHojasNaranja() {mostrarHojasColor(COLOR_HOJA_N);}

function mostrarHojasAzul() {mostrarHojasColor(COLOR_HOJA_A);}

function ocultarHojasNaranja() {ocultarHojasColor(COLOR_HOJA_N);}

function ocultarHojasAzul() {ocultarHojasColor(COLOR_HOJA_A);}

function mostrarHojasColor(color) {

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var n = 0;
  
  hojas.map(function(h){
  
    if (h.getTabColor() == color && h.isSheetHidden()) {h.showSheet(); n++}})

  if (n == 0) {SpreadsheetApp.getUi().alert('No se han encontrado hojas ocultas con la pestaña de ese color.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
}

function ocultarHojasColor(color) {

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var n = 0;
  
  hojas.map(function(h){
  
    if (h.getTabColor() == color && !h.isSheetHidden()) {h.hideSheet(); n++}})

  if (n == 0) {SpreadsheetApp.getUi().alert('No se han encontrado hojas visibles con la pestaña de ese color.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
}

function eliminarHojasOcultas() {

  var hdc = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = hdc.getSheets();
  var nombreHojas = '';
  
  hojas = hojas.filter(function(h) {
    
    if (h.isSheetHidden()) {nombreHojas += h.getName()+'\n';}
    return h.isSheetHidden();})
 
  if (nombreHojas) {  
    if (SpreadsheetApp.getUi().alert('⚠️ ¿Eliminar hojas ocultas?',
      'Se van a eliminar las hojas: \n\n' + nombreHojas + '\n' +
      'Puedes revertir el proceso utilizando tantas veces\n' +
      'como sea necesario la función de deshacer de la\n' +
      'hoja de cálculo.'
      ,SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK)
        {hojas.map(function(h) {hdc.deleteSheet(h);})
    }
  }
  else {SpreadsheetApp.getUi().alert('No se han encontrado hojas ocultas.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
}

function eliminarHojas() {

  var hdc = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = hdc.getSheets();
  var hojaActualId = SpreadsheetApp.getActiveSheet().getSheetId();
  var nombreHojas = '';
  
  hojas = hojas.filter(function(h) {
    
    if (h.getSheetId() != hojaActualId) {nombreHojas += h.getName()+'\n';}
    return h.getSheetId() != hojaActualId;})
 
  if (nombreHojas) {  
 
    if (SpreadsheetApp.getUi().alert('⚠️ ¿Eliminar hojas?',
      'Se van a eliminar las hojas: \n\n' + nombreHojas + '\n' +
      'Puedes revertir el proceso utilizando tantas veces\n' +
      'como sea necesario la función de deshacer de la\n' +
      'hoja de cálculo.'
      ,SpreadsheetApp.getUi().ButtonSet.OK_CANCEL) == SpreadsheetApp.getUi().Button.OK)
        {hojas.map(function(h) {hdc.deleteSheet(h);})
    }
  }
  else {SpreadsheetApp.getUi().alert('No hay más hojas que eliminar.', SpreadsheetApp.getUi().ButtonSet.OK);}

}

function ocultarHojas() {

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var hojaActualId = SpreadsheetApp.getActiveSheet().getSheetId();
  
  hojas.map(function(h){
      
    if (!h.isSheetHidden() && h.getSheetId() != hojaActualId) {h.hideSheet();}})

}

function mostrarHojas() {

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var hojaActual = SpreadsheetApp.getActiveSheet();
  var n = 0;
  
  hojas.map(function(h){
  
    if (h.isSheetHidden()) {h.showSheet(); n++}})

  if (n == 0) {SpreadsheetApp.getUi().alert('No se han encontrado hojas ocultas.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
}

function mostrarNoActual() {

  var hojas = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var hojaActual = SpreadsheetApp.getActiveSheet();
  var n = 0;
  
  hojas.map(function(h){
  
    if (h.isSheetHidden()) {h.showSheet(); n++}})

  hojaActual.hideSheet();
  
}