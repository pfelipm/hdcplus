/**
 * Cuenta las celdas con un color de texto o fondo
 * determinado.
 *
 * @param {"A2:C20"} rango_cadena
 * Rango de celdas sobre el que realizar la cuenta, entre comillas dobles,
 * con / sin indicación de la hoja (Ej: "Hoja 2!A2:C20").
 * 
 * @param {"fondo" | "fuente"} objeto
 * Indicación de cuenta de color de "fondo" o "fuente" de la celda.
 *
 * @param {"#FF0000" | "como"} color
 * Indicación del color en hexadecimal #RRGGBB (#FF0000)
 * o "como", lo que supone facilitar en el siguiente parámetro
 * la referencia a una celda de la que tomar el formato
 * para realizar el recuento.
 * 
 * @param {"A1"} celda_cadena
 * (Opcional) Referencia a la celda, entre comillas
 * dobles, cuyo atributo de color de fuente o fondo se
 * utilizará como modelo, con / sin indicación de la
 * hoja (Ej: "Hoja 2!A1").
 *
 * @return nº de celdas del color indicado
 *
 * @customfunction
 */

function CONTARCOLOR(rango_cadena, objeto, color, celda_cadena) {

  // Control de parámetros
  
  if (typeof rango_cadena == 'undefined' || typeof(objeto) == 'undefined' || typeof color == 'undefined') { 
    return '!Faltan argumentos';
  }
  if (typeof rango_cadena != 'string') {
    return '!El rango debe ser una cadena de texto';
  }
  if (typeof objeto != 'string' ||
      (typeof objeto == 'string' && (objeto.toLowerCase() != 'fondo' && objeto.toLowerCase() != 'fuente'))) {
    return '!Objeto para tipo coincidencia inválido';
  }
  if (typeof color != 'string') {
    return '!Color inválido';
  }
  if (color.toLowerCase() == 'como' && (typeof celda_cadena != 'string'
      || (typeof celda_cadena == 'string' && !celda_cadena))) {
    return '!Celda modelo incorrecta o falta';
  }
  
  // Los parámetros parecen correctos ¡adelante!

  var resultado;
  var rango = SpreadsheetApp.getActiveSpreadsheet().getRange(rango_cadena);
  var nf = rango.getNumRows();
  var nc = rango.getNumColumns();
  var cuenta = 0;
  var f, c, colores;
  var cuenta;
  
  // ¿Fondo o fuente?
  if (objeto.toLowerCase() == 'fondo') {
    colores = rango.getBackgrounds();
  } else { // objeto == 'fuente'
    colores = rango.getFontColors();
  }
  
  // ¿Color específico o referencia a celda?
  
  if (color.toLowerCase() == 'como') {   
    if (objeto.toLowerCase() == 'fondo') {
      color = SpreadsheetApp.getActiveSpreadsheet().getRange(celda_cadena).getBackground();
    }
    else { // objeto = 'fuente'
      color = SpreadsheetApp.getActiveSpreadsheet().getRange(celda_cadena).getFontColor();
    }
  }
  
  for (f = 0; f < nf; f++) {
    for (c = 0; c < nc; c++) {
      if (colores[f][c] == color.toLowerCase()) {cuenta++;}
    }
  }
  resultado = cuenta; 
    
  return resultado;
}