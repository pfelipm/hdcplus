// Envoltorio para invocar la función CONTARCOLOR como HDCP_...()
/**
 * Realiza un recuento, calcula la suma o el promedio de los valores de las celdas
 * que tienen un color de texto o fondo determinado. También puede simplemente devolver
 * los valores que cumplan la condición de color para que puedan ser utilizados en
 * cualquier fórmula.
 *  
 * @param {"A2:C20"} rango_cadena
 * Rango de celdas sobre el que realizar la cuenta, entre comillas dobles,
 * con o sin indicación de la hoja (Ej: "Hoja 2!A2:C20").
 * 
 * @param {"fondo" | "fuente"} modo_color
 * Modo de coincidencia de color: "FONDO" o "FUENTE" de la celda.
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
 * utilizará como modelo, con o sin indicación de la
 * hoja (Ej: "Hoja 2!A1").
 * 
 * @param {"PROMEDIO"} tipo_calculo
 * (Opcional) RECUENTO: recuento de celdas. 
 * SUMA: suma de los valores de las celdas.
 * PROMEDIO: promedio de los valores de las celdas.
 * VALORES: contenido de las celdas que cumplen la condición de color.
 * Si se omite se asume RECUENTO.
 *
 * @return Nº de celdas, suma o promedio de los valores de las celdas del color indicado
 *         o contenido de las celdas que cumplen la condición de color.
 *
 * @customfunction
 */
function HDCP_CONTARCOLOR(rango_cadena, modo_color, color, celda_cadena, tipo_calculo = 'RECUENTO') {
   return CONTARCOLOR(...arguments);
}

/**
 * Realiza un recuento, calcula la suma o el promedio de los valores de las celdas
 * que tienen un color de texto o fondo determinado. También puede simplemente devolver
 * los valores que cumplan la condición de color para que puedan ser utilizados en
 * cualquier fórmula.
 * 
 * Esta función es un alias de la función HDCP_CONTARCOLOR().
 *  
 * @param {"A2:C20"} rango_cadena
 * Rango de celdas sobre el que realizar la cuenta, entre comillas dobles,
 * con o sin indicación de la hoja (Ej: "Hoja 2!A2:C20").
 * 
 * @param {"fondo" | "fuente"} modo_color
 * Modo de coincidencia de color: "FONDO" o "FUENTE" de la celda.
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
 * utilizará como modelo, con o sin indicación de la
 * hoja (Ej: "Hoja 2!A1").
 * 
 * @param {"PROMEDIO"} tipo_calculo
 * (Opcional) RECUENTO: recuento de celdas, SUMA: suma de los valores de las celdas,
 * PROMEDIO: promedio de los valores de las celdas. Si se omite se asume RECUENTO,
 * VALORES: contenido de las celdas que cumplen la condición de color.
 *
 * @return Nº de celdas, suma o promedio de los valores de las celdas del color indicado
 *         o contenido de las celdas que cumplen la condición de color.
 * 
 * @customfunction
 */

function CONTARCOLOR(rango_cadena, modo_color, color, celda_cadena, tipo_calculo = 'RECUENTO') {

  const TIPOS_CALCULO = {
    recuento: 'RECUENTO',
    suma: 'SUMA',
    promedio: 'PROMEDIO',
    valores: 'VALORES'
  };

  const MODOS_COLOR = {
    fuente: 'FUENTE',
    fondo: 'FONDO',
  }

  const COMO = 'COMO';

  /**
   * Devuelve una representación de texto para el color del objeto,
   * como #RRGGBB o una enumeración de tipo ThemeColorType. Cuando
   * un objeto de color es de tipo Theme al parecer puede tener una
   * representación hexadecimal o de tipo ThemeColorType ¿?
   * @param {SpreadSheetApp.color} color Objeto de color
   * @return {string}
   */
  function obtenerColorObjeto(color) {

    // Según la documentación, asHexString() puede devolver una cadena hexadecimal tipo CSS de 7 caracteres (#rrggbb) o de 9 caracteres
    // con canal alfa (#aarrggbb), en mis pruebas solo me ha ocurrido con el negro puro (#FF000000), así que siempre rellenaremos por la izq.
    if(color.getColorType() == SpreadsheetApp.ColorType.RGB) return color.asRgbColor().asHexString().toUpperCase().slice(1).padStart(9, '#FF');
    else if (color.getColorType() == SpreadsheetApp.ColorType.THEME)  {
      color = color.asThemeColor();
      switch(color.getColorType()) {
        case SpreadsheetApp.ColorType.RGB: return color.asRgasHexString().asHexString().toUpperCase().slice(1).padStart(9, '#FF'); break;
        case SpreadsheetApp.ColorType.THEME: return color.getThemeColorType().toString(); break;
        default: return undefined;
      }
    } else return undefined;
    
  }

  // Control de parámetros
  
  if (rango_cadena == undefined || modo_color == undefined || color == undefined) throw 'Faltan argumentos';
  if (typeof rango_cadena != 'string') throw 'El rango debe ser una cadena de texto';
  modo_color = modo_color.toUpperCase?.();
  if (!modo_color || !Object.values(MODOS_COLOR).includes(modo_color.toUpperCase())) throw 'Modo de coincidencia de color inválido';
  color = color.toUpperCase?.();
  if (!color) throw 'Color de coincidencia no especificado';
  if (color == COMO && (typeof celda_cadena != 'string' || celda_cadena == '')) throw 'Celda modelo incorrecta o no indicada';
  tipo_calculo = tipo_calculo.toUpperCase?.();
  if (!Object.values(TIPOS_CALCULO).includes(tipo_calculo)) throw 'Tipo de cálculo no admitido';

  // Los parámetros parecen correctos ¡adelante!

  // Variables para almacenar el resultado de la función
  let resultado;
  let recuento = 0;
  let suma = 0;
  const matrizResultado = []; 

  const rango = SpreadsheetApp.getActiveSpreadsheet().getRange(rango_cadena);
  const valores = rango.getValues();
  const nf = rango.getNumRows();
  const nc = rango.getNumColumns();
  let f, c, colores;
  
  // ¿Fondo o fuente?
  if (modo_color == MODOS_COLOR.fondo) {
    // Aún no se indica que getBackground() esté obsoleto, pero uso ya en su lugar el método que devuelve objeto de color
    colores = rango.getBackgroundObjects()
      .map(colorObjectRow => colorObjectRow.map(colorObject => obtenerColorObjeto(colorObject)));
  } else { // objeto == 'fuente'
    colores = rango.getFontColorObjects()
      .map(colorObjectRow => colorObjectRow.map(colorObject => obtenerColorObjeto(colorObject)));
  }

  // ¿Color específico o referencia a celda?
  
  try {
    if (color == COMO) {
      // De nuevo prefiero el método que devuelve objeto de color
      if (modo_color == MODOS_COLOR.fondo) color = obtenerColorObjeto(SpreadsheetApp.getActiveSpreadsheet().getRange(celda_cadena).getBackgroundObject());
      else color = obtenerColorObjeto(SpreadsheetApp.getActiveSpreadsheet().getRange(celda_cadena).getFontColorObject());
      // Añadir '#' si el usuario no la ha usado al especificar el color
    } else color = (color.charAt(0) != '#' ? '#' + color : color).slice(1).padStart(9, '#FF');
  } catch { throw 'La referencia al intervalo de celdas o a la celda modelo es inválida'; }
 
  // Cálculo sobre el intervalo de celdas: recuento, suma o promedio

  let resultadoFila;
  colores.forEach((coloresFila, fila) => {
      resultadoFila = [];
      coloresFila.forEach((colorCelda, columna) => {
      if (colorCelda == color && colorCelda != undefined && color != undefined) {
        switch(tipo_calculo) {
          case TIPOS_CALCULO.recuento: recuento++; break;
          case TIPOS_CALCULO.suma:
          case TIPOS_CALCULO.promedio:  {
            if (typeof valores[fila][columna] == 'number') {
              recuento++;
              suma+= valores[fila][columna];
            }
            break;
          }
          case TIPOS_CALCULO.valores: resultadoFila.push(valores[fila][columna]);
        }
      }
    });
    if (resultadoFila.length > 0) matrizResultado.push(resultadoFila);
  });

  switch (tipo_calculo) {
    case TIPOS_CALCULO.recuento: resultado = recuento; break;
    case TIPOS_CALCULO.suma: resultado = suma; break;
    case TIPOS_CALCULO.promedio: {
      if (recuento == 0) throw '#¡DIV/0! La evaluación de la función CONTARCOLOR ha provocado un error de división por cero';
      resultado = suma / recuento;
      break;
    }
    case TIPOS_CALCULO.valores: resultado = matrizResultado;
  } 
  
  // Evitar excepción tipo "no se ha encontrado el intervalo"
  if (tipo_calculo == TIPOS_CALCULO.valores && resultado.length == 0) throw 'No se han encontrado celdas que cumplan la condición de color';
  return resultado;

}