// Funciones de transformación del texto

function minusculas() {
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => typeof celda == 'string' ? celda.toLowerCase() : celda));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }    
}

function mayusculas() {
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => typeof celda == 'string' ? celda.toUpperCase() : celda));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Pasa a mayúsculas la letra inicial de las celdas de los intervalos seleccionados.
 * Utiliza expresiones regulares Unicode (\p{L}) para detectar la primera letra en cualquier idioma.
 * 
 * MODIFICADO 10/05/26: Refactorización ES6 + Unicode.
 */
function inicialMays_() {
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();  
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => {
        return typeof celda === 'string' ? celda.replace(/\p{L}/u, l => l.toUpperCase()) : celda;
      }));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Capitaliza las celdas de los intervalos seleccionados (Primera Letra De Cada Palabra).
 * Utiliza propiedades Unicode para detectar límites de palabras internacionales.
 * 
 * MODIFICADO 10/05/26: Refactorización ES6 + Unicode.
 */
function inicialesMays_() {
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();  
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => {
        if (typeof celda !== 'string') return celda;
        // Convierte a minúsculas y capitaliza la primera letra de cada secuencia de letras Unicode
        return celda.toLowerCase().replace(/(^|[^\p{L}])\p{L}/gu, l => l.toUpperCase());
      }));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// --- VERSIONES HISTÓRICAS (PRESESERVADAS POR MOTIVOS DIDÁCTICOS) ---

function inicialMays_anterior_() {
  
  try {

    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();  
    rangos.map(rango => {
      
      const matriz = rango.getValues().map(fila => {
      
        return fila.map(celda => {
      
          // Solo se procesan cadenas de texto
          if (typeof celda != 'string') return celda;

          let pos = 0;
          while (celda.charAt(pos).toLowerCase() == celda.charAt(pos).toUpperCase() && pos < celda.length) pos++;
          if (pos < celda.length) return  celda.slice(0, pos) + celda.charAt(pos).toUpperCase() + celda.slice(pos + 1);
          else return celda;
          
        });
      
      });
      rango.setValues(matriz);

    });
  
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }
}

function inicialesMays_anterior_() {

  try {

    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();  
    rangos.map(rango => {
      
      const matriz = rango.getValues().map(fila => {
      
        return fila.map(celda => {
      
          // Solo se procesan cadenas de texto
          if (typeof celda != 'string') return celda;

          // [1] Se identifican los caracteres no estándar, pueden incluir vocales con tilde, eñes, etc.
          const coincidencias = celda.matchAll(/\W/gi);
          let noAptos = [];
          for (let coincidencia of coincidencias) {
            noAptos.push( { indice: coincidencia.index, caracter: coincidencia[0]});
          }
          // Descartar espacios
          noAptos = noAptos.filter(noApto => noApto.caracter != ' ');
          
          // [2] Ahora los eliminamos todos de la cadena original a procesar, se convierte a vector
          //     para poder modificar elementos en el siguiente paso.
          const prep = [...celda.replaceAll(/\W/gi, ' ')];
          
          // [3] Antes de que split().join() hagan su magia hay que reinsertar en su lugar los caracteres que sí deseamos procesar
          noAptos.forEach(coincidencia => {
            // Si hay diferencia entre mayúsculas / minúsculas es que el carácter debe ser tratado y debemos reinsertarlo ahora
            if (coincidencia.caracter.toLowerCase() != coincidencia.caracter.toUpperCase()) {
              prep[coincidencia.indice] =  coincidencia.caracter;
            }
          });

          // [4] Aplicamos la transformación, solo iniciales de palabras a mayúsculas, resto a minúsculas
          const nuevoVectorCadena = [
            ...prep.join('')
            .toLowerCase()
            .split(' ')
            // split() devuelve n-1 caracteres [''] al trocear una secuencia de n espacios consecutivos
            .map(palabra => palabra.at(0) ? palabra.at(0).toUpperCase() + palabra.slice(1) : '')
            .join(' ')
          ];

          // [5] Resinsertar en su lugar el resto de caracteres especiales que no debían ser tenidos en cuenta
          noAptos.forEach(caracter => {
            // Si no hay diferencia entre mayúsculas / minúsculas es que el carácter debe ser reinsertado
            if (caracter.caracter.toLowerCase() == caracter.caracter.toUpperCase()) {
              nuevoVectorCadena[caracter.indice] = caracter.caracter;
            }
          });    

          // Se transforma el vector en cadena de nuevo
          return nuevoVectorCadena.join('');
          
        });
      
      });
      rango.setValues(matriz);

    });
  
  } catch (e) {
    const ui = SpreadsheetApp.getUi();
    ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }
}

// --- RESTO DE FUNCIONES MODERNIZADAS ---

/**
 * Reemplaza las ocurrencias de un patrón de expresión regular con un carácter especificado.
 */
function sustituir(exp, caracter) {
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();  
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => typeof celda === 'string' ? celda.replace(exp, caracter) : celda));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// Funciones de atajos
const eliminarEspacios = () => sustituir(/\s+/g, '');
const eliminarEspaciosComas = () => sustituir(/,\s+/g, ',');
const eliminarSaltos = () => sustituir(/\n/g, '');
const comasEspacios = () => sustituir(/,/g, ' ');
const comasSaltos = () => sustituir(/,/g, '\n');
const espaciosComas = () => sustituir(/\s/g, ',');
const espaciosSaltos = () => sustituir(/\s/g, '\n');
const saltosEspacios = () => sustituir(/\n/g, ' ');
const saltosComas = () => sustituir(/\n/g, ',');

function latinizar() {
  const latin_map={"·":"-","ß":"ss","Á":"A","Ă":"A","Ắ":"A","Ặ":"A","Ằ":"A","Ẳ":"A","Ẵ":"A","Ǎ":"A","Â":"A","Ấ":"A","Ậ":"A","Ầ":"A","Ẩ":"A","Ẫ":"A","Ä":"A","Ǟ":"A","Ȧ":"A","Ǡ":"A","Ạ":"A","Ȁ":"A","À":"A","Ả":"A","Ȃ":"A","Ā":"A","Ą":"A","Å":"A","Ǻ":"A","Ḁ":"A","Ⱥ":"A","Ã":"A","Ꜳ":"AA","Æ":"AE","Ǽ":"AE","Ǣ":"AE","Ꜵ":"AO","Ꜷ":"AU","Ꜹ":"AV","Ꜻ":"AV","Ꜽ":"AY","Ḃ":"B","Ḅ":"B","Ɓ":"B","Ḇ":"B","Ƀ":"B","Ƃ":"B","Ć":"C","Č":"C","Ç":"C","Ḉ":"C","Ĉ":"C","Ċ":"C","Ƈ":"C","Ȼ":"C","Ď":"D","Ḑ":"D","Ḓ":"D","Ḋ":"D","Ḍ":"D","Ɗ":"D","Ḏ":"D","ǲ":"D","ǅ":"D","Đ":"D","Ƌ":"D","Ǳ":"DZ","Ǆ":"DZ","É":"E","Ĕ":"E","Ě":"E","Ȩ":"E","Ḝ":"E","Ê":"E","Ế":"E","Ệ":"E","Ề":"E","Ể":"E","Ễ":"E","Ḙ":"E","Ë":"E","Ė":"E","Ẹ":"E","Ȅ":"E","È":"E","Ẻ":"E","Ȇ":"E","Ē":"E","Ḗ":"E","Ḕ":"E","Ę":"E","Ɇ":"E","Ẽ":"E","Ḛ":"E","Ꝫ":"ET","Ḟ":"F","Ƒ":"F","Ǵ":"G","Ğ":"G","Ǧ":"G","Ģ":"G","Ĝ":"G","Ġ":"G","Ɠ":"G","Ḡ":"G","Ǥ":"G","Ḫ":"H","Ȟ":"H","Ḩ":"H","Ĥ":"H","Ⱨ":"H","Ḧ":"H","Ḣ":"H","Ḥ":"H","Ħ":"H","Í":"I","Ĭ":"I","Ǐ":"I","Î":"I","Ï":"I","Ḯ":"I","İ":"I","Ị":"I","Ȉ":"I","Ì":"I","Ỉ":"I","Ȋ":"I","Ī":"I","Į":"I","Ɨ":"I","Ĩ":"I","Ḭ":"I","Ꝺ":"D","Ꝼ":"F","Ᵹ":"G","Ꞃ":"R","Ꞅ":"S","Ꞇ":"T","Ꝭ":"IS","Ĵ":"J","Ɉ":"J","Ḱ":"K","Ǩ":"K","Ķ":"K","Ⱪ":"K","Ꝃ":"K","Ḳ":"K","Ƙ":"K","Ḵ":"K","Ꝁ":"K","Ꝅ":"K","Ĺ":"L","Ƚ":"L","Ľ":"L","Ļ":"L","Ḽ":"L","Ḷ":"L","Ḹ":"L"," ":"L","Ꝉ":"L","Ḻ":"L","Ŀ":"L","Ɫ":"L","ǈ":"L","Ł":"L","Ǉ":"LJ","Ḿ":"M","Ṁ":"M","Ṃ":"M","Ɱ":"M","Ń":"N","Ň":"N","Ņ":"N","Ṋ":"N","Ṅ":"N","Ṇ":"N","Ǹ":"N","Ɲ":"N","Ṉ":"N","Ƞ":"N","ǋ":"N","Ñ":"N","Ǌ":"NJ","Ó":"O","Ŏ":"O","Ǒ":"O","Ô":"O","Ố":"O","Ộ":"O","Ồ":"O","Ổ":"O","Ỗ":"O","Ö":"O","Ȫ":"O","Ȯ":"O","Ȱ":"O","Ọ":"O","Ő":"O","Ȍ":"O","Ò":"O","Ỏ":"O","Ơ":"O","Ớ":"O","Ợ":"O","Ờ":"O","Ở":"O","Ỡ":"O","Ȏ":"O","Ꝋ":"O","Ꝍ":"O","Ō":"O","Ṓ":"O","Ṑ":"O","Ɵ":"O","Ǫ":"O","Ǭ":"O","Ø":"O","Ǿ":"O","Õ":"O","Ṍ":"O","Ṏ":"O","Ȭ":"O","Ƣ":"OI","Ꝏ":"OO","Ɛ":"E","Ɔ":"O","Ȣ":"OU","Ṕ":"P","Ṗ":"P","Ꝓ":"P","Ƥ":"P","Ꝕ":"P"," ":"P","Ꝑ":"P","Ꝙ":"Q","Ꝗ":"Q","Ŕ":"R","Ř":"R","Ŗ":"R","Ṙ":"R","Ṛ":"R","Ṝ":"R","Ȑ":"R","Á":"A","Ă":"A","Ắ":"A","Ặ":"A","Ằ":"A","Ẳ":"A","Ẵ":"A","Ǎ":"A","Â":"A","Ấ":"A","Ậ":"A","Ầ":"A","Ẩ":"A","Ẫ":"A","Ä":"A","Ǟ":"A","Ȧ":"A","Ǡ":"A","Ạ":"A","Ȁ":"A","À":"A","Ả":"A","Ȃ":"A","Ā":"A","Ą":"A","Å":"A","Ǻ":"A","Ḁ":"A","Ⱥ":"A","Ã":"A","Ꜳ":"AA","Æ":"AE","Ǽ":"AE","Ǣ":"AE","Ꜵ":"AO","Ꜷ":"AU","Ꜹ":"AV","Ꜻ":"AV","Ꜽ":"AY","Ḃ":"B","Ḅ":"B","Ɓ":"B","Ḇ":"B","Ƀ":"B","Ƃ":"B","Ć":"C","Č":"C","Ç":"C","Ḉ":"C","Ĉ":"C","Ċ":"C","Ƈ":"C","Ȼ":"C","Ď":"D","Ḑ":"D","Ḓ":"D","Ḋ":"D","Ḍ":"D","Ɗ":"D","Ḏ":"D","ǲ":"D","ǅ":"D","Đ":"D","Ƌ":"D","Ǳ":"DZ","Ǆ":"DZ","É":"E","Ĕ":"E","Ě":"E","Ȩ":"E","Ḝ":"E","Ê":"E","Ế":"E","Ệ":"E","Ề":"E","Ể":"E","Ễ":"E","Ḙ":"E","Ë":"E","Ė":"E","Ẹ":"E","Ȅ":"E","È":"E","Ẻ":"E","Ȇ":"E","Ē":"E","Ḗ":"E","Ḕ":"E","Ę":"E","Ɇ":"E","Ẽ":"E","Ḛ":"E","Ꝫ":"ET","Ḟ":"F","Ƒ":"F","Ǵ":"G","Ğ":"G","Ǧ":"G","Ģ":"G","Ĝ":"G","Ġ":"G","Ɠ":"G","Ḡ":"G","Ǥ":"G","Ḫ":"H","Ȟ":"H","Ḩ":"H","Ĥ":"H","Ⱨ":"H","Ḧ":"H","Ḣ":"H","Ḥ":"H","Ħ":"H","Í":"I","Ĭ":"I","Ǐ":"I","Î":"I","Ï":"I","Ḯ":"I","İ":"I","Ị":"I","Ȉ":"I","Ì":"I","Ỉ":"I","Ȋ":"I","Ī":"I","Į":"I","Ɨ":"I","Ĩ":"I","Ḭ":"I","Ꝺ":"D","Ꝼ":"F","Ᵹ":"G","Ꞃ":"R","Ꞅ":"S","Ꞇ":"T","Ꝭ":"IS","Ĵ":"J","Ɉ":"J","Ḱ":"K","Ǩ":"K","Ķ":"K","Ⱪ":"K","Ꝃ":"K","Ḳ":"K","Ƙ":"K","Ḵ":"K","Ꝁ":"K","Ꝅ":"K","Ĺ":"L","Ƚ":"L","Ľ":"L","Ļ":"L","Ḽ":"L","Ḷ":"L","Ḹ":"L"," ":"L","Ꝉ":"L","Ḻ":"L","Ŀ":"L","Ɫ":"L","ǈ":"L","Ł":"L","Ǉ":"LJ","Ḿ":"M","Ṁ":"M","Ṃ":"M","Ɱ":"M","Ń":"N","Ň":"N","Ņ":"N","Ṋ":"N","Ṅ":"N","Ṇ":"N","Ǹ":"N","Ɲ":"N","Ṉ":"N","Ƞ":"N","ǋ":"N","Ñ":"N","Ǌ":"NJ","Ó":"O","Ŏ":"O","Ǒ":"O","Ô":"O","Ố":"O","Ộ":"O","Ồ":"O","Ổ":"O","Ỗ":"O","Ö":"O","Ȫ":"O","Ȯ":"O","Ȱ":"O","Ọ":"O","Ő":"O","Ȍ":"O","Ò":"O","Ỏ":"O","Ơ":"O","Ớ":"O","Ợ":"O","Ờ":"O","Ở":"O","Ỡ":"O","Ȏ":"O","Ꝋ":"O","Ꝍ":"O","Ō":"O","Ṓ":"O","Ṑ":"O","Ɵ":"O","Ǫ":"O","Ǭ":"O","Ø":"O","Ǿ":"O","Õ":"O","Ṍ":"O","Ṏ":"O","Ȭ":"O","Ƣ":"OI","Ꝏ":"OO","Ɛ":"E","Ɔ":"O","Ȣ":"OU","Ṕ":"P","Ṗ":"P","Ꝓ":"P","Ƥ":"P","Ꝕ":"P"," ":"P","Ꝑ":"P","Ꝙ":"Q","Ꝗ":"Q","Ŕ":"R","Ř":"R","Ŗ":"R","Ṙ":"R","Ṛ":"R","Ṝ":"R","Ȑ":"R"};
  
  try {
    const rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    rangos.forEach(r => {
      const matriz = r.getValues().map(fila => fila.map(celda => {
        return typeof celda == 'string' ? celda.replace(/[^A-Za-z0-9\[\] ]/g, a => latin_map[a] || a) : celda;
      }));
      r.setValues(matriz);
    });
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al procesar las celdas.\n\n⚠️ ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}