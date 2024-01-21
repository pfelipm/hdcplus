/**
 * HdC+
 * Una colección de pequeñas utilidades para la edición de hojas de cálculo
 * Copyright (C) Pablo Felip (@pfelipm). Se distribuye bajo licencia GNU GPL v3.
 *
 * @OnlyCurrentDoc
 */

const VERSION = 'Versión: 1.8 (enero 2024)';

// Para mostrar / ocultar pestañas por color
const COLORES_HOJAS = {
  naranja: { nombre: 'naranja', hex:'#ff9900' },
  azul: { nombre: 'azul', hex:'#0000ff' }
};

const urlAyudaFxPersonalizadas = 'https://bit.ly/funciones-hdcplus';

function onInstall(e) {
  
  // Otras cosas que se deben hacer siempre
  
  onOpen(e);
  
}

function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addSubMenu(ui.createMenu('🔀 Barajar datos')
      .addItem('Barajar por columnas', 'desordenarFil')
      .addItem('Barajar por filas', 'desordenarCol'))
    .addSubMenu(ui.createMenu('✅ Casillas de verificación')
      .addItem('✔️️ Activar seleccionadas', 'check')
      .addItem('❌ Desactivar seleccionadas ', 'uncheck')
      .addItem('➖ Invertir seleccionadas ', 'recheck'))
    .addSubMenu(ui.createMenu('🧮 Estructura datos')
      .addItem('Consolidar dimensiones (despivotar)', 'unpivot')
      .addItem('Transponer (☢️ destructivo)', 'transponer'))
    .addSubMenu(ui.createMenu('📐 Estructura hoja de cálculo')
      .addItem('Eliminar celdas no seleccionadas', 'eliminarFyCNoSeleccionadas')
      .addItem('Eliminar F/C sobrantes', 'eliminarFyC')
      .addSeparator()
      .addItem('Insertar F/C nuevas', 'insertarFyC'))
    .addSubMenu(ui.createMenu('👁️‍🗨️ Gestión hojas')
      .addItem('Ocultar resto de hojas', 'ocultarHojas')
      .addItem('Mostrar hojas ocultas', 'mostrarHojas')
      .addItem('Mostrar todas excepto activa', 'mostrarTodasMenosActual')
      .addSeparator()
      .addItem('🔹Mostrar hojas color azul', 'mostrarHojasAzul')        
      .addItem('🔹Ocultar hojas color azul', 'ocultarHojasAzul')        
      .addItem('🔸Mostrar hojas color naranja', 'mostrarHojasNaranja')        
      .addItem('🔸Ocultar hojas color naranja', 'ocultarHojasNaranja')
      .addSeparator()
      .addItem('Eliminar hojas ocultas', 'eliminarHojasOcultas')
      .addItem('Eliminar todas excepto activa', 'eliminarHojas'))
    .addSubMenu(ui.createMenu('🧠 Generar')
      .addItem('NIFs', 'generarNIF')
      .addItem('Nombres y apellidos', 'generarNombres'))
    .addSubMenu(ui.createMenu('🕶️ Ofuscar')
      .addItem('Codificar texto en base64 ', 'base64')
      .addItem('Sustituir por hash MD2 (b64) ', 'hashMD2')
      .addItem('Sustituir por hash MD5 (b64) ', 'hashMD5')
      .addItem('Sustituir por hash SHA-1 (b64)', 'hashSHA1')
      .addItem('Sustituir por hash SHA-256 (b64)', 'hashSHA256')
      .addItem('Sustituir por hash SHA-384 (b64)', 'hashSHA384')
      .addItem('Sustituir por hash SHA-512 (b64)', 'hashSHA512'))
    .addSubMenu(ui.createMenu('⚡ Transformar')
      .addItem('❌👽 Eliminar caracteres especiales', 'latinizar')
      .addItem('❌➖ Eliminar espacios', 'eliminarEspacios')
      .addItem('❌🔺➖ Eliminar espacios tras comas', 'eliminarEspaciosComas')
      .addItem('❌↩️ Eliminar saltos de línea ', 'eliminarSaltos')
      .addItem('🔺➖️ Comas a espacios', 'comasEspacios')
      .addItem('🔺↩️ Comas a saltos de línea', 'comasSaltos')
      .addItem('➖🔺 Espacios a comas', 'espaciosComas')      
      .addItem('➖↩️ Espacios a saltos de línea', 'espaciosSaltos')      
      .addItem('↩️➖ Saltos de línea a espacios', 'saltosEspacios')
      .addItem('↩️🔺 Saltos de línea a comas', 'saltosComas')
      .addItem('↗️ Iniciales a mayúsculas', 'inicialesMays_')
      .addItem('⬆️ Inicial a mayúsculas', 'inicialMays_')
      .addItem('🔠 Todo a mayúsculas', 'mayusculas')
      .addItem('🔤 Todo a minúsculas', 'minusculas'))
    .addSeparator()
    .addItem('🔄 Forzar recálculo de hoja', 'forzarRecalculo')
    .addSeparator()
    .addItem('🛟 Ayuda fx personalizadas (sitio externo)', 'abrirWebExterna')
    .addSeparator()
    .addItem('💡 Acerca de HdC+', 'acercaDe')
    .addToUi();
}

// Abre el cuadro de diálogo de información de HdC+
function acercaDe() {

  // Presentación del complemento
  var panel = HtmlService.createTemplateFromFile('acercaDe');
  panel.version = VERSION;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(350), '💡 ¿Qué es HdC+?');
}

// Abre la web externa de documentación de las funciones personalizadas de HdC+
// https://www.labnol.org/open-webpage-google-sheets-220507
function abrirWebExterna() {

  const htmlTemplate = HtmlService.createTemplateFromFile('ayudaFxPersonalizadas.html');
  htmlTemplate.url = urlAyudaFxPersonalizadas;
  SpreadsheetApp.getUi().showModelessDialog(
    htmlTemplate.evaluate().setHeight(20).setWidth(300),
    '🌐 Abriendo sitio web...'
  );
  // No parece ser necesario
  Utilities.sleep(2000);

};

// Funciones de transformación del texto

function minusculas()  {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return typeof c == 'string' ? c.toLowerCase() : c;
      })
    
    })
    r.setValues(matriz);
  })
    
}

function mayusculas() {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return typeof c == 'string' ? c.toUpperCase() : c;

      })
    
    })
    r.setValues(matriz);
  })
    
}

/**
 * Pasa a mayúsculas la letra inicial de las celdas de los intervalos
 * seleccionados. Tiene en cuenta los caracteres especiales del castellano (vocales
 * con tilda, ñ, etc.) y también el resto de signos de puntuación y
 * caracteres especiales (que no son tratados).
 */
function inicialMays_() {
  
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
  
  } catch(e) { throw `HdC+ dice: Error inesperado, contacta con soporte // ${e}`; }

}

/**
 * Capitaliza las celdas de los intervalos seleccionados, pasando
 * las iniciales de cada palabra a mayúsculas y el resto a minúsculas.
 * Tiene en cuenta los caracteres especiales del castellano (vocales
 * con tilda, ñ, etc.) y también el resto de signos de puntuación y
 * caracteres especiales (que no son tratados). Pretende ser un clon
 * de la función integrada NOMPROPIO().
 * 03/03/23
 */
function inicialesMays_() {

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
  
  } catch(e) { throw `HdC+ dice: Error inesperado, contacta con soporte // ${e}`; }

}

// Función genérica que sustituye las coincidencias de la expresión regular
// exp por caracter en todas las celdas de los rangos seleccionados

function sustituir(exp, caracter) {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return c.replace(exp, caracter);
      })
    
    })
    r.setValues(matriz);
  })
    
}

function eliminarEspacios() {sustituir(/\s+/g, '');}
function eliminarEspaciosComas() {sustituir(/,\s+/g, ',');}
function eliminarSaltos() {sustituir(/\n/g, '');}
function comasEspacios() {sustituir(/,/g, ' ');}
function comasSaltos() {sustituir(/,/g, '\n');}
function espaciosComas() {sustituir(/\s/g, ',');}
function espaciosSaltos() {sustituir(/\s/g, '\n');}
function saltosEspacios() {sustituir(/\n/g, ' ');}
function saltosComas() {sustituir(/\n/g, ',');}


function latinizar() {

  var latinise={};latinise.latin_map={"·":"-","ß":"ss","Á":"A","Ă":"A","Ắ":"A","Ặ":"A","Ằ":"A","Ẳ":"A","Ẵ":"A","Ǎ":"A","Â":"A","Ấ":"A","Ậ":"A","Ầ":"A","Ẩ":"A","Ẫ":"A","Ä":"A","Ǟ":"A","Ȧ":"A","Ǡ":"A","Ạ":"A","Ȁ":"A","À":"A","Ả":"A","Ȃ":"A","Ā":"A","Ą":"A","Å":"A","Ǻ":"A","Ḁ":"A","Ⱥ":"A","Ã":"A","Ꜳ":"AA","Æ":"AE","Ǽ":"AE","Ǣ":"AE","Ꜵ":"AO","Ꜷ":"AU","Ꜹ":"AV","Ꜻ":"AV","Ꜽ":"AY","Ḃ":"B","Ḅ":"B","Ɓ":"B","Ḇ":"B","Ƀ":"B","Ƃ":"B","Ć":"C","Č":"C","Ç":"C","Ḉ":"C","Ĉ":"C","Ċ":"C","Ƈ":"C","Ȼ":"C","Ď":"D","Ḑ":"D","Ḓ":"D","Ḋ":"D","Ḍ":"D","Ɗ":"D","Ḏ":"D","ǲ":"D","ǅ":"D","Đ":"D","Ƌ":"D","Ǳ":"DZ","Ǆ":"DZ","É":"E","Ĕ":"E","Ě":"E","Ȩ":"E","Ḝ":"E","Ê":"E","Ế":"E","Ệ":"E","Ề":"E","Ể":"E","Ễ":"E","Ḙ":"E","Ë":"E","Ė":"E","Ẹ":"E","Ȅ":"E","È":"E","Ẻ":"E","Ȇ":"E","Ē":"E","Ḗ":"E","Ḕ":"E","Ę":"E","Ɇ":"E","Ẽ":"E","Ḛ":"E","Ꝫ":"ET","Ḟ":"F","Ƒ":"F","Ǵ":"G","Ğ":"G","Ǧ":"G","Ģ":"G","Ĝ":"G","Ġ":"G","Ɠ":"G","Ḡ":"G","Ǥ":"G","Ḫ":"H","Ȟ":"H","Ḩ":"H","Ĥ":"H","Ⱨ":"H","Ḧ":"H","Ḣ":"H","Ḥ":"H","Ħ":"H","Í":"I","Ĭ":"I","Ǐ":"I","Î":"I","Ï":"I","Ḯ":"I","İ":"I","Ị":"I","Ȉ":"I","Ì":"I","Ỉ":"I","Ȋ":"I","Ī":"I","Į":"I","Ɨ":"I","Ĩ":"I","Ḭ":"I","Ꝺ":"D","Ꝼ":"F","Ᵹ":"G","Ꞃ":"R","Ꞅ":"S","Ꞇ":"T","Ꝭ":"IS","Ĵ":"J","Ɉ":"J","Ḱ":"K","Ǩ":"K","Ķ":"K","Ⱪ":"K","Ꝃ":"K","Ḳ":"K","Ƙ":"K","Ḵ":"K","Ꝁ":"K","Ꝅ":"K","Ĺ":"L","Ƚ":"L","Ľ":"L","Ļ":"L","Ḽ":"L","Ḷ":"L","Ḹ":"L","Ⱡ":"L","Ꝉ":"L","Ḻ":"L","Ŀ":"L","Ɫ":"L","ǈ":"L","Ł":"L","Ǉ":"LJ","Ḿ":"M","Ṁ":"M","Ṃ":"M","Ɱ":"M","Ń":"N","Ň":"N","Ņ":"N","Ṋ":"N","Ṅ":"N","Ṇ":"N","Ǹ":"N","Ɲ":"N","Ṉ":"N","Ƞ":"N","ǋ":"N","Ñ":"N","Ǌ":"NJ","Ó":"O","Ŏ":"O","Ǒ":"O","Ô":"O","Ố":"O","Ộ":"O","Ồ":"O","Ổ":"O","Ỗ":"O","Ö":"O","Ȫ":"O","Ȯ":"O","Ȱ":"O","Ọ":"O","Ő":"O","Ȍ":"O","Ò":"O","Ỏ":"O","Ơ":"O","Ớ":"O","Ợ":"O","Ờ":"O","Ở":"O","Ỡ":"O","Ȏ":"O","Ꝋ":"O","Ꝍ":"O","Ō":"O","Ṓ":"O","Ṑ":"O","Ɵ":"O","Ǫ":"O","Ǭ":"O","Ø":"O","Ǿ":"O","Õ":"O","Ṍ":"O","Ṏ":"O","Ȭ":"O","Ƣ":"OI","Ꝏ":"OO","Ɛ":"E","Ɔ":"O","Ȣ":"OU","Ṕ":"P","Ṗ":"P","Ꝓ":"P","Ƥ":"P","Ꝕ":"P","Ᵽ":"P","Ꝑ":"P","Ꝙ":"Q","Ꝗ":"Q","Ŕ":"R","Ř":"R","Ŗ":"R","Ṙ":"R","Ṛ":"R","Ṝ":"R","Ȑ":"R","Ȓ":"R","Ṟ":"R","Ɍ":"R","Ɽ":"R","Ꜿ":"C","Ǝ":"E","Ś":"S","Ṥ":"S","Š":"S","Ṧ":"S","Ş":"S","Ŝ":"S","Ș":"S","Ṡ":"S","Ṣ":"S","Ṩ":"S","Ť":"T","Ţ":"T","Ṱ":"T","Ț":"T","Ⱦ":"T","Ṫ":"T","Ṭ":"T","Ƭ":"T","Ṯ":"T","Ʈ":"T","Ŧ":"T","Ɐ":"A","Ꞁ":"L","Ɯ":"M","Ʌ":"V","Ꜩ":"TZ","Ú":"U","Ŭ":"U","Ǔ":"U","Û":"U","Ṷ":"U","Ü":"U","Ǘ":"U","Ǚ":"U","Ǜ":"U","Ǖ":"U","Ṳ":"U","Ụ":"U","Ű":"U","Ȕ":"U","Ù":"U","Ủ":"U","Ư":"U","Ứ":"U","Ự":"U","Ừ":"U","Ử":"U","Ữ":"U","Ȗ":"U","Ū":"U","Ṻ":"U","Ų":"U","Ů":"U","Ũ":"U","Ṹ":"U","Ṵ":"U","Ꝟ":"V","Ṿ":"V","Ʋ":"V","Ṽ":"V","Ꝡ":"VY","Ẃ":"W","Ŵ":"W","Ẅ":"W","Ẇ":"W","Ẉ":"W","Ẁ":"W","Ⱳ":"W","Ẍ":"X","Ẋ":"X","Ý":"Y","Ŷ":"Y","Ÿ":"Y","Ẏ":"Y","Ỵ":"Y","Ỳ":"Y","Ƴ":"Y","Ỷ":"Y","Ỿ":"Y","Ȳ":"Y","Ɏ":"Y","Ỹ":"Y","Ź":"Z","Ž":"Z","Ẑ":"Z","Ⱬ":"Z","Ż":"Z","Ẓ":"Z","Ȥ":"Z","Ẕ":"Z","Ƶ":"Z","Ĳ":"IJ","Œ":"OE","ᴀ":"A","ᴁ":"AE","ʙ":"B","ᴃ":"B","ᴄ":"C","ᴅ":"D","ᴇ":"E","ꜰ":"F","ɢ":"G","ʛ":"G","ʜ":"H","ɪ":"I","ʁ":"R","ᴊ":"J","ᴋ":"K","ʟ":"L","ᴌ":"L","ᴍ":"M","ɴ":"N","ᴏ":"O","ɶ":"OE","ᴐ":"O","ᴕ":"OU","ᴘ":"P","ʀ":"R","ᴎ":"N","ᴙ":"R","ꜱ":"S","ᴛ":"T","ⱻ":"E","ᴚ":"R","ᴜ":"U","ᴠ":"V","ᴡ":"W","ʏ":"Y","ᴢ":"Z","á":"a","ă":"a","ắ":"a","ặ":"a","ằ":"a","ẳ":"a","ẵ":"a","ǎ":"a","â":"a","ấ":"a","ậ":"a","ầ":"a","ẩ":"a","ẫ":"a","ä":"a","ǟ":"a","ȧ":"a","ǡ":"a","ạ":"a","ȁ":"a","à":"a","ả":"a","ȃ":"a","ā":"a","ą":"a","ᶏ":"a","ẚ":"a","å":"a","ǻ":"a","ḁ":"a","ⱥ":"a","ã":"a","ꜳ":"aa","æ":"ae","ǽ":"ae","ǣ":"ae","ꜵ":"ao","ꜷ":"au","ꜹ":"av","ꜻ":"av","ꜽ":"ay","ḃ":"b","ḅ":"b","ɓ":"b","ḇ":"b","ᵬ":"b","ᶀ":"b","ƀ":"b","ƃ":"b","ɵ":"o","ć":"c","č":"c","ç":"c","ḉ":"c","ĉ":"c","ɕ":"c","ċ":"c","ƈ":"c","ȼ":"c","ď":"d","ḑ":"d","ḓ":"d","ȡ":"d","ḋ":"d","ḍ":"d","ɗ":"d","ᶑ":"d","ḏ":"d","ᵭ":"d","ᶁ":"d","đ":"d","ɖ":"d","ƌ":"d","ı":"i","ȷ":"j","ɟ":"j","ʄ":"j","ǳ":"dz","ǆ":"dz","é":"e","ĕ":"e","ě":"e","ȩ":"e","ḝ":"e","ê":"e","ế":"e","ệ":"e","ề":"e","ể":"e","ễ":"e","ḙ":"e","ë":"e","ė":"e","ẹ":"e","ȅ":"e","è":"e","ẻ":"e","ȇ":"e","ē":"e","ḗ":"e","ḕ":"e","ⱸ":"e","ę":"e","ᶒ":"e","ɇ":"e","ẽ":"e","ḛ":"e","ꝫ":"et","ḟ":"f","ƒ":"f","ᵮ":"f","ᶂ":"f","ǵ":"g","ğ":"g","ǧ":"g","ģ":"g","ĝ":"g","ġ":"g","ɠ":"g","ḡ":"g","ᶃ":"g","ǥ":"g","ḫ":"h","ȟ":"h","ḩ":"h","ĥ":"h","ⱨ":"h","ḧ":"h","ḣ":"h","ḥ":"h","ɦ":"h","ẖ":"h","ħ":"h","ƕ":"hv","í":"i","ĭ":"i","ǐ":"i","î":"i","ï":"i","ḯ":"i","ị":"i","ȉ":"i","ì":"i","ỉ":"i","ȋ":"i","ī":"i","į":"i","ᶖ":"i","ɨ":"i","ĩ":"i","ḭ":"i","ꝺ":"d","ꝼ":"f","ᵹ":"g","ꞃ":"r","ꞅ":"s","ꞇ":"t","ꝭ":"is","ǰ":"j","ĵ":"j","ʝ":"j","ɉ":"j","ḱ":"k","ǩ":"k","ķ":"k","ⱪ":"k","ꝃ":"k","ḳ":"k","ƙ":"k","ḵ":"k","ᶄ":"k","ꝁ":"k","ꝅ":"k","ĺ":"l","ƚ":"l","ɬ":"l","ľ":"l","ļ":"l","ḽ":"l","ȴ":"l","ḷ":"l","ḹ":"l","ⱡ":"l","ꝉ":"l","ḻ":"l","ŀ":"l","ɫ":"l","ᶅ":"l","ɭ":"l","ł":"l","ǉ":"lj","ſ":"s","ẜ":"s","ẛ":"s","ẝ":"s","ḿ":"m","ṁ":"m","ṃ":"m","ɱ":"m","ᵯ":"m","ᶆ":"m","ń":"n","ň":"n","ņ":"n","ṋ":"n","ȵ":"n","ṅ":"n","ṇ":"n","ǹ":"n","ɲ":"n","ṉ":"n","ƞ":"n","ᵰ":"n","ᶇ":"n","ɳ":"n","ñ":"n","ǌ":"nj","ó":"o","ŏ":"o","ǒ":"o","ô":"o","ố":"o","ộ":"o","ồ":"o","ổ":"o","ỗ":"o","ö":"o","ȫ":"o","ȯ":"o","ȱ":"o","ọ":"o","ő":"o","ȍ":"o","ò":"o","ỏ":"o","ơ":"o","ớ":"o","ợ":"o","ờ":"o","ở":"o","ỡ":"o","ȏ":"o","ꝋ":"o","ꝍ":"o","ⱺ":"o","ō":"o","ṓ":"o","ṑ":"o","ǫ":"o","ǭ":"o","ø":"o","ǿ":"o","õ":"o","ṍ":"o","ṏ":"o","ȭ":"o","ƣ":"oi","ꝏ":"oo","ɛ":"e","ᶓ":"e","ɔ":"o","ᶗ":"o","ȣ":"ou","ṕ":"p","ṗ":"p","ꝓ":"p","ƥ":"p","ᵱ":"p","ᶈ":"p","ꝕ":"p","ᵽ":"p","ꝑ":"p","ꝙ":"q","ʠ":"q","ɋ":"q","ꝗ":"q","ŕ":"r","ř":"r","ŗ":"r","ṙ":"r","ṛ":"r","ṝ":"r","ȑ":"r","ɾ":"r","ᵳ":"r","ȓ":"r","ṟ":"r","ɼ":"r","ᵲ":"r","ᶉ":"r","ɍ":"r","ɽ":"r","ↄ":"c","ꜿ":"c","ɘ":"e","ɿ":"r","ś":"s","ṥ":"s","š":"s","ṧ":"s","ş":"s","ŝ":"s","ș":"s","ṡ":"s","ṣ":"s","ṩ":"s","ʂ":"s","ᵴ":"s","ᶊ":"s","ȿ":"s","ɡ":"g","ᴑ":"o","ᴓ":"o","ᴝ":"u","ť":"t","ţ":"t","ṱ":"t","ț":"t","ȶ":"t","ẗ":"t","ⱦ":"t","ṫ":"t","ṭ":"t","ƭ":"t","ṯ":"t","ᵵ":"t","ƫ":"t","ʈ":"t","ŧ":"t","ᵺ":"th","ɐ":"a","ᴂ":"ae","ǝ":"e","ᵷ":"g","ɥ":"h","ʮ":"h","ʯ":"h","ᴉ":"i","ʞ":"k","ꞁ":"l","ɯ":"m","ɰ":"m","ᴔ":"oe","ɹ":"r","ɻ":"r","ɺ":"r","ⱹ":"r","ʇ":"t","ʌ":"v","ʍ":"w","ʎ":"y","ꜩ":"tz","ú":"u","ŭ":"u","ǔ":"u","û":"u","ṷ":"u","ü":"u","ǘ":"u","ǚ":"u","ǜ":"u","ǖ":"u","ṳ":"u","ụ":"u","ű":"u","ȕ":"u","ù":"u","ủ":"u","ư":"u","ứ":"u","ự":"u","ừ":"u","ử":"u","ữ":"u","ȗ":"u","ū":"u","ṻ":"u","ų":"u","ᶙ":"u","ů":"u","ũ":"u","ṹ":"u","ṵ":"u","ᵫ":"ue","ꝸ":"um","ⱴ":"v","ꝟ":"v","ṿ":"v","ʋ":"v","ᶌ":"v","ⱱ":"v","ṽ":"v","ꝡ":"vy","ẃ":"w","ŵ":"w","ẅ":"w","ẇ":"w","ẉ":"w","ẁ":"w","ⱳ":"w","ẘ":"w","ẍ":"x","ẋ":"x","ᶍ":"x","ý":"y","ŷ":"y","ÿ":"y","ẏ":"y","ỵ":"y","ỳ":"y","ƴ":"y","ỷ":"y","ỿ":"y","ȳ":"y","ẙ":"y","ɏ":"y","ỹ":"y","ź":"z","ž":"z","ẑ":"z","ʑ":"z","ⱬ":"z","ż":"z","ẓ":"z","ȥ":"z","ẕ":"z","ᵶ":"z","ᶎ":"z","ʐ":"z","ƶ":"z","ɀ":"z","ﬀ":"ff","ﬃ":"ffi","ﬄ":"ffl","ﬁ":"fi","ﬂ":"fl","ĳ":"ij","œ":"oe","ﬆ":"st","ₐ":"a","ₑ":"e","ᵢ":"i","ⱼ":"j","ₒ":"o","ᵣ":"r","ᵤ":"u","ᵥ":"v","ₓ":"x"};
  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return typeof c == 'string' ? c.replace(/[^A-Za-z0-9\[\] ]/g,function(a){return latinise.latin_map[a]||a}) : c;
      })
    
    })
    r.setValues(matriz);
  })
    
}

function unpivot() {

  var ui = HtmlService.createTemplateFromFile('panelUnpivot')
      .evaluate()
      .setTitle('Consolidar dimensiones');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function unpivot_capturarRango(modo) {

  var hdc = SpreadsheetApp.getActiveSheet();

  if (modo == 'rango') {
    var rango = hdc.getActiveRange();
    if (rango.getNumColumns() >= 2 && rango.getNumRows() >= 2) {
      return hdc.getName() + '!' + rango.getA1Notation();
    }
    else {
      return '¡Selecciona más celdas!';
    }
  }
  else if (modo == 'celda') {
    return hdc.getName() + '!' + hdc.getActiveCell().getA1Notation();
  }
  
}

function unpivot_core(e) {

  var hdc = SpreadsheetApp.getActiveSpreadsheet();
  
  // Recibir parámetros panel unpivot
  
  var rango = hdc.getRange(e.rango);
  var destino = hdc.getRange(e.destino); 
  var colFijas = +e.numCol;
  var nEncabezados = colFijas + 2;
  var encabezados = e.encabezados.split('\n'); // Dividir cadena por saltos de línea
  var nEncabezadosSplit = encabezados.length;
  var matrizConsolidada = [];
  var encabezadosUnpivot = [];
  var i;
    
  // Todo en bloque try..catch para cazar problemas
  
  try {
    
    // Construir vector de encabezados, primero tomamos los introducidos por el usuario
    // Se trata el caso de que el usuario haya introducido intros de más en caja de texto
    
    for (i = 0; i < nEncabezadosSplit && i < nEncabezados; i++) {
      if (encabezados[i].length == 0) {
        encabezadosUnpivot.push('Columna ' + (i + 1));
      }
      else {
        encabezadosUnpivot.push(encabezados[i]);
      }
    }
    
    // Si faltan, completamos con "Columna n"
    
    for (i = nEncabezadosSplit; i < nEncabezados; i++) {
      encabezadosUnpivot.push('Columna ' + (i + 1));
    }
    
    // UNPIVOT espera un vector fila con los encabezados
    
    var encabezadosUnpivotH = [];
    encabezadosUnpivotH.push(encabezadosUnpivot);
    
    // Aplicar UNPIVOT
    
    matrizConsolidada = UNPIVOT(rango.getValues(), colFijas, encabezadosUnpivotH);
   
    // Expandir rango destino para recibir tabla consolidada y rellenar celdas
    
    destino.offset(0, 0, matrizConsolidada.length, matrizConsolidada[0].length).setValues(matrizConsolidada);

  }

  catch(e) {
  
      SpreadsheetApp.getUi().alert(e);
      return (e);
      
  }
  
}

function insertarFyC() {

  var ui = HtmlService.createTemplateFromFile('panelCrearFyC')
      .evaluate()
      .setTitle('Insertar filas y/o columnas');
  SpreadsheetApp.getUi().showSidebar(ui);
  
}

function insertarFyC_core(e) {
  
  var nFil = +e.numFil;
  var nCol = +e.numCol;
  var modo = +e.modo;
  var hdc = SpreadsheetApp.getActiveSheet();
  
  // Fila y Columna posición superior izquierda rango seleccionado
  
  var fil = hdc.getActiveCell().getRow();
  var col = hdc.getActiveCell().getColumn();
  
  // Fila y Columna máximas
  
  var filMax = hdc.getMaxRows();
  var colMax = hdc.getMaxColumns();
    
  switch (modo) {
  
    case 1: // al principio de la hoja
      
      if (nFil > 0) {hdc.insertRowsBefore(1, nFil);}
      if (nCol > 0) {hdc.insertColumnsBefore(1, nCol);}
      break;
    
    case 2: // al final de la hoja
    
      if (nFil > 0) {hdc.insertRowsAfter(filMax, nFil);}
      if (nCol > 0) {hdc.insertColumnsAfter(colMax, nCol);}
      break;
    
    case 3: // antes de celda
      
      if (nFil > 0) {hdc.insertRowsBefore(fil, nFil);}
      if (nCol > 0) {hdc.insertColumnsBefore(col, nCol);}
      break;
      
    case 4: // después de celda
      if (nFil > 0) {hdc.insertRowsAfter(fil, nFil);}
      if (nCol > 0) {hdc.insertColumnsAfter(col, nCol);}
      break;    
  }
}

function eliminarFyC() {

  var ui = HtmlService.createTemplateFromFile('panelEliminarFyC')
      .evaluate()
      .setTitle('Eliminar filas y/o columnas sobrantes');
  SpreadsheetApp.getUi().showSidebar(ui);
  
}

function eliminarFyC_core(e) {
  
  var modo = +e.modo;
  var global = e.global == 'on' ? true : false;
    
  var nFilas, nMaxFilas, nColumnas, nMaxColumnas;
  var hojas = global ? SpreadsheetApp.getActiveSpreadsheet().getSheets() : SpreadsheetApp.getActiveSheet();
    
  if (global) { // se procesan todas las hojas de la hdc
  
    hojas.map(function(h) {
      
      nFilas = h.getLastRow();
      nMaxFilas = h.getMaxRows();
      nColumnas = h.getLastColumn();
      nMaxColumnas = h.getMaxColumns();
      if (nFilas > 0 && nColumnas > 0) {
        
        switch (modo) {
        
          case 1: // Columnas
          
            if (nMaxColumnas > nColumnas) {h.deleteColumns(nColumnas+1, nMaxColumnas - nColumnas);}
            break;
          
          case 2: // Filas
      
            if (nMaxFilas > nFilas) {h.deleteRows(nFilas+1, nMaxFilas - nFilas);}
            break;   
            
          case 3: // Filas y columnas 
          
            if (nMaxColumnas > nColumnas) {h.deleteColumns(nColumnas+1, nMaxColumnas - nColumnas);}
            if (nMaxFilas > nFilas) {h.deleteRows(nFilas+1, nMaxFilas - nFilas);}          
            break;
        }      
      }
  })}
  else { // solo se procesa la hoja actual
  
    nFilas = hojas.getLastRow();
    nMaxFilas = hojas.getMaxRows();
    nColumnas = hojas.getLastColumn();
    nMaxColumnas = hojas.getMaxColumns();
    if (nFilas == 0 || nColumnas == 0) {
      SpreadsheetApp.getUi().alert('💡 La hoja de datos está vacía, no se eliminará nada.');
    }
    else {
    
      switch (modo) {
      
        case 1: // Columnas
        
          if (nMaxColumnas > nColumnas) {hojas.deleteColumns(nColumnas+1, nMaxColumnas - nColumnas);}
          break;
        
        case 2: // Filas
    
          if (nMaxFilas > nFilas) {hojas.deleteRows(nFilas+1, nMaxFilas - nFilas);}
          break;   
          
        case 3: // Filas y columnas 
        
          if (nMaxColumnas > nColumnas) {hojas.deleteColumns(nColumnas+1, nMaxColumnas - nColumnas);}
          if (nMaxFilas > nFilas) {hojas.deleteRows(nFilas+1, nMaxFilas - nFilas);}          
          break;
      }
    }
  }
  
}

/**
 * Elimina las filas y columnas de la hoja actual que quedan fuera del intervalo (único) de celdas seleccionado
 */
function eliminarFyCNoSeleccionadas() {

  const ui = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSheet();
  const rangosSeleccionados = hoja.getActiveRangeList().getRanges();
  
  console.info(rangosSeleccionados.length);

  if (rangosSeleccionados.length > 1) {
    ui.alert('Selecciona un único intervalo de datos.');
  } else {

    const filSup = rangosSeleccionados[0].getRow();
    const colIzq = rangosSeleccionados[0].getColumn();
    const numFil = rangosSeleccionados[0].getNumRows();
    const numCol = rangosSeleccionados[0].getNumColumns();

    try {

      // Eliminar filas anteriores sobrantes
      if (filSup > 1) hoja.deleteRows(1, filSup - 1);

      // Eliminar filas posteriores sobrantes
      if (numFil < hoja.getMaxRows()) hoja.deleteRows(numFil + 1, hoja.getMaxRows() - numFil);

      // Eliminar columnas a la izquierda sobrantes
      if (colIzq > 1) hoja.deleteColumns(1, colIzq - 1);

      // Eliminar columnas a la derecha sobrantes
      if (numCol < hoja.getMaxColumns()) hoja.deleteColumns(numCol + 1, hoja.getMaxColumns() - numCol);

      // Deseleccionar intervalo (hace activa la celda A1 a piñón fijo)
      hoja.setActiveRange(hoja.getRange('A1'));
    
    } catch (e) {
      ui.alert(`Se ha producido un error inesperado al ajustar eliminar las filas y columnas no seleccionadas.
        
        ⚠️ ${e.message}`, ui.ButtonSet.OK);
    }
  }

}


function base64()  {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return Utilities.base64Encode(c, Utilities.Charset.UTF_8);
      })
    
    })
    r.setValues(matriz);
  })
    
}

function hashGenerico(rango, tipoHash) {

  // Tomado de https://stackoverflow.com/questions/7994410/hash-of-a-cell-text-in-google-spreadsheet

  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
 for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var valor = matriz[f][c];      
      matriz[f][c] = Utilities.base64Encode(Utilities.computeDigest(tipoHash, valor, Utilities.Charset.UTF_8));
    }
  }
  
  rango.setValues(matriz);
  
}

// Funciones auxiliares para hashes y base64

// Convierte hash binario en hexadecimal

function bin2hex(rawHash) {

  var txtHash = '';
  var hashVal;
  
  for (var j = 0; j <rawHash.length; j++) {
  
    hashVal = rawHash[j];
    if (hashVal < 0)
      hashVal += 256; 
    if (hashVal.toString(16).length == 1)
      txtHash += "0";
    txtHash += hashVal.toString(16);
  }      
  return txtHash;

}

// Convierte cadena hexadecimal en vector de bytes

function hex2bytes(hex) {
    for (var bytes = [], c = 0; c < hex.length; c += 2)
    bytes.push(parseInt(hex.substr(c, 2), 16));
    return bytes;
}

// Invocación de hashGenerico() para cada tipo

function hashSHA256() {
  
  SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {
    
    hashGenerico(r, Utilities.DigestAlgorithm.SHA_256);})
  
}

function hashSHA384() {

  SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {
  
    hashGenerico(r, Utilities.DigestAlgorithm.SHA_384);})  
  
}

function hashSHA512() {

SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {

  hashGenerico(r, Utilities.DigestAlgorithm.SHA_512);})  
  
}

function hashSHA1() {

SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {

  hashGenerico(r, Utilities.DigestAlgorithm.SHA_1);})
  
}

function hashMD2() {

SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {

  hashGenerico(r, Utilities.DigestAlgorithm.MD2);})
  
}

function hashMD5() {

SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges().map(function(r) {

  hashGenerico(r, Utilities.DigestAlgorithm.MD5);})
  
}

function transponer() {

  var hojaActual = SpreadsheetApp.getActiveSheet();
  var rango = hojaActual.getActiveRange();
  var matriz = rango.getValues();
  var nCol = matriz[0].length;
  var nFil = matriz.length;
  var filDestino = rango.getRow();
  var colDestino = rango.getColumn();
  var rangoLibre;
  
  // Copiar transponiendo con todo (formato celda / número / condicional, validación, notas...
  
  rango.copyTo(hojaActual.getRange(filDestino, colDestino), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, true)
  
  // Ahora limpiamos el rango que queda libre al transponer
  
  if (nCol != nFil) {
  
      // El rango es rectangular, hay que "limpiar" el espacio libre que deja tras la transposición
    
    if (nCol > nFil) {
      rangoLibre = hojaActual.getRange(filDestino, colDestino + nFil, nFil, nCol - nFil);
    }
    else {  // nCol < nFil
      rangoLibre = hojaActual.getRange(filDestino + nCol, colDestino, nFil - nCol, nCol);  
    }
    
    rangoLibre.clear();
    rangoLibre.clearNote();
    rangoLibre.clearDataValidations();
  }
}

function desordenarFil(){

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
   
  // Comprobaciones previas
 
  if (matriz.length <= 1) {SpreadsheetApp.getUi().alert('💡 La selección debe tener más de 1 fila.');}
  
  var ncol = matriz[0].length;
  var nfil = matriz.length;

  for (var c = 0; c < ncol; c++) {
    for (var f = nfil - 1; f > 0; f--) { 
      var j = Math.floor(Math.random() * (f + 1));
      var temp = matriz[f][c];
      matriz[f][c] = matriz[j][c];
      matriz[j][c] = temp;
    }    
  }
  
  rango.setValues(matriz);
  
}

function desordenarCol(){

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
   
  // Comprobaciones previas
 
  if (matriz[0].length <= 1) {SpreadsheetApp.getUi().alert('💡 La selección debe tener más de 1 columna.');}
  
  var ncol = matriz[0].length;
  var nfil = matriz.length;

  for (var f = 0; f < nfil; f++) {
    for (var c = ncol - 1; c > 0; c--) { 
      var j = Math.floor(Math.random() * (c + 1));
      var temp = matriz[f][c];
      matriz[f][c] = matriz[f][j];
      matriz[f][j] = temp;
    }    
  }
  
  rango.setValues(matriz);
  
}

function generarNIF()  {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  var dni;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
        var dni = '';
        for (var n = 0; n < 8; n++) {
          dni = dni + Math.floor(Math.random()*10).toString();
        }
        dni = dni + "TRWAGMYFPDXBNJZSQVHLCKE".charAt(+dni % 23);
      
        return dni;    
      })
    
    })
    r.setValues(matriz);
  })
    
}  

function generarNombres() {

  var nombresTop = ['Guillermo','José Luis','Nuria','José','David','Victor','Santiago','Francisco José','Marcos','Yolanda',
                'Isabel','César','Carolina','Nerea','Lucas','Alba','Gabriel','Salvador','Ignacio','Manuela','María Carmen',
                'Domingo','Rosario','Albert','María José','Samuel','María Ángeles','Juan Antonio','Laura','Xavier','Lidia',
                'Fátima','Paula','José Carlos','Joaquin','Marina','Alejandro','Juan José','Alicia','José Manuel','Pilar',
                'Luisa','Encarnación','Alfredo','Felipe','Luis Miguel','Sergio','Antonia','Juan','José Ignacio','Beatriz',
                'Alvaro','Angel','Ruben','Lucía','Luis','Rocio','Celia','Juana','Rafael','Irene','Tomas','Joan','José Maria',
                'Cristina','Aitor','Josep','Francisco','Rosa','María Mar','Veronica','Ana','Carla','María Concepción','Vicente',
                'Fernando','Ramon','Daniel','Alex','Susana','Sofia','Alejandra','Felix','Adrian','Jorge','Silvia','Gloria','Josefa',
                'Ana Isabel','Inmaculada','Miriam','Catalina','Hugo','José Francisco','Javier','Alfonso','Aurora','Ainhoa',
                'Juan Manuel','Sandra','Concepción','Manuel','Amparo','Daniela','María Rosa','Noelia','Rodrigo','Ismael','Martín',
                'María Jesús','Juan Luis','Claudia','María Isabel','María','José Antonio','Emilio','Andrés','José Angel',
                'María Antonia','Pablo','Juan Francisco','Consuelo','Martina','Raúl','Antonio','Esther','Dolores','María Dolores',
                'Héctor','Carlos','Olga','Miguel Ángel','Gregorio','Andrea','Miguel','María Teresa','Diego','Rosa María','María Pilar',
                'Mercedes','Elena','Maria Rosario','Marc','Victoria','Victor Manuel','Nicolas','Patricia','Alberto','Roberto',
                'Gonzalo','Inés','María Josefa','Mónica','José Ramon','Oscar','Ángela','Emilia','María Luísa','Ana María',
                'María Mercedes','Jaime','Lorena','Agustín','Julián','Raquel','Eva María','Marta','Pedro','Íker','Margarita',
                'Mariano','Mario','Ana Belén','María Elena','Francisco Javier','Francisca','María Nieves','María Soledad','Carmen',
                'Enrique','Julia','Teresa','Sebastián','Ángeles','Juan Carlos','Montserrat','Cristian','Jesús','María Victoria',
                'Natalia','José Miguel','Eva','Eduardo','Jordi','Iván','Mohamed','Julio','Ricardo','Sonia'];
 
 var apellidosTop = ['González','Rodríguez','Fernández','López','Martínez','Sánchez','Pérez','Gómez','Martín','Jiménez','Ruíz',
                      'Hernández','Díaz','Moreno','Muñoz','Álvarez','Romero','Alonso','Gutiérrez','Navarro','Torres','Domínguez',
                      'Vázquez','Ramos','Gil','Ramírez','Serrano','Blanco','Molina','Morales','Suárez','Ortega','Delgado','Castro',
                      'Ortiz','Rubio','Marín','Sanz','Núñez','Iglesias','Medina','Garrido','Cortés','Castillo','Santos','Lozano',
                      'Guerrero','Cano','Prieto','Méndez','Cruz','Calvo','Gallego','Herrera','Márquez','León','Vidal','Peña',
                      'Flores','Cabrera','Campos','Vega','Fuentes','Carrasco','Díez','Reyes','Caballero','Nieto','Aguilar',
                      'Pascual','Santana','Herrero','Montero','Lorenzo','Hidalgo','Giménez','Ibáñez','Ferrer','Durán',
                      'Santiago','Benítez','Vargas','Mora','Vicente','Arias','Carmona','Crespo','Román','Pastor','Soto',
                      'Sáez','Velasco','Moya','Soler','Parra','Esteban','Bravo','Gallardo','Rojas', 'Felip'];
       
  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  var ncol, nfil;
  
  rangos.map(function(r) {
  
    matriz = r.getValues();
    ncol = matriz[0].length;
    nfil = matriz.length;
  
    if (ncol == 3) {
    
      // El usuario desea apellido1, apellido2 y nombre en columnas contiguas
      
      for (var f = 0; f < nfil; f++) {
        matriz[f][0] = apellidosTop[Math.floor(Math.random()*apellidosTop.length)];
        matriz[f][1] = apellidosTop[Math.floor(Math.random()*apellidosTop.length)];
        matriz[f][2] = nombresTop[Math.floor(Math.random()*nombresTop.length)];
      }
    }
    
    else
    {
      
      // El usuario desea apellido1, apellido2 y nombre concatenados en todo el rango
  
      for (var c = 0; c < ncol; c++) {
        for (var f = 0; f < nfil; f++) {
          matriz[f][c] = apellidosTop[Math.floor(Math.random()*apellidosTop.length)] + ' ' +
                         apellidosTop[Math.floor(Math.random()*apellidosTop.length)] + ', ' +
                         nombresTop[Math.floor(Math.random()*nombresTop.length)];
        }
      }
    } 
    
    r.setValues(matriz);
  })

}                       

function forzarRecalculo() {
 
  // Hack para forzar recálculo de funciones personalizadas
  // No funciona insertando fila por debajo, ocultando, moviendo a primera posición y eliminando :-/
  // Ídem con columnas
  SpreadsheetApp.getActiveSheet().insertRowBefore(1);
  SpreadsheetApp.getActiveSheet().setRowHeight(1, 1);
  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSheet().deleteRow(1);
  
}

// Funciones de manipulación de casillas de verificación

function check() {

  procesarCheck(true);

}


function uncheck() {

  procesarCheck(false);

}

/**
 * Procesa el rango seleccionado, ajustando el valor de 
 * las celdas con valores TRUE o FALSE al valor T/F indicado
 */

function procesarCheck(valor)  {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
        return (typeof c == 'boolean' ? valor : c);    
      })
    
    })
    r.setValues(matriz);
  })
    
}

/**
 * Procesa el rango seleccionado, invirtiendo el valor de 
 * las celdas con valores TRUE o FALSE 
 */

function recheck(valor)  {

  var rangos = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
  var matriz;
  
  rangos.map(function(r) {
  
    matriz = r.getValues().map(function(c) {
      
      return c.map(function(c) {
      
       return (typeof c == 'boolean' ? !c : c);    
      })
    
    })
    r.setValues(matriz);
  })
    
}