/**
 * HdC+
 * Una colección de pequeñas utilidades para la edición de hojas de cálculo
 * Copyright (C) Pablo Felip (@pfelipm) · Se distribuye bajo licencia GNU GPL v3.
 *
 * @OnlyCurrentDoc
 */

var VERSION = 'Versión: 1.1 (enero 2020)';

// Para mostrar / ocultar pestaña por color

var COLOR_HOJA_N = "#ff9900"
var COLOR_HOJA_A = "#0000ff"

function onInstall(e) {
  
  // Otras cosas que se deben hacer siempre
  
  onOpen(e);
  
}

function onOpen() {

  var ss = SpreadsheetApp.getUi().createAddonMenu()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔀 Barajar datos')
      .addItem('Barajar por columnas', 'desordenarFil')
      .addItem('Barajar por filas', 'desordenarCol'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('✅ Casillas de verificación')
      .addItem('✔️️ Activar seleccionadas', 'check')
      .addItem('❌ Desactivar seleccionadas ', 'uncheck')
      .addItem('➖ Invertir seleccionadas ', 'recheck'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧮 Estructura datos')
      .addItem('Consolidar dimensiones (despivotar)', 'unpivot')
      .addItem('Transponer (☢️ destructivo)', 'transponer'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📐 Estructura hoja de cálculo')
      .addItem('Eliminar F/C sobrantes', 'eliminarFyC')
      .addItem('Insertar F/C nuevas', 'insertarFyC')
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Manipular hojas')
        .addItem('Ocultar resto de hojas', 'ocultarHojas')
        .addItem('Mostrar hojas ocultas', 'mostrarHojas')
        .addItem('Mostrar todas menos actual', 'mostrarNoActual')        
        .addItem('Eliminar hojas ocultas', 'eliminarHojasOcultas')
        .addItem('Eliminar todas menos actual', 'eliminarHojas')
        .addItem('🔸Mostrar hojas color naranja', 'mostrarHojasNaranja')        
        .addItem('🔸Ocultar hojas color naranja', 'ocultarHojasNaranja')
        .addItem('🔹Mostrar hojas color azul', 'mostrarHojasAzul')        
        .addItem('🔹Ocultar hojas color azul', 'ocultarHojasAzul')))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧠 Generar')
      .addItem('NIFs', 'generarNIF')
      .addItem('Nombres y apellidos', 'generarNombres'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🕶️ Ofuscar')
      .addItem('Codificar texto en base64 ', 'base64')
      .addItem('Sustituir por hash MD2 (b64) ', 'hashMD2')
      .addItem('Sustituir por hash MD5 (b64) ', 'hashMD5')
      .addItem('Sustituir por hash SHA-1 (b64)', 'hashSHA1')
      .addItem('Sustituir por hash SHA-256 (b64)', 'hashSHA256')
      .addItem('Sustituir por hash SHA-384 (b64)', 'hashSHA384')
      .addItem('Sustituir por hash SHA-512 (b64)', 'hashSHA512'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('⚡ Transformar')
      .addItem('Eliminar caracteres especiales', 'latinizar')
      .addItem('Iniciales a mayúsculas', 'inicialesMays')
      .addItem('Saltos de línea como espacios', 'saltosEspacios')
      .addItem('Suprimir saltos de línea ', 'eliminarEspacios')
      .addItem('Todo a mayúsculas', 'mayusculas')
      .addItem('Todo a minúsculas', 'minusculas'))
    .addSeparator()
    .addItem('🔄 Forzar recálculo de hoja', 'forzarRecalculo')
    .addSeparator()
    .addItem('💡 Acerca de HdC+', 'acercaDe')
    .addToUi();
}

function acercaDe() {

  // Presentación del complemento
  var panel = HtmlService.createTemplateFromFile('acercaDe');
  panel.version = VERSION;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(220), '💡 ¿Qué es HdC+?');
}

function eliminarEspacios() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
    
  matriz = matriz.map(function(c) {
  
     return c.map(function(c) {
    
       return c.replace(/\s+/g, '');   
    })
  })
  
  rango.setValues(matriz);

}

function saltosEspacios() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
    
  matriz = matriz.map(function(c) {
  
     return c.map(function(c) {
    
       return c.replace(/\n/g, ' ');    
    })
  })
  
  rango.setValues(matriz);

}

function eliminarSaltos() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
    
  matriz = matriz.map(function(c) {
 
     return c.map(function(c) {
    
       return c.replace(/\n/g, '');  
    })
  })
  
  rango.setValues(matriz);

}

function latinizar() {

  // Adaptado de https://stackoverflow.com/questions/990904/remove-accents-diacritics-in-a-string-in-javascript
  
  var latinise={};latinise.latin_map={"·":"-","ß":"ss","Á":"A","Ă":"A","Ắ":"A","Ặ":"A","Ằ":"A","Ẳ":"A","Ẵ":"A","Ǎ":"A","Â":"A","Ấ":"A","Ậ":"A","Ầ":"A","Ẩ":"A","Ẫ":"A","Ä":"A","Ǟ":"A","Ȧ":"A","Ǡ":"A","Ạ":"A","Ȁ":"A","À":"A","Ả":"A","Ȃ":"A","Ā":"A","Ą":"A","Å":"A","Ǻ":"A","Ḁ":"A","Ⱥ":"A","Ã":"A","Ꜳ":"AA","Æ":"AE","Ǽ":"AE","Ǣ":"AE","Ꜵ":"AO","Ꜷ":"AU","Ꜹ":"AV","Ꜻ":"AV","Ꜽ":"AY","Ḃ":"B","Ḅ":"B","Ɓ":"B","Ḇ":"B","Ƀ":"B","Ƃ":"B","Ć":"C","Č":"C","Ç":"C","Ḉ":"C","Ĉ":"C","Ċ":"C","Ƈ":"C","Ȼ":"C","Ď":"D","Ḑ":"D","Ḓ":"D","Ḋ":"D","Ḍ":"D","Ɗ":"D","Ḏ":"D","ǲ":"D","ǅ":"D","Đ":"D","Ƌ":"D","Ǳ":"DZ","Ǆ":"DZ","É":"E","Ĕ":"E","Ě":"E","Ȩ":"E","Ḝ":"E","Ê":"E","Ế":"E","Ệ":"E","Ề":"E","Ể":"E","Ễ":"E","Ḙ":"E","Ë":"E","Ė":"E","Ẹ":"E","Ȅ":"E","È":"E","Ẻ":"E","Ȇ":"E","Ē":"E","Ḗ":"E","Ḕ":"E","Ę":"E","Ɇ":"E","Ẽ":"E","Ḛ":"E","Ꝫ":"ET","Ḟ":"F","Ƒ":"F","Ǵ":"G","Ğ":"G","Ǧ":"G","Ģ":"G","Ĝ":"G","Ġ":"G","Ɠ":"G","Ḡ":"G","Ǥ":"G","Ḫ":"H","Ȟ":"H","Ḩ":"H","Ĥ":"H","Ⱨ":"H","Ḧ":"H","Ḣ":"H","Ḥ":"H","Ħ":"H","Í":"I","Ĭ":"I","Ǐ":"I","Î":"I","Ï":"I","Ḯ":"I","İ":"I","Ị":"I","Ȉ":"I","Ì":"I","Ỉ":"I","Ȋ":"I","Ī":"I","Į":"I","Ɨ":"I","Ĩ":"I","Ḭ":"I","Ꝺ":"D","Ꝼ":"F","Ᵹ":"G","Ꞃ":"R","Ꞅ":"S","Ꞇ":"T","Ꝭ":"IS","Ĵ":"J","Ɉ":"J","Ḱ":"K","Ǩ":"K","Ķ":"K","Ⱪ":"K","Ꝃ":"K","Ḳ":"K","Ƙ":"K","Ḵ":"K","Ꝁ":"K","Ꝅ":"K","Ĺ":"L","Ƚ":"L","Ľ":"L","Ļ":"L","Ḽ":"L","Ḷ":"L","Ḹ":"L","Ⱡ":"L","Ꝉ":"L","Ḻ":"L","Ŀ":"L","Ɫ":"L","ǈ":"L","Ł":"L","Ǉ":"LJ","Ḿ":"M","Ṁ":"M","Ṃ":"M","Ɱ":"M","Ń":"N","Ň":"N","Ņ":"N","Ṋ":"N","Ṅ":"N","Ṇ":"N","Ǹ":"N","Ɲ":"N","Ṉ":"N","Ƞ":"N","ǋ":"N","Ñ":"N","Ǌ":"NJ","Ó":"O","Ŏ":"O","Ǒ":"O","Ô":"O","Ố":"O","Ộ":"O","Ồ":"O","Ổ":"O","Ỗ":"O","Ö":"O","Ȫ":"O","Ȯ":"O","Ȱ":"O","Ọ":"O","Ő":"O","Ȍ":"O","Ò":"O","Ỏ":"O","Ơ":"O","Ớ":"O","Ợ":"O","Ờ":"O","Ở":"O","Ỡ":"O","Ȏ":"O","Ꝋ":"O","Ꝍ":"O","Ō":"O","Ṓ":"O","Ṑ":"O","Ɵ":"O","Ǫ":"O","Ǭ":"O","Ø":"O","Ǿ":"O","Õ":"O","Ṍ":"O","Ṏ":"O","Ȭ":"O","Ƣ":"OI","Ꝏ":"OO","Ɛ":"E","Ɔ":"O","Ȣ":"OU","Ṕ":"P","Ṗ":"P","Ꝓ":"P","Ƥ":"P","Ꝕ":"P","Ᵽ":"P","Ꝑ":"P","Ꝙ":"Q","Ꝗ":"Q","Ŕ":"R","Ř":"R","Ŗ":"R","Ṙ":"R","Ṛ":"R","Ṝ":"R","Ȑ":"R","Ȓ":"R","Ṟ":"R","Ɍ":"R","Ɽ":"R","Ꜿ":"C","Ǝ":"E","Ś":"S","Ṥ":"S","Š":"S","Ṧ":"S","Ş":"S","Ŝ":"S","Ș":"S","Ṡ":"S","Ṣ":"S","Ṩ":"S","Ť":"T","Ţ":"T","Ṱ":"T","Ț":"T","Ⱦ":"T","Ṫ":"T","Ṭ":"T","Ƭ":"T","Ṯ":"T","Ʈ":"T","Ŧ":"T","Ɐ":"A","Ꞁ":"L","Ɯ":"M","Ʌ":"V","Ꜩ":"TZ","Ú":"U","Ŭ":"U","Ǔ":"U","Û":"U","Ṷ":"U","Ü":"U","Ǘ":"U","Ǚ":"U","Ǜ":"U","Ǖ":"U","Ṳ":"U","Ụ":"U","Ű":"U","Ȕ":"U","Ù":"U","Ủ":"U","Ư":"U","Ứ":"U","Ự":"U","Ừ":"U","Ử":"U","Ữ":"U","Ȗ":"U","Ū":"U","Ṻ":"U","Ų":"U","Ů":"U","Ũ":"U","Ṹ":"U","Ṵ":"U","Ꝟ":"V","Ṿ":"V","Ʋ":"V","Ṽ":"V","Ꝡ":"VY","Ẃ":"W","Ŵ":"W","Ẅ":"W","Ẇ":"W","Ẉ":"W","Ẁ":"W","Ⱳ":"W","Ẍ":"X","Ẋ":"X","Ý":"Y","Ŷ":"Y","Ÿ":"Y","Ẏ":"Y","Ỵ":"Y","Ỳ":"Y","Ƴ":"Y","Ỷ":"Y","Ỿ":"Y","Ȳ":"Y","Ɏ":"Y","Ỹ":"Y","Ź":"Z","Ž":"Z","Ẑ":"Z","Ⱬ":"Z","Ż":"Z","Ẓ":"Z","Ȥ":"Z","Ẕ":"Z","Ƶ":"Z","Ĳ":"IJ","Œ":"OE","ᴀ":"A","ᴁ":"AE","ʙ":"B","ᴃ":"B","ᴄ":"C","ᴅ":"D","ᴇ":"E","ꜰ":"F","ɢ":"G","ʛ":"G","ʜ":"H","ɪ":"I","ʁ":"R","ᴊ":"J","ᴋ":"K","ʟ":"L","ᴌ":"L","ᴍ":"M","ɴ":"N","ᴏ":"O","ɶ":"OE","ᴐ":"O","ᴕ":"OU","ᴘ":"P","ʀ":"R","ᴎ":"N","ᴙ":"R","ꜱ":"S","ᴛ":"T","ⱻ":"E","ᴚ":"R","ᴜ":"U","ᴠ":"V","ᴡ":"W","ʏ":"Y","ᴢ":"Z","á":"a","ă":"a","ắ":"a","ặ":"a","ằ":"a","ẳ":"a","ẵ":"a","ǎ":"a","â":"a","ấ":"a","ậ":"a","ầ":"a","ẩ":"a","ẫ":"a","ä":"a","ǟ":"a","ȧ":"a","ǡ":"a","ạ":"a","ȁ":"a","à":"a","ả":"a","ȃ":"a","ā":"a","ą":"a","ᶏ":"a","ẚ":"a","å":"a","ǻ":"a","ḁ":"a","ⱥ":"a","ã":"a","ꜳ":"aa","æ":"ae","ǽ":"ae","ǣ":"ae","ꜵ":"ao","ꜷ":"au","ꜹ":"av","ꜻ":"av","ꜽ":"ay","ḃ":"b","ḅ":"b","ɓ":"b","ḇ":"b","ᵬ":"b","ᶀ":"b","ƀ":"b","ƃ":"b","ɵ":"o","ć":"c","č":"c","ç":"c","ḉ":"c","ĉ":"c","ɕ":"c","ċ":"c","ƈ":"c","ȼ":"c","ď":"d","ḑ":"d","ḓ":"d","ȡ":"d","ḋ":"d","ḍ":"d","ɗ":"d","ᶑ":"d","ḏ":"d","ᵭ":"d","ᶁ":"d","đ":"d","ɖ":"d","ƌ":"d","ı":"i","ȷ":"j","ɟ":"j","ʄ":"j","ǳ":"dz","ǆ":"dz","é":"e","ĕ":"e","ě":"e","ȩ":"e","ḝ":"e","ê":"e","ế":"e","ệ":"e","ề":"e","ể":"e","ễ":"e","ḙ":"e","ë":"e","ė":"e","ẹ":"e","ȅ":"e","è":"e","ẻ":"e","ȇ":"e","ē":"e","ḗ":"e","ḕ":"e","ⱸ":"e","ę":"e","ᶒ":"e","ɇ":"e","ẽ":"e","ḛ":"e","ꝫ":"et","ḟ":"f","ƒ":"f","ᵮ":"f","ᶂ":"f","ǵ":"g","ğ":"g","ǧ":"g","ģ":"g","ĝ":"g","ġ":"g","ɠ":"g","ḡ":"g","ᶃ":"g","ǥ":"g","ḫ":"h","ȟ":"h","ḩ":"h","ĥ":"h","ⱨ":"h","ḧ":"h","ḣ":"h","ḥ":"h","ɦ":"h","ẖ":"h","ħ":"h","ƕ":"hv","í":"i","ĭ":"i","ǐ":"i","î":"i","ï":"i","ḯ":"i","ị":"i","ȉ":"i","ì":"i","ỉ":"i","ȋ":"i","ī":"i","į":"i","ᶖ":"i","ɨ":"i","ĩ":"i","ḭ":"i","ꝺ":"d","ꝼ":"f","ᵹ":"g","ꞃ":"r","ꞅ":"s","ꞇ":"t","ꝭ":"is","ǰ":"j","ĵ":"j","ʝ":"j","ɉ":"j","ḱ":"k","ǩ":"k","ķ":"k","ⱪ":"k","ꝃ":"k","ḳ":"k","ƙ":"k","ḵ":"k","ᶄ":"k","ꝁ":"k","ꝅ":"k","ĺ":"l","ƚ":"l","ɬ":"l","ľ":"l","ļ":"l","ḽ":"l","ȴ":"l","ḷ":"l","ḹ":"l","ⱡ":"l","ꝉ":"l","ḻ":"l","ŀ":"l","ɫ":"l","ᶅ":"l","ɭ":"l","ł":"l","ǉ":"lj","ſ":"s","ẜ":"s","ẛ":"s","ẝ":"s","ḿ":"m","ṁ":"m","ṃ":"m","ɱ":"m","ᵯ":"m","ᶆ":"m","ń":"n","ň":"n","ņ":"n","ṋ":"n","ȵ":"n","ṅ":"n","ṇ":"n","ǹ":"n","ɲ":"n","ṉ":"n","ƞ":"n","ᵰ":"n","ᶇ":"n","ɳ":"n","ñ":"n","ǌ":"nj","ó":"o","ŏ":"o","ǒ":"o","ô":"o","ố":"o","ộ":"o","ồ":"o","ổ":"o","ỗ":"o","ö":"o","ȫ":"o","ȯ":"o","ȱ":"o","ọ":"o","ő":"o","ȍ":"o","ò":"o","ỏ":"o","ơ":"o","ớ":"o","ợ":"o","ờ":"o","ở":"o","ỡ":"o","ȏ":"o","ꝋ":"o","ꝍ":"o","ⱺ":"o","ō":"o","ṓ":"o","ṑ":"o","ǫ":"o","ǭ":"o","ø":"o","ǿ":"o","õ":"o","ṍ":"o","ṏ":"o","ȭ":"o","ƣ":"oi","ꝏ":"oo","ɛ":"e","ᶓ":"e","ɔ":"o","ᶗ":"o","ȣ":"ou","ṕ":"p","ṗ":"p","ꝓ":"p","ƥ":"p","ᵱ":"p","ᶈ":"p","ꝕ":"p","ᵽ":"p","ꝑ":"p","ꝙ":"q","ʠ":"q","ɋ":"q","ꝗ":"q","ŕ":"r","ř":"r","ŗ":"r","ṙ":"r","ṛ":"r","ṝ":"r","ȑ":"r","ɾ":"r","ᵳ":"r","ȓ":"r","ṟ":"r","ɼ":"r","ᵲ":"r","ᶉ":"r","ɍ":"r","ɽ":"r","ↄ":"c","ꜿ":"c","ɘ":"e","ɿ":"r","ś":"s","ṥ":"s","š":"s","ṧ":"s","ş":"s","ŝ":"s","ș":"s","ṡ":"s","ṣ":"s","ṩ":"s","ʂ":"s","ᵴ":"s","ᶊ":"s","ȿ":"s","ɡ":"g","ᴑ":"o","ᴓ":"o","ᴝ":"u","ť":"t","ţ":"t","ṱ":"t","ț":"t","ȶ":"t","ẗ":"t","ⱦ":"t","ṫ":"t","ṭ":"t","ƭ":"t","ṯ":"t","ᵵ":"t","ƫ":"t","ʈ":"t","ŧ":"t","ᵺ":"th","ɐ":"a","ᴂ":"ae","ǝ":"e","ᵷ":"g","ɥ":"h","ʮ":"h","ʯ":"h","ᴉ":"i","ʞ":"k","ꞁ":"l","ɯ":"m","ɰ":"m","ᴔ":"oe","ɹ":"r","ɻ":"r","ɺ":"r","ⱹ":"r","ʇ":"t","ʌ":"v","ʍ":"w","ʎ":"y","ꜩ":"tz","ú":"u","ŭ":"u","ǔ":"u","û":"u","ṷ":"u","ü":"u","ǘ":"u","ǚ":"u","ǜ":"u","ǖ":"u","ṳ":"u","ụ":"u","ű":"u","ȕ":"u","ù":"u","ủ":"u","ư":"u","ứ":"u","ự":"u","ừ":"u","ử":"u","ữ":"u","ȗ":"u","ū":"u","ṻ":"u","ų":"u","ᶙ":"u","ů":"u","ũ":"u","ṹ":"u","ṵ":"u","ᵫ":"ue","ꝸ":"um","ⱴ":"v","ꝟ":"v","ṿ":"v","ʋ":"v","ᶌ":"v","ⱱ":"v","ṽ":"v","ꝡ":"vy","ẃ":"w","ŵ":"w","ẅ":"w","ẇ":"w","ẉ":"w","ẁ":"w","ⱳ":"w","ẘ":"w","ẍ":"x","ẋ":"x","ᶍ":"x","ý":"y","ŷ":"y","ÿ":"y","ẏ":"y","ỵ":"y","ỳ":"y","ƴ":"y","ỷ":"y","ỿ":"y","ȳ":"y","ẙ":"y","ɏ":"y","ỹ":"y","ź":"z","ž":"z","ẑ":"z","ʑ":"z","ⱬ":"z","ż":"z","ẓ":"z","ȥ":"z","ẕ":"z","ᵶ":"z","ᶎ":"z","ʐ":"z","ƶ":"z","ɀ":"z","ﬀ":"ff","ﬃ":"ffi","ﬄ":"ffl","ﬁ":"fi","ﬂ":"fl","ĳ":"ij","œ":"oe","ﬆ":"st","ₐ":"a","ₑ":"e","ᵢ":"i","ⱼ":"j","ₒ":"o","ᵣ":"r","ᵤ":"u","ᵥ":"v","ₓ":"x"};

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
  for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var valor = matriz[f][c];
      if (typeof valor == 'string') {
        matriz[f][c] = valor.replace(/[^A-Za-z0-9\[\] ]/g,function(a){return latinise.latin_map[a]||a})};
    }
  }
  
  rango.setValues(matriz);
  
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

function base64() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  
  matriz = matriz.map(function(c) {
  
     return c.map(function(c) {
    
       return Utilities.base64Encode(c, Utilities.Charset.UTF_8);    
    })
  })

  rango.setValues(matriz);
  
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

// Convierte hash binario en hexadecimal
// Utilizada también por fx hash()

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


// Invocación de hashGenerico() para cada tipo

function hashSHA256() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.SHA_256);    
  
}

function hashSHA384() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.SHA_384);  
  
}

function hashSHA512() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.SHA_512);  
  
}


function hashSHA1() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.SHA_1);
  
}

function hashMD2() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.MD2);
  
}

function hashMD5() {
  
  hashGenerico(SpreadsheetApp.getActiveSheet().getActiveRange(), Utilities.DigestAlgorithm.MD5);
  
}

function inicialesMays() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
 for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var valor = matriz[f][c];
      if (typeof valor == 'string') {
        matriz[f][c] = valor.toLowerCase().replace(/\b\w/gi,(function(c) {return c.toUpperCase()}));
      }
    }
  }
  
  rango.setValues(matriz);
  
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
  
      // El rango es rectangular, yay que "limpiar" el espacio libre que deja tras la transposición
    
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

function minusculas() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
 for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var valor = matriz[f][c];
      if (typeof valor == 'string') {matriz[f][c] = valor.toLowerCase();}
    }
  }
  
  rango.setValues(matriz);
  
}

function mayusculas() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
 for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var valor = matriz[f][c];
      if (typeof valor == 'string') {matriz[f][c] = valor.toUpperCase();}
    }
  }
  
  rango.setValues(matriz);
  
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


function generarNIF() {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length;
  
 for (var c = 0; c < ncol; c++) {
    for (var f = 0; f < nfil; f++) {
      var dni = '';
      for (var n = 0; n < 8; n++) {
        dni = dni + Math.floor(Math.random()*10).toString();
      }
      dni = dni + "TRWAGMYFPDXBNJZSQVHLCKE".charAt(+dni % 23);
      matriz[f][c] = dni;
    }
  }
  
  rango.setValues(matriz);
  
}

function generarNombres() {

  var nombresTop = ['Guillermo','Jose Luis','Nuria','José','David','Victor','Santiago','Francisco José','Marcos','Yolanda',
                'Isabel','César','Carolina','Nerea','Lucas','Alba','Gabriel','Salvador','Ignacio','Manuela','María Carmen',
                'Domingo','Rosario','Albert','María José','Samuel','María Ángeles','Juan Antonio','Laura','Xavier','Lidia',
                'Fatima','Paula','Jose Carlos','Joaquin','Marina','Alejandro','Juan Jose','Alicia','Jose Manuel','Pilar',
                'Luisa','Encarnación','Alfredo','Felipe','Luis Miguel','Sergio','Antonia','Juan','Jose Ignacio','Beatriz',
                'Alvaro','Angel','Ruben','Lucia','Luis','Rocio','Celia','Juana','Rafael','Irene','Tomas','Joan','Jose Maria',
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
                'Natalia','Jose Miguel','Eva','Eduardo','Jordi','Iván','Mohamed','Julio','Ricardo','Sonia'];
 
 var apellidosTop = ['González','Rodríguez','Fernández','López','Martínez','Sánchez','Pérez','Gómez','Martín','Jiménez','Ruíz',
                      'Hernández','Díaz','Moreno','Muñoz','Álvarez','Romero','Alonso','Gutiérrez','Navarro','Torres','Domínguez',
                      'Vázquez','Ramos','Gil','Ramírez','Serrano','Blanco','Molina','Morales','Suárez','Ortega','Delgado','Castro',
                      'Ortiz','Rubio','Marín','Sanz','Núñez','Iglesias','Medina','Garrido','Cortés','Castillo','Santos','Lozano',
                      'Guerrero','Cano','Prieto','Méndez','Cruz','Calvo','Gallego','Herrera','Márquez','León','Vidal','Peña',
                      'Flores','Cabrera','Campos','Vega','Fuentes','Carrasco','Díez','Reyes','Caballero','Nieto','Aguilar',
                      'Pascual','Santana','Herrero','Montero','Lorenzo','Hidalgo','Giménez','Ibáñez','Ferrer','Durán',
                      'Santiago','Benítez','Vargas','Mora','Vicente','Arias','Carmona','Crespo','Román','Pastor','Soto',
                      'Sáez','Velasco','Moya','Soler','Parra','Esteban','Bravo','Gallardo','Rojas', 'Felip'];
       
  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
  var ncol = matriz[0].length;
  var nfil = matriz.length
  
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
  
  rango.setValues(matriz);

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

function procesarCheck(valor) {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
    
  matriz = matriz.map(function(c) {
  
     return c.map(function(c) {
    
       return (typeof c == 'boolean' ? valor : c);    
    })
  })
  
  rango.setValues(matriz);
  
}

/**
 * Procesa el rango seleccionado, invirtiendo el valor de 
 * las celdas con valores TRUE o FALSE 
 */

function recheck(valor) {

  var rango = SpreadsheetApp.getActiveSheet().getActiveRange();
  var matriz = rango.getValues();
    
  matriz = matriz.map(function(c) {
  
     return c.map(function(c) {
    
       return (typeof c == 'boolean' ? !c : c);    
    })
  })
  
  rango.setValues(matriz);
  
}