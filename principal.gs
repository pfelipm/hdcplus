/**
 * HdC+
 * Una colección de pequeñas utilidades para la edición de hojas de cálculo
 * Copyright (C) Pablo Felip (@pfelipm). Se distribuye bajo licencia GNU GPL v3.
 *
 * @OnlyCurrentDoc
 */

const VERSION = 'Versión: 1.91 (abril 2025)';

const ENCABEZADO_ALERTAS = 'HdC+';

// Para mostrar / ocultar pestañas por color
const COLORES_HOJAS = {
  azul: { nombre: 'azul', hex:'#0000ff', icono: '🟦' },
  morado: { nombre: 'morado', hex:'#9900ff', icono: '🟪'  },
  verde: { nombre: 'verde', hex:'#00ff00', icono: '🟩'  },
  naranja: { nombre: 'naranja', hex:'#ff9900', icono: '🟧'  },
  rojo: { nombre: 'rojo', hex:'#ff0000', icono: '🟥'  }
};

// Constantes utilizadas en las funciones de protección de celdas con fórmulas
const PROTECCION = {
  modo: { advertencia: 'advertencia', soloYo: 'soloyo' },
  descripcion: '~Protegido por HdC+'
};

const URL_EXTERNAS = {
  ayudaFxPersonalizadas: 'https://pfelipm.notion.site/fxpersonalizadashdcplus',
  kitFuncionesNombreAyuda: 'https://kitfuncioneshdc.notion.site',
  kitFuncionesNombrePlantillaEjemplos: 'https://docs.google.com/spreadsheets/d/17hvFqSAcrztcdVuN_5dB1mK56P5AeoZflFSTes7g0vM/copy',
  kitFuncionesNombrePlantillaVacia: 'https://docs.google.com/spreadsheets/d/1UB4AAtNid5OvbIipKvSb7lFSh7pfffpZNn_eNQIQzgw/copy'
};

function onInstall(e) {
  // Otras cosas que se deben hacer siempre
  onOpen(e);
}

function onOpen() {

  const ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    // Acondicionar texto
    .addSubMenu(ui.createMenu('🔤 Acondicionar el texto de las celdas')
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
    // Casillas de verificación
    .addSubMenu(ui.createMenu('⚡ Ajustar casillas de verificación')
      .addItem('✔️️ Activar seleccionadas', 'check')
      .addItem('❌ Desactivar seleccionadas ', 'uncheck')
      .addItem('➖ Invertir seleccionadas ', 'recheck'))
    // Anotaciones
    .addSubMenu(ui.createMenu('📝 Anotar celdas')
      .addItem('Insertar nota con fecha', 'notaFecha')
      .addItem('Insertar nota con usuario', 'notaUsuario')
      .addItem('Insertar nota con fecha y hora', 'notaFechaHora')
      .addItem('Insertar nota con fecha y usuario', 'notaFechaUsuario')
      .addItem('Insertar nota con fecha, hora y usuario', 'notaFechaHoraUsuario'))
    // Desordenar filas / columnas
    .addSubMenu(ui.createMenu('🔀 Barajar datos')
      .addItem('Barajar por columnas', 'desordenarFil')
      .addItem('Barajar por filas', 'desordenarCol'))
    // Estructura de datos
    .addSubMenu(ui.createMenu('📐 Estructurar datos')
      .addItem('Consolidar dimensiones (despivotar)', 'unpivot_')
      .addItem('Transponer (destructivo)', 'transponer'))
    // Manipular hojas
    .addSubMenu(ui.createMenu('📋 Gestionar hojas')
      .addItem('Ordenar hojas (A → Z)', 'ordenarHojasAsc')
      .addItem('Ordenar hojas (Z → A)', 'ordenarHojasDesc')
      .addItem('Desordenar hojas', 'desordenarHojas')
      .addSeparator()
      .addItem('Mostrar solo hoja activa', 'mostrarSoloActiva')
      .addItem('Mostrar todas excepto activa', 'mostrarTodasMenosActual')
      .addItem('Conmutar visibilidad hojas', 'conmutarVisibilidadHojas')
      .addItem('Mostrar todas las hojas', 'mostrarHojas')
      .addSeparator()
      .addItem('Eliminar hojas ocultas', 'eliminarHojasOcultas')
      .addItem('Eliminar todas excepto activa', 'eliminarHojas')
      .addSeparator()
      // Manipular hojas → Colores
      .addSubMenu(ui.createMenu('🟦 Hojas color azul')
        .addItem('Ocultar hojas color azul', 'ocultarHojasAzul')        
        .addItem('Mostrar hojas color azul', 'mostrarHojasAzul')
        .addItem('Mostrar solo hojas color azul', 'mostrarSoloHojasAzul')
        .addSeparator()
        .addItem('Eliminar hojas color azul', 'eliminarHojasAzul'))
      .addSubMenu(ui.createMenu('🟪 Hojas color morado')
        .addItem('Ocultar hojas color morado', 'ocultarHojasMorado')
        .addItem('Mostrar hojas color morado', 'mostrarHojasMorado')
        .addItem('Mostrar solo hojas color morado', 'mostrarSoloHojasMorado')
        .addSeparator()
        .addItem('Eliminar hojas color morado', 'eliminarHojasMorado'))
      .addSubMenu(ui.createMenu('🟩 Hojas color verde')
        .addItem('Ocultar hojas color verde', 'ocultarHojasVerde')
        .addItem('Mostrar hojas color verde', 'mostrarHojasVerde')
        .addItem('Mostrar solo hojas color verde', 'mostrarSoloHojasVerde')
        .addSeparator()
        .addItem('Eliminar hojas color verde', 'eliminarHojasVerde'))
      .addSubMenu(ui.createMenu('🟧 Hojas color naranja')
        .addItem('Ocultar hojas color naranja', 'ocultarHojasNaranja')
        .addItem('Mostrar hojas color naranja', 'mostrarHojasNaranja')
        .addItem('Mostrar solo hojas color naranja', 'mostrarSoloHojasNaranja')
        .addSeparator()
        .addItem('Eliminar hojas color naranja', 'eliminarHojasNaranja'))
      .addSubMenu(ui.createMenu('🟥 Hojas color rojo')
        .addItem('Ocultar hojas color rojo', 'ocultarHojasRojo')
        .addItem('Mostrar hojas color rojo', 'mostrarHojasRojo')
        .addItem('Mostrar solo hojas color rojo', 'mostrarSoloHojasRojo')
        .addSeparator()
        .addItem('Eliminar hojas color rojo', 'eliminarHojasRojo')))
    // Filas y columnas
    .addSubMenu(ui.createMenu('🗜️ Insertar y eliminar filas/columnas')
      .addItem('Eliminar celdas no seleccionadas', 'eliminarFyCNoSeleccionadas')
      .addItem('Eliminar filas/columnas sobrantes', 'eliminarFyC')
      .addSeparator()
      .addItem('Insertar más filas/columnas', 'insertarFyC'))
    // Generación
    .addSubMenu(ui.createMenu('✨ Generar datos falsos')
      .addItem('NIFs', 'generarNIF')
      .addItem('Nombres y apellidos', 'generarNombres'))
    // Ofuscar
    .addSubMenu(ui.createMenu('🕶️ Ofuscar información')
      .addItem('Codificar texto en Base64 ', 'base64_')
      .addItem('Sustituir por hash MD2 (b64) ', 'hashMD2')
      .addItem('Sustituir por hash MD5 (b64) ', 'hashMD5')
      .addItem('Sustituir por hash SHA-1 (b64)', 'hashSHA1')
      .addItem('Sustituir por hash SHA-256 (b64)', 'hashSHA256')
      .addItem('Sustituir por hash SHA-384 (b64)', 'hashSHA384')
      .addItem('Sustituir por hash SHA-512 (b64)', 'hashSHA512'))
    // Proteger celdas
    .addSubMenu(ui.createMenu('🔏 Proteger celdas con fórmulas')
      .addItem('Proteger fx hoja (advertencia)', 'protegerFxHojaAdvertencia')
      .addItem('Proteger fx hoja (solo tú)', 'protegerFxHojaSoloYo')
      .addSeparator()
      .addItem('Eliminar protecciones HdC+ de hoja', 'eliminarProteccionesFxHdCPlus')
      .addSeparator()
      .addItem('Eliminar todas las protecciones de hoja', 'eliminarProteccionesFxTodas'))
    .addSeparator()
    // Otros
    .addItem('⚙️ Forzar recálculo de fórmulas hoja', 'forzarRecalculo')
    .addSeparator()
    .addSubMenu(ui.createMenu('🧰 Kit funciones con nombre')
      .addItem('Obtener plantilla vacía', 'kfnPlantillaVacia')
      .addItem('Obtener plantilla con ejemplos', 'kfnPlantillaEjemplos')
      .addItem('Ayuda', 'kfnAyuda'))
    .addSeparator()
    .addItem('🛟 Ayuda fx personalizadas', 'ayudaFxPersonalizadas')
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


// Comandos de menú de apertura de sitios externos
function ayudaFxPersonalizadas() {abrirWebExterna_(URL_EXTERNAS.ayudaFxPersonalizadas);}
function kfnPlantillaVacia() {abrirWebExterna_(URL_EXTERNAS.kitFuncionesNombrePlantillaVacia);}
function kfnPlantillaEjemplos() {abrirWebExterna_(URL_EXTERNAS.kitFuncionesNombrePlantillaEjemplos);}
function kfnAyuda() {abrirWebExterna_(URL_EXTERNAS.kitFuncionesNombreAyuda);}

// Abre un sitio web externo
// https://www.labnol.org/open-webpage-google-sheets-220507
function abrirWebExterna_(urlSitioExterno) {

  console.info(urlSitioExterno)

  const htmlTemplate = HtmlService.createTemplateFromFile('dialogoAbrirPaginaExterna.html');
  htmlTemplate.url = urlSitioExterno;
  SpreadsheetApp.getUi().showModelessDialog(
    htmlTemplate.evaluate().setHeight(85).setWidth(400),
    '🌐 Abriendo sitio web...'
  );
  // No parece ser necesario
  Utilities.sleep(2000);

};

function unpivot_() {

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


function base64_()  {

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