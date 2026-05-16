/**
 * HdC+
 * Una colección de pequeñas utilidades para la edición de hojas de cálculo
 * Copyright (C) Pablo Felip (@pfelipm). Se distribuye bajo licencia GNU GPL v3.
 *
 * @OnlyCurrentDoc
 */

const VERSION = 'Versión: 2.0 (mayo 2026)';

const ENCABEZADO_ALERTAS = 'HdC+';

const NOMBRE_INDICE = 'Índice HdC+';

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
    // Anotaciones
    .addSubMenu(ui.createMenu('📝 Anotar celdas')      .addItem('Insertar nota con fecha', 'notaFecha')
      .addItem('Insertar nota con usuario', 'notaUsuario')
      .addItem('Insertar nota con fecha y hora', 'notaFechaHora')
      .addItem('Insertar nota con fecha y usuario', 'notaFechaUsuario')
      .addItem('Insertar nota con fecha, hora y usuario', 'notaFechaHoraUsuario'))
    // Desordenar filas / columnas
    .addSubMenu(ui.createMenu('🔀 Barajar datos')
      .addItem('Barajar por columnas', 'desordenarFil')
      .addItem('Barajar por filas', 'desordenarCol'))
    // Manipular hojas
    .addSubMenu(ui.createMenu('📋 Gestionar hojas')
      .addItem('📁 Gestionar hojas...', 'abrirConsolaPestañas')
      .addSeparator()
      .addItem('📑 Generar pestaña de índice', 'generarIndiceCompleto')
      .addItem('📌 Insertar índice aquí (solo enlaces)', 'insertarIndiceLigero')
      .addSeparator()
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
      .addItem('🖼️ Crear marco de color...', 'abrirDialogoMarcoColor')
      .addSeparator()
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
    // Estructura de datos
    .addSubMenu(ui.createMenu('📐 Manipular intervalos de datos')
      .addItem('⚡ Invertir casillas de verificación', 'invertirCasillas')
      .addItem('☑️ Convertir texto a casillas', 'textoACasillas')
      .addItem('⬇️ Rellenar celdas vacías hacia abajo', 'fillDown')
      .addSubMenu(ui.createMenu('🗜️ Compactar selección')
        .addItem('Filas vacías', 'compactarFilas')
        .addItem('Columnas vacías', 'compactarColumnas')
        .addItem('Filas y columnas vacías', 'compactarAmbas'))
      .addSubMenu(ui.createMenu('↕️ Invertir y transponer')
        .addItem('Invertir filas (solo valores)', 'invertirOrdenFil')
        .addItem('Invertir columnas (solo valores)', 'invertirOrdenCol')
        .addItem('Transponer (todo, destructivo)', 'transponer'))
      .addItem('🔗 Extraer URLs de enlaces', 'extraerURLs')
      .addSeparator()
      .addItem('Consolidar dimensiones (despivotar)', 'unpivot_'))
    // Proteger celdas
    .addSubMenu(ui.createMenu('🔏 Proteger celdas con fórmulas')
      .addSubMenu(ui.createMenu('📄 Hoja actual')
        .addItem('Proteger (advertencia)', 'protegerFxHojaAdvertencia')
        .addItem('Proteger (solo tú)', 'protegerFxHojaSoloYo')
        .addSeparator()
        .addItem('Eliminar protecciones HdC+', 'eliminarProteccionesFxHdCPlus')
        .addItem('Eliminar todas las protecciones', 'eliminarProteccionesFxTodas'))
      .addSubMenu(ui.createMenu('📚 Todas las hojas')
        .addItem('Proteger (advertencia)', 'protegerFxLibroAdvertencia')
        .addItem('Proteger (solo tú)', 'protegerFxLibroSoloYo')
        .addSeparator()
        .addItem('Eliminar protecciones HdC+', 'eliminarProteccionesFxLibroHdCPlus')
        .addItem('Eliminar todas las protecciones', 'eliminarProteccionesFxLibroTodas'))
      .addSeparator()
      .addItem('🛡️ Seleccionar hojas...', 'abrirPanelProteccion'))
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
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(420).setHeight(600), '💡 ¿Qué es HdC+?');}


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
  
      SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, e, SpreadsheetApp.getUi().ButtonSet.OK);
      return (e);
      
  }
  
}

function insertarFyC() {

  var ui = HtmlService.createTemplateFromFile('panelCrearFyC')
      .evaluate()
      .setTitle('Insertar filas y/o columnas');
  SpreadsheetApp.getUi().showSidebar(ui);
  
}

const insertarFyC_core = (e) => {
  
  const nFil = +e.numFil;
  const nCol = +e.numCol;
  const modo = +e.modo;
  const hdc = SpreadsheetApp.getActiveSheet();
  
  // Fila y Columna posición superior izquierda rango seleccionado
  const fil = hdc.getActiveCell().getRow();
  const col = hdc.getActiveCell().getColumn();
  
  // Fila y Columna máximas
  const filMax = hdc.getMaxRows();
  const colMax = hdc.getMaxColumns();
    
  switch (modo) {
    case 1: // al principio de la hoja
      if (nFil > 0) hdc.insertRowsBefore(1, nFil);
      if (nCol > 0) hdc.insertColumnsBefore(1, nCol);
      break;
    
    case 2: // al final de la hoja
      if (nFil > 0) hdc.insertRowsAfter(filMax, nFil);
      if (nCol > 0) hdc.insertColumnsAfter(colMax, nCol);
      break;
    
    case 3: // antes de celda
      if (nFil > 0) hdc.insertRowsBefore(fil, nFil);
      if (nCol > 0) hdc.insertColumnsBefore(col, nCol);
      break;
      
    case 4: // después de celda
      if (nFil > 0) hdc.insertRowsAfter(fil, nFil);
      if (nCol > 0) hdc.insertColumnsAfter(col, nCol);
      break;    
  }

  return `✅ Operación completada con éxito.`;
}

function eliminarFyC() {

  var ui = HtmlService.createTemplateFromFile('panelEliminarFyC')
      .evaluate()
      .setTitle('Eliminar filas y/o columnas sobrantes');
  SpreadsheetApp.getUi().showSidebar(ui);
  
}

const eliminarFyC_core = (e) => {
  
  const modo = +e.modo;
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  
  // Si viene un nombre de hoja específico (modo granular)
  if (e.nombreHoja) {
    const hoja = hdc.getSheetByName(e.nombreHoja);
    if (!hoja) throw new Error(`Pestaña no encontrada: ${e.nombreHoja}`);
    procesarRecorteHoja_(hoja, modo, e.encuadre === 'on');
    return { nombre: e.nombreHoja, completado: true };
  }

  // Modo directo (Hoja actual o todas las hojas en una sola llamada - OBSOLETO por UI granular)
  const global = e.global === 'on';
  const encuadre = e.encuadre === 'on';
  const hojas = global ? hdc.getSheets() : [hdc.getActiveSheet()];
    
  hojas.forEach(h => procesarRecorteHoja_(h, modo, encuadre));

  return `✅ Recorte completado en ${hojas.length} pestañas.`;
}

/**
 * Función interna de procesamiento de recorte para una hoja
 * @private
 */
function procesarRecorteHoja_(h, modo, encuadre) {
  // (A) RECORTE POR ARRIBA E IZQUIERDA (ENCUADRE)
  if (encuadre) {
    const vals = h.getDataRange().getValues();
    let firstRow = 0, firstCol = 0;

    for (let r = 0; r < vals.length; r++) {
      if (vals[r].join("").length > 0) {
        firstRow = r + 1;
        break;
      }
    }

    if (firstRow > 0) {
      let minCol = vals[0].length;
      vals.forEach(fila => {
        for (let c = 0; c < fila.length; c++) {
          if (fila[c] !== "") {
            minCol = Math.min(minCol, c);
            break;
          }
        }
      });
      firstCol = minCol + 1;
    }

    if (firstCol > 1) h.deleteColumns(1, firstCol - 1);
    if (firstRow > 1) h.deleteRows(1, firstRow - 1);
    SpreadsheetApp.flush();
  }

  // (B) RECORTE POR ABAJO Y DERECHA (SOBRANTES)
  const nFilas = h.getLastRow();
  const nMaxFilas = h.getMaxRows();
  const nColumnas = h.getLastColumn();
  const nMaxColumnas = h.getMaxColumns();

  if (nFilas > 0 && nColumnas > 0) {
    switch (modo) {
      case 1: if (nMaxColumnas > nColumnas) h.deleteColumns(nColumnas + 1, nMaxColumnas - nColumnas); break;
      case 2: if (nMaxFilas > nFilas) h.deleteRows(nFilas + 1, nMaxFilas - nFilas); break;   
      case 3: 
        if (nMaxColumnas > nColumnas) h.deleteColumns(nColumnas + 1, nMaxColumnas - nColumnas);
        if (nMaxFilas > nFilas) h.deleteRows(nFilas + 1, nMaxFilas - nFilas);          
        break;
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
    ui.alert(ENCABEZADO_ALERTAS, 'Selecciona un único intervalo de datos.', ui.ButtonSet.OK);
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
      ui.alert(ENCABEZADO_ALERTAS, `Se ha producido un error inesperado al ajustar eliminar las filas y columnas no seleccionadas.
        
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
 
  if (matriz.length <= 1) {SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, '💡 La selección debe tener más de 1 fila.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
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
 
  if (matriz[0].length <= 1) {SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, '💡 La selección debe tener más de 1 columna.', SpreadsheetApp.getUi().ButtonSet.OK);}
  
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

const compactarAmbas = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    let tFil = 0, tCol = 0;

    ranges.forEach(r => {
      let vals = r.getValues();
      const nRowsOri = vals.length;
      const nColsOri = vals[0].length;

      // 1. Compactar filas
      vals = vals.filter(f => f.join('').trim() !== '');
      if (vals.length === 0) {
        tFil += nRowsOri;
        tCol += nColsOri;
        r.clearContent();
        return;
      }
      tFil += (nRowsOri - vals.length);

      // 2. Compactar columnas
      const nRows = vals.length;
      const nCols = vals[0].length;
      const indicesNoVacios = [];
      for (let c = 0; c < nCols; c++) {
        let colVacia = true;
        for (let f = 0; f < nRows; f++) {
          if (vals[f][c] !== '') {
            colVacia = false;
            break;
          }
        }
        if (!colVacia) indicesNoVacios.push(c);
      }
      tCol += (nCols - indicesNoVacios.length);

      const matrizFinal = vals.map(fila => indicesNoVacios.map(idx => fila[idx]));
      r.clearContent();
      r.offset(0, 0, matrizFinal.length, matrizFinal[0].length).setValues(matrizFinal);
    });

    const msg = (tFil > 0 && tCol > 0) ? `Se han compactado ${tFil} filas y ${tCol} columnas.` :
                (tFil > 0) ? `Se han compactado ${tFil} filas.` :
                (tCol > 0) ? `Se han compactado ${tCol} columnas.` : 'No se encontraron huecos vacíos.';
    
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '🗜️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al compactar selección: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};


const compactarColumnas = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    let tCol = 0;

    ranges.forEach(r => {
      const vals = r.getValues();
      const nRows = vals.length;
      const nCols = vals[0].length;
      const indicesNoVacios = [];
      for (let c = 0; c < nCols; c++) {
        let colVacia = true;
        for (let f = 0; f < nRows; f++) {
          if (vals[f][c] !== '') {
            colVacia = false;
            break;
          }
        }
        if (!colVacia) indicesNoVacios.push(c);
      }
      if (indicesNoVacios.length === nCols) return;
      
      tCol += (nCols - indicesNoVacios.length);
      const nuevaMatriz = vals.map(fila => indicesNoVacios.map(idx => fila[idx]));
      r.clearContent();
      if (indicesNoVacios.length > 0) {
        r.offset(0, 0, nRows, indicesNoVacios.length).setValues(nuevaMatriz);
      }
    });
    
    const msg = tCol > 0 ? `Se han compactado ${tCol} columnas.` : 'No se encontraron columnas vacías.';
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '🗜️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al compactar columnas: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};


const extraerURLs = () => {
  const ui = SpreadsheetApp.getUi();
  const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();

  if (ui.alert(ENCABEZADO_ALERTAS, '⚠️ ¿Extraer URLs?\n\nEsta acción sustituirá el texto visible por las direcciones URL de los enlaces encontrados en la selección.', ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;

  try {
    ranges.forEach(r => {
      const rtMatrix = r.getRichTextValues();
      const output = rtMatrix.map(fila => fila.map(celda => {
        if (!celda) return '';
        const urls = celda.getRuns()
          .map(run => run.getLinkUrl())
          .filter((url, i, self) => url && self.indexOf(url) === i); // Únicos no nulos
        return urls.length > 0 ? urls.join('\n') : celda.getText();
      }));
      r.setValues(output);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('URLs extraídas con éxito.', '🔗 HdC+');
  } catch (e) {
    ui.alert(ENCABEZADO_ALERTAS, `Error al extraer URLs: ${e.message}`, ui.ButtonSet.OK);
  }
};

const invertirOrdenCol = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    ranges.forEach(r => {
      const vals = r.getValues();
      const output = vals.map(fila => fila.reverse());
      r.setValues(output);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Orden de columnas invertido.', '↔️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al invertir columnas: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

const invertirOrdenFil = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    ranges.forEach(r => r.setValues(r.getValues().reverse()));
    SpreadsheetApp.getActiveSpreadsheet().toast('Orden de filas invertido.', '↕️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al invertir filas: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};


const compactarFilas = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    let tFil = 0;

    ranges.forEach(r => {
      const vals = r.getValues();
      const nRowsOri = vals.length;
      const filtrados = vals.filter(f => f.join('').trim() !== '');
      
      tFil += (nRowsOri - filtrados.length);
      
      if (filtrados.length === nRowsOri) return;
      
      r.clearContent();
      if (filtrados.length > 0) {
        r.offset(0, 0, filtrados.length, vals[0].length).setValues(filtrados);
      }
    });

    const msg = tFil > 0 ? `Se han compactado ${tFil} filas.` : 'No se encontraron filas vacías.';
    SpreadsheetApp.getActiveSpreadsheet().toast(msg, '🗜️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al compactar: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};


const fillDown = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    ranges.forEach(r => {
      const vals = r.getValues();
      const nCols = vals[0].length;
      const nRows = vals.length;

      for (let c = 0; c < nCols; c++) {
        let ultimoValor = '';
        for (let f = 0; f < nRows; f++) {
          if (vals[f][c] !== '') ultimoValor = vals[f][c];
          else if (ultimoValor !== '') vals[f][c] = ultimoValor;
        }
      }
      r.setValues(vals);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Relleno completado.', '⬇️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error en Fill Down: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

const textoACasillas = () => {
  const regSI = /^(sí|si|v|verdadero|1|true|yes|y|ok|x)$/i;
  const regNO = /^(no|f|falso|0|false|n)$/i;

  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    ranges.forEach(r => {
      const vals = r.getValues().map(fila => fila.map(c => {
        const s = String(c).trim();
        if (regSI.test(s)) return true;
        if (regNO.test(s)) return false;
        return c;
      }));
      r.setValues(vals).insertCheckboxes();
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Conversión a casillas finalizada.', '☑️ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al convertir: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};

const invertirCasillas = () => {
  try {
    const ranges = SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges();
    ranges.forEach(r => {
      const vals = r.getValues().map(fila => fila.map(c => typeof c === 'boolean' ? !c : c));
      r.setValues(vals);
    });
    SpreadsheetApp.getActiveSpreadsheet().toast('Casillas invertidas.', '⚡ HdC+');
  } catch (e) {
    SpreadsheetApp.getUi().alert(ENCABEZADO_ALERTAS, `Error al invertir casillas: ${e.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
};