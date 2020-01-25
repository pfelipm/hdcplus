/**
 * Rellena un rango de celdas con un elemento determinado.
 *
 * @param {"Calificación"} elemento
 * Elemento (número, cadena o referencia a celda única) con el que rellenar rango de celdas.
 * 
 * @param {5} nfilas
 * Nº de filas a rellenar.
 *
 * @param {3} ncolumnas
 * Nº de columnas a rellenar.
 *
 * @return vector fila, columna o matriz
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */

function RELLENAR(elemento, nfilas, ncolumnas) {
    
  var resultadoC = [],
      resultado  = [],
      i;
  
  // Control de parámetros

  if (typeof elemento == 'undefined' || typeof(nfilas) == 'undefined' || typeof ncolumnas == 'undefined') { 
    return '!Faltan argumentos';
  }
  if (typeof elemento == 'object') {
    return '!Elemento debe ser único';
  }
  if (!elemento) {
    return null;
  }
  if ((typeof nfilas == 'number' && nfilas < 1) || !nfilas) { 
    return null;
  }
  if ((typeof ncolumnas == 'number' && ncolumnas < 1) || !ncolumnas) { 
    return null;
  }
  if (typeof nfilas != 'number') {
    return  '!Nº filas inválido';
  }
  if (typeof ncolumnas != 'number') {
    return  '!Nº columnas inválido';
  }

  // Los parámetros parecen correctos ¡adelante!

  for (i = 0; i < ncolumnas; i++) {
    resultadoC.push (elemento);
  } 
  
  for (i = 0; i < nfilas; i++) {          
    resultado.push (resultadoC);
  }
  
  return resultado;   

}

/**
 * Repite un elemento el nº de veces indicado en filas y columnas contiguas.
 *
 * @param {"Calificación"} elemento
 * Elemento a repetir (numérico, cadena o referencia a celda única).
 * 
 * @param {5} veces
 * Nº de veces a repetir.
 *
 * @param {"fil" | "col"} modo
 *
 *
 * @return vector fila o columna
 *
 * @customfunction
 *
 * (c) Pablo Felip @pfelipm GNU GPL v3
 */

function REPETIRFC(elemento, veces, modo) {
  
  var resultado;
  var i;
  
  // Control de parámetros

  if (typeof elemento == 'undefined' || typeof(veces) == 'undefined' || typeof modo == 'undefined') { 
    return '!Faltan argumentos';
  }
  if (typeof elemento == 'object') {
    return '!Elemento debe ser único';
  }
  if (!elemento) {
    return null;
  }
  if ((typeof veces == 'number' && veces < 1) || !veces) { 
    return null;
  }
  if (typeof veces != 'number') {
    return  '!Nº repeticiones inválido';
  }
 
  // Los parámetros parecen correctos ¡adelante!
    
  switch (modo.toLowerCase()) {
      
    // extender en filas
      
    case 'fil':
      
      resultado =[];     
      for (i = 0; i < veces; i++) {          
        resultado.push (elemento);
      }
      break;
      
      // extender en columnas
      
    case 'col':
      
      resultado =[[]];      
      for (i = 0; i < veces; i++) {
        resultado[0].push (elemento);
      } 
      break;
      
    default:
      
      resultado = 'Modo desconocido';   
  }
  
  return resultado; 

}