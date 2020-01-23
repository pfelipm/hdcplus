/**
 * Rellena una rango de celdas con un elemento determinado.
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
  
  if (!elemento) { 
    resultado = 'Falta elemento';
  }
  else if (typeof nfilas != 'number') {
    resultado = 'Nº filas no es número';
  }
  else if (nfilas < 1 && typeof nfilas == 'number') {
    resultado = 'Nº filas debe ser > 0';
  }
  else if (typeof ncolumnas != 'number') {
    resultado = 'Nº columnas no es número';
  }
  else if (ncolumnas < 1 && typeof ncolumnas == 'number') {
    resultado = 'Nº columnas debe ser > 0';
  }
  else {
  
    // ¡Adelante!

    for (i = 0; i < ncolumnas; i++) {
      resultadoC.push (elemento);
    } 
    
    for (i = 0; i < nfilas; i++) {          
      resultado.push (resultadoC);
    }
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
  
  if (!elemento) { 
    resultado = 'Falta elemento';
  }
  else if (veces < 2 && typeof veces == 'number') {
    resultado = 'Repeticiones debe ser > 1';
  }
  else if (typeof veces != 'number') {
    resultado = 'Repeticiones no es número';
  }
  else if (!modo) {
    resultado = 'Falta modo';
  }
  else {
  
    // ¡Adelante!
    
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
  }
  
  return resultado; 

}