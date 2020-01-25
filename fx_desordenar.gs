/**
 * Ordena aleatoriamente las *columnas* de un intervalo de modo
 * que los datos de distintas *filas* queden desemparejados
 * utilizando el algoritmo de Durstenfeld (coste O(n)).
 *
 * @param {A2:C20} intervalo
 * Intervalo de datos a barajar (ofuscar)
 *
 * @return rango de datos barajado
 *
 * @customfunction
 */

function BARAJARDATOSFIL(intervalo){

  // Control de parámetros
   
  if (typeof intervalo == 'undefined') {return '!Faltan argumentos';}
  if (Array.isArray(intervalo) == false) {return '!Se esperaba una matriz de más de 1 columna';}
  if (intervalo[0].length <= 1) {return '!Se esperaba una matriz de más de 1 columna';}
  
  // Los parámetros parecen correctos ¡adelante!
 
  var ncol = intervalo[0].length;
  var nfil = intervalo.length;

  for (var f = 0; f < nfil; f++) {
    for (var c = ncol - 1; c > 0; c--) { 
      var j = Math.floor(Math.random() * (c + 1));
      var temp = intervalo[f][c];
      intervalo[f][c] = intervalo[f][j];
      intervalo[f][j] = temp;
    }     
  }
  return intervalo;  
}

/**
 * Ordena aleatoriamente las *filas* de un intervalo de modo
 * que los datos de distintas *columnas* queden desemparejados
 * utilizando el algoritmo de Durstenfeld (coste O(n)).
 *
 * @param {A2:C20} intervalo
 * Rango de datos a barajar (ofuscar)
 *
 * @return rango de datos barajado
 *
 * @customfunction
 */

function BARAJARDATOSCOL(intervalo){

  // Control de parámetros
  
  if (typeof intervalo == 'undefined') {return '!Faltan argumentos';}
  if (Array.isArray(intervalo) == false) {return '!Se esperaba una matriz de más de 1 fila';}
  if (intervalo.length <= 1) {return '!Se esperaba una matriz de más de 1 fila';}

  // Los parámetros parecen correctos ¡adelante!

  var ncol = intervalo[0].length;
  var nfil = intervalo.length;

  for (var c = 0; c < ncol; c++) {
    for (var f = nfil - 1; f > 0; f--) { 
      var j = Math.floor(Math.random() * (f + 1));
      var temp = intervalo[f][c];
      intervalo[f][c] = intervalo[j][c];
      intervalo[j][c] = temp;
    }    
  }
  return intervalo;  
}