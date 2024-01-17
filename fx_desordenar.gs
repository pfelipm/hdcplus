/**
 * Ordena aleatoriamente los datos contenidos en cada una de las *filas* de un intervalo
 * de manera independiente, utilizando el algoritmo de Durstenfeld (coste O(n)),
 * para entremezclar (barajar) la información.
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
 * Ordena aleatoriamente los datos contenidos en cada una de las *columnas* de un intervalo
 * de manera independiente, utilizando el algoritmo de Durstenfeld (coste O(n)),
 * para entremezclar (barajar) la información.
 *
 * @param {A2:C20} intervalo
 * Intervalo de datos a barajar (ofuscar)
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