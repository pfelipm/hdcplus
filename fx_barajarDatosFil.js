/**
 * Ordena aleatoriamente las filas de una matriz de modo
 * que los datos de distintas *filas* queden desemparejados
 * utilizando el algoritmo de Durstenfeld (coste O(n)).
 *
 * @param {A2:C20} matriz
 * Rango de datos a barajar (ofuscar)
 *
 * @return rango de datos barajado
 *
 * @customfunction
 *
 * CC BY-NC-SA Pablo Felip @pfelipm
 */

function BARAJARDATOSFIL(matriz){

  // Comprobaciones previas
  
  // ¿Tenemos parámetro?
  if (matriz == null) {return 'Parámetros incorrectos';}
 
  // ¿Es una matriz?
  if (Array.isArray(matriz) == false) {return 'Se esperaba una matriz';}
  
  // ¿Tiene al menos 2 filas?
  if (matriz.length <= 1) {return 'Se esperaba una matriz de más de 1 fila';}
  
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
  return matriz;  
}