/**
 * Consolida los datos reagrupando por dimensión (despivotar)
 *
 * @return Datos agrupados (disposición horizontal >> vertical)
 * @param {A1:D10} matriz Matriz bidimensional, con fila de encabezado.
 * @param {1} numCol Número de columnas fijas a la izquierda.
 * @param {"Nombre"\"Pregunta"\"Respuesta"} encabezados Nuevas etiquetas de encabezado de cada columna.
 *
 * @customfunction
 */

// Tomado de https://stackoverflow.com/a/55869948

function UNPIVOT(matriz, numCol, encabezados) {
  var out = matriz.reduce(function(acc, row) {
    var left = row.splice(0, numCol); //static columns on left
    row.forEach(function(col, i) {
      acc.push(left.concat([acc[0][i + numCol], col])); //concat left and unpivoted right and push as new array to accumulator
    });
    return acc;
  }, matriz.splice(0, 1));//headers in arr as initial value
  encabezados ? out.splice(0, 1, encabezados[0]) : null; //use custom headers, if present.
  return out;
}