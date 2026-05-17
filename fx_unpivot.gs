/**
 * Consolida los datos reagrupando por dimensión (despivotaje).
 *
 * @return Datos agrupados (disposición horizontal >> vertical)
 * @param {A1:D10} intervalo Intervalo de datos bidimensional, con fila de encabezado.
 * @param {1} numCol Número de columnas fijas a la izquierda.
 * @param {"Nombre"\"Pregunta"\"Respuesta"} encabezadosNuevos [Opcional] Nuevas etiquetas de encabezado de cada columna.
 *
 * @customfunction
 */

function UNPIVOT(intervalo, numCol, encabezadosNuevos) {

  // Control de parámetros
  if (typeof intervalo == 'undefined' || typeof numCol == 'undefined') { 
    return '!Faltan argumentos';
  }
  if (typeof numCol != 'number') {
    return  '!Nº columnas fijas inválido';
  }
 
  // 1. Extraer y guardar cabeceras originales
  var originalHeaders = intervalo.splice(0, 1)[0];
  var out = [];
  
  // 2. Construir la nueva fila de encabezados (ancho: numCol + 2)
  var finalHeaders = [];
  if (encabezadosNuevos && encabezadosNuevos[0] && encabezadosNuevos[0].length >= (numCol + 2)) {
    // Si el usuario pasó etiquetas suficientes (vía panel o manual), las usamos todas
    finalHeaders = encabezadosNuevos[0].slice(0, numCol + 2);
  } else {
    // Si no, reutilizamos las fijas originales y añadimos genéricas/parciales
    finalHeaders = originalHeaders.slice(0, numCol);
    var extra = (encabezadosNuevos && encabezadosNuevos[0]) ? encabezadosNuevos[0] : [];
    
    // Atributo (nombre de la columna pivotada)
    finalHeaders.push(extra[numCol] || "Atributo");
    // Valor
    finalHeaders.push(extra[numCol + 1] || "Valor");
  }
  out.push(finalHeaders);

  // 3. Procesar datos
  intervalo.forEach(function(row) {
    var left = row.splice(0, numCol); // Columnas estáticas
    row.forEach(function(col, i) {
      // Por cada columna a pivotar, creamos una nueva fila: [fijas] + [nombre_col_original] + [valor]
      out.push(left.concat([originalHeaders[i + numCol], col]));
    });
  });
  
  return out;
}