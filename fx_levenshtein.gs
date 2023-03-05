/**
 * Dos funciones personalizadas para hojas de cálculo de Google basadas 
 * que utilizan la distancia de Levenshtein.
 * 
 * MIT License
 * Copyright (c) 2023 Pablo Felip Monferrer (@pfelipm)
 */



/**
 * Calcula la distancia de Levenshtein entre dos cadenas de texto. Opcionalmente, puede
 * diferenciar o no mayúsculas y minúsculas, asi como permitir transposiciones de caracteres
 * no adyacentes (versión simple de la distancia de Damerau-Levenshtein
 * https://en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance#Optimal_string_alignment_distance).

 * @param {"Pablo"}   c1                  La primera cadena o intervalo de cadenas de texto.
 * @param {"Paula"}   c2                  La segunda cadena o intervalo de cadenas de texto.
 * @param {VERDADERO} permiteTrans        Indica si se admiten transposiciones ([VERDADERO] | FALSO).
 * @param {VERDADERO} distingueMayusculas Indica si deben diferenciarse mayúsculas de minúsculas ([VERDADERO] | FALSO).
 * @param {VERDADERO} fuerzaTexto         Indica si se deben tratar de forzar a texto los valores numéricos ([VERDADERO] | FALSO).
 * @param {1}         costeInsercion      El valor numérico que representa el coste de la operación de inserción [1].
 * @param {1}         costeEliminacion    El valor numérico que representa el coste de la operación de eliminación [1].
 * @param {1}         costeSustitucion    Valor numérico que representa el coste de la operación de sustitución [1].
 * @param {1}         costeTransposicion  Valor numérico que representa el coste de la operación de transposicion [1].
 *
 * @return                                La distancia o distancias de edición entre las cadenas comparadas.
 *
 * @customfunction
 */
function DISTANCIA_EDICION(
  c1, c2,
  permiteTrans, maysImportan, fuerzaTexto,
  costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
) {

  // Comprobación de parámetros y valores por defecto, se hace de este modo para cazar el uso de [;]
  // para "saltar" parámetros (se reciben como ''),
  permiteTrans = typeof permiteTrans != 'boolean' ? true : permiteTrans;
  distingueMayusculas = typeof distingueMayusculas != 'boolean' ? true : distingueMayusculas;
  fuerzaTexto = typeof fuerzaTexto != 'boolean' ? true : fuerzaTexto;
  costeInsercion = typeof costeInsercion != 'number' ? 1 : costeInsercion;
  costeEliminacion = typeof costeEliminacion != 'number' ? 1 : costeEliminacion;
  costeSustitucion = typeof costeSustitucion != 'number' ? 1 : costeSustitucion;
  costeTransposicion = typeof costeTransposicion != 'number' ? 1 : costeTransposicion;
  
  if (Array.isArray(c1) && Array.isArray(c2)) {
    
    // Intervalo / Intervalo
    if (c1.length != c2.length || c1[0].length != c2[0].length) throw 'Las dimensiones de los intervalos que contiene las cadenas de texto no coinciden'
    return c1.map((vectorFil, fil) => vectorFil.map((cadena, col) => distanciaLevenshtein_(
      cadena, c2[fil][col],
      permiteTrans, maysImportan, fuerzaTexto,
      costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
    )));

  } else if (Array.isArray(c1) && typeof c2 == 'string') {
    
    // Intervalo / Cadena
    return c1.map(vectorFil => vectorFil.map(cadena => distanciaLevenshtein_(
      cadena, c2,
      permiteTrans, maysImportan, fuerzaTexto,
      costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
    )));

  } else if (typeof c1 == 'string' && Array.isArray(c2)) {
    
    // Cadena / Intervalo
    return c2.map(vectorFil => vectorFil.map(cadena => distanciaLevenshtein_(
      c2, cadena,
      permiteTrans, maysImportan, fuerzaTexto, 
      costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
    )));

  } else if (typeof c1 == 'string' && typeof c2 == 'string' || fuerzaTexto) {
   
    // Cadena / Cadena
    return distanciaLevenshtein_(
      c1, c2,
      permiteTrans, maysImportan, fuerzaTexto,
      costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
    );

  } else throw 'Error en los parámetros, revisa las cadenas o intervalos de cadenas de texto cuyas distancias deseas medir';

}



/**
 * Devuelve la cadena o cadenas más próximas a una o más específicas, de acuerdo con la distancia de Levenshtein,
 * de entre un conjunto de secuencias de texto de referencia candidatas. Opcionalmente, puede diferenciar o no mayúsculas
 * y minúsculas, así como permitir transposiciones de caracteres no adyacentes, (versión simple de la distancia de Damerau-Levenshtein
 * https://en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance#Optimal_string_alignment_distance).
 *
 * @param {"Pablo"}   cadena              La primera cadena o intervalo de cadenas de texto.
 * @param {B2:B10}    referencia          El intervalo de cadenas de texto de referencia.
 * @param {3}         numCadenas          Nº de cadenas que se devolverán [1]
 * @param {FALSO}     devuelveDistancia   Indica si también se debe devolver la distancia calculada para cada valor de referencia (VERDADERO | [FALSO]).
 * @param {VERDADERO} permiteTrans        Indica si se admiten transposiciones ([VERDADERO] | FALSO).
 * @param {FALSO}     distingueMayusculas Indica si deben diferenciarse mayúsculas de minúsculas ([VERDADERO] | FALSO).
 * @param {VERDADERO} fuerzaTexto         Indica si se deben tratar de forzar a texto los valores numéricos ([VERDADERO] | FALSO).
 * @param {1}         costeInsercion      El valor numérico que representa el coste de la operación de inserción [1].
 * @param {1}         costeEliminacion    El valor numérico que representa el coste de la operación de eliminación [1].
 * @param {1}         costeSustitucion    El valor numérico que representa el coste de la operación de sustitución [1].
 * @param {1}         costeTransposicion  El valor numérico que representa el coste de la operación de transposicion [1].
 *
 * @return                                Lista de cadenas similares de la longitud indicada, ordenada por distancia (creciente) + distancia numérica (opcional).
 *
 * @customfunction
 */
function DISTANCIA_EDICION_MINIMA(
  cadena, referencia, numCadenas, devuelveDistancia,
  permiteTrans, maysImportan, fuerzaTexto,
  costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
) {

  // Comprobación de parámetros y valores por defecto, se hace de este modo para cazar el uso de [;]
  // para "saltar" parámetros (se reciben como ''),
  numCadenas = Math.abs(Math.trunc((typeof numCadenas != 'number' || numCadenas < 1 ? 1 : numCadenas)));
  devuelveDistancia = typeof devuelveDistancia != 'boolean' ? false : devuelveDistancia;
  permiteTrans = typeof permiteTrans != 'boolean' ? true : permiteTrans;
  distingueMayusculas = typeof distingueMayusculas != 'boolean' ? true : distingueMayusculas;
  fuerzaTexto = typeof fuerzaTexto != 'boolean' ? true : fuerzaTexto;
  costeInsercion = typeof costeInsercion != 'number' ? 1 : costeInsercion;
  costeEliminacion = typeof costeEliminacion != 'number' ? 1 : costeEliminacion;
  costeSustitucion = typeof costeSustitucion != 'number' ? 1 : costeSustitucion;
  costeTransposicion = typeof costeTransposicion != 'number' ? 1 : costeTransposicion;

  /**
   * Función auxiliar, devuelve las distancias medidas entre una cadena de texto dada y cada una
   * de las cadenas de texto facilitadas en un array de candidatas, en orden creciente y, en caso
   * de empate, de posición en el vector.
   * 
   * @param {string}      cadena      Cadena a comprobar
   * @param {string[][]}  referencia  Matriz de cadenas candidatas
   * 
   * @return {Object[]}               Vector de objetos { cadena_candidata, distancia } 
   */
  function obtenerDistancias(cadena, referencia) {

    // Se admiten que el intervalo que contiene las celdas con las cadenas de referencia sea bidimensional
    return referencia.map(vectorFil => vectorFil.map(candidata => (
      {
        cadena: candidata,
        distancia: distanciaLevenshtein_(
          cadena, candidata,
          permiteTrans, maysImportan, costeInsercion, costeEliminacion, costeSustitucion, costeTransposicion
        )
      }
    ))).flat()
      // Ordenar por distancia (creciente) y recortar longitud de respuesta
      .sort((cadena1, cadena2) => cadena1.distancia - cadena2.distancia)
      .slice(0, numCadenas);

  }

  if (Array.isArray(cadena) && Array.isArray(referencia)) {

    // Cadena / Intervalo
    // Se utilizan 2 reduce() anidados para construir la matriz resultado dado que puede tener dimensiones
    // distintas a la matriz de cadenas a tratar según parámetro cadena, dependiendo de parámetros numCadenas y devuelveDistancia.
    const resultado = cadena.reduce((resultado, vectorFil) => resultado.concat(
      vectorFil.reduce((resultadoFil, cadena) => {
        const resultadoCadena = obtenerDistancias(cadena, referencia)
          .map(candidata => devuelveDistancia ? [candidata.cadena, candidata.distancia] : [candidata.cadena]);
        // Expandir matriz resultado parcial añadiendo las columnas necesarias
        return resultadoFil.length == 0
          ? resultadoCadena
          : resultadoFil.map((fila, nFil) => fila.concat(resultadoCadena[nFil]));
      }, [])
    ), []);

    // Versión inicial con 2 forEach() anidados, ¿mejora la inteligibilidad?
    /*
    let resultado = [];
    cadena.forEach(vectorFil => {

      let resultadoFil = [];
      vectorFil.forEach(cadena => {

        const resultadoCadena = obtenerDistancias(cadena, referencia)
          .map(candidata => devuelveDistancia ? [candidata.cadena, candidata.distancia] : [candidata.cadena]);  
        // Expandir matriz resultado parcial añadiendo las columnas necesarias
        resultadoFil = resultadoFil.length == 0
          ? resultadoCadena
          : resultadoFil.map((fila, nFil) => fila.concat(resultadoCadena[nFil]));
      
      });
      // Expandir matriz resultado general añadiendo filas
      resultado = resultado.concat(resultadoFil);

    });
    */
    
    return resultado;

  } else if (typeof cadena == 'string' && Array.isArray(referencia)) {
    // Intervalo / Intervalo
    return obtenerDistancias(cadena, referencia).map(candidata => devuelveDistancia ? [candidata.cadena, candidata.distancia] : candidata.cadena);

  } else throw 'Error en los parámetros, revisa las cadenas o intervalos de cadenas de texto cuyas distancias deseas medir';

}

/**
 * Calcula la distancia de Levenshtein entre dos cadenas de texto. Soporta operaciones de sustitución, eliminación,
 * adición y transposiciones no adyacentes de caracteres (versión simple de la distancia de Damerau-Levenshtein
 * https://en.wikipedia.org/wiki/Damerau%E2%80%93Levenshtein_distance#Optimal_string_alignment_distance). También puede
 * diferenciar o no mayúsculas y minúsculas. Solo utiliza 3 vectores, en lugar de una matriz completa.
 * Basada en una ensoñación de ChatGPT (aunque no sin una larga conversación previa).
 * 
 * Se trata de una función auxiliar privada invocada desde DISTANCIA_EDICION() y MENOR_DISTANCIA_EDICION()
 *
 * @param   {string}    c1                  La primera cadena de texto a comparar.
 * @param   {string}    c2                  La segunda cadena de texto a comparar.
 * @param   {boolean}   [permiteTrans=true] Admite operaciones de transposición de caracteres. Por defecto, verdadero.
 * @param   {boolean}   [maysImportan=true] Diferencia mayúsculas de minúsculas. Por defecto, verdadero.
 * @param   {boolean}   [fuerzaTexto=true]  Indica si se deben tratar de forzar a texto los valores numéricos ([VERDADERO] | FALSO).
 * @param   {number}    [costeIns=1]        El valor numérico que representa el coste de la operación de inserción [1].
 * @param   {number}    [costeEli=1]        El valor numérico que representa el coste de la operación de eliminación [1].
 * @param   {number}    [costeSus=1]        El valor numérico que representa el coste de la operación de sustitución [1].
 * @param   {number}    [costeTrans=1]      El valor numérico que representa el coste de la operación de transposición [1].
 * 
 * @return  {number}                        La distancia de Levenshtein entre las dos cadenas de texto.
 */
function distanciaLevenshtein_(
  c1, c2,
  permiteTrans = true, maysImportan = true, fuerzaTexto = true, 
  costeIns = 1, costeEli = 1, costeSus = 1, costeTrans = 1
) {

  // Se desea forzar conversión a texto de valores numéricos?
  if (fuerzaTexto) {
    if (typeof c1 == 'number') c1 = String(c1);
    if (typeof c2 == 'number') c2 = String(c2);
  }

  // Si no tenemos cadenas no devolveremos nada de manera explícita
  if (typeof c1 == 'string' && typeof c2 == 'string') {

    const n = c1.length;
    const m = c2.length;

    // Casos elementales
    if (c1 == c2) return 0;
    if (n == 0) return m * costeIns;
    if (m == 0) return n * costeEli;

    // Tratamiento de mayúsculas
    if (!maysImportan) { c1 = c1.toUpperCase(); c2 = c2.toUpperCase(); }

    // Vectores para el cálculo, necesitamos v2 para soportar transposiciones
    const v0 = [...Array(m + 1)].map((_, i) => i);
    const v1 = Array(m + 1).fill(0);
    const v2 = Array(m + 1).fill(0);

    for (let i = 0; i < n; i++) {
      v1[0] = i + 1;
      v2[0] = i + 1;
      for (let j = 0; j < m; j++) {
        v1[j + 1] = Math.min(
          v1[j] + costeIns,
          v0[j + 1] + costeEli,
          v0[j] + (c1[i] == c2[j] ? 0 : costeSus));

        if (permiteTrans && i > 0 && j > 0 && c1[i] == c2[j - 1] && c1[i - 1] == c2[j]) {
          v1[j + 1] = Math.min(v1[j + 1], v2[j - 1] + costeTrans);
        }
      }
      for (let j = 0; j <= m; j++) { v0[j] = v1[j]; v1[j] = v2[j]; v2[j] = 0; }
    }

    return v0[m];

  }
  
}