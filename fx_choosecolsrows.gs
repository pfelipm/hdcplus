/**
 * Las funciones integradas CHOOSEROWS() y CHOOSECOLS() permiten seleccionar un subconjunto de filas
 * o columnas, respectivamente, de un intervalo de datos, pero es necesario enumerarlas individualmente.
 * Para facilitar el trabajo con conjuntos de datos extensos, HDCP_CHOOSECOLSROWS() integra ambas funciones
 * en una sola y ofrece una mayor flexibilidad a la hora de indicar las filas y columnas que se desean extraer.
 *  
 * Esta función crea una matriz a partir de las filas y columnas seleccionadas de un intervalo.
 * Para designar las filas y columnas es posible utilizar: [I] una lista de valores numéricos que representan
 * su posición desde el extremo superior izquierdo del intervalo de entrada (valores positivos) o inferior
 * derecho (valores negativos), comenzando siempre por 1, por ejemplo: {1;2;4;-3;1}; [II] una lista de
 * expresiones de texto entrecomilladas ("..-..") que denotan intervalos, abiertos o cerrados, por ejemplo:
 * "3-12" (desde la fila/columna 3 a la 12), "-12" (desde la fila/columna 1 hasta la 12), "3-" o simplemente "3"
 * (desde la fila/columna 3 hasta la última).
 * 
 * Para seleccionar las columnas también es posible [III] utilizar una lista de etiquetas de texto, que se
 * cotejarán con las que se encuentran en la fila de encabezado del intervalo de datos. Estas etiquetas pueden
 * ser valores literales, referencias a otras celdas o una combinación de ambas cosas, por ejemplo:
 * {"Nombre";"Edad";"Grupo"} o {B1:D1\I1:K1\"Grupo"}.
 * 
 * En todos los casos [I-III] las listas de elementos pueden construirse usando tanto vectores fila (como {1\2\3}) como
 * columna (como {1;2;3}), pero no utilizando una combinación de ambos.
 * 
 * Es posible designar una fila o columna más de una vez, lo que hará que se muestre repetida en la
 * matriz devuelta por la función. También se admite el cambio de orden de las filas y columnas con respecto
 * a sus posiciones en el intervalo de datos original.
 * 
 * IMPORTANTE: Los caracteres " [ " y " ] " que se muestran en la fórmula de EJEMPLO deben sustituirse por los
 * símbolos " { " y " } " propios de la sintaxis de matrices de Google Sheets.
 * de Google.
 * 
 * @param {A2:P20}                intervalo       Intervalo de datos del que extraer un subconjunto de filas y columnas.
 * @param {VERDADERO}             usa_etiquetas   Activa el modo de selección de columnas usando las etiquetas de la fila de encabezado del intervalo. 
 *                                                La comparación es estricta y diferencia mayúsculas de minúsculas (VERDADERO o [FALSO]).
 * @param {["-3";6;7;"10-12"]}    columnas        Vector ({..\.. } o {..;..}) de índices de columnas, comenzando por 1, intervalos de índices como cadenas de
 *                                                texto ("..-..") o etiquetas de encabezado (requiere usa_etiquetas = VERDADERO).
 *                                                Se seleccionarán todas las columnas si se omite.
 * @param {["1-5";6;7;"-12"]}     filas           Vector ({..\.. } o {..;..}) de índices de filas, comenzando por 1, o intervalos de índices como cadenas de
 *                                                texto ("..-.."). Se seleccionarán todas las filas si se omite.

 *                                             
 * @return Matriz de datos compuesta únicamente por las filas y columnas seleccionadas.
 * 
 * @customfunction
 */
function HDCP_CHOOSECOLSROWS(intervalo, usa_etiquetas = false, columnas, filas) {

  // Comprobaciones iniciales sobre los argumentos
  if (!intervalo.map) throw "Intervalo no válido";
  if (columnas?.map && columnas.length > 1 && columnas[0].length > 1) throw 'El argumento «columnas» debe ser un valor único o un vector fila o columna';
  if (filas?.map && filas.length > 1 && filas[0].length > 1) throw 'El argumento «filas» debe ser un valor único o un vector fila o columna';
  if (!typeof usa_etiquetas == 'boolean') throw 'El argumento «etiquetas» debe ser un valor VERDADERO o FALSO';
  if (!typeof encabezados == 'boolean') throw 'El argumento «encabezados» debe ser un valor VERDADERO o FALSO';
  
  /**
   * Función auxiliar para procesar los argumentos de selección de de filas y columnas.
   * Interpreta la lista de filas o columnas seleccionadas y devuelve
   * un vector con los numeros naturales correspondientes.
   * @param   {Array.<number>|Array.<string>} vector
   * @param   {number}                        longitud
   * @param   {string}                        contexto  Literal 'filas' o 'columnas' para tratamiento excepciones y análisis en modo encabezados.
   * @return  {Array.<number>}
   */
  function interpretarVectorIndices(vector, longitud, contexto) {

    // Si el argumento se ha omitido o es un valor único 0, seleccionar todas las filas/columnas
    if (!vector) vector = [...Array(longitud).keys()].map(indice => indice + 1);

    // ¡Si el argumento es un array de 1 solo elemento, HDCP_CHOOSECOLSROWS recibirá un valor que no es de tipo array!
    if (!vector.map) vector = [vector];

    // Se usa flat() para admitir tanto vectores filas {\} como columna {;} como lista de índices de fila/columna
    return vector.flat().reduce((resultado, indice) => {

      if (typeof indice == 'number') {

        // ### El selector de fila/columna es un valor numérico positivo o negativo ###

        if(Math.abs(indice) > longitud) throw `Índice ${indice} fuera de rango en argumento «${contexto}»`;
        // Índices negativos suponen comenzar a contar desde la última fila o columna del intervalo
        else return resultado.concat(indice < 0 ? longitud + indice + 1 : indice);
        
      } else if (typeof indice == 'string' && (!usa_etiquetas || contexto == 'filas')) {

        // ### El selector de fila/columna es una expresión de texto que indica un intervalo de filas/columnas ###
        
        let [inicio, fin] = indice.split('-');

        // Sustituir por primera o última fil/col en caso de intervalos abiertos (ej; "-4" o "5-")
        inicio = Number(!inicio ? 1 : inicio);
        fin = Number(!fin ? longitud : fin);

        // Corrección en el caso de índices cruzados (ej: "4-1"), opto por desactivarla dado que ocasiones mensajes de error confusos cuando
        // solo se indica un índice de fila o columna superior a la longitud, dado que pasa a ser interpretado como índice final al ser intercambiado por
        // el valor del índice máximo. En su lugar lanzo excepción cuando los índices están cruzados. 
        // [inicio, fin] = [Math.min(inicio, fin), Math.max(inicio, fin)];
        
        // Excepción si índices fuera de rango
        if (inicio < 1 || inicio > longitud) throw `Índice inicial ${inicio} fuera de rango en expresión de intervalo "${indice}" de argumento «${contexto}»`;
        if (fin < 1 || fin > longitud) throw `Índice final ${fin} fuera de rango en expresión de intervalo "${indice}" de argumento «${contexto}»`;
        if (inicio > fin) throw `El índice final ${fin} es anterior al inicial ${inicio} en expresión de intervalo "${indice}" de argumento «${contexto}»`;

        // Excepción si no han podido determinarse índices numéricos a apartir de la expresión del selector de fila/columna
        if(Number.isNaN(inicio)) throw `El índice inicial en expresión de intervalo "${indice}" de argumento «${contexto}» no es un número`;
        if(Number.isNaN(fin)) throw `El índice final en expresión de intervalo "${indice}" de argumento «${contexto}» no es un número`;

        // Devolver secuencia de índices consecutivos entre inicio y fin
        return resultado.concat([...Array(fin - inicio + 1).keys()].map(indiceIntervalo => inicio + indiceIntervalo));

      } else if (typeof indice == 'string' && contexto == 'columnas' && usa_etiquetas) {

        // ### El selector de fila/columna es una cadena de texto que enumera los encabezados de las columnas que se desean devolver ###

        // Identificar la posición de la columna cuyo texto de encabezado es coincidente con el valor de la variable "indice"
        const indiceColumna = intervalo[0].indexOf(indice);
        if (indiceColumna == -1) throw `La etiqueta de encabezado ${indice} no se encuentra en la primera fila del intervalo`;
        else return resultado.concat(indiceColumna + 1);

      } else throw `Error inesperado, revisa el argumento «${contexto}»`;

    }, []);

  }

  try {

    // Interpretar argumentos de selección de filas y columnas
    const filasSeleccionadas = interpretarVectorIndices(filas, intervalo.length, 'filas')
    const columnasSeleccionadas = interpretarVectorIndices(columnas, intervalo[0].length, 'columnas');

    // Devolver matriz de datos con las filas y columnas seleccionadas
    return filasSeleccionadas
      // Filetear filas
      .reduce((matrizReducida, fila) => matrizReducida.concat([intervalo[fila - 1]]), [])
      // Filetear columnas
      .map(fila => columnasSeleccionadas.reduce((filaReducida, columna) => filaReducida.concat(fila[columna - 1]), []));

  } catch (e) {
    throw typeof e == 'string'
      ? e
      : `Error: "${e}". Puede que hayas combinado los símbolos ";" y "\\"
        de manera incorrecta en los argumentos de selección de columnas y filas {..} `;
  }

}