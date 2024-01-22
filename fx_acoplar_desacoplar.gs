/**
 * Esta función desacopla las filas de un intervalo de datos que contiene valores múltiples, delimitados por la 
 * secuencia de caracteres indicada, en algunas de sus columnas. Se ha diseñado principalmente para facilitar el
 * tratamiento estadístico de las respuestas a un formulario cuando algunas de sus preguntas admiten múltiples opciones 
 * (casillas de verificación), que en ese caso están separadas por la secuencia delimitadora ", " (coma espacio).
 * Tras ser desacopladas, las respuestas (filas) se repiten en el intervalo resultante para cada combinación posible
 * de los valores múltiples únicos que se encuentran en las columnas especificadas.
 * 
 * @param {A1:D10}    intervalo       Intervalo de datos.
 * @param {VERDADERO} encabezado      Indica si el rango tiene una fila de encabezado con etiquetas para cada columna ([VERDADERO] | FALSO).
 * @param {", "}      separador       Secuencia de caracteres que separa los valores múltiples. Opcional, si se omite se utiliza ", " (coma espacio).
 * @param {VERDADERO} forzarNum       Indica si los valores múltiples identificados como numéricos se devolverán como números en lugar de texto ([VERDADERO] | FALSO).
 * @param {2}         columna         Número de orden, desde la izquierda, de la columna dentro del intervalo que contiene valores múltiples a descoplar,
 *                                    si se indican valores negativos para las columnas se generarán filas desacopladas también para los valores duplicados.
 * @param {4}         [más_columnas]  Columnas adicionales, opcionales, que contienen valores múltiples a desacoplar, separadas por ";".
 *
 * @return                            Intervalo de datos desacoplados
 *
 * @customfunction
 *
 * Artículo:    https://pablofelip.online/desacoplar-acoplar/
 * Repositorio: https://github.com/pfelipm/fxdesacoplar-acoplar
 * 
 * MODIFICADO 01/02/23:
 *  - Si el nº de columna es positivo únicamente se duplicarán filas para los valores únicos, si es negativo
 *    también se generarán duplicados para cada elemento repetido.
 *  - Se ignoran las filas del intervalo de entrada s que están totalmente vacías.
 *  - Se devuelven valores numéricos cuando pueden ser convertidos a números (¡mejor parametrizar!)
 * 
 * MIT License
 * Copyright (c) 2023 Pablo Felip Monferrer (@pfelipm)
 */
function DESACOPLAR(intervalo, encabezado, separador, forzarNum, columna, ...masColumnas) {

  // Control de parámetros inicial

  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto';
  if (typeof forzarNum != 'boolean') forzarNum = true;
  
  // Montar vector de columnas, truncar números no enteros, si los hay
  const columnas = (typeof columna != 'undefined' ? [columna, ...masColumnas] : [...masColumnas])
    .map(columna => typeof columna == 'number' ? Math.trunc(columna) : columna);
  if (columnas.length == 0) throw 'No se han indicado columnas a descoplar';
  if (columnas.some(col => typeof col != 'number' || col == 0)) throw 'Las columnas deben indicarse mediante números enteros distintos de cero';
  if (Math.max(...columnas) > intervalo[0].length
    || Math.min(...columnas) < -intervalo[0].length) throw 'Al menos una columna está fuera del intervalo de datos a desacoplar';

  // En la fx original se utiliza un Set para evitar duplicados, pero ahora el nº de columna puede ser positivo o negativo,
  // lo que cuando hay valores de columna (+/-) podría hacer que la función devolviera las columnas desordenadas.
  const colArray = [];
  columnas.sort((col1, col2) => Math.abs(col1) - Math.abs(col2))
    .forEach(col => {
      // No se tienen en cuenta valores de columna repetidos
      if (!(colArray.includes(col) || colArray.includes(-col))) colArray.push(col);
    });

  // Listos para comenzar

  if (encabezado) encabezado = intervalo.shift();

  const intervaloDesacoplado = [];

  // Recorramos el intervalo fila a fila

  intervalo.forEach(fila => {

    // No trataremos filas en blanco
    if (fila.some(celda => celda != '')) {

      // Enumerar los valores únicos en cada columna que se ha indicado contiene datos múltiples

      const opciones = [];
      colArray.forEach(col => {

        // Eliminar opciones duplicadas, si las hay, en cada columna gracias al uso de un nuevo conjunto
        if (col < 0) {
          // Se admiten valores repetidos
          const opcionesVector = [];
          String(fila[Math.abs(col) - 1]).split(separador).forEach(opcion => opcionesVector.push(opcion));
          opciones.push(opcionesVector);
        } else {
          // No se admiten valores repetidos
          const opcionesSet = new Set();
          String(fila[col - 1]).split(separador).forEach(opcion => opcionesSet.add(opcion)); // split solo funciona con string, convertimos números
          opciones.push([...opcionesSet]); // también opciones.push(Array.from(opcionesSet))
        }

      });

      // Ahora desacoplamos la respuesta (fila) mediante una IIFE recursiva 🔄
      // que genera un vector para todas las posibles combinaciones (vectores) de respuestas
      // de las columnas con valores múltiples
      // Ej:
      //     ENTRADA: vector = [ [a, b], [1, 2] ]
      //     SALIDA:  combinaciones = [ [a, 1], [a, 2], [b, 1], [b, 2] ]

      const combinaciones = (function combinar(vector) {

        if (vector.length == 1) {

          // Fin del proceso recursivo

          const resultado = [];
          vector[0].forEach(opcion => resultado.push([opcion]));
          return resultado;
        }

        else {

          // El resultado se calcula recursivamente

          const resultado = [];
          const subvector = vector.splice(0, 1)[0];
          const subresultado = combinar(vector);

          // Composición de resultados en la secuencia recursiva >> generación de vector de combinaciones

          subvector.forEach(e1 => subresultado.forEach(e2 => resultado.push([e1, ...e2])));

          return resultado;

        }
      })(opciones);

      // Ahora hay que generar las filas repetidas para cada combinación de datos múltiples
      // Ej:
      //     ENTRADA: combinaciones = [ [a, 1], [a, 2], [b, 1], [b, 2] ]
      //     SALIDA:  respuestaDesacoplada = [ [Pablo, a, 1, Tarde], [Pablo, a, 2, Tarde], [Pablo, b, 1, Tarde], [Pablo, b, 2, Tarde] ]

      const respuestaDesacoplada = combinaciones.map(combinacion => {

        let colOpciones = 0;
        const filaDesacoplada = [];
        fila.forEach((valor, columna) => {

          // Tomar columna de la fila original o combinación de datos generada anteriormente
          // correspondiente a cada una de las columnas con valores múltiples

          if (!colArray.includes(columna + 1) && !colArray.includes(-(columna + 1))) filaDesacoplada.push(valor);
          else filaDesacoplada.push(combinacion[colOpciones++]);

        });
        // Devolver valores encontrados en las columnas a desacoplar que pueden convertirse a números como tales (puede ser delicado, se ha parametrizado)
        // https://dev.to/sanchithasr/7-ways-to-convert-a-string-to-number-in-javascript-4l
        return filaDesacoplada.map((valor, columna) => (isNaN(valor) || !forzarNum || !colArray.find(col => Math.abs(col) == columna + 1)) ? valor : Number(valor));
      });

      // Se desestructura (...) respuestaDesacoplada dado que combinaciones.map es [[]]

      intervaloDesacoplado.push(...respuestaDesacoplada);

    }

  });

  // Si hay fila de encabezados, colocar en 1ª posición en la matriz de resultados

  return encabezado.map ? [encabezado, ...intervaloDesacoplado] : intervaloDesacoplado;

}

/**
 * Esta función acopla (combina) las filas de un intervalo de datos que corresponden a una misma entidad. Para ello, 
 * se debe indicar la columna (o columnas) *clave* que identifican los datos de cada entidad única. Los valores registrados
 * en el resto de columnas se agruparán, para cada una de ellas, utilizando como delimitador la secuencia de caracteres 
 * indicada. Se trata de una función que realiza una operación complementaria a DESACOPLAR(), aunque no perfectamente simétrica.
 * 
 * @param {A1:D10}    intervalo         Intervalo de datos.
 * @param {VERDADERO} encabezado        Indica si el rango tiene una fila de encabezado con etiquetas para cada columna ([VERDADERO] | FALSO).
 * @param {", "}      separador         Secuencia de caracteres a emplear como separador de los valores múltiples. Opcional, si se omite se utiliza ", " (coma espacio).
 * @param {VERDADERO} permitirRepetidos Indica si se permitirán valores consolidados repetidos dentro de las celdas de una fila (VERDADERO | [FALSO]).
 * @param {1}         columna           Número de orden, desde la izquierda, de la columna clave que identifica los datos de la fila como únicos.
 * @param {2}         [más_columnas]    Columnas clave adicionales, opcionales, que actúan como identificadores únicos, separadas por ";".
 *
 * @return                              Intervalo de datos desacoplados
 *
 * @customfunction
 * 
 * MODIFICADO 06/02/23:
 *  - Se añade un parámetro para agrupar o no los valores repetidos dentro de una celda con datos consolidados.
 *  - Se ignoran las filas del intervalo de entrada s que están totalmente vacías.
 *
 * MIT License
 * Copyright (c) 2020 Pablo Felip Monferrer (@pfelipm)
 */
function ACOPLAR(intervalo, encabezado, separador, permitirRepetidos, columna, ...masColumnas) {

  // Control de parámetros inicial

  if (typeof intervalo == 'undefined' || !Array.isArray(intervalo)) throw 'No se ha indicado un intervalo';
  if (typeof encabezado != 'boolean') encabezado = true;
  if (intervalo.length == 1 && encabezado) throw 'El intervalo es demasiado pequeño, añade más filas';
  separador = separador || ', ';
  if (typeof separador != 'string') throw 'El separador no es del tipo correcto';
  if (typeof permitirRepetidos != 'boolean') forzarNum = false;
  
  // Montar vector de columnas, truncar números no enteros, si los hay
  const columnas = (typeof columna != 'undefined' ? [columna, ...masColumnas] : [...masColumnas])
    .map(columna => typeof columna == 'number' ? Math.trunc(columna) : columna);
  if (columnas.length == 0) throw 'No se han indicado columnas clave';
  if (columnas.some(col => typeof col != 'number' || col < 1)) throw 'Las columnas clave deben indicarse mediante números enteros positivos';
  if (Math.max(...columnas) > intervalo[0].length) throw 'Al menos una columna clave está fuera del intervalo de datos a acoplar';

  // Se construye un conjunto (set) para evitar automáticamente duplicados en columnas CLAVE

  const colSet = new Set();
  columnas.forEach(col => colSet.add(col - 1));

  // ...y en este conjunto se identifican las columnas susceptibles de contener valores que deben concatenarse

  const colNoClaveSet = new Set();
  for (let col = 0; col < intervalo[0].length; col++) {

    if (!colSet.has(col)) colNoClaveSet.add(col);

  }

  // Listos para comenzar

  if (encabezado) encabezado = intervalo.shift();

  const intervaloAcoplado = [];

  // 1ª pasada: recorremos el intervalo fila a fila para identificar entidades (concatenación de columnas clave) únicas

  const entidadesClave = new Set();
  intervalo.forEach(fila => {
  
    // No trataremos filas en blanco
    if (fila.some(celda => celda != '')) {

      const clave = [];
      // ⚠️ A la hora de diferenciar dos entidades únicas (filas) usando una serie de columnas clave:
      //    a) No basta con concatenar los valores de las columnas clave como cadenas y simplemente compararlas. Ejemplo:
      //       clave fila 1 → col1 = 'pablo' col2 = 'felip'     >> Clave compuesta: 'pablofelip'
      //       clave fila 2 → col1 = 'pa'    col2 = 'blofelip'  >> Clave compuesta: 'pablofelip'
      //       ✖️ Misma clave compuesta, pero entidades diferentes
      //    b) No basta con con unir los valores de las columnas clave como cadenas utilizando un carácter delimitador. Ejemplo ('/'):
      //       clave fila 1 → col1 = 'pablo/' col2 = 'felip'    >> Clave compuesta: 'pablo//felip' 
      //       clave fila 2 → col1 = 'pablo'  col2 = '/felip'   >> Clave compuesta: 'pablo//felip'
      //       ✖️ Misma clave compuesta, pero entidades diferentes
      //    c) No es totalmente apropiado eliminar espacios antes y después de valores clave y unirlos usando un espacio delimitador (' '):
      //       clave fila 1 → col1 = ' pablo' col2 = 'felip'    >> Clave compuesta: 'pablo felip'
      //       clave fila 2 → col1 = 'pablo'  col2 = 'felip'    >> Clave compuesta: 'pablo felip'
      //       ✖️ Misma clave compuesta, pero entidades estrictamente diferentes (a menos que espacios anteriores y posteriores no importen)
      // 💡 En su lugar, se generan vectores con valores de columnas clave y se comparan sus versiones transformadas en cadenas JSON.
      for (const col of colSet) clave.push(String(fila[col]))
      entidadesClave.add(JSON.stringify(clave));
    
    }

  });

  // 2ª pasada: obtener filas para cada clave única, combinar columnas no-clave y generar filas resultado

  for (const clave of entidadesClave) {

    const filasEntidad = intervalo.filter(fila => {

      const claveActual = [];
      for (const col of colSet) claveActual.push(String(fila[col]));
      return clave == JSON.stringify(claveActual);

    });

    // Acoplar todas las filas de cada entidad, concatenando valores en columnas no-clave con separador indicado

    const filaAcoplada = filasEntidad[0];  // Se toma la 1ª fila del grupo como base
    
    // Para identificar los valores a consolidar usaremos un array si se admiten duplicados o un conjunto en caso contrario
    const noClaveSets = [...Array(colNoClaveSet.size)].map(fila => permitirRepetidos ? [] : new Set());
    
    filasEntidad.forEach(fila => {

      let conjunto = 0;
      for (const col of colNoClaveSet) {
        if (permitirRepetidos) noClaveSets[conjunto++].push(String(fila[col]));
        else noClaveSets[conjunto++].add(String(fila[col]));
      }

    });

    // Set >> Vector >> Cadena única con separador

    let conjunto = 0;
    for (const col of colNoClaveSet) {
      if (permitirRepetidos) filaAcoplada[col] = noClaveSets[conjunto++].join(separador);
      else filaAcoplada[col] = [...noClaveSets[conjunto++]].join(separador);
    }

    intervaloAcoplado.push(filaAcoplada);

  }

  return encabezado.map ? [encabezado, ...intervaloAcoplado] : intervaloAcoplado;

}