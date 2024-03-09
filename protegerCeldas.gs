// Envoltorios para protegerFxHoja()
function protegerFxHojaAdvertencia() { protegerFxHoja(PROTECCION.modo.advertencia) }
function protegerFxHojaSoloYo() { protegerFxHoja(PROTECCION.modo.soloYo) }

// Envoltorios para eliminarProteccionesFx()
function eliminarProteccionesFxHdCPlus() { eliminarProteccionesFx(true) }
function eliminarProteccionesFxTodas() { eliminarProteccionesFx(false) }

/**
 * Protege las celdas que contienen fórmulas, tratando de agruparlas
 * en un número mínimo de intervalos (reglas de protección)
 * @param {string}  modoProteccion  "advertencia" o "soloyo"
 */
function protegerFxHoja(modoProteccion) {

  // Utilizado para etiquetar las reglas de protección de intervalos
  const opcionesFechaHora = {
    year: 'numeric',
    month: 'numeric',
    day: 'numeric',
    timeZone: Session.getScriptTimeZone(),
    timeZoneName: 'short',
    hour: 'numeric',
    minute: 'numeric',
    hour12: false
  };

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = SpreadsheetApp.getActiveSheet();
  let numIntervalosProtegidos = 0;
  let iniciado = false;

  hdc.toast('Buscando intervalos con fórmulas...', '👀 HdC+ dice:', -1);

  // (Ia) Obtiene una lista de {fila, columna} (índices comienzan en 1),
  // las celdas quedan ordenadas de 1º arriba abajo y 2º izquierda a derecha,
  // lo que puede favorecer el proceso de agrupación de celdas con fórmulas
  // dado que estas suelen disponerse en columnas.
  const listaCeldasFx = [];
  const intervaloFx = hoja.getDataRange().getFormulas();
  for (let nCol = 0; nCol < intervaloFx[0].length; nCol++) {
    for (let nFil = 0; nFil < intervaloFx.length; nFil++) {
      const contenidoCelda = intervaloFx[nFil][nCol];
      if (contenidoCelda) listaCeldasFx.push(
        {
          inicio: { fila: nFil + 1, columna: nCol + 1 },
          fin: { fila: nFil + 1, columna: nCol + 1 }
        }
      );
    }

  }

  /* ALTERNATIVA NO USADA (se considera que Ia optimizará el funcionamiento del proceso).
  // (Ib) Obtiene una lista de {fila, columna} (índices comienzan en 1),
  // las celdas quedan ordenadas 1º de izquierda a derecha y 2º de arriba abajo.
  const listaCeldasFx = hoja.getDataRange().getFormulas().reduce((intervalosFx, fila, nFil) => {
    const fxFila = fila.reduce((subIntervalosFx, celda, nCol) => {
      return celda
        ? subIntervalosFx.concat({ inicio: { fila: nFil + 1, columna: nCol + 1 }, fin: { fila: nFil + 1, columna: nCol + 1 } })
        : subIntervalosFx;
    }, []);
    return fxFila
      ? intervalosFx.concat(fxFila)
      : intervalosFx
  }, []);
  */

  if (listaCeldasFx.length == 0) ui.alert('No se han encontrado fórmulas en esta hoja.', ui.ButtonSet.OK);
  else try {

    /**
     * Función auxiliar.
     * Devuelve el intervalo resultante de la unión de dos intervalos disjuntos, o null si no es posible su unión.
     * @typedef {Object} coordenada
     *    @property {number} fila         Fila de la celda
     *    @property {number} columna      Columna de la celda
     * @typedef {Object} intervalo
     *    @property {coordenada} inicio   Fila y columna de inicio (arriba - izquierda)
     *    @property {coordenada} fin      Fila y columna de fin (abajo - derecha)
     * @param   {intervalo} intervalo1
     * @param   {intervalo} intervalo2
     * @return  {intervalo|null}
     */
    function unirIntervalos(intervalo1, intervalo2) {

      if (
        // Intervalo 1 conincide en filas con intervalo 2
        intervalo1.inicio.fila == intervalo2.inicio.fila &&
        intervalo1.fin.fila == intervalo2.fin.fila
      ) {
        if (intervalo2.fin.columna + 1 == intervalo1.inicio.columna) {
          // Intervalo 2 inmediatamente a la izquierda de intervalo 1
          return {
            inicio: { fila: intervalo1.inicio.fila, columna: intervalo2.inicio.columna },
            fin: { fila: intervalo1.fin.fila, columna: intervalo1.fin.columna }
          };
        }
        else if (intervalo1.fin.columna + 1 == intervalo2.inicio.columna) {
          // Intervalo 2 inmediatamente a la derecha de intervalo 1
          return {
            inicio: { fila: intervalo1.inicio.fila, columna: intervalo1.inicio.columna },
            fin: { fila: intervalo1.fin.fila, columna: intervalo2.fin.columna }
          };

        } else return null;
      }

      if (
        // Intervalo 2 coincide en columnas con intervalo 1
        intervalo1.inicio.columna == intervalo2.inicio.columna &&
        intervalo1.fin.columna == intervalo2.fin.columna
      ) {
        if (intervalo2.fin.fila + 1 == intervalo1.inicio.fila) {
          // Intervalo 2 inmediatamente encima de intervalo 1
          return {
            inicio: { fila: intervalo2.inicio.fila, columna: intervalo1.inicio.columna },
            fin: { fila: intervalo1.fin.fila, columna: intervalo1.fin.columna }
          };
        }
        else if (intervalo1.fin.fila + 1 == intervalo2.inicio.fila) {
          // Intervalo 2 inmediatamente debajo de intervalo 1
          return {
            inicio: { fila: intervalo1.inicio.fila, columna: intervalo1.inicio.columna },
            fin: { fila: intervalo2.fin.fila, columna: intervalo1.fin.columna }
          };

        } else return null;
      }

    }

    // (II) Obtiene el número ¿mínimo? de intervalos rectangulares no solapados de celdas con fórmulas

    let numIntervalos;
    do {
      // Repetir hasta que el vector de intervalos con fórmulas sea estable (no disminuya de tamaño),
      // el coste computacional no es bueno asintóticamente,seguramente esto podría optimizarse.    
      numIntervalos = listaCeldasFx.length;    
      let i = 0, j = 1;
      // Bucle externo para comparar cada elemento de la lista (vector)...
      while (i < listaCeldasFx.length - 1) {
        j = i + 1;
        // ...con los que tiene a la derecha (bucle interno)
        while (j < listaCeldasFx.length) {
          const superIntervalo = unirIntervalos(listaCeldasFx[i], listaCeldasFx[j]);
          // Si es posible, ambos intervalos se unen en el primero (el que queda más a la izq.)
          if (superIntervalo) { 
            listaCeldasFx[i] = { ...superIntervalo };
            // En ese caso el 2º intervalo debe eliminarse de la lista
            listaCeldasFx.splice(j, 1);
            // Solo en caso contrario (no se ha eliminado ningún elemento) incrementamos el índice
          } else j++;
        }
        i++;
      }

    // Se compara el tamaño de la lista tras una iteración con el que tenía en la iteración anterior,
    // si no se ha reducido quiere decir que ya no podemos realizar más agrupaciones.
    } while (listaCeldasFx.length < numIntervalos);

    if (ui.alert('⚠️ ¿Proteger intervalos con fórmulas?',
      `Se van a proteger ${listaCeldasFx.length} intervalos en la hoja actual. Los intervalos protegidos previamente por HdC+ se regenerarán.
      
      ${modoProteccion == PROTECCION.modo.advertencia
        ? 'Se mostrará una advertencia al tratar de editar las celdas de los intervalos protegidos.'
        : `Solo tú (${Session.getActiveUser().getEmail()}) podrás editar las celdas de los intervalos protegidos.`} `, ui.ButtonSet.OK_CANCEL) == ui.Button.OK) {

      iniciado = true;
      hdc.toast(`Protegiendo ${listaCeldasFx.length} intervalos...`, '🔒 HdC+ dice:', -1);

      // (III) Proteger intervalos reducidos de celdas
      // A tener en cuenta: celdas que deben dejar de estar protegidas, protecciones múltiples

      // Elimina reglas de protección previas creadas por HdC+, puede que algunas celdas protegidas ya no contengan fórmulas
      // Mejora: identificar y eliminar solo las que no forman parte del conjunto actual a crear
      hoja.getProtections(SpreadsheetApp.ProtectionType.RANGE)
        .filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion)
        .forEach(regla => regla.remove());

      // Crea nuevas reglas de protección
      listaCeldasFx.forEach(intervalo => {
        const proteccion = hoja.getRange(
          intervalo.inicio.fila,
          intervalo.inicio.columna,
          intervalo.fin.fila - intervalo.inicio.fila + 1,
          intervalo.fin.columna - intervalo.inicio.columna + 1,
        ).protect();
        const descripcion = `${PROTECCION.descripcion} ${new Date().toLocaleString(Session.getActiveUserLocale(), opcionesFechaHora)}`;
        switch (modoProteccion) {
          case PROTECCION.modo.advertencia: proteccion.setWarningOnly(true).setDescription(descripcion); break;
          case PROTECCION.modo.soloYo: proteccion.addEditor(Session.getActiveUser()).setDescription(descripcion); break;      
        }
        numIntervalosProtegidos++;
      });

    }
  
  } catch (e) {
    error = true;
    ui.alert(`Se ha producido un error inesperado al proteger las fórmulas de la hoja activa, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

  hdc.toast(
    `Se han protegido ${numIntervalosProtegidos} intervalos.`,
    `${listaCeldasFx.length == 0 || !iniciado ? '🟠' : listaCeldasFx.length != numIntervalosProtegidos ? '🔴' : '🟢'} HdC+ dice:`
  );

}

/**
 * Eliminar todos los intervalos protegidos previamente por HdC+ (o todos los existentes) de la hoja activa
 * @param {boolean} [soloHdCPlus]
 */
function eliminarProteccionesFx(soloHdCPlus = true) {

  const ui = SpreadsheetApp.getUi();
  const hdc = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = SpreadsheetApp.getActiveSheet();
  let numIntervalosDesprotegidos = 0;
  let iniciado = false;

  hdc.toast('Buscando intervalos protegidos...', '👀 HdC+ dice:', -1);
  
  let protecciones = SpreadsheetApp.getActiveSheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  if (soloHdCPlus) protecciones = protecciones.filter(regla => regla.getDescription().substring(0, PROTECCION.descripcion.length) == PROTECCION.descripcion);
  
  if (protecciones.length == 0) ui.alert(`En esta hoja no hay intervalos protegidos ${soloHdCPlus ? 'creados por HdC+ ' : ''}que eliminar.`, ui.ButtonSet.OK);
  else try {
    if (ui.alert('⚠️ ¿Eliminar intervalos protegidos por HdC+?',
      `Se van a eliminar ${protecciones.length} intervalos protegidos en la hoja actual.`,
    ui.ButtonSet.OK_CANCEL) == ui.Button.OK) {
        
        iniciado = true;
        hdc.toast(`Desprotegiendo ${protecciones.length} intervalos...`, '🔓 HdC+ dice:', -1);
        
        protecciones.forEach(regla => {
          regla.remove();
          numIntervalosDesprotegidos++;
        });
    }
  } catch (e) {
    ui.alert(`Se ha producido un error inesperado al eliminar los intervalos protegidos, inténtalo de nuevo.
      
      ⚠️ ${e.message}`, ui.ButtonSet.OK);
  }

  hdc.toast(
    `Se han desprotegido ${numIntervalosDesprotegidos} intervalos.`,
    `${protecciones.length == 0 || !iniciado ? '🟠' : protecciones.length != numIntervalosDesprotegidos ? '🔴' : '🟢'} HdC+ dice:`
  );


}