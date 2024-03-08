// Envoltorios para insertarNota()
function notaFechaHoraUsuario() { insertarNota(true, true, true); }
function notaFechaUsuario() { insertarNota(true, false, true); }
function notaFechaHora() { insertarNota(true, true, false); }
function notaFecha() { insertarNota(true, false, false); }
function notaUsuario() { insertarNota(false, false, true); }


/**
 * Crea una nota en la celda actual e inserta de manera parametrizada
 * fecha, hora y dirección de correo del usuario activo.
 * Si ya existe una nota, la prefija con los elementos mencionados
 * @param {boolean} fecha   VERDADERO si se desea insertar la fecha actual (DD/MM/AAAA)
 * @param {boolean} hora    VERDADERO si se desea inserta la hora actual (HH:MM)
 * @param {boolean} email   VERDADERO si se desea insertar la dirección de correo del usuario
 */
function insertarNota(fecha, hora, email) {

    const hoja = SpreadsheetApp.getActiveSheet();
    const ui = SpreadsheetApp.getUi();
    const celdaActual = hoja.getCurrentCell();

    const opcionesFechaHora = {};
    if (fecha) {
        opcionesFechaHora.year = 'numeric';
        opcionesFechaHora.month = 'numeric';
        opcionesFechaHora.day = 'numeric';
    }
    if (hora) {
        opcionesFechaHora.timeZone = Session.getScriptTimeZone();
        opcionesFechaHora.timeZoneName = 'short'
        opcionesFechaHora.hour = 'numeric';
        opcionesFechaHora.minute = 'numeric';
        opcionesFechaHora.hour12 = false;
    }

    const respuesta = ui.prompt('Introduce anotación:', ui.ButtonSet.OK_CANCEL);
    if (respuesta.getSelectedButton() == ui.Button.OK) {

        // Si no hay nota, devuelve cadena vacía (longitud = 0)
        const textoAnotacionActual = celdaActual.getNote();
        // Se insertará automáticamente un salto de línea si el [prefijo] de la anotación contiene 2 o 3 elementos.
        const textoaAnotacionNueva = '[' +
            (fecha || hora ? ' ' + new Date().toLocaleString(Session.getActiveUserLocale(), opcionesFechaHora) : '') + 
            (email ? ' ' + Session.getActiveUser().getEmail() : '') + (fecha + hora + email > 1 ? ' ]\n' : '] ') +
            respuesta.getResponseText();
        celdaActual.setNote(`${textoaAnotacionNueva}\n\n${textoAnotacionActual}`); 
    
    }
 
}