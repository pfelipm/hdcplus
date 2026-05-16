# Plan de Implementación: Reorganización A-Z y Nuevas Inversiones

## Objetivo
1. Reorganizar el menú principal `onOpen` en `principal.gs` para que los submenús sigan un estricto orden alfabético. Esto implica mover "Manipular intervalos de datos" justo debajo de "Ofuscar información".
2. Renombrar la función "Invertir orden de filas (flip)" a "Invertir orden de filas".
3. Añadir una nueva función "Invertir orden de columnas" (`invertirOrdenCol`).
4. Agrupar las funciones de inversión y transposición en un nuevo submenú anidado llamado "↕️ Invertir y transponer" para mantener la interfaz limpia.

## Archivos Afectados
*   `principal.gs`: Menú `onOpen` y lógica de las funciones `invertirOrden...`.

## Pasos de Implementación

### 1. Reordenación del Menú (`principal.gs`)
*   Cortar el bloque `.addSubMenu(ui.createMenu('📐 Manipular intervalos de datos')...)` de su posición actual (entre "Barajar datos" y "Gestionar hojas").
*   Pegarlo justo después del bloque `.addSubMenu(ui.createMenu('🕶️ Ofuscar información')...)` y antes de "Proteger celdas con fórmulas".

### 2. Estructuración del Submenú `Manipular intervalos de datos`
*   Reorganizar los ítems internos.
*   Crear un nuevo submenú `.addSubMenu(ui.createMenu('↕️ Invertir y transponer'))`.
*   Añadir dentro:
    *   `Invertir orden de filas` (apunta a `invertirOrdenFil`)
    *   `Invertir orden de columnas` (apunta a `invertirOrdenCol`)
    *   `Transponer (destructivo)` (apunta a `transponer`)

### 3. Lógica Backend (`principal.gs`)
*   Renombrar la función actual `invertirOrden` a `invertirOrdenFil`.
*   Actualizar el toast de `invertirOrdenFil` para que diga "Orden de filas invertido."
*   Crear la función `invertirOrdenCol`:
    *   Lógica: Iterar por cada rango, obtener la matriz.
    *   Transponer la matriz temporalmente, aplicar `reverse()` a cada fila (que ahora son columnas), y volver a transponer (o simplemente aplicar `reverse()` internamente a cada array de fila). *En JavaScript, una matriz bidimensional es un array de filas. Invertir columnas significa aplicar `fila.reverse()` a cada fila.*
    *   Volcar con `setValues()`.
    *   Toast: "Orden de columnas invertido."

## Verificación
1. Abrir Google Sheets y verificar que "Manipular intervalos de datos" aparece entre "Ofuscar información" y "Proteger celdas con fórmulas".
2. Ejecutar "Invertir orden de columnas" sobre un rango de prueba y verificar que las columnas (izquierda/derecha) se intercambian como en un espejo sin alterar las filas.
3. Verificar el correcto funcionamiento del renombrado "Invertir orden de filas".