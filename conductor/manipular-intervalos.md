# Plan de Implementación: Suite de Manipulación de Intervalos

## Objetivo
Reorganizar el menú "Estructurar datos" hacia un concepto más amplio ("Manipular intervalos de datos") e incorporar una suite de 6 herramientas rápidas para la transformación y limpieza de rangos seleccionados.

## Archivos Afectados
*   `principal.gs`: Estructura del menú y lógica principal de las nuevas funciones.

## Pasos de Implementación

### 1. Reorganización del Menú (`principal.gs`)
*   Eliminar el submenú actual `⚡ Ajustar casillas de verificación`.
*   Renombrar el submenú `📐 Estructurar datos` a `📐 Manipular intervalos de datos`.
*   Añadir los siguientes ítems al nuevo submenú:
    *   `⚡ Invertir casillas de verificación` (`invertirCasillas`)
    *   `☑️ Convertir texto a casillas` (`textoACasillas`)
    *   `⬇️ Rellenar celdas vacías hacia abajo (Fill Down)` (`fillDown`)
    *   `🗜️ Compactar filas vacías en selección` (`compactarFilas`)
    *   `↕️ Invertir orden de filas (Flip)` (`invertirOrden`)
    *   `🔗 Extraer URLs de enlaces` (`extraerURLs`)
    *   `[Separador]`
    *   (Mantener `Consolidar dimensiones` y `Transponer`)

### 2. Desarrollo de Funciones Backend (`principal.gs`)

Todas las funciones operarán sobre los rangos actualmente seleccionados (`SpreadsheetApp.getActiveSheet().getActiveRangeList().getRanges()`) y utilizarán un bloque `try...catch` con el estándar de alertas de HdC+.

#### A. Invertir Casillas (`invertirCasillas`)
*   *Lógica:* Iterar sobre los valores de **cada rango seleccionado**. Si `typeof celda === 'boolean'`, devolver `!celda`.
*   *Rendimiento:* Usar `getValues()` y `setValues()` por cada rango en la lista.

#### B. Convertir Texto a Casillas (`textoACasillas`)
*   *Lógica:* Identificar textos comunes afirmativos y negativos usando expresiones regulares en **cada rango seleccionado**.
*   Convertir esos textos a booleanos puros (`true` / `false`).
*   Volcar los valores y luego aplicar `rango.insertCheckboxes()` en cada rango procesado.

#### C. Rellenar hacia abajo (`fillDown`)
*   *Lógica:* Recorrer la matriz columna por columna (de arriba a abajo) del rango seleccionado.
*   Guardar en una variable temporal el último valor no vacío encontrado.
*   Si la celda actual está vacía (`''`), asignarle el valor temporal.
*   Sobreescribir el rango con la matriz resultante. *(Nota: Por su naturaleza relacional, operará sobre el rango activo principal, o iterará independientemente sobre cada rango si se seleccionan varios)*.

#### D. Compactar Filas Vacías (`compactarFilas`)
*   *Lógica:* Operar *solo* dentro de la selección.
*   Obtener valores. Filtrar el array de filas descartando aquellas completamente vacías.
*   Limpiar el contenido del rango original (`rango.clearContent()`).
*   Volcar el array filtrado (compactado) en la parte superior del rango original. *(Nota: Al igual que Fill Down, operará iterando independientemente por cada rango)*.

#### E. Invertir orden (`invertirOrden`)
*   *Lógica:* Obtener la matriz bidimensional de valores con `getValues()` para **cada rango seleccionado**.
*   Aplicar `.reverse()` al array de filas para darles la vuelta.
*   Volcar los datos con `setValues()`.

#### F. Extraer URLs (`extraerURLs`)
*   *Lógica:* Procesar **múltiples rangos seleccionados**. Dado que una celda puede contener varios enlaces (Rich Text con múltiples *runs*), se analizará cada fragmento de texto.
*   *Estrategia (basada en el artículo referenciado):* Usar `getRichTextValues()`. Para cada celda, iterar sobre `getRuns()`. Si un run tiene un enlace (`getLinkUrl()`), guardarlo.
*   Se extraerán **todos** los URLs de la celda y se concatenarán (ej. separados por comas o saltos de línea). Si no hay enlaces, se mantiene el valor original.
*   *Consideración:* Lanzar un `ui.alert` de confirmación previa avisando de que el texto visible será sustituido por las URLs extraídas.

## Verificación
1.  **Casillas:** Validar la inversión y la creación de casillas desde texto.
2.  **Fill Down:** Comprobar que rellena huecos correctamente por columnas independientes.
3.  **Compactar:** Asegurar que las fórmulas adyacentes fuera de la selección no se rompen.
4.  **Flip:** Validar la inversión del orden.
5.  **URLs:** Comprobar que extrae la URL subyacente correctamente.