# Plan de Implementación: Modernización de Herramientas de Dimensión

## Objetivo
Actualizar la interfaz gráfica y potenciar la lógica del backend de las herramientas de inserción y eliminación de filas y columnas (`panelCrearFyC` y `panelEliminarFyC`). El objetivo es homogeneizar la estética con el resto de la v2.0 (Materialize CSS) e introducir una nueva funcionalidad de "encuadre" (recorte integral).

## Archivos Afectados
*   `principal.gs`: Lógica de backend (`insertarFyC_core`, `eliminarFyC_core`).
*   `panelCrearFyC.html` / `panelCrearFyC_js.html`: Diálogo de inserción.
*   `panelEliminarFyC.html` / `panelEliminarFyC_js.html`: Diálogo de recorte.

## Pasos de Implementación

### 1. Refactorización Backend (`principal.gs`)
*   **Modernización de código:** Actualizar `insertarFyC_core` y `eliminarFyC_core` a sintaxis ES6 (funciones flecha, `const`/`let`).
*   **Lógica de Encuadre en `eliminarFyC_core`:**
    *   Actualmente el script recorta por la derecha y por abajo basándose en `getLastRow()` y `getLastColumn()`.
    *   Se añadirá un nuevo parámetro boolean `encuadre`.
    *   Si `encuadre` es true, el script identificará también el inicio de los datos usando `sheet.getDataRange().getRow()` y `sheet.getDataRange().getColumn()`.
    *   Se eliminarán las filas vacías superiores (de 1 a `firstRow - 1`) y las columnas vacías izquierdas (de 1 a `firstCol - 1`).
    *   *Nota técnica:* Al eliminar por arriba/izquierda, los índices cambian, por lo que la eliminación inferior/derecha debe recalcularse o hacerse primero.

### 2. UI: Panel de Inserción (`panelCrearFyC.html`)
*   Eliminar la dependencia del antiguo `panelLateral_css.html`.
*   Implementar CDN de Materialize CSS e Iconos.
*   Rediseñar los inputs numéricos (Filas y Columnas) con un layout más compacto y moderno.
*   Convertir el desplegable (Select) de "Modo de inserción" (Inicio, Final, Antes, Después) en un grupo de **Botones de Radio (Radio Buttons)** para agilizar la selección con un solo clic.
*   Implementar el patrón de bloqueo de UI (Spinner + deshabilitación de botones) durante la llamada al servidor.
*   Añadir notificaciones `M.toast()` para éxito y error (eliminando alerts nativos).

### 3. UI: Panel de Recorte (`panelEliminarFyC.html`)
*   Aplicar el mismo rediseño visual base (Materialize, independización de CSS antiguo).
*   Sustituir el menú desplegable actual (Filas, Columnas, Ambas) por **Botones de Radio**.
*   Añadir un nuevo **Checkbox: "Encuadre total (Recortar también por arriba/izquierda)"** debajo de la selección de modo.
*   Añadir el checkbox de "Aplicar a todas las pestañas" (existente pero rediseñado).
*   Implementar el bloqueo de UI y notificaciones `toast` al igual que en el otro panel.

## Consideraciones de UX
*   Los paneles deben mantener un diseño limpio, con un título claro (`section-title` como en el gestor de pestañas) y botones de acción en un footer anclado.
*   Se unificará el estilo de los botones principales (Color azul, iconos a la izquierda).

## Verificación
1.  Comprobar que el "Encuadre" no elimina datos accidentalmente si la hoja ya empieza en A1.
2.  Verificar que la inserción de filas con índices relativos sigue funcionando tras la refactorización.