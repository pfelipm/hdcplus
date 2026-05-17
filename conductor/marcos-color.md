# Plan de Implementación: Marcos de Color para Hojas

## Objetivo
Crear una herramienta visual que genere un marco de color (bordes estéticos) alrededor del área de datos de una pestaña de Google Sheets. Este marco se compondrá de filas y columnas coloreadas, ofreciendo opciones avanzadas de personalización geométrica.

## Archivos Afectados
*   `principal.gs`: Añadir comando al menú y lógica principal del backend (`aplicarMarcoColor`).
*   `dialogoMarcoColor.html`: Nueva interfaz modal de usuario.

## Componentes de la Interfaz (UI)
El diálogo será un modal (no una barra lateral) que contendrá:
1.  **Selector de Color:** Una cuadrícula visual idéntica a la usada en la Consola de Pestañas (12 opciones, incluyendo el selector nativo y la opción de 'Sin Color' para limpiar el marco).
2.  **Grosor del Marco:** 
    *   Input numérico para Anchura (Columnas a izq/der).
    *   Input numérico para Altura (Filas arriba/abajo).
3.  **Margen Interior (Padding):**
    *   Checkbox para habilitar el margen en blanco entre los datos y el marco.
    *   Input numérico (Celdas de margen).
4.  **Combinación de Celdas (Merge):**
    *   Checkbox para fusionar las celdas que componen el marco.
    *   Selector de dirección de fusión para las esquinas críticas: "Dominancia Horizontal" (las bandas superior/inferior ocupan todo el ancho) o "Dominancia Vertical" (las bandas laterales ocupan toda la altura).

## Algoritmo de Construcción (Backend)

La función `aplicarMarcoColor` recibirá la configuración y ejecutará:

1.  **Identificación del Área de Trabajo:**
    *   Detectar el `getDataRange()` actual. Si la hoja está completamente vacía, aplicar un área predeterminada (ej. A1:J20).
    *   Limpiar posibles marcos anteriores (eliminar filas/columnas extra o descombinar celdas exteriores).
    
2.  **Cálculo de Inserciones:**
    *   Calcular el número total de filas y columnas a insertar en los cuatro costados, sumando el *Grosor del Marco* y el *Margen Interior*.

3.  **Inserción Geométrica:**
    *   Insertar filas arriba (`insertRowsBefore(1, totalF)`) y abajo (`insertRowsAfter(lastRow, totalF)`).
    *   Insertar columnas a la izquierda (`insertColumnsBefore(1, totalC)`) y a la derecha (`insertColumnsAfter(lastCol, totalC)`).

4.  **Aplicación de Color:**
    *   Definir los 4 rangos rectangulares que componen el marco (superior, inferior, izquierdo, derecho), calculando sus coordenadas exactas teniendo en cuenta si hay margen en blanco o no.
    *   Aplicar `setBackground(color)` a esos 4 rangos.

5.  **Combinación (Merge) Condicional:**
    *   Si el usuario solicita combinación, calcular los rangos a fusionar dependiendo de la *Dominancia*.
    *   *Dominancia Horizontal:* El rango superior va desde la col 1 hasta la última. Los rangos laterales empiezan una fila más abajo y terminan una fila antes del final.
    *   *Dominancia Vertical:* El rango izquierdo va desde la fila 1 hasta la última. Los rangos superior/inferior empiezan una columna más a la derecha.

## Verificación
*   Confirmar que las esquinas no generan conflicto de "overlapping merges".
*   Verificar que la inserción de columnas/filas no altera las fórmulas internas del `getDataRange`.
*   Asegurar que aplicar un marco de "Sin Color" deshace eficazmente cualquier marco existente, devolviendo la hoja a su estado original.