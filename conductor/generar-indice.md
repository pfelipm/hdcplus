# Plan de Implementación: Generador de Índices Dual

## Objetivo
Añadir una funcionalidad dual al complemento HdC+ para generar índices de navegación de las pestañas del documento, adaptándose a diferentes necesidades del usuario (índice completo independiente o índice ligero integrado en un dashboard).

## Archivos Clave
*   `principal.gs`: Para añadir los nuevos comandos al menú, separados visualmente.
*   `hojas.gs`: Para implementar la lógica central de generación de enlaces y metadatos.

## Pasos de Implementación

### 1. Actualización del Menú (`principal.gs`)
*   Dentro del submenú `📋 Gestionar hojas`, añadir un separador (`.addSeparator()`) para agrupar visualmente las nuevas funciones.
*   Añadir dos comandos distintos:
    1.  `📑 Generar índice (nueva pestaña)`: Apunta a `generarIndiceCompleto()`.
    2.  `📌 Insertar índice aquí (solo enlaces)`: Apunta a `insertarIndiceLigero()`.

### 2. Lógica Principal (`hojas.gs`)

#### A. Función de utilidad: `obtenerDatosIndice(incluirMetadatos)`
*   Recopilar todas las hojas del documento (excluyendo la hoja llamada "Índice HdC+" si existe).
*   Si `incluirMetadatos` es true:
    *   Leer la configuración de grupos desde `PropertiesService`.
    *   Devolver un array bidimensional: `[ [Enlace(RichText), Estado(👁️/🙈), Grupo], ... ]`.
*   Si `incluirMetadatos` es false:
    *   Devolver un array unidimensional (como columnas): `[ [Enlace(RichText)], ... ]`.
*   *Nota sobre enlaces:* Utilizar `SpreadsheetApp.newRichTextValue()` para crear hipervínculos internos (`#gid=ID_HOJA`).

#### B. Comando 1: `generarIndiceCompleto()`
*   Buscar una pestaña llamada "Índice HdC+".
    *   Si existe: Borrar su contenido (`clear()`).
    *   Si no existe: Crear una nueva e insertarla en la posición 0.
*   Escribir cabeceras en negrita: `['Pestaña', 'Visibilidad', 'Grupo']`.
*   Obtener datos llamando a `obtenerDatosIndice(true)`.
*   Volcar los datos (batch) a partir de la fila 2.
*   Aplicar formato automático de ancho de columnas y añadir una nota en la celda A1 indicando la fecha de última actualización.

#### C. Comando 2: `insertarIndiceLigero()`
*   Identificar la celda activa actual.
*   Obtener datos llamando a `obtenerDatosIndice(false)`.
*   Comprobar si hay espacio suficiente hacia abajo. Si lo hay, volcar los datos (solo enlaces) a partir de la celda activa.

## Verificación y Pruebas
1.  **Índice Completo:** Comprobar que crea/actualiza la pestaña "Índice HdC+" con la tabla completa y enlaces funcionales.
2.  **Índice Ligero:** Comprobar que inserta una sola columna de enlaces en la celda activa sin sobrescribir datos valiosos de forma descontrolada.
3.  **Seguridad de Enlaces:** Asegurar que los enlaces RichText funcionan incluso si la pestaña destino cambia de nombre.