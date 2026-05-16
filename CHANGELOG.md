# Registro de cambios (Changelog) - HdC+

Todos los cambios notables en este proyecto serán documentados en este archivo. El formato está basado en [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [2.0 (mayo 2026)] - 2026-05-10

### Añadido
- Creación de este archivo `CHANGELOG.md` para el seguimiento de cambios en la rama `hdcplus2`.
- Vinculación del proyecto local con el backend de Apps Script mediante `clasp`.
- **Nueva función de Marcos de Color**: Herramienta visual para generar bordes estéticos con opciones de grosor (píxeles), margen interior (padding) y combinación de celdas inteligente.
- **Generador de Índices Dual**: Capacidad para crear una pestaña dedicada 'Índice HdC+' enriquecida con metadatos o insertar una lista de enlaces ligeros en cualquier celda.
- **Easter Egg**: Atajo oculto en el selector de color personalizado (clic derecho) para activar el verde oficial de Google Sheets con efectos visuales.
- Archivo `.gitignore` configurado para excluir archivos de desarrollo (.clasp.json, .gdsheet, node_modules).

### Modificado
- **principal.gs**: 
    - Actualizada la constante `VERSION` a "2.0 (mayo 2026)".
    - Rediseño integral del menú con una estructura jerárquica y funcional mejorada.
    - Estandarización global de `ui.alert()` utilizando la constante `ENCABEZADO_ALERTAS` y corrección de errores de firma de método (unificación de parámetros de texto).
    - Refactorización de la lógica de dimensiones (`insertarFyC_core`, `eliminarFyC_core`) a ES6, implementando el "Encuadre total".
- **hojas.gs**:
    - Implementación de la infraestructura de backend para la "Consola de Pestañas" y gestión de grupos mediante `PropertiesService`.
    - Lógica de aplicación de cambios masivos optimizada y gestión inteligente de visibilidad de la hoja activa.
    - Inclusión del motor geométrico para Marcos de Color con soporte para dimensiones en píxeles.
- **panelCrearFyC.html / panelEliminarFyC.html**:
    - Modernización visual completa con Materialize v2.0 y sustitución de menús por botones de radio.
    - Implementación de procesamiento secuencial para acciones masivas con feedback granular ("Pestaña X/Y").
    - Ajustados los valores predeterminados de inserción a 10 filas/columnas.
- **fx_acoplar_desacoplar.gs**: 
    - Optimización crítica de rendimiento en `ACOPLAR`: paso de complejidad $O(N^2)$ a lineal $O(N)$ mediante el uso de `Map`.
- **fx_repetir.gs**:
    - Marcadas las funciones `RELLENAR` y `REPETIRFC` como obsoletas en favor de `MAKEARRAY()`.
    - Corregido el bug de referencias compartidas en memoria mediante el operador spread.
- **protegerCeldas.gs**:
    - Optimización crítica de rendimiento en `protegerFxHoja`: paso de complejidad $O(N^2)$ a $O(N \log N)$ mediante algoritmo de dos pasadas.
    - Soporte multioja y feedback visual mejorado con "toasts" de larga duración y cierre automático.
- **dialogoGestionarHojas.html**:
    - Nueva consola centralizada con **Modo Live** (sincronización en tiempo real), gestión de grupos y activación directa de pestañas.
- **dialogoMarcoColor.html**:
    - Interfaz modal avanzada con soporte para encuadre granular (anterior/posterior) y selector de color con paleta sincronizada.
- **acondicionarTexto.gs**:
    - Refactorización moderna utilizando expresiones regulares Unicode (`\p{L}`) y ES6, manteniendo versiones históricas para fines didácticos.
- **acercaDe.html**:
    - Rediseño estético total con Materialize CSS para una experiencia de usuario moderna.

### A estudiar
- Migración a la API avanzada de Google Sheets (`batchUpdate`) para operaciones instantáneas.
- Detección programática de pestañas seleccionadas por el usuario.
- Interfaz unificada de texto con encadenamiento (chaining) de transformaciones.
