# Registro de cambios (Changelog) - HdC+

Todos los cambios notables en este proyecto serán documentados en este archivo. El formato está basado en [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [2.0 (mayo 2026)] - 2026-05-10

### Añadido
- Creación de este archivo `CHANGELOG.md` para el seguimiento de cambios en la rama `hdcplus2`.
- Vinculación del proyecto local con el backend de Apps Script mediante `clasp`.

### Modificado
- **principal.gs**: 
    - Actualizada la constante `VERSION` a "2.0 (mayo 2026)".
    - Rediseño del menú "Proteger celdas con fórmulas" con una estructura jerárquica: Hoja actual, Todas las hojas y Selección personalizada.
    - Estandarización global de `ui.alert()` para utilizar siempre la constante `ENCABEZADO_ALERTAS` como título y corregir firmas de método erróneas.
    - Ajuste de dimensiones del diálogo "Acerca de" para el nuevo diseño.
- **fx_acoplar_desacoplar.gs**: 
    - Optimización crítica de la función `ACOPLAR`. Se ha sustituido la lógica de doble pasada con complejidad $O(N^2)$ por una agrupación basada en un objeto `Map` con complejidad lineal $O(N)$.
    - Refactorización para evitar mutaciones accidentales del intervalo original mediante el uso del operador spread (`[...]`).
    - Restauración y mejora de los comentarios técnicos sobre la importancia de `JSON.stringify()` para evitar ambigüedad en las claves compuestas.
- **fx_repetir.gs**:
    - Marcadas las funciones `RELLENAR` y `REPETIRFC` como obsoletas (`@deprecated`) en favor de funciones nativas como `MAKEARRAY()`.
    - Corregido el bug de referencias compartidas en `RELLENAR` mediante el uso del operador spread (`[...]`).
    - Limpieza estética: eliminados espacios innecesarios entre métodos y paréntesis (ej. `push (...)` -> `push(...)`).
- **protegerCeldas.gs**:
    - Optimización crítica de rendimiento en `protegerFxHoja`: se ha sustituido el algoritmo de comparación exhaustiva $O(N^2)$ por un algoritmo de agrupación de dos pasadas (Vertical -> Horizontal) con complejidad $O(N \log N)$.
    - Soporte para protección/desprotección multioja: se han añadido funciones para aplicar reglas en toda la hoja de cálculo o en una selección personalizada.
    - Mejora de fluidez en la UI: sustitución de `ui.alert` finales por "toasts" de larga duración (10s) con cierre automático.
    - Mejora de UX en acciones masivas: implementación de conteo previo de intervalos y sistema de notificaciones de progreso pestaña por pestaña mediante "toasts".
    - Nueva lógica de backend para procesamiento granular: añadida función `procesarAccionHojaIndividual` para permitir feedback en tiempo real desde el cliente.
- **hojas.gs**:
    - Implementación de la infraestructura de backend para la "Consola de Pestañas" (Gestión avanzada).
    - Gestión de persistencia de grupos de hojas mediante `PropertiesService`.
    - Funciones de obtención de metadatos detallados (ID, color, visibilidad, grupo) para todas las pestañas.
    - Lógica de aplicación de cambios masivos (orden, visibilidad, color, grupos) optimizada, incluyendo gestión inteligente de la hoja activa al ocultar pestañas.
    - Añadida función `activarHojaPorId()` para navegación rápida desde la interfaz.
    - Corrección de errores en la sintaxis de alertas masivas (unificación de parámetros de texto).
- **dialogoGestionarHojas.html**:
    - Nueva interfaz avanzada para la administración centralizada de pestañas basada en Materialize CSS.
    - **Modo Live**: Implementación de sincronización en tiempo real que aplica cambios automáticamente sin necesidad de guardar manualmente.
    - **Activación de Hojas**: Nuevo botón `launch` en cada fila para convertir una pestaña en la hoja activa del documento al instante.
    - Soporte para reordenación mediante arrastrar y soltar (Drag & Drop) utilizando SortableJS.
    - Sistema de creación, gestión y filtrado por grupos personalizados con separación visual ("Mis Grupos").
    - Buscador en tiempo real con botón de limpieza rápida y filtros de visibilidad.
    - Acciones masivas de visibilidad avanzada: "Mostrar solo seleccionadas" y "Mostrar todas menos seleccionadas".
    - Paleta de colores ampliada (12 opciones) sincronizada con HdC+ e incluyendo selector de color personalizado.
    - Bloqueo de UI (freeze) durante procesos activos para garantizar la integridad de las operaciones.
- **acondicionarTexto.gs**:
    - Refactorización integral utilizando estándares modernos de ES6 y expresiones regulares con soporte Unicode (`\p{L}`).
    - Mejora drástica de `inicialesMays_()` (capitalización de nombres propios), reduciendo su complejidad y aumentando la robustez ante caracteres internacionales.
    - Preservación de la lógica histórica mediante las funciones `inicialMays_anterior_()` e `inicialesMays_anterior_()` con fines didácticos.
    - Restauración íntegra de la tabla de mapeo `latin_map` para asegurar la compatibilidad total de la función `latinizar()`.
- **acercaDe.html**:
    - Modernización completa del diseño utilizando tarjetas y componentes de Materialize CSS, mejorando la legibilidad y la estética general.
    - Preservación de la identidad visual mediante la integración de la cabecera original en base64.

### A estudiar
- Migración de las funciones de gestión de hojas (ordenación, visibilidad masiva) en `hojas.gs` a la API avanzada de Google Sheets (`batchUpdate`) para lograr un rendimiento instantáneo, evaluando la necesidad de ampliar los alcances (scopes) de OAuth.
- Uso de la API avanzada para detectar programáticamente las pestañas seleccionadas por el usuario en la interfaz de Sheets, permitiendo acciones rápidas de protección/desprotección sin necesidad de selección manual en el panel lateral.
- Creación de una interfaz unificada para el acondicionamiento de texto que permita el encadenamiento (chaining) de múltiples transformaciones (ej. pasar a minúsculas, latinizar y eliminar espacios) en una sola operación.
