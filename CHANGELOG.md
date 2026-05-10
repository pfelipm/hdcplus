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
- **panelProteccion.html**:
    - Nuevo panel lateral interactivo desarrollado con Materialize CSS.
    - Implementación de procesamiento secuencial de pestañas: feedback granular ("Procesando 2/5 [Nombre]") que mejora la UX y evita problemas de 'timeout'.
    - Eliminación de alertas y confirmaciones nativas del navegador, sustituyéndolas por un sistema de mensajes de estado interno y un modal de confirmación personalizado.
    - Incorporación de barra de progreso animada y bloqueos de seguridad en todos los controles durante la ejecución.
    - Estilo mixto de progreso: spinner circular para carga de datos y barra horizontal para ejecución de acciones.
    - Listado dinámico con indicadores visuales para hojas ocultas y hoja activa, con botón de recarga rápida y controles de selección masiva.

### A estudiar
- Migración de las funciones de gestión de hojas (ordenación, visibilidad masiva) en `hojas.gs` a la API avanzada de Google Sheets (`batchUpdate`) para lograr un rendimiento instantáneo, evaluando la necesidad de ampliar los alcances (scopes) de OAuth.
- Uso de la API avanzada para detectar programáticamente las pestañas seleccionadas por el usuario en la interfaz de Sheets, permitiendo acciones rápidas de protección/desprotección sin necesidad de selección manual en el panel lateral.
