# Setup y Sincronización Inicial de hdcplus2

## Objetivo
Configurar el entorno de desarrollo local con `clasp` para sincronizar la codebase de `hdcplus2` con el nuevo proyecto de Google Apps Script.

## Contexto
El usuario ha solicitado salir del modo de planificación para ejecutar comandos de `clasp` directamente. Este plan puente documenta la transición.

## Pasos de Implementación
1.  **Configurar clasp**: Crear archivo `.clasp.json` con el `scriptId` proporcionado.
2.  **Sincronizar**: Ejecutar `npx @google/clasp push` para subir el código al servidor.