# Operación Inssitu de Microsoft 365 (servidor MCP)

Este archivo describe cómo está montado en Inssitu el MCP (el servidor que conecta las apps de Microsoft 365). No cambia cómo corre el servicio: el código, Docker y las claves quedan igual. La guía genérica de Softeria (README y [deployment.md](deployment.md)) no describe este despliegue.

Lo de abajo es el mapa de lo que corre hoy. Los procedimientos largos siguen en los manuales del box y en los scripts del watchdog; aquí no se copian enteros.

## Qué corre en el box

En la máquina del box hay dos rutas de este fork:

| Ruta | Rol |
| --- | --- |
| `/home/box/opt/ms365-mcp-fork` | Clone de este fork (`Calrornds/ms-365-mcp-server`). |
| `/home/box/opt/ms365-mcp-http` | Adaptador que pone el servidor disponible por la red (HTTP). |

El arranque vigilado no vive en este repositorio. Vive en [Calrornds/inssitu-wa-bridge](https://github.com/Calrornds/inssitu-wa-bridge), carpeta `ops/box-watchdog/`:

- `run-ms365-http.sh` — proceso HTTP del MCP.
- `run-cloudflared-ms365.sh` — túnel Cloudflare de este MCP.
- `README.md` de esa carpeta — nombra el hostname público.

Este repo no edita esos scripts. Un cambio de arranque, túnel o adaptador se hace en `inssitu-wa-bridge` o en el box, en un cambio aparte.

## URL pública

El hostname público es `ms365.inssitu.com.pe` (el README de `ops/box-watchdog` lo registra así).

El túnel de WhatsApp y el de este MCP están los dos en el manual del box `MANUAL-TUNELES-CLOUDFLARED.md`. Ese manual no se reproduce aquí.

## Subida a OneDrive

La guía que va con este fork, en inglés, ya está en el árbol:

[docs/examples/onedrive-upload-instructions.md](examples/onedrive-upload-instructions.md)

Para subir un archivo a OneDrive usas las instrucciones de ese archivo. Si algo falla, el mensaje de error debería citar la carpeta o el archivo; si no, revisa el manual del box.

Nota para quien configura el servidor: `driveItemId` va sin codificar (`root:/folder/my file.docx:`); la sesión de subida es `create-upload-session`; el `PUT` de los bytes va a `uploadUrl` con `Content-Range` y sin cabecera `Authorization`. Ese archivo se pasa con `--extra-instructions-file` / `MS365_MCP_EXTRA_INSTRUCTIONS_FILE` (ver «Extra MCP instructions» en el README).

En el box, el mismo tema está desarrollado en `manual-subida-onedrive-grok.md` y `spec-ms365-binary-upload.md`. Esos dos no se pegan aquí: la versión que queda en git es la de `docs/examples/`.

Si un script de `ops/box-watchdog/` o un manual del box dice otra cosa, manda esa fuente. Este archivo no lista puertos, variables ni tokens.
