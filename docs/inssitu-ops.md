# Operación Inssitu del MCP MS365

Este archivo describe cómo está montado el servidor en Inssitu. No cambia el runtime: el código, Docker y los secretos siguen como están. La guía genérica de Softeria (README y [deployment.md](deployment.md)) no describe este despliegue.

Lo de abajo es el mapa de lo que corre hoy. Los procedimientos largos siguen en los manuales del box y en los scripts del watchdog; aquí no se copian enteros.

## Qué corre en el box

En la máquina del box hay dos rutas de este fork:

| Ruta | Rol |
| --- | --- |
| `/home/box/opt/ms365-mcp-fork` | Clone de este fork (`Calrornds/ms-365-mcp-server`). |
| `/home/box/opt/ms365-mcp-http` | Wrap que levanta el servidor en HTTP. |

El arranque vigilado no vive en este repositorio. Vive en [Calrornds/inssitu-wa-bridge](https://github.com/Calrornds/inssitu-wa-bridge), carpeta `ops/box-watchdog/`:

- `run-ms365-http.sh` — proceso HTTP del MCP.
- `run-cloudflared-ms365.sh` — túnel Cloudflare de este MCP.
- `README.md` de esa carpeta — nombra el hostname público.

Este repo no edita esos scripts. Un cambio de arranque, túnel o wrap se hace en `inssitu-wa-bridge` o en el box, en un cambio aparte.

## URL pública

El hostname público es `ms365.inssitu.com.pe` (el README de `ops/box-watchdog` lo registra así).

El túnel de WhatsApp y el de este MCP están los dos en el manual del box `MANUAL-TUNELES-CLOUDFLARED.md`. Ese manual no se reproduce aquí.

## Subida a OneDrive

La guía que va con este fork, en inglés, ya está en el árbol:

[docs/examples/onedrive-upload-instructions.md](examples/onedrive-upload-instructions.md)

Sirve como archivo de `--extra-instructions-file` / `MS365_MCP_EXTRA_INSTRUCTIONS_FILE` (ver «Extra MCP instructions» en el README). En corto: `driveItemId` va sin codificar (`root:/folder/my file.docx:`); la sesión de subida es `create-upload-session`; el `PUT` de los bytes va a `uploadUrl` con `Content-Range` y sin cabecera `Authorization`.

En el box, el mismo tema está desarrollado en `manual-subida-onedrive-grok.md` y `spec-ms365-binary-upload.md`. Esos dos no se pegan aquí: la versión que queda en git es la de `docs/examples/`.

Si un script de `ops/box-watchdog/` o un manual del box dice otra cosa, manda esa fuente. Este archivo no lista puertos, variables ni tokens.
