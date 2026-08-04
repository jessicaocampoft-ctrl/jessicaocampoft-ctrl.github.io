# Editor aislado del Pasaporte

Este Worker solo permite editar un pasaporte buscado por nombre y teléfono. Valida Cloudflare Access y crea una sesión de Apps Script exclusivamente en el servidor; el navegador nunca recibe ese token.

Las variables y secretos se configuran exclusivamente en Cloudflare. No se guardan valores sensibles en este repositorio.
