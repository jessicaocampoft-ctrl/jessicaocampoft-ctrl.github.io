# Mapa de despliegues — Pasaporte

Fecha de auditoría: 2026-08-03

| Componente | Ambiente | URL o identificador | Uso |
| --- | --- | --- | --- |
| Worker `cuidandote-admin-login-proxy` | Pruebas | `https://login-pruebas.cuidandotefisioterapia.com` | Cloudflare Access, `/session`, `/bootstrap`, `/admin-test`, `/api` y cierre de sesión. |
| Administrador | Producción | `https://admin.cuidandotefisioterapia.com` | Panel administrativo publicado. |
| Apps Script público | Producción | `AKfycbwPE46Z_7mnCPTE7KcthlluGXsuqwAofVBHS0jGXv7C8ekLpmPzHu5x0jFYB7yxEquw` | Implementación pública estable. |
| Apps Script administrador | Producción | `AKfycbzLWv3DYB5WxsiT7jMCv5lnfc6oX8IvQJPMU6dGyW3_ZGU7Jj-WfXIWeEuPcI41mQ5L` | Implementación administrativa existente. |
| Apps Script de auditoría | Pruebas | `AKfycbx7biQkVS9l1nU4AYQeOmQzbPcKebOUJ5UmX97vCJDaXg5s-9y0-mgSrE0ANZXZJ8Hd` | Usada por `/bootstrap`, `/api` y por la vista pública de Pasaporte en esta rama. |

## Fuente de configuración

- `/bootstrap`: Worker de pruebas → implementación de Apps Script de auditoría.
- Acciones administrativas del Pasaporte en `/admin-test`: Worker `/api` → implementación de Apps Script de auditoría.
- `pasaporte.html` en esta rama: implementación de Apps Script de auditoría.
- Las tres rutas anteriores usan el mismo proyecto Apps Script y, por ello, el mismo Spreadsheet vinculado.

No se incluyen secretos, tokens de sesión, tokens de pacientes ni datos personales.
