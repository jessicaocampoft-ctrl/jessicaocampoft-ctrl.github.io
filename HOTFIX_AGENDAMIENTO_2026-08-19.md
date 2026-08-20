# Hotfix de agendamiento público — 19/08/2026

## Problema confirmado
La página pública enviaba la reserva mediante `fetch(..., mode: 'no-cors')` y mostraba la pantalla de éxito inmediatamente, sin poder comprobar la respuesta de Apps Script. Esto permitía mostrar “Reserva temporal creada” aunque el backend no hubiera guardado la cita.

Además, la verificación de disponibilidad permitía continuar cuando fallaba la consulta al servidor.

## Cambio
- La verificación de disponibilidad ahora falla de forma segura: si no puede consultarse, no permite avanzar.
- Antes del POST se vuelve a verificar que el horario siga disponible.
- El POST espera una respuesta del backend y solo muestra éxito con `ok:true` e identificador de cita.
- Un rechazo explícito muestra el error y no modifica la interfaz como si hubiera éxito.
- Una respuesta ambigua por red, timeout, CORS o JSON inválido nunca muestra éxito y bloquea el reenvío ciego; ofrece verificar el código por WhatsApp para evitar duplicados.
- No se modifica Google Apps Script en este hotfix.

## Seguridad
No crear citas reales durante la validación. Las pruebas usan respuestas simuladas.