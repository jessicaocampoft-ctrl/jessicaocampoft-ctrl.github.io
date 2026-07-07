# Activación de automatizaciones

El código ya está preparado. Google requiere una autorización inicial porque las tareas usan Calendar, Gmail, Drive, Formularios y Hojas de cálculo.

## Activar

1. Actualiza el proyecto de Google Apps Script con `google-apps-script.js` y `appsscript.json`.
2. Crea una versión nueva de la aplicación web conservando la misma URL.
3. Abre el panel de administración y entra en **Más herramientas → Automatizaciones**.
4. Pulsa **Activar programación** y acepta los permisos de Google si los solicita.
5. Pulsa **Ejecutar tareas diarias** para comprobar la primera ejecución.

## Horarios

- Todos los días, 7:00 a. m.: recordatorios, seguimientos y alertas de cobro.
- Todos los días, 10:00 p. m.: cierre de citas e historial de indicadores.
- Lunes, 8:00 a. m.: reactivación, reporte semanal, calidad de datos y respaldo.

## Datos creados

El sistema crea automáticamente estas hojas cuando las necesita:

- `AutomationLog`: historial de ejecuciones.
- `ColaMensajes`: mensajes de WhatsApp pendientes.
- `ListaEspera`: lista sincronizada entre dispositivos.
- `KPIHistory`: cierres mensuales de indicadores.

## WhatsApp

Los mensajes aparecen preparados en el panel y se envían con un clic. El envío completamente automático requiere conectar una cuenta oficial de WhatsApp Business API; no se incluyen credenciales ni se realizan envíos silenciosos desde números personales.
