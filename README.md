# Vacaciones del equipo · elección anónima de días

Aplicación web gratuita (Google Sheets + Apps Script) para que un equipo elija sus vacaciones **sin solapes y sin ver quién ha cogido cada día**. La encargada ve todo en una hoja de Google y recibe un aviso por correo cada vez que alguien guarda.

## Qué hace

**Para cada persona**
- Entra con un enlace, elige su nombre y crea su propio PIN de 4 números la primera vez.
- Toca los días en el calendario y pulsa **Confirmar vacaciones**.
- Los días ya cogidos por otras personas aparecen como **"Ocupado"**, sin nombre.
- Al confirmar ve un **comprobante** (con opción de descargarlo como imagen), marcado como *pendiente de aprobación por RRHH*.

**Reglas que comprueba**
- 23 días laborables por persona (lunes a viernes; fines de semana y festivos no cuentan).
- Meses bloqueados (por defecto enero, febrero, noviembre y diciembre).
- Reparto por semestre: máximo 14 días en un semestre (13/10 o 10/13 con un día de margen).
- 4 semanas completas de lunes a viernes + los días restantes juntos (pegados a una semana o dentro de otra semana). Si una semana tiene festivo, ese día va el lunes siguiente o junto a los días sueltos.
- Máximo 2 semanas con festivo por persona.
- Dos personas no pueden coger el mismo día aunque guarden a la vez.

**Para la encargada**
- Vista de administración con nombres, días por semestre y si cada persona cumple las reglas.
- Correo automático cada vez que alguien guarda cambios.
- Calendario anual listo para imprimir (pestaña *Imprimir* o PDF en Drive).
- Se puede cerrar la elección poniendo `NO` en *Ajustes › B2*.

## Archivos

| Archivo | Qué es |
|---|---|
| `Code.gs` | La parte de servidor: lee y escribe en la hoja, comprueba las reglas, envía correos y genera el PDF. |
| `Index.html` | La página que ven las personas (calendario, PIN, comprobante, vista de encargada). |
| `appsscript.json` | Ajustes del proyecto (zona horaria Europe/Madrid). Solo hace falta si usas `clasp`. |

## Instalación (unos 10 minutos)

1. En Google Drive: **Nuevo → Hojas de cálculo**. Llámala, por ejemplo, *Vacaciones 2027*.
2. En la hoja: **Extensiones → Apps Script**.
3. Borra el contenido de `Código.gs` y pega el de **`Code.gs`**.
4. Pulsa **+ → HTML**, llámalo `Index` y pega el contenido de **`Index.html`**. Guarda.
5. Arriba elige la función **`preparar`** y pulsa **Ejecutar**. Acepta los permisos (Configuración avanzada → Ir al proyecto → Permitir).
6. En la hoja, pestaña **Personas**: escribe los nombres del equipo (deja la columna PIN vacía). En **Festivos** revisa los festivos locales. En **Ajustes** tienes el código de encargada y el correo para avisos.
7. En Apps Script: **Implementar → Nueva implementación → Aplicación web**
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
8. Copia el enlace y envíalo al equipo.

Cada vez que cambies el código: **Implementar → Gestionar implementaciones → ✏️ → Nueva versión**. El enlace no cambia.

## Configuración

Al principio de `Code.gs`:

```js
const ANIO = 2027;
const DIAS_POR_PERSONA = 23;
const MESES_BLOQUEADOS = [1, 2, 11, 12];
const MAX_SEMESTRE = 14;
const MAX_SEMANAS_FESTIVO = 2;
```

Los festivos se editan en la pestaña **Festivos** de la hoja (formato `AAAA-MM-DD`).

## Usarlo en otra tienda

Haz **Archivo → Hacer una copia** de la hoja (la copia lleva el código), vacía *Personas* y *Reservas*, cambia los festivos locales, ejecuta `preparar` y publica la copia como aplicación web nueva. Cada copia tiene su propio enlace y sus propios datos.

## Privacidad

- Los datos (nombres, PIN y días) se guardan solo en la hoja de Google de quien la instala.
- Este repositorio no contiene datos de personas ni códigos: se generan al ejecutar `preparar`.
- El PIN es una protección sencilla para un equipo pequeño, no un sistema de seguridad.
