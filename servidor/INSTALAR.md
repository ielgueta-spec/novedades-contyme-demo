# Instalar el servidor (Google Apps Script)

Una sola vez, con la cuenta de Google de Contyme que será dueña de los datos. Toma unos 10 minutos.

Cuando un cliente toca «Enviar a Contyme», el servidor:
- guarda el envío y sus novedades en una **Google Sheet** («Novedades App · datos»);
- sube los respaldos (fotos de licencias, carnets) a una **carpeta privada de Drive** («Novedades App · respaldos»),
  compartida solo con la encargada de esa empresa;
- manda un **correo a la encargada** con el resumen y el link a la carpeta, y una **confirmación al cliente**.

## Pasos

1. Entra a https://script.google.com con la cuenta de Contyme → **Nuevo proyecto**. Ponle de nombre «Novedades App».
2. Borra lo que trae el editor y pega todo el contenido de `servidor/Code.gs`. Guarda (Ctrl + S).
3. **Configuración del proyecto** (el engranaje, a la izquierda) → **Propiedades de la secuencia de comandos** →
   **Agregar propiedad**:
   - `ADMINS`: los correos de administrador, separados por coma (los que entran a la bandeja).
   - `CLAVE_INICIAL`: la clave con que las empresas entran la primera vez. **No** uses la de la demo, y no la
     escribas en correos, chats ni notas (se entrega como diga Contyme).
4. Vuelve al editor, elige la función **`instalar`** arriba y aprieta **Ejecutar**. Google pide permisos (Sheets,
   Drive y enviar correos): acepta. En el registro debe decir «Listo».
5. **Implementar** → **Nueva implementación** → tipo **Aplicación web**:
   - Ejecutar como: **Yo**
   - Quién tiene acceso: **Cualquier usuario**
   → **Implementar** y copia la **URL de la aplicación web** (termina en `/exec`).
6. Esa URL va en `index.html`, en `CONFIG.servidor`. No es secreta: sin clave ni código no deja ver nada.

Si en el paso 5 no aparece «Cualquier usuario», el administrador de Google Workspace de Contyme tiene que permitir
apps web públicas.

## Recordatorios automáticos de plazo (una vez)

1. En el editor elige la función **`activarRecordatorios`** → **Ejecutar** → acepta el permiso nuevo (activadores).
2. Desde ahí, todos los días a las 9 la app revisa el plazo: 3 días antes, el día anterior y el mismo día manda un
   recordatorio a cada empresa que no ha enviado, y el día anterior un resumen a cada encargada.
3. El plazo es el día **25** de cada mes. Para cambiarlo, agrega la propiedad **`PLAZO_DIA`** (por ejemplo `20`).

## Cuando cambie `Code.gs`

Pega el código nuevo → **Implementar** → **Gestionar implementaciones** → lápiz → Versión: **Nueva versión** →
**Implementar**. La URL no cambia.

## Seguridad

- Las claves de las empresas se guardan solo como huella (hash con sal), nunca la clave.
- Los administradores entran con un **código de 6 números que llega a su correo** (vence en 10 minutos).
- La clave inicial, la lista de administradores y los IDs de la Sheet y la carpeta están en las propiedades del
  script, no en el código.
- Los avisos de envíos solo se mandan a correos `@contyme.cl`.
- Bloqueo tras 5 intentos fallidos y sesiones que vencen (2 h empresas, 8 h administradores).
- Límites de Google: unos 1.500 correos al día y 30 segundos por envío (respaldos de hasta 8 MB).
