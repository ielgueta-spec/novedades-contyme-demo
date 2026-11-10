# Novedades Contyme (versión piloto)

App web instalable (PWA) para que las **empresas clientes** envíen a Contyme las novedades de sus trabajadores del mes
(horas extra, bonos, licencias con respaldo, vacaciones, ausencias, anticipos, ingresos, términos o «sin novedades»), y
para que el equipo de **Remuneraciones** las revise y haga las liquidaciones.

> **Versión piloto, sin servidor.** Parte sin datos: Contyme habilita las empresas. Todo lo que se ingresa queda
> guardado **solo en el navegador de cada dispositivo** (todavía no llega a Contyme) y el acceso se revisa en el
> navegador, así que no es seguridad real. No ingresar aquí licencias médicas ni datos sensibles reales. La versión
> real va en Remutyme (módulo «App»), con servidor y base de datos.
>
> Los correos de administrador no están escritos en el código: se guarda solo su huella SHA-256 (`ADMINS` en
> `index.html`). Para agregar uno: `sha256("novedades-admin|" + correo en minúsculas)`.

## Qué hace

- **Empresas:** entran con **RUT + clave**. La primera vez crean su propia clave (o usan un link de un solo uso; se
  elige en Admin → Ajustes). Agregan novedades paso a paso, revisan un resumen y las envían. Ven «En qué va» y sus
  liquidaciones.
- **Contyme (administradores):**
  - bandeja por empresa (no ha enviado, por revisar, revisada, ingresada);
  - pedir corrección, recordar a las que no han enviado y copiar para Excel;
  - habilitar empresas en forma masiva (RUT pegados o CSV);
  - restablecer o bloquear accesos;
  - registro de actividad.
- Funciona **sin internet** una vez abierta (service worker) y se instala en el celular.

## Probarla en el PC

```
python -m http.server 8090
```

Abrir http://localhost:8090. «Equipo Contyme» abre el ingreso de administradores (correo autorizado). En
Administración → Empresas se habilitan con «RUT; nombre; correo de RRHH». Para volver a la demo con datos ficticios:
`CONFIG.demo = true` y una semilla en `semilla()`.

## Archivos

| Archivo | Para qué |
|---|---|
| `index.html` | La app completa (HTML + CSS + JS, sin dependencias). `CONFIG.demo = false` quita todo lo de prueba. |
| `manifest.webmanifest` | Nombre, colores e íconos para instalarla. |
| `sw.js` | Guarda la app para abrirla sin internet. Al cambiar algo, subir `VERSION`. |
| `iconos/`, `fuentes/` | Íconos de la app y fuentes (Outfit y Source Sans 3, licencia OFL) incluidas. |

## Para la versión real (Remutyme, módulo «App»)

Toda la lectura y escritura pasa por la capa `srv` de `index.html` (`srv.cli.*` para la empresa, `srv.adm.*` para
administradores): cada función se convierte en un endpoint del servidor.

- Sesión en cookie `HttpOnly`/`Secure`, CSRF, claves con hash en el servidor, límite de intentos por RUT e IP.
- Administradores con Google del dominio de Contyme y lista de administradores.
- Respaldos (licencias médicas = datos de salud) en almacenamiento privado, enlaces que vencen y plazo de borrado.
- Correos reales (aviso a la encargada, confirmación, recordatorio, activación) sin novedades ni claves.
- Período, plazo, empresas y trabajadores desde Remutyme.
- La clave inicial real nunca puede ser la de esta demo.
