/**
 * Novedades Remuneraciones (Contyme) · servidor en Google Apps Script.
 *
 * Recibe lo que manda la app (index.html), guarda empresas, claves (solo su huella), envíos y novedades en una Google
 * Sheet, sube los respaldos a una carpeta privada de Drive y manda los correos de aviso.
 *
 * Instalación (una vez, con la cuenta de Google de Contyme): ver servidor/INSTALAR.md.
 *
 * Nada sensible va escrito en este archivo (es público en GitHub). Va en «Propiedades del script»:
 *   ADMINS          correos con acceso de administrador, separados por coma
 *   CLAVE_INICIAL   clave genérica con que las empresas entran la primera vez
 *   PLAZO_DIA       (opcional) día del mes en que vence el envío de novedades; si no está, 25
 *   (SHEET_ID y CARPETA_ID los completa instalar())
 */

var APP_URL = 'https://ielgueta-spec.github.io/novedades-remuneraciones/'; // link de los botones de los correos
var DOMINIO_ENCARGADAS = 'contyme.cl';   // los avisos de envíos solo van a correos de este dominio
var MAX_INTENTOS = 5;
var ITERACIONES = 2000;                   // vueltas del hash de las claves
var HORAS_SESION_CLIENTE = 2;
var HORAS_SESION_ADMIN = 8;
var DIAS_RECORDAR = 30;                 // «Recuérdame en este dispositivo»
var MAX_ARCHIVO_MB = 8;
var HOJAS = {
  Empresas: ['rut', 'nombre', 'correoEncargada', 'habilitada', 'hash', 'sal', 'intentos', 'contacto', 'pideReset', 'actualizada'],
  Trabajadores: ['rut', 'nombre', 'rutTrabajador', 'baja'],   // baja = último período en que aparece (término)
  Envios: ['id', 'periodo', 'rut', 'empresa', 'fecha', 'cantidad', 'sinNovedades', 'estado', 'nota', 'carpeta'],
  Novedades: ['envioId', 'periodo', 'rut', 'trabajador', 'tipo', 'detalle', 'comentario', 'respaldo', 'respaldoUrl', 'datos', 'rutTrabajador'],
  Actividad: ['fecha', 'rut', 'tipo', 'texto'],
  Historial: ['fecha', 'periodo', 'rut', 'evento', 'resumen', 'de', 'autor', 'texto']   // seguimiento de cada envío (y su mensaje)
};

/* ---------- Instalación ---------- */

/** Crea la Sheet y la carpeta de Drive y guarda sus IDs. Se corre una vez desde el editor. */
function instalar() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_ID')) {
    var ss = SpreadsheetApp.create('Novedades App · datos');
    Object.keys(HOJAS).forEach(function (n, i) {
      var h = i === 0 ? ss.getSheets()[0].setName(n) : ss.insertSheet(n);
      h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
      h.setFrozenRows(1);
    });
    p.setProperty('SHEET_ID', ss.getId());
  }
  if (!p.getProperty('CARPETA_ID')) p.setProperty('CARPETA_ID', DriveApp.createFolder('Novedades App · respaldos').getId());
  var faltan = ['ADMINS', 'CLAVE_INICIAL'].filter(function (k) { return !p.getProperty(k); });
  Logger.log('Sheet: ' + p.getProperty('SHEET_ID') + ' · Carpeta: ' + p.getProperty('CARPETA_ID') +
    (faltan.length ? ' · FALTA configurar: ' + faltan.join(', ') : ' · Listo'));
}

/* ---------- Entrada ---------- */

function doPost(e) {
  var r;
  try {
    _mem = {}; var t0 = Date.now();
    var q = JSON.parse(e.postData.contents);
    var f = ACCIONES[q.accion];
    r = f ? f(q) : { error: 'Acción no válida.' };
    // La app pide los datos de administración en la misma respuesta (una ida y vuelta menos).
    if (q.conDatos && r && !r.error && (r.token || /^adm-/.test(q.accion)) && q.accion !== 'adm-datos') {
      var d = ACCIONES['adm-datos'](r.token ? { token: r.token, periodo: q.periodo } : q);
      if (d && d.ok) r.datosAdm = d;
    }
    if (r) r._ms = Date.now() - t0;
  } catch (err) {
    console.error(err && err.stack || err);
    r = { error: 'Hubo un problema en el servidor de Contyme. Intenta de nuevo en unos minutos.' };
  }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // ?calentar=1 lo manda la app al abrirse: deja las hojas en la caché mientras la persona escribe su RUT y clave.
  // No devuelve datos (solo «ok»), así que no expone nada.
  if (e && e.parameter && e.parameter.calentar) { try { _mem = {}; precargar(['Empresas', 'Trabajadores', 'Envios', 'Novedades']); } catch (err) {} }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'novedades-contyme' })).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Utilidades ---------- */

/* Velocidad: _mem guarda lo leído durante una consulta; las hojas además quedan en CacheService (rápido) con una
 * versión que cambia en cada escritura, así nunca se lee una copia vieja. */
var _mem = {};
function prop(k) { if (!_mem.p) _mem.p = PropertiesService.getScriptProperties().getProperties(); return _mem.p[k] || ''; }
function libro() { return _mem.libro || (_mem.libro = SpreadsheetApp.openById(prop('SHEET_ID'))); }
function hoja(n) {
  _mem.h = _mem.h || {}; if (_mem.h[n]) return _mem.h[n];
  var h = libro().getSheetByName(n);
  if (!h && HOJAS[n]) { h = libro().insertSheet(n); h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold'); h.setFrozenRows(1); }
  return (_mem.h[n] = h);
}
/* Período = mes en curso; plazo = PLAZO_DIA (25 si no está) de ese mes. */
function dos(n) { return ('0' + n).slice(-2); }
function periodoActual() { var d = new Date(); return d.getFullYear() + '-' + dos(d.getMonth() + 1); }
function diaPlazo() { var n = Math.round(Number(prop('PLAZO_DIA'))); return n >= 1 && n <= 28 ? n : 25; }
function plazoDe(periodo) { return periodo + '-' + dos(diaPlazo()); }
/* Historial del envío del mes: enviado / agregó / corrigió / revisada / ingresada / listas / reabierto / por-revisar. */
function historialDe(rut, periodo, conAutor) {
  return filas('Historial').filter(function (x) { return x.rut === rut && x.periodo === periodo; })
    .map(function (x) { return { t: x.fecha, evento: x.evento, resumen: String(x.resumen || ''), de: x.de, autor: conAutor ? x.autor : (x.de === 'contyme' ? 'Contyme' : ''), texto: String(x.texto || '') }; });
}
function anotar(rut, periodo, evento, resumen, de, autor, texto) {
  agregar('Historial', { fecha: ahora(), periodo: periodo, rut: rut, evento: evento, resumen: resumen || '', de: de, autor: autor || '', texto: String(texto || '').trim().slice(0, 1000) });
}
function valores(n) {
  _mem.v = _mem.v || {};
  if (!_mem.v[n]) precargar([n]);
  return _mem.v[n];
}
/* Trae varias hojas de la caché de una vez (2 o 3 consultas en total; antes eran 2 por hoja). Las hojas grandes se
 * guardan en trozos, así Novedades y Actividad siguen en caché aunque crezcan mes a mes (antes, pasados ~95.000
 * caracteres, se leían enteras desde la planilla en cada consulta). */
var TROZO = 30000; // caracteres por trozo: la caché acepta 100 KB por valor y así cabe aunque haya tildes y ñ
function precargar(nombres) {
  _mem.v = _mem.v || {};
  var faltan = nombres.filter(function (n) { return !_mem.v[n]; });
  if (!faltan.length) return;
  var c = cache(), vers = c.getAll(faltan.map(function (n) { return 'ver:' + n; })) || {}, nuevas = {};
  faltan.forEach(function (n) { if (!vers['ver:' + n]) nuevas['ver:' + n] = vers['ver:' + n] = String(Date.now()) + Math.random().toString(36).slice(2, 6); });
  if (Object.keys(nuevas).length) c.putAll(nuevas, 21600);
  var clave = function (n) { return 'hoja:' + n + ':' + vers['ver:' + n]; };
  var cabezas = c.getAll(faltan.map(clave)) || {}, trozos = [];
  faltan.forEach(function (n) { var m = /^trozos:(\d+)$/.exec(cabezas[clave(n)] || ''); if (m) for (var i = 0; i < Number(m[1]); i++) trozos.push(clave(n) + ':' + i); });
  var partes = trozos.length ? (c.getAll(trozos) || {}) : {};
  faltan.forEach(function (n) {
    var k = clave(n), x = cabezas[k], m = /^trozos:(\d+)$/.exec(x || ''), v = null;
    try {
      if (m) { var j = ''; for (var i = 0; i < Number(m[1]); i++) { if (partes[k + ':' + i] == null) { j = ''; break; } j += partes[k + ':' + i]; } if (j) v = JSON.parse(j); }
      else if (x) v = JSON.parse(x);
    } catch (e) { v = null; }
    if (!v) { v = hoja(n).getDataRange().getValues().map(function (fila) { return fila.map(function (y) { return y instanceof Date ? y.toISOString() : y; }); }); guardarCache(k, v); }
    _mem.v[n] = v;
  });
}
function guardarCache(k, v) {
  try {
    var j = JSON.stringify(v), c = cache();
    if (j.length <= TROZO) { c.put(k, j, 21600); return; }
    var o = {}, i = 0, n = 0;
    while (i < j.length) {
      var f = Math.min(i + TROZO, j.length), u = j.charCodeAt(f - 1);
      if (f < j.length && u >= 0xD800 && u <= 0xDBFF) f--; // no cortar un emoji por la mitad
      o[k + ':' + n++] = j.slice(i, f); i = f;
    }
    if (n > 100) return; // más de ~3 MB: se lee de la planilla
    c.putAll(o, 21600); c.put(k, 'trozos:' + n, 21600); // la cabeza va al final: nadie lee trozos incompletos
  } catch (e) {}
}
function cambio(n) { // después de escribir: nueva versión (la copia anterior queda descartada)
  if (_mem.v) delete _mem.v[n];
  cache().put('ver:' + n, String(Date.now()) + Math.random().toString(36).slice(2, 6), 21600);
}
function filas(n) {
  var v = valores(n), cab = HOJAS[n], out = [];
  for (var i = 1; i < v.length; i++) { var o = { _fila: i + 1 }; cab.forEach(function (c, j) { o[c] = v[i][j]; }); out.push(o); }
  return out;
}
function escribir(n, o) { hoja(n).getRange(o._fila, 1, 1, HOJAS[n].length).setValues([HOJAS[n].map(function (c) { return o[c] === undefined ? '' : o[c]; })]); cambio(n); }
function agregar(n, o) { hoja(n).appendRow(HOJAS[n].map(function (c) { return o[c] === undefined ? '' : o[c]; })); cambio(n); }
/* Varias filas en una sola escritura (antes, una llamada a la planilla por novedad). Solo dentro de conCandado. */
function asegurarCabecera(n) {
  var c = cache(), k = 'cab:' + n + ':' + HOJAS[n].length; if (c.get(k)) return;
  try { // solo cosmético: nunca debe impedir un envío
    var h = hoja(n), v = h.getRange(1, 1, 1, HOJAS[n].length).getValues()[0];
    if (v.join('|') !== HOJAS[n].join('|')) h.getRange(1, 1, 1, HOJAS[n].length).setValues([HOJAS[n]]).setFontWeight('bold');
    c.put(k, '1', 21600);
  } catch (err) { console.error(err); }
}
function agregarVarias(n, lista) {
  if (!lista.length) return;
  var h = hoja(n), m = lista.map(function (o) { return HOJAS[n].map(function (c) { var v = o[c] === undefined ? '' : o[c]; return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }); }); // = + - @ al inicio: queda como texto, nunca como fórmula
  var desde = h.getLastRow() + 1, hasta = desde + m.length - 1;
  if (hasta > h.getMaxRows()) h.insertRowsAfter(h.getMaxRows(), hasta - h.getMaxRows() + 100); // appendRow agregaba filas solo
  h.getRange(desde, 1, m.length, HOJAS[n].length).setValues(m);
  cambio(n);
}
function ahora() { return new Date().toISOString(); }
function hex(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function sha(t) { return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)); }
function azar() { return sha(Utilities.getUuid() + Utilities.getUuid() + Date.now()).slice(0, 40); }
function esPre(p) { return /^[0-9a-f]{64}$/.test(String(p || '')); }
function huellaV2(pre, sal) { return sha(sal + '|' + pre); }
function claveCorrecta(e, q) {
  if (String(e.sal).indexOf('v2:') === 0) return esPre(q.pre) && huellaV2(q.pre, e.sal) === e.hash;
  return huella(String(q.clave || ''), e.sal) === e.hash; // clave antigua: se pasa a v2 al entrar
}
function huella(clave, sal) { var h = sal + '|' + clave; for (var i = 0; i < ITERACIONES; i++) h = sha(sal + h); return h; }
function normRut(t) {
  var x = String(t || '').replace(/[^0-9kK]/g, '').toUpperCase().replace(/^0+/, '');
  return x.length > 1 && /^\d+$/.test(x.slice(0, -1)) ? x.slice(0, -1) + '-' + x.slice(-1) : '';
}
function dvValido(r) {
  var p = String(r).split('-'); if (p.length !== 2) return false;
  var s = 0, m = 2; for (var i = p[0].length - 1; i >= 0; i--) { s += Number(p[0][i]) * m; m = m === 7 ? 2 : m + 1; }
  var dv = 11 - (s % 11); dv = dv === 11 ? '0' : dv === 10 ? 'K' : String(dv); return dv === p[1];
}
var RE_CORREO = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
function periodoOk(p) { return /^\d{4}-\d{2}$/.test(String(p || '')); }
function nombreMes(p) { var m = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']; return m[Number(p.slice(5, 7)) - 1] + ' ' + p.slice(0, 4); }
function admins() { return prop('ADMINS').split(',').map(function (x) { return x.trim().toLowerCase(); }).filter(Boolean); }
function cache() { return CacheService.getScriptCache(); }
/* Sesión firmada: s1|rol|rut o correo|huella corta de la clave|vence|firma. La firma usa TOKEN_SECRETO (se crea solo).
 * Si la empresa cambia o restablece su clave, o un correo deja de ser administrador, la sesión deja de servir. */
function secreto() {
  var k = prop('TOKEN_SECRETO');
  if (!k) { k = azar() + azar(); PropertiesService.getScriptProperties().setProperty('TOKEN_SECRETO', k); if (_mem.p) _mem.p.TOKEN_SECRETO = k; }
  return k;
}
function firmar(t) { return hex(Utilities.computeHmacSha256Signature(t, secreto())); }
function sesionDe(q, rol) {
  var tk = String(q.token || '');
  if (tk.indexOf('s1|') === 0) {
    var p = tk.split('|'); if (p.length !== 6) return null;
    if (p[1] !== rol || !(Number(p[4]) > Date.now()) || firmar(p.slice(0, 5).join('|')) !== p[5]) return null;
    if (rol === 'cliente') { var e = empresa(p[2]); return e && e.hash && String(e.hash).slice(0, 12) === p[3] ? { rol: rol, rut: p[2] } : null; }
    return admins().indexOf(p[2]) >= 0 ? { rol: rol, correo: p[2] } : null;
  }
  var v = tk && cache().get('t:' + tk); if (!v) return null; // sesiones antiguas (antes de la firma)
  var s = JSON.parse(v); return s.rol === rol ? s : null;
}
function nuevaSesion(datos, horas, recordar) {
  var vence = Date.now() + (recordar ? DIAS_RECORDAR * 24 : horas) * 3600000;
  var quien = datos.rol === 'cliente' ? datos.rut : datos.correo;
  var huellaCorta = datos.rol === 'cliente' ? String((empresa(datos.rut) || {}).hash || '').slice(0, 12) : '';
  var t = ['s1', datos.rol, quien, huellaCorta, vence].join('|');
  return t + '|' + firmar(t);
}
function registrar(rut, tipo, texto) { agregar('Actividad', { fecha: ahora(), rut: rut || '', tipo: tipo, texto: String(texto).slice(0, 500) }); }
/**
 * Manda un correo con versión HTML (diseño Contyme) y versión de texto.
 * para: uno o varios correos separados por coma. m: { titulo, intro, filas, tabla, nota, boton: {txt, url}, boton2, responderA }
 */
function correo(para, asunto, m) {
  var lista = String(para || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return RE_CORREO.test(x); });
  if (!lista.length) return false;
  if (typeof m === 'string') m = { intro: m };
  var o = { to: lista.join(','), subject: asunto, name: 'Novedades Remuneraciones · Contyme', body: textoCorreo(m), htmlBody: htmlCorreo(m) };
  if (m.responderA && RE_CORREO.test(m.responderA)) o.replyTo = m.responderA;
  try { MailApp.sendEmail(o); return true; }
  catch (e) { console.error('Correo no enviado (' + asunto + '): ' + (e && e.message || e)); return false; } // los datos ya quedaron guardados
}
function hx(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function textoCorreo(m) {
  var t = [];
  if (m.titulo) t.push(m.titulo, '');
  if (m.intro) t.push(m.intro, '');
  (m.filas || []).forEach(function (f) { t.push(f[0] + ': ' + f[1]); });
  if (m.filas && m.filas.length) t.push('');
  if (m.tabla && m.tabla.length) {
    m.tabla.forEach(function (n) { t.push('• ' + n.trabajador + ' — ' + n.tipo + ': ' + n.detalle + (n.comentario ? ' («' + n.comentario + '»)' : '') + (n.url ? ' [respaldo: ' + n.url + ']' : '')); });
    t.push('');
  }
  if (m.nota) t.push(m.nota, '');
  if (m.boton) t.push(m.boton.txt + ': ' + m.boton.url);
  if (m.boton2) t.push(m.boton2.txt + ': ' + m.boton2.url);
  t.push('', '—', 'Novedades Remuneraciones · Contyme · Asesores · Auditores · Contables', m.responderA ? 'Si respondes este correo, le llega a tu encargada en Contyme.' : 'Mensaje automático.');
  return t.join('\n');
}
function htmlCorreo(m) {
  var V = '#0E7C7B', OSC = '#0B3B3A', GRIS = '#5B6B6A', BORDE = '#DCE6E5';
  var x = '<div style="margin:0;padding:24px 12px;background:#F2F6F6;font-family:Arial,Helvetica,sans-serif;color:#1F2A2A">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;margin:0 auto;background:#fff;border:1px solid ' + BORDE + ';border-radius:12px;overflow:hidden">' +
    '<tr><td style="background:' + V + ';padding:18px 24px;color:#fff"><img src="' + APP_URL + 'iconos/logo-contyme-blanco-520.png" width="170" alt="CONTYME" style="display:block;width:170px;height:auto;border:0;color:#fff;font-size:20px;font-weight:bold">' +
    '<div style="font-size:13px;opacity:.9;margin-top:6px">Novedades Remuneraciones</div></td></tr><tr><td style="padding:24px">';
  if (m.titulo) x += '<h1 style="margin:0 0 12px;font-size:20px;color:' + OSC + '">' + hx(m.titulo) + '</h1>';
  if (m.intro) x += '<p style="margin:0 0 16px;font-size:15px;line-height:1.5">' + hx(m.intro).replace(/\n/g, '<br>') + '</p>';
  if (m.filas && m.filas.length) {
    x += '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 16px;font-size:14px">';
    m.filas.forEach(function (f) { x += '<tr><td style="padding:3px 16px 3px 0;color:' + GRIS + '">' + hx(f[0]) + '</td><td style="padding:3px 0;font-weight:bold">' + hx(f[1]) + '</td></tr>'; });
    x += '</table>';
  }
  if (m.tabla && m.tabla.length) {
    var th = 'text-align:left;padding:8px;border-bottom:2px solid ' + V + ';font-size:12px;color:' + GRIS + ';text-transform:uppercase';
    var td = 'padding:8px;border-bottom:1px solid ' + BORDE + ';vertical-align:top';
    x += '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin:0 0 16px;font-size:14px">' +
      '<tr><th style="' + th + '">Trabajador</th><th style="' + th + '">Novedad</th><th style="' + th + '">Detalle</th></tr>';
    m.tabla.forEach(function (n) {
      x += '<tr><td style="' + td + ';font-weight:bold">' + hx(n.trabajador) + (n.rutTrabajador ? '<div style="font-size:12px;font-weight:normal;color:' + GRIS + '">' + rutPuntos(n.rutTrabajador) + '</div>' : '') + '</td><td style="' + td + '">' + hx(n.tipo) + '</td><td style="' + td + '">' + hx(n.detalle) +
        (n.comentario ? '<div style="color:' + GRIS + ';font-style:italic;margin-top:4px">«' + hx(n.comentario) + '»</div>' : '') +
        (n.url ? '<div style="margin-top:4px"><a href="' + hx(n.url) + '" style="color:' + V + '">📎 ' + hx(n.respaldo || 'Ver respaldo') + '</a></div>' : '') + '</td></tr>';
    });
    x += '</table>';
  }
  if (m.nota) x += '<div style="margin:0 0 16px;padding:12px 14px;background:#EAF4F4;border-left:4px solid ' + V + ';border-radius:6px;font-size:14px;line-height:1.5">' + hx(m.nota).replace(/\n/g, '<br>') + '</div>';
  [m.boton, m.boton2].forEach(function (b, i) {
    if (!b) return;
    x += '<a href="' + hx(b.url) + '" style="display:inline-block;margin:4px 8px 4px 0;padding:11px 20px;border-radius:8px;font-size:14px;font-weight:bold;text-decoration:none;' +
      (i ? 'color:' + V + ';border:1px solid ' + V : 'background:' + V + ';color:#fff') + '">' + hx(b.txt) + '</a>';
  });
  x += '</td></tr><tr><td style="padding:14px 24px;background:#F7FAFA;border-top:1px solid ' + BORDE + ';font-size:12px;color:' + GRIS + '">CONTYME · Asesores · Auditores · Contables<br>' +
    (m.responderA ? 'Si respondes este correo, le llega a tu encargada en Contyme.' : 'Mensaje automático de Novedades Remuneraciones.') + '</td></tr></table></div>';
  return x;
}
function rutPuntos(r) { var p = String(r).split('-'); return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + (p[1] ? '-' + p[1] : ''); }
function deContyme(c) { return RE_CORREO.test(c || '') && String(c).toLowerCase().split('@')[1] === DOMINIO_ENCARGADAS; }
function conCandado(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { _mem.v = {}; return fn(); } finally { l.releaseLock(); } } // con el candado se relee todo: nunca se escribe sobre una copia vieja
/* Carpeta privada del envío: <raíz>/<período>/<RUT> <empresa>, compartida con la encargada. Se llama con candado. */
function carpetaEnvio(periodo, rut, e) {
  var claveCar = 'car:' + periodo + '|' + rut + '|' + (e.correoEncargada || ''), idCar = PropertiesService.getScriptProperties().getProperty(claveCar), car = null;
  if (idCar) { try { car = DriveApp.getFolderById(idCar); } catch (err) { car = null; } }
  if (!car) {
    var raiz = DriveApp.getFolderById(prop('CARPETA_ID'));
    var fp = raiz.getFoldersByName(periodo), carPer = fp.hasNext() ? fp.next() : raiz.createFolder(periodo);
    var nomCar = rut + ' ' + e.nombre, fe = carPer.getFoldersByName(nomCar); car = fe.hasNext() ? fe.next() : carPer.createFolder(nomCar);
    if (deContyme(e.correoEncargada)) { try { car.addViewer(e.correoEncargada); } catch (err) {} }
    PropertiesService.getScriptProperties().setProperty(claveCar, car.getId());
  }
  return car;
}
function empresa(rut) { return filas('Empresas').filter(function (x) { return x.rut === rut; })[0] || null; }
function siNo(v) { return v === true || v === 'TRUE' || v === 'true' || v === 1; }
function envioDe(rut, periodo) { return filas('Envios').filter(function (x) { return x.rut === rut && x.periodo === periodo; })[0] || null; }
function novedadesDe(id) {
  return filas('Novedades').filter(function (x) { return x.envioId === id; }).map(function (x) {
    var d = {}; try { d = JSON.parse(x.datos || '{}'); } catch (e) {}
    return { trabajador: x.trabajador, tipo: x.tipo, detalle: x.detalle, comentario: x.comentario, datos: d,
      adjunto: x.respaldo ? { nombre: x.respaldo, url: x.respaldoUrl } : null };
  });
}
function mesDe(rut, periodo, conUrls) {
  var e = envioDe(rut, periodo);
  if (!e) return { enviado: null, revisado: false, ingresado: false, nota: null, novedades: null };
  var nov = novedadesDe(e.id);
  if (e.estado === 'reabierto') { // vuelven como borrador para corregir (en cualquier dispositivo)
    if (!conUrls) nov.forEach(function (n) { if (n.adjunto) delete n.adjunto.url; });
    return { enviado: null, revisado: false, ingresado: false, nota: e.nota || null, novedades: nov, sinNovedades: siNo(e.sinNovedades), reabierto: true };
  }
  if (!conUrls) nov.forEach(function (n) { if (n.adjunto) delete n.adjunto.url; });
  return { enviado: e.fecha, revisado: ['revisada', 'ingresada', 'listas'].indexOf(e.estado) >= 0, ingresado: e.estado === 'ingresada' || e.estado === 'listas', listas: e.estado === 'listas',
    nota: null, novedades: nov, sinNovedades: siNo(e.sinNovedades), carpeta: conUrls ? e.carpeta : undefined };
}
function trabajadorVigente(x, periodo) { return !(x.baja && periodo && periodo > String(x.baja)); }
function trabajadoresDe(rut, periodo) {
  var vistos = {}, out = [];
  filas('Trabajadores').forEach(function (x) { if (x.rut === rut && trabajadorVigente(x, periodo) && !vistos[x.nombre]) { vistos[x.nombre] = 1; out.push(x.nombre); } });
  return out;
}
function rutsTrabajadores(rut) { var m = {}; filas('Trabajadores').forEach(function (x) { if (x.rut === rut && x.rutTrabajador) m[x.nombre] = x.rutTrabajador; }); return m; }
function datosCliente(rut, periodo) {
  precargar(['Empresas', 'Trabajadores', 'Envios', 'Novedades', 'Historial']);
  var e = empresa(rut);
  return { rut: rut, nombre: e.nombre, contacto: e.contacto, plazo: periodoOk(periodo) ? plazoDe(periodo) : '', historial: periodoOk(periodo) ? historialDe(rut, periodo, false) : [], trabajadores: trabajadoresDe(rut, periodo), mes: periodoOk(periodo) ? mesDe(rut, periodo, false) : null };
}
function claveDebil(clave, rut) {
  var k = clave.toLowerCase(), ini = prop('CLAVE_INICIAL').toLowerCase();
  if (clave.length < 8) return 'La clave debe tener al menos 8 caracteres.';
  if (!/[0-9]/.test(clave) || !/[a-zñ]/i.test(clave)) return 'Usa letras y números.';
  if ((ini && k.indexOf(ini) >= 0) || k.indexOf('contyme') >= 0 || k.indexOf(rut.split('-')[0]) >= 0) return 'Es muy fácil de adivinar (no uses la clave inicial, el RUT ni "contyme"). Elige otra.';
  return '';
}

var GENERICO = 'El RUT o la clave no son correctos. Si tu empresa todavía no usa la app, pídele a tu encargada de remuneraciones en Contyme que la habilite.';
var BLOQUEO = 'Por seguridad, el acceso quedó bloqueado tras ' + MAX_INTENTOS + ' intentos fallidos. Pídele a tu encargada en Contyme que restablezca la clave.';

/* ---------- Acciones ---------- */

var ACCIONES = {
  /* --- Empresas --- */
  'ingresar': function (q) {
    var rut = normRut(q.rut), clave = String(q.clave || '');
    if (!rut || !dvValido(rut)) return { campo: 'rut', error: 'Revisa el RUT: falta un número o el dígito verificador no calza.' };
    if (!clave) return { campo: 'clave', error: 'Escribe la clave.' };
    return conCandado(function () {
      precargar(['Empresas', 'Trabajadores', 'Envios', 'Novedades']); // todo lo que usa el ingreso, en una sola ida a la caché
      var e = empresa(rut);
      if (!e) { Utilities.sleep(300); return { campo: 'clave', error: GENERICO }; }
      if (!siNo(e.habilitada)) return { error: 'El acceso de esta empresa está bloqueado. Habla con tu encargada de remuneraciones en Contyme.' };
      if (Number(e.intentos) >= MAX_INTENTOS) return { error: BLOQUEO };
      var ok = e.hash ? claveCorrecta(e, q) : (prop('CLAVE_INICIAL') !== '' && clave === prop('CLAVE_INICIAL'));
      if (!ok) {
        e.intentos = Number(e.intentos || 0) + 1; escribir('Empresas', e);
        if (e.intentos >= MAX_INTENTOS) { registrar(rut, 'seguridad', e.nombre + ': acceso bloqueado por ' + MAX_INTENTOS + ' intentos fallidos.'); return { error: BLOQUEO }; }
        var quedan = MAX_INTENTOS - e.intentos;
        return { campo: 'clave', error: GENERICO + (quedan <= 2 ? ' Te quedan ' + quedan + (quedan === 1 ? ' intento.' : ' intentos.') : '') };
      }
      var cambiar = Number(e.intentos || 0) !== 0;
      if (e.hash && String(e.sal).indexOf('v2:') !== 0 && esPre(q.pre)) { e.sal = 'v2:' + azar(); e.hash = huellaV2(q.pre, e.sal); cambiar = true; }
      if (cambiar) { e.intentos = 0; escribir('Empresas', e); }
      if (!e.hash) { var t = azar(); cache().put('k:' + t, rut, 600); return { paso: 'crear', ticket: t, nombre: e.nombre, rut: rut }; }
      return { ok: true, rut: rut, token: nuevaSesion({ rol: 'cliente', rut: rut }, HORAS_SESION_CLIENTE, !!q.recordar), datos: datosCliente(rut, q.periodo) };
    });
  },

  'crear-clave': function (q) {
    var rut = q.ticket && cache().get('k:' + q.ticket);
    if (!rut) return { error: 'Pasó mucho rato. Vuelve a empezar el ingreso.', reiniciar: true };
    var nueva = String(q.nueva || ''), contacto = String(q.correo || '').trim().toLowerCase();
    var d = claveDebil(nueva, rut); if (d) return { campo: 'nueva', error: d };
    if (!RE_CORREO.test(contacto)) return { campo: 'contacto', error: 'Escribe un correo de contacto válido, por ejemplo nombre@tuempresa.cl.' };
    return conCandado(function () {
      var e = empresa(rut); if (!e || !siNo(e.habilitada)) return { error: 'El acceso de esta empresa está bloqueado.', reiniciar: true };
      if (esPre(q.pre)) { e.sal = 'v2:' + azar(); e.hash = huellaV2(q.pre, e.sal); } else { e.sal = azar(); e.hash = huella(nueva, e.sal); }
      e.contacto = contacto; e.intentos = 0; e.pideReset = false; e.actualizada = ahora();
      escribir('Empresas', e); cache().remove('k:' + q.ticket);
      registrar(rut, 'seguridad', e.nombre + ' creó su clave.');
      return { ok: true, rut: rut, token: nuevaSesion({ rol: 'cliente', rut: rut }, HORAS_SESION_CLIENTE, !!q.recordar), datos: datosCliente(rut, q.periodo) };
    });
  },

  'pedir-restablecer': function (q) {
    var rut = normRut(q.rut);
    if (rut && dvValido(rut)) conCandado(function () {
      var e = empresa(rut);
      if (e && !siNo(e.pideReset)) {
        e.pideReset = true; escribir('Empresas', e);
        registrar(rut, 'correo', e.nombre + ' pidió restablecer su clave.');
        var para = deContyme(e.correoEncargada) ? e.correoEncargada : admins()[0];
        correo(para, 'Novedades · ' + e.nombre + ' pide restablecer su clave', {
          titulo: e.nombre + ' pide restablecer su clave',
          filas: [['Empresa', e.nombre], ['RUT', rutPuntos(rut)]],
          nota: 'Entra a la app como administrador → Empresas → «Restablecer clave». Después avísale a la empresa por la vía de siempre (nunca mandes la clave por correo).',
          boton: { txt: 'Abrir la app', url: APP_URL } });
      }
    });
    return { ok: true }; // misma respuesta siempre: no revela si el RUT es cliente
  },

  /* --- Cliente (con sesión) --- */
  'cli-datos': function (q) {
    var s = sesionDe(q, 'cliente'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    precargar(['Empresas', 'Trabajadores', 'Envios', 'Novedades']);
    var e = empresa(s.rut); if (!e || !siNo(e.habilitada) || !e.hash) return { error: 'Tu acceso cambió. Vuelve a entrar.', salir: true };
    return { ok: true, datos: datosCliente(s.rut, q.periodo) };
  },

  'enviar': function (q) {
    var s = sesionDe(q, 'cliente'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar; tu borrador sigue guardado.', salir: true };
    if (!periodoOk(q.periodo)) return { error: 'Período no válido.' };
    var lista = Array.isArray(q.novedades) ? q.novedades.slice(0, 200) : [];
    var archivos = Array.isArray(q.archivos) ? q.archivos.slice(0, 50) : [];
    if (!lista.length && !q.sinNovedades) return { error: 'Agrega una novedad o marca que no hubo novedades.' };
    for (var a = 0; a < archivos.length; a++) if (String(archivos[a].base64 || '').length > MAX_ARCHIVO_MB * 1.4 * 1048576) return { error: 'Un respaldo pesa más de ' + MAX_ARCHIVO_MB + ' MB.' };
    var bloqueada = { error: 'El acceso de esta empresa está bloqueado.', salir: true };
    var yaEnviaste = { error: 'Ya enviaste las novedades de ' + nombreMes(q.periodo) + '. Si te faltó algo, escríbele a tu encargada.' };
    // Revisión rápida antes de subir nada (se repite con candado más abajo).
    precargar(['Empresas', 'Envios']);
    var e = empresa(s.rut); if (!e || !siNo(e.habilitada)) return bloqueada;
    var antes = envioDe(s.rut, q.periodo); if (antes && antes.estado !== 'reabierto' && !q.agregar) return yaEnviaste;
    // 1) Respaldos a Drive FUERA del candado: es lo más lento y así no frena el ingreso ni los envíos de otras empresas.
    //    Si no hay respaldos no se toca Drive (antes se creaba y compartía la carpeta igual: segundos perdidos).
    var car = null, urls = {}, subidos = [];
    if (archivos.length) {
      car = conCandado(function () { return carpetaEnvio(q.periodo, s.rut, e); });
      archivos.forEach(function (f) {
        var mime = /^(image\/[a-z0-9.+-]+|application\/pdf)$/i.test(f.mime || '') ? f.mime : 'application/octet-stream';
        var nombre = String(f.nombre || 'respaldo').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80);
        var arch = car.createFile(Utilities.newBlob(Utilities.base64Decode(String(f.base64 || '')), mime, nombre));
        subidos.push(arch); urls[f.idx] = { nombre: nombre, url: arch.getUrl() };
      });
    }
    // 2) Planilla con candado (un par de segundos), revisando de nuevo por si entró otro envío mientras se subían los archivos.
    var r = conCandado(function () {
      precargar(['Empresas', 'Envios', 'Trabajadores']);
      e = empresa(s.rut); if (!e || !siNo(e.habilitada)) return bloqueada;
      var previo = envioDe(s.rut, q.periodo);
      if (previo && previo.estado !== 'reabierto' && !q.agregar) return yaEnviaste;
      var agrego = !!(previo && q.agregar && previo.estado !== 'reabierto');
      var id = previo ? previo.id : azar().slice(0, 12);
      // Si fue reabierto, se reemplazan las novedades anteriores (borrando de a bloques de filas seguidas).
      // Los respaldos que ya estaban en Drive se conservan: la app los manda solo por nombre y aquí se recupera su link.
      var enDrive = {};
      if (previo) {
        var h = hoja('Novedades'), viejas = filas('Novedades').filter(function (x) { return x.envioId === id; });
        viejas.forEach(function (x) { if (x.respaldoUrl) enDrive[String(x.respaldo)] = x.respaldoUrl; });
        var fl = viejas.map(function (x) { return x._fila; }).sort(function (a, b) { return b - a; });
        for (var k = 0; k < fl.length; k++) { var fin = fl[k], ini = fin; while (k + 1 < fl.length && fl[k + 1] === ini - 1) { ini--; k++; } h.deleteRows(ini, fin - ini + 1); }
        cambio('Novedades');
      }
      asegurarCabecera('Trabajadores'); asegurarCabecera('Novedades');
      var conocidos = trabajadoresDe(s.rut), rutsT = rutsTrabajadores(s.rut), lineas = [], filasNov = [], nuevos = [], paraApp = [], bajas = [];
      lista.forEach(function (n, i) {
        var trab = String(n.trabajador || '').slice(0, 120), nomAdj = n.adjunto ? String(n.adjunto.nombre || '') : '', adj = urls[i] || (nomAdj && enDrive[nomAdj] ? { nombre: nomAdj, url: enDrive[nomAdj] } : null), det = String(n.detalle || '').slice(0, 300), com = String(n.comentario || '').slice(0, 500);
        var respaldo = adj ? adj.nombre : (n.adjunto ? n.adjunto.nombre + ' (no llegó el archivo)' : ''), datosTxt = JSON.stringify(n.datos || {}).slice(0, 2000), d = {};
        var rutT = rutsT[trab] || '';
        if (n.tipo === 'Ingreso' && n.datos && n.datos.rut) { var rI = normRut(n.datos.rut); if (rI && dvValido(rI)) rutT = rI; }
        filasNov.push({ envioId: id, periodo: q.periodo, rut: s.rut, trabajador: trab, tipo: String(n.tipo || '').slice(0, 40),
          detalle: det, comentario: com, respaldo: respaldo, respaldoUrl: adj ? adj.url : '', datos: datosTxt, rutTrabajador: rutT });
        if (trab && (n.trabajadorNuevo || n.tipo === 'Ingreso') && conocidos.indexOf(trab) < 0) { nuevos.push({ rut: s.rut, nombre: trab, rutTrabajador: rutT, baja: '' }); conocidos.push(trab); }
        if (trab && n.tipo === 'Término') bajas.push(trab);
        lineas.push({ rutTrabajador: rutT, trabajador: trab, nueva: agrego && !!n.nueva, tipo: (agrego && n.nueva ? '🆕 ' : '') + String(n.tipo || ''), detalle: String(n.detalle || ''), comentario: String(n.comentario || ''), respaldo: adj ? adj.nombre : '', url: adj ? adj.url : '' });
        try { d = JSON.parse(datosTxt); } catch (err) {}
        paraApp.push({ trabajador: trab, tipo: String(n.tipo || '').slice(0, 40), detalle: det, comentario: com, datos: d, adjunto: respaldo ? { nombre: respaldo } : null });
      });
      agregarVarias('Novedades', filasNov); agregarVarias('Trabajadores', nuevos);
      if (bajas.length) filas('Trabajadores').forEach(function (x) { if (x.rut === s.rut && bajas.indexOf(x.nombre) >= 0 && !x.baja) { x.baja = q.periodo; escribir('Trabajadores', x); } });
      var env = { id: id, periodo: q.periodo, rut: s.rut, empresa: e.nombre, fecha: ahora(), cantidad: lista.length, sinNovedades: !lista.length, estado: 'enviado', nota: '', carpeta: car ? car.getUrl() : (previo && previo.carpeta) || '' };
      if (previo) { env._fila = previo._fila; escribir('Envios', env); } else agregar('Envios', env);
      var cuantas = lista.length ? lista.length + (lista.length === 1 ? ' novedad' : ' novedades') : 'sin novedades';
      var nuevasN = lista.filter(function (n) { return n.nueva; }).length;
      var comentarioEnvio = String(q.comentarioEnvio || '').trim().slice(0, 1000);
      anotar(s.rut, q.periodo, agrego ? 'agrego' : previo ? 'corrigio' : 'enviado', (agrego ? nuevasN + (nuevasN === 1 ? ' nueva · ' : ' nuevas · ') : '') + cuantas, 'cliente', e.contacto || '', comentarioEnvio);
      registrar(s.rut, 'accion', e.nombre + (agrego ? ' agregó ' + nuevasN + (nuevasN === 1 ? ' novedad' : ' novedades') + ' (' + cuantas + ' en total)' : ' envió ' + cuantas) + ' de ' + nombreMes(q.periodo) + '.');
      // La respuesta se arma con lo recién escrito (antes se releían 3 hojas completas desde la planilla).
      return { ok: true, previo: !!previo, agrego: agrego, nuevasN: nuevasN, comentarioEnvio: comentarioEnvio, cuantas: cuantas, lineas: lineas, carpeta: env.carpeta, conRespaldos: lineas.some(function (l) { return !!l.url; }),
        datos: { rut: s.rut, nombre: e.nombre, contacto: e.contacto, trabajadores: conocidos, plazo: plazoDe(q.periodo), historial: historialDe(s.rut, q.periodo, false),
          mes: { enviado: env.fecha, revisado: false, ingresado: false, nota: null, novedades: paraApp, sinNovedades: !lista.length } } };
    });
    if (r.error) { subidos.forEach(function (f) { try { f.setTrashed(true); } catch (err) {} }); return r; } // no deja respaldos sueltos
    // 3) Correos, fuera del candado.
    var previo = r.previo, cuantas = r.cuantas, lineas = r.lineas;
    var para = deContyme(e.correoEncargada) ? e.correoEncargada : admins().join(',');
    var mes = nombreMes(q.periodo), Mes = mes.charAt(0).toUpperCase() + mes.slice(1);
    var sinEnc = !deContyme(e.correoEncargada);
    var agrego = r.agrego, nuevasTxt = r.nuevasN + (r.nuevasN === 1 ? ' nueva' : ' nuevas');
    var okAviso = correo(para, (agrego ? '[Más novedades] ' : previo ? '[Corregidas] ' : '') + 'Novedades ' + mes + ' · ' + e.nombre + ' (' + (agrego ? nuevasTxt + ', ' : '') + cuantas + ')', {
      titulo: e.nombre + (agrego ? ' agregó ' + nuevasTxt + ' novedad' + (r.nuevasN === 1 ? '' : 'es') : previo ? ' corrigió sus novedades' : ' envió sus novedades'),
      filas: [['Empresa', e.nombre], ['RUT', rutPuntos(s.rut)], ['Período', Mes], ['Novedades', lista.length ? String(lista.length) + (agrego ? ' (' + nuevasTxt + ', marcadas con 🆕)' : '') : 'Sin novedades este mes'], ['Enviado por', e.contacto || '—']],
      tabla: lineas,
      nota: (r.comentarioEnvio ? 'Mensaje de la empresa: «' + r.comentarioEnvio + '»\n' : '') + (lista.length ? '' : 'La empresa avisó que este mes no hubo novedades: se repite la liquidación del mes anterior.\n') +
        (sinEnc ? 'Esta empresa no tiene encargada @contyme.cl: el aviso llegó a los administradores. Asígnala en Empresas (carga masiva).' : ''),
      boton: { txt: 'Abrir la bandeja', url: APP_URL },
      boton2: r.conRespaldos && r.carpeta ? { txt: 'Ver respaldos en Drive', url: r.carpeta } : null });
    var okConf = !!e.contacto && correo(e.contacto, 'Recibimos tus novedades de ' + mes + ' · ' + e.nombre, {
      titulo: agrego ? '¡Listo! Recibimos tus novedades nuevas' : '¡Listo! Recibimos tus novedades',
      intro: 'Contyme recibió las novedades de ' + mes + ' de ' + e.nombre + '. Este es el resumen de lo que enviaste:',
      filas: [['Empresa', e.nombre], ['Período', Mes], ['Novedades', lista.length ? String(lista.length) : 'Sin novedades este mes']],
      tabla: lineas.map(function (n) { return { trabajador: n.trabajador, tipo: n.tipo, detalle: n.detalle, comentario: n.comentario, respaldo: n.respaldo }; }),
      nota: 'Si te faltó algo, entra a la app y toca «Agregar más novedades». Para corregir algo ya enviado, responde este correo.',
      boton: { txt: 'Ver en qué va', url: APP_URL },
      responderA: sinEnc ? '' : e.correoEncargada });
    registrar(s.rut, 'correo', (okAviso ? 'Aviso a ' + para : 'NO se pudo avisar a ' + para + ' (revisar cuota de correos)') + (okConf ? ' y confirmación al cliente.' : ' · sin confirmación al cliente.'));
    return { ok: true, datos: r.datos };
  },

  /* --- Administradores --- */
  'admin-codigo': function (q) {
    var c = String(q.correo || '').trim().toLowerCase();
    if (admins().indexOf(c) >= 0) {
      var cod = String(Math.floor(100000 + Math.random() * 900000));
      cache().put('c:' + c, JSON.stringify({ cod: cod, n: 0 }), 600);
      correo(c, 'Tu código para entrar a Novedades Remuneraciones: ' + cod, {
        titulo: 'Tu código es ' + cod,
        intro: 'Escríbelo en la app para entrar como administrador. Vence en 10 minutos.',
        nota: 'Si no lo pediste, ignora este correo: nadie puede entrar sin el código.' });
    } else Utilities.sleep(400);
    return { ok: true }; // misma respuesta siempre
  },

  'admin-entrar': function (q) {
    var c = String(q.correo || '').trim().toLowerCase(), v = cache().get('c:' + c);
    if (!v || admins().indexOf(c) < 0) return { error: 'El código venció o no es válido. Pide uno nuevo.' };
    var x = JSON.parse(v);
    if (String(q.codigo || '').trim() !== x.cod) {
      x.n++; if (x.n >= 5) { cache().remove('c:' + c); return { error: 'Demasiados intentos. Pide un código nuevo.' }; }
      cache().put('c:' + c, JSON.stringify(x), 600); return { campo: 'a-codigo', error: 'El código no coincide.' };
    }
    cache().remove('c:' + c);
    registrar('', 'seguridad', 'Ingreso de administrador: ' + c + '.');
    return { ok: true, correo: c, token: nuevaSesion({ rol: 'admin', correo: c }, HORAS_SESION_ADMIN, !!q.recordar) };
  },

  'adm-datos': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var per = periodoOk(q.periodo) ? q.periodo : '';
    precargar(['Empresas', 'Trabajadores', 'Envios', 'Novedades', 'Actividad', 'Historial']);
    var trab = filas('Trabajadores'), envios = filas('Envios').filter(function (x) { return x.periodo === per; });
    var empresas = filas('Empresas').map(function (e) {
      return { rut: e.rut, nombre: e.nombre, correoEncargada: e.correoEncargada, habilitada: siNo(e.habilitada), registrada: !!e.hash,
        intentos: Number(e.intentos || 0), contacto: e.contacto, pideReset: siNo(e.pideReset),
        trabajadores: trab.filter(function (t) { return t.rut === e.rut && trabajadorVigente(t, per); }).map(function (t) { return t.nombre; }) };
    });
    var meses = {};
    envios.forEach(function (x) { meses[x.rut] = mesDe(x.rut, per, true); if (x.estado === 'reabierto') meses[x.rut].nota = x.nota; });
    var hist = {}; if (per) filas('Historial').forEach(function (x) { if (x.periodo === per) (hist[x.rut] = hist[x.rut] || []).push({ t: x.fecha, evento: x.evento, resumen: String(x.resumen || ''), de: x.de, autor: x.autor, texto: String(x.texto || '') }); });
    var act = filas('Actividad').slice(-200).map(function (a) { return { t: a.fecha, rut: a.rut || null, tipo: a.tipo, txt: a.texto }; });
    return { ok: true, empresas: empresas, meses: meses, actividad: act, historial: hist, plazo: per ? plazoDe(per) : '' };
  },

  'adm-habilitar': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var lista = Array.isArray(q.empresas) ? q.empresas.slice(0, 2000) : [], n = 0;
    conCandado(function () {
      var todas = filas('Empresas');
      lista.forEach(function (x) {
        var rut = normRut(x.rut); if (!rut || !dvValido(rut)) return;
        var e = todas.filter(function (y) { return y.rut === rut; })[0];
        var enc = String(x.correo || '').trim().toLowerCase();
        if (e) { e.habilitada = true; if (enc) e.correoEncargada = enc; if (x.nombre) e.nombre = String(x.nombre).slice(0, 120); e.actualizada = ahora(); escribir('Empresas', e); }
        else { if (!x.nombre) return; agregar('Empresas', { rut: rut, nombre: String(x.nombre).slice(0, 120), correoEncargada: enc, habilitada: true, hash: '', sal: '', intentos: 0, contacto: '', pideReset: false, actualizada: ahora() }); }
        registrar(rut, 'seguridad', (x.nombre || rut) + ': habilitada en la app por ' + s.correo + '.'); n++;
      });
    });
    return { ok: true, n: n };
  },

  /* Carga masiva de trabajadores: [{ rut (empresa), nombre, rutTrabajador }]. Si ya existe (mismo RUT de trabajador, o
     mismo nombre sin RUT) se actualiza y vuelve a quedar vigente; si no, se agrega. */
  'adm-trabajadores': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var lista = Array.isArray(q.trabajadores) ? q.trabajadores.slice(0, 5000) : [], nuevos = 0, actualizados = 0, malos = 0;
    conCandado(function () {
      asegurarCabecera('Trabajadores');
      var empresas = {}; filas('Empresas').forEach(function (e) { empresas[e.rut] = e.nombre; });
      var todas = filas('Trabajadores'), agregar = [], porEmpresa = {};
      lista.forEach(function (x) {
        var rut = normRut(x.rut), rt = normRut(x.rutTrabajador), nom = String(x.nombre || '').trim().replace(/\s+/g, ' ').slice(0, 120);
        if (!rut || !empresas[rut] || !rt || !dvValido(rt) || nom.length < 3) { malos++; return; }
        var e = todas.filter(function (y) { return y.rut === rut && (y.rutTrabajador === rt || (!y.rutTrabajador && String(y.nombre).toLowerCase() === nom.toLowerCase())); })[0];
        if (e) { e.nombre = nom; e.rutTrabajador = rt; e.baja = ''; escribir('Trabajadores', e); actualizados++; }
        else if (!agregar.some(function (y) { return y.rut === rut && y.rutTrabajador === rt; })) { agregar.push({ rut: rut, nombre: nom, rutTrabajador: rt, baja: '' }); nuevos++; }
        porEmpresa[rut] = (porEmpresa[rut] || 0) + 1;
      });
      agregarVarias('Trabajadores', agregar);
      Object.keys(porEmpresa).forEach(function (r) { registrar(r, 'accion', empresas[r] + ': ' + porEmpresa[r] + ' trabajadores cargados por ' + s.correo + '.'); });
    });
    return { ok: true, nuevos: nuevos, actualizados: actualizados, malos: malos };
  },

  'adm-bloqueo': function (q) { return conEmpresaAdm(q, function (e, s) { e.habilitada = !!q.habilitada; registrar(e.rut, 'seguridad', e.nombre + ': acceso ' + (e.habilitada ? 'habilitado' : 'bloqueado') + ' por ' + s.correo + '.'); }); },
  'adm-restablecer': function (q) { return conEmpresaAdm(q, function (e, s) { e.hash = ''; e.sal = ''; e.intentos = 0; e.pideReset = false; registrar(e.rut, 'seguridad', e.nombre + ': clave restablecida por ' + s.correo + '.'); }); },

  'adm-marcar': function (q) {
    return conEnvioAdm(q, function (x, e, s) {
      x.estado = q.estado === 'listas' ? 'listas' : q.estado === 'ingresada' ? 'ingresada' : q.estado === 'revisada' ? 'revisada' : 'enviado';
      var nota = String(q.nota || '').trim().slice(0, 1000);
      anotar(e.rut, x.periodo, x.estado === 'enviado' ? 'por-revisar' : x.estado, '', 'contyme', s.correo, nota);
      registrar(e.rut, 'accion', e.nombre + ': marcada «' + x.estado + '» por ' + s.correo + '.');
      if (x.estado === 'listas' || (nota && x.estado !== 'enviado')) {
        var tit = { revisada: 'Revisamos tus novedades', ingresada: 'Ya estamos haciendo tus liquidaciones', listas: '¡Tus liquidaciones están listas!' }[x.estado];
        correo(e.contacto, tit + ' · ' + e.nombre + ' · ' + nombreMes(x.periodo), {
          titulo: tit, intro: x.estado === 'listas' ? 'Terminamos las liquidaciones de ' + nombreMes(x.periodo) + ' de ' + e.nombre + '. Tu encargada te las hace llegar por la vía de siempre.' : 'Novedades de ' + nombreMes(x.periodo) + ' de ' + e.nombre + '.',
          nota: nota ? 'Nota de Contyme: «' + nota + '»' : '', boton: { txt: 'Ver el historial en la app', url: APP_URL }, responderA: deContyme(e.correoEncargada) ? e.correoEncargada : '' });
      }
    });
  },
  'adm-reabrir': function (q) {
    return conEnvioAdm(q, function (x, e, s) {
      x.estado = 'reabierto'; x.nota = String(q.nota || '').slice(0, 500);
      anotar(e.rut, x.periodo, 'reabierto', '', 'contyme', s.correo, x.nota);
      registrar(e.rut, 'correo', 'A ' + e.nombre + ': Contyme pidió revisar sus novedades (' + s.correo + ').');
      correo(e.contacto, 'Revisa tus novedades de ' + nombreMes(x.periodo) + ' · ' + e.nombre, {
        titulo: 'Te pedimos revisar tus novedades',
        intro: 'Tu encargada en Contyme reabrió el envío de ' + nombreMes(x.periodo) + ' de ' + e.nombre + ' con este mensaje:',
        nota: '«' + x.nota + '»',
        boton: { txt: 'Corregir y volver a enviar', url: APP_URL },
        responderA: deContyme(e.correoEncargada) ? e.correoEncargada : '' });
    });
  },
  'adm-recordar': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var e = empresa(normRut(q.rut)); if (!e || !periodoOk(q.periodo)) return { error: 'Empresa no encontrada.' };
    var x = envioDe(e.rut, q.periodo); if (x && x.estado !== 'reabierto') return { ok: true };
    var enviado = correo(e.contacto, 'Recuerda enviar tus novedades de ' + nombreMes(q.periodo) + ' · ' + e.nombre, {
      titulo: 'Recuerda enviar tus novedades',
      intro: 'Todavía no recibimos las novedades de ' + nombreMes(q.periodo) + ' de ' + e.nombre + (q.plazo ? '. El plazo es el ' + q.plazo : '') + '.',
      nota: 'Horas extra, bonos, licencias médicas, vacaciones, ingresos y términos. Si no hubo cambios, basta con tocar «Este mes no hubo novedades».',
      boton: { txt: 'Enviar mis novedades', url: APP_URL },
      responderA: deContyme(e.correoEncargada) ? e.correoEncargada : '' });
    registrar(e.rut, 'correo', 'Recordatorio a ' + e.nombre + (enviado ? '' : ' (no tiene correo de contacto: no se envió)') + '.');
    return enviado ? { ok: true } : { error: e.nombre + ' todavía no tiene correo de contacto (se registra al crear su clave).' };
  }
};

function conEmpresaAdm(q, fn) {
  var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
  return conCandado(function () { var e = empresa(normRut(q.rut)); if (!e) return { error: 'Empresa no encontrada.' }; fn(e, s); e.actualizada = ahora(); escribir('Empresas', e); return { ok: true }; });
}
function conEnvioAdm(q, fn) {
  var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
  if (!periodoOk(q.periodo)) return { error: 'Período no válido.' };
  return conCandado(function () {
    var e = empresa(normRut(q.rut)), x = e && envioDe(e.rut, q.periodo);
    if (!x) return { error: 'Esa empresa no tiene envío en el período.' };
    fn(x, e, s); escribir('Envios', x); return { ok: true };
  });
}

/* ---------- Recordatorios automáticos de plazo ----------
 * activarRecordatorios() se corre UNA vez desde el editor (pide permiso para activadores). Después, cada día a las 9:
 * - 3 días antes, 1 día antes y el día del plazo: correo a cada empresa habilitada que no ha enviado el mes;
 * - 1 día antes: resumen a cada encargada (o a los administradores) con las empresas que faltan. */
function activarRecordatorios() {
  ScriptApp.getProjectTriggers().forEach(function (tr) { if (tr.getHandlerFunction() === 'recordatoriosDiarios') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('recordatoriosDiarios').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Recordatorios activados: todos los días a las 9. Plazo: día ' + diaPlazo() + ' de cada mes.');
}
function recordatoriosDiarios() {
  _mem = {};
  var per = periodoActual(), plazo = plazoDe(per), hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var p = plazo.split('-'), fin = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])), dias = Math.round((fin - hoy) / 86400000);
  if ([3, 1, 0].indexOf(dias) < 0) return;
  precargar(['Empresas', 'Envios']);
  var faltan = filas('Empresas').filter(function (e) { if (!siNo(e.habilitada)) return false; var x = envioDe(e.rut, per); return !x || x.estado === 'reabierto'; });
  var cuando = dias === 0 ? 'hoy' : dias === 1 ? 'mañana' : 'en ' + dias + ' días', fecha = p[2] + '-' + p[1] + '-' + p[0];
  faltan.forEach(function (e) {
    var k = 'rec:' + per + ':' + dias + ':' + e.rut; if (cache().get(k)) return;
    var ok = correo(e.contacto, (dias === 0 ? 'Hoy vence' : 'Vence ' + cuando) + ' el plazo de novedades de ' + nombreMes(per) + ' · ' + e.nombre, {
      titulo: dias === 0 ? 'Hoy vence el plazo' : 'Quedan ' + (dias === 1 ? '1 día' : dias + ' días'),
      intro: 'Todavía no recibimos las novedades de ' + nombreMes(per) + ' de ' + e.nombre + '. El plazo vence ' + cuando + ' (' + fecha + ').',
      nota: 'Si no hubo cambios, basta con tocar «No tuve novedades este mes» y enviar.',
      boton: { txt: 'Enviar mis novedades', url: APP_URL }, responderA: deContyme(e.correoEncargada) ? e.correoEncargada : '' });
    cache().put(k, '1', 93600);
    registrar(e.rut, 'correo', 'Recordatorio automático (plazo ' + cuando + ')' + (ok ? '' : ': sin correo de contacto, no se envió') + '.');
  });
  if (dias === 1 && faltan.length) {
    var porEnc = {};
    faltan.forEach(function (e) { var k = deContyme(e.correoEncargada) ? e.correoEncargada : admins().join(','); (porEnc[k] = porEnc[k] || []).push(e); });
    Object.keys(porEnc).forEach(function (para) {
      var kr = 'recres:' + per + ':' + para; if (cache().get(kr)) return; cache().put(kr, '1', 93600);
      correo(para, 'Mañana vence el plazo: ' + porEnc[para].length + ' empresa' + (porEnc[para].length === 1 ? '' : 's') + ' sin enviar novedades', {
        titulo: 'Empresas que aún no envían ' + nombreMes(per),
        intro: 'El plazo vence mañana (' + fecha + '). A cada una ya le llegó un recordatorio automático.',
        tabla: porEnc[para].map(function (e) { return { trabajador: e.nombre, tipo: rutPuntos(e.rut), detalle: e.contacto ? 'Recordatorio enviado' : 'Sin correo de contacto: avisar por otra vía' }; }),
        boton: { txt: 'Abrir la bandeja', url: APP_URL } });
    });
  }
}
