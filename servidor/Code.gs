/**
 * Novedades Contyme · servidor en Google Apps Script.
 *
 * Recibe lo que manda la app (index.html), guarda empresas, claves (solo su huella), envíos y novedades en una Google
 * Sheet, sube los respaldos a una carpeta privada de Drive y manda los correos de aviso.
 *
 * Instalación (una vez, con la cuenta de Google de Contyme): ver servidor/INSTALAR.md.
 *
 * Nada sensible va escrito en este archivo (es público en GitHub). Va en «Propiedades del script»:
 *   ADMINS          correos con acceso de administrador, separados por coma
 *   CLAVE_INICIAL   clave genérica con que las empresas entran la primera vez
 *   (SHEET_ID y CARPETA_ID los completa instalar())
 */

var APP_URL = 'https://ielgueta-spec.github.io/novedades-contyme-demo/'; // link de los botones de los correos
var DOMINIO_ENCARGADAS = 'contyme.cl';   // los avisos de envíos solo van a correos de este dominio
var MAX_INTENTOS = 5;
var ITERACIONES = 2000;                   // vueltas del hash de las claves
var HORAS_SESION_CLIENTE = 2;
var HORAS_SESION_ADMIN = 8;
var MAX_ARCHIVO_MB = 8;
var HOJAS = {
  Empresas: ['rut', 'nombre', 'correoEncargada', 'habilitada', 'hash', 'sal', 'intentos', 'contacto', 'pideReset', 'actualizada'],
  Trabajadores: ['rut', 'nombre'],
  Envios: ['id', 'periodo', 'rut', 'empresa', 'fecha', 'cantidad', 'sinNovedades', 'estado', 'nota', 'carpeta'],
  Novedades: ['envioId', 'periodo', 'rut', 'trabajador', 'tipo', 'detalle', 'comentario', 'respaldo', 'respaldoUrl', 'datos'],
  Actividad: ['fecha', 'rut', 'tipo', 'texto']
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

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, app: 'novedades-contyme' })).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Utilidades ---------- */

/* Velocidad: _mem guarda lo leído durante una consulta; las hojas además quedan en CacheService (rápido) con una
 * versión que cambia en cada escritura, así nunca se lee una copia vieja. */
var _mem = {};
function prop(k) { if (!_mem.p) _mem.p = PropertiesService.getScriptProperties().getProperties(); return _mem.p[k] || ''; }
function libro() { return _mem.libro || (_mem.libro = SpreadsheetApp.openById(prop('SHEET_ID'))); }
function hoja(n) { _mem.h = _mem.h || {}; return _mem.h[n] || (_mem.h[n] = libro().getSheetByName(n)); }
function valores(n) {
  _mem.v = _mem.v || {};
  if (_mem.v[n]) return _mem.v[n];
  var c = cache(), ver = c.get('ver:' + n);
  if (!ver) { ver = String(Date.now()) + Math.random().toString(36).slice(2, 6); c.put('ver:' + n, ver, 21600); }
  var guardado = c.get('hoja:' + n + ':' + ver), v = null;
  if (guardado) { try { v = JSON.parse(guardado); } catch (e) { v = null; } }
  if (!v) {
    v = hoja(n).getDataRange().getValues().map(function (fila) { return fila.map(function (x) { return x instanceof Date ? x.toISOString() : x; }); });
    var j = JSON.stringify(v); if (j.length < 95000) { try { c.put('hoja:' + n + ':' + ver, j, 21600); } catch (e) {} }
  }
  return (_mem.v[n] = v);
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
function sesionDe(q, rol) {
  var v = q.token && cache().get('t:' + q.token); if (!v) return null;
  var s = JSON.parse(v); return s.rol === rol ? s : null;
}
function nuevaSesion(datos, horas) { var t = azar(); cache().put('t:' + t, JSON.stringify(datos), horas * 3600); return t; }
function registrar(rut, tipo, texto) { agregar('Actividad', { fecha: ahora(), rut: rut || '', tipo: tipo, texto: String(texto).slice(0, 500) }); }
/**
 * Manda un correo con versión HTML (diseño Contyme) y versión de texto.
 * para: uno o varios correos separados por coma. m: { titulo, intro, filas, tabla, nota, boton: {txt, url}, boton2, responderA }
 */
function correo(para, asunto, m) {
  var lista = String(para || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return RE_CORREO.test(x); });
  if (!lista.length) return false;
  if (typeof m === 'string') m = { intro: m };
  var o = { to: lista.join(','), subject: asunto, name: 'Novedades Contyme', body: textoCorreo(m), htmlBody: htmlCorreo(m) };
  if (m.responderA && RE_CORREO.test(m.responderA)) o.replyTo = m.responderA;
  MailApp.sendEmail(o);
  return true;
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
  t.push('', '—', 'Novedades Contyme · Asesores · Auditores · Contables', m.responderA ? 'Si respondes este correo, le llega a tu encargada en Contyme.' : 'Mensaje automático.');
  return t.join('\n');
}
function htmlCorreo(m) {
  var V = '#0E7C7B', OSC = '#0B3B3A', GRIS = '#5B6B6A', BORDE = '#DCE6E5';
  var x = '<div style="margin:0;padding:24px 12px;background:#F2F6F6;font-family:Arial,Helvetica,sans-serif;color:#1F2A2A">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:640px;margin:0 auto;background:#fff;border:1px solid ' + BORDE + ';border-radius:12px;overflow:hidden">' +
    '<tr><td style="background:' + V + ';padding:18px 24px;color:#fff"><div style="font-size:20px;font-weight:bold;letter-spacing:1px">CONTYME</div>' +
    '<div style="font-size:12px;opacity:.85">Novedades de remuneraciones</div></td></tr><tr><td style="padding:24px">';
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
      x += '<tr><td style="' + td + ';font-weight:bold">' + hx(n.trabajador) + '</td><td style="' + td + '">' + hx(n.tipo) + '</td><td style="' + td + '">' + hx(n.detalle) +
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
    (m.responderA ? 'Si respondes este correo, le llega a tu encargada en Contyme.' : 'Mensaje automático de la app de Novedades.') + '</td></tr></table></div>';
  return x;
}
function rutPuntos(r) { var p = String(r).split('-'); return p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + (p[1] ? '-' + p[1] : ''); }
function deContyme(c) { return RE_CORREO.test(c || '') && String(c).toLowerCase().split('@')[1] === DOMINIO_ENCARGADAS; }
function conCandado(fn) { var l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
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
  if (!e || e.estado === 'reabierto') return { enviado: null, revisado: false, ingresado: false, nota: e ? e.nota || null : null, novedades: null };
  var nov = novedadesDe(e.id);
  if (!conUrls) nov.forEach(function (n) { if (n.adjunto) delete n.adjunto.url; });
  return { enviado: e.fecha, revisado: e.estado === 'revisada' || e.estado === 'ingresada', ingresado: e.estado === 'ingresada',
    nota: null, novedades: nov, sinNovedades: siNo(e.sinNovedades), carpeta: conUrls ? e.carpeta : undefined };
}
function trabajadoresDe(rut) { return filas('Trabajadores').filter(function (x) { return x.rut === rut; }).map(function (x) { return x.nombre; }); }
function datosCliente(rut, periodo) {
  var e = empresa(rut);
  return { rut: rut, nombre: e.nombre, contacto: e.contacto, trabajadores: trabajadoresDe(rut), mes: periodoOk(periodo) ? mesDe(rut, periodo, false) : null };
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
      return { ok: true, rut: rut, token: nuevaSesion({ rol: 'cliente', rut: rut }, HORAS_SESION_CLIENTE), datos: datosCliente(rut, q.periodo) };
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
      return { ok: true, rut: rut, token: nuevaSesion({ rol: 'cliente', rut: rut }, HORAS_SESION_CLIENTE), datos: datosCliente(rut, q.periodo) };
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
    return conCandado(function () {
      var e = empresa(s.rut); if (!e || !siNo(e.habilitada)) return { error: 'El acceso de esta empresa está bloqueado.', salir: true };
      var previo = envioDe(s.rut, q.periodo);
      if (previo && previo.estado !== 'reabierto') return { error: 'Ya enviaste las novedades de ' + nombreMes(q.periodo) + '. Si te faltó algo, escríbele a tu encargada.' };
      var id = previo ? previo.id : azar().slice(0, 12);
      // Carpeta privada del envío: <raíz>/<período>/<RUT> <empresa>
      var claveCar = 'car:' + q.periodo + '|' + s.rut + '|' + (e.correoEncargada || ''), idCar = PropertiesService.getScriptProperties().getProperty(claveCar), car = null;
      if (idCar) { try { car = DriveApp.getFolderById(idCar); } catch (err) { car = null; } }
      if (!car) {
        var raiz = DriveApp.getFolderById(prop('CARPETA_ID'));
        var fp = raiz.getFoldersByName(q.periodo), carPer = fp.hasNext() ? fp.next() : raiz.createFolder(q.periodo);
        var nomCar = s.rut + ' ' + e.nombre, fe = carPer.getFoldersByName(nomCar); car = fe.hasNext() ? fe.next() : carPer.createFolder(nomCar);
        if (deContyme(e.correoEncargada)) { try { car.addViewer(e.correoEncargada); } catch (err) {} }
        PropertiesService.getScriptProperties().setProperty(claveCar, car.getId());
      }
      var urls = {};
      archivos.forEach(function (f) {
        var mime = /^(image\/[a-z0-9.+-]+|application\/pdf)$/i.test(f.mime || '') ? f.mime : 'application/octet-stream';
        var nombre = String(f.nombre || 'respaldo').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80);
        var arch = car.createFile(Utilities.newBlob(Utilities.base64Decode(String(f.base64 || '')), mime, nombre));
        urls[f.idx] = { nombre: nombre, url: arch.getUrl() };
      });
      // Si fue reabierto, se reemplazan las novedades anteriores
      if (previo) { var h = hoja('Novedades'), fl = filas('Novedades').filter(function (x) { return x.envioId === id; }).map(function (x) { return x._fila; }).sort(function (a, b) { return b - a; }); fl.forEach(function (f) { h.deleteRow(f); }); cambio('Novedades'); }
      var conocidos = trabajadoresDe(s.rut), lineas = [];
      lista.forEach(function (n, i) {
        var trab = String(n.trabajador || '').slice(0, 120), adj = urls[i] || null;
        agregar('Novedades', { envioId: id, periodo: q.periodo, rut: s.rut, trabajador: trab, tipo: String(n.tipo || '').slice(0, 40),
          detalle: String(n.detalle || '').slice(0, 300), comentario: String(n.comentario || '').slice(0, 500),
          respaldo: adj ? adj.nombre : (n.adjunto ? n.adjunto.nombre + ' (no llegó el archivo)' : ''), respaldoUrl: adj ? adj.url : '', datos: JSON.stringify(n.datos || {}).slice(0, 2000) });
        if (trab && n.trabajadorNuevo && conocidos.indexOf(trab) < 0) { agregar('Trabajadores', { rut: s.rut, nombre: trab }); conocidos.push(trab); }
        lineas.push({ trabajador: trab, tipo: String(n.tipo || ''), detalle: String(n.detalle || ''), comentario: String(n.comentario || ''), respaldo: adj ? adj.nombre : '', url: adj ? adj.url : '' });
      });
      var env = { id: id, periodo: q.periodo, rut: s.rut, empresa: e.nombre, fecha: ahora(), cantidad: lista.length, sinNovedades: !lista.length, estado: 'enviado', nota: '', carpeta: car.getUrl() };
      if (previo) { env._fila = previo._fila; escribir('Envios', env); } else agregar('Envios', env);
      var cuantas = lista.length ? lista.length + (lista.length === 1 ? ' novedad' : ' novedades') : 'sin novedades';
      registrar(s.rut, 'accion', e.nombre + ' envió ' + cuantas + ' de ' + nombreMes(q.periodo) + '.');
      var para = deContyme(e.correoEncargada) ? e.correoEncargada : admins().join(',');
      var mes = nombreMes(q.periodo), Mes = mes.charAt(0).toUpperCase() + mes.slice(1);
      var sinEnc = !deContyme(e.correoEncargada);
      correo(para, (previo ? '[Corregidas] ' : '') + 'Novedades ' + mes + ' · ' + e.nombre + ' (' + cuantas + ')', {
        titulo: e.nombre + (previo ? ' corrigió sus novedades' : ' envió sus novedades'),
        filas: [['Empresa', e.nombre], ['RUT', rutPuntos(s.rut)], ['Período', Mes], ['Novedades', lista.length ? String(lista.length) : 'Sin novedades este mes'], ['Enviado por', e.contacto || '—']],
        tabla: lineas,
        nota: (lista.length ? '' : 'La empresa avisó que este mes no hubo novedades: se repite la liquidación del mes anterior.\n') +
          (sinEnc ? 'Esta empresa no tiene encargada @contyme.cl: el aviso llegó a los administradores. Asígnala en Empresas (carga masiva).' : ''),
        boton: { txt: 'Abrir la bandeja', url: APP_URL },
        boton2: archivos.length ? { txt: 'Ver respaldos en Drive', url: car.getUrl() } : null });
      if (e.contacto) correo(e.contacto, 'Recibimos tus novedades de ' + mes + ' · ' + e.nombre, {
        titulo: '¡Listo! Recibimos tus novedades',
        intro: 'Contyme recibió las novedades de ' + mes + ' de ' + e.nombre + '. Este es el resumen de lo que enviaste:',
        filas: [['Empresa', e.nombre], ['Período', Mes], ['Novedades', lista.length ? String(lista.length) : 'Sin novedades este mes']],
        tabla: lineas.map(function (n) { return { trabajador: n.trabajador, tipo: n.tipo, detalle: n.detalle, comentario: n.comentario, respaldo: n.respaldo }; }),
        nota: 'Si te faltó algo, responde este correo o escríbele a tu encargada y ella reabre tu envío.',
        boton: { txt: 'Ver en qué va', url: APP_URL },
        responderA: sinEnc ? '' : e.correoEncargada });
      registrar(s.rut, 'correo', 'Aviso a ' + para + ' y confirmación al cliente.');
      return { ok: true, datos: datosCliente(s.rut, q.periodo) };
    });
  },

  /* --- Administradores --- */
  'admin-codigo': function (q) {
    var c = String(q.correo || '').trim().toLowerCase();
    if (admins().indexOf(c) >= 0) {
      var cod = String(Math.floor(100000 + Math.random() * 900000));
      cache().put('c:' + c, JSON.stringify({ cod: cod, n: 0 }), 600);
      correo(c, 'Tu código para entrar a Novedades Contyme: ' + cod, {
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
    return { ok: true, correo: c, token: nuevaSesion({ rol: 'admin', correo: c }, HORAS_SESION_ADMIN) };
  },

  'adm-datos': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var per = periodoOk(q.periodo) ? q.periodo : '';
    var trab = filas('Trabajadores'), envios = filas('Envios').filter(function (x) { return x.periodo === per; });
    var empresas = filas('Empresas').map(function (e) {
      return { rut: e.rut, nombre: e.nombre, correoEncargada: e.correoEncargada, habilitada: siNo(e.habilitada), registrada: !!e.hash,
        intentos: Number(e.intentos || 0), contacto: e.contacto, pideReset: siNo(e.pideReset),
        trabajadores: trab.filter(function (t) { return t.rut === e.rut; }).map(function (t) { return t.nombre; }) };
    });
    var meses = {};
    envios.forEach(function (x) { meses[x.rut] = mesDe(x.rut, per, true); if (x.estado === 'reabierto') meses[x.rut].nota = x.nota; });
    var act = filas('Actividad').slice(-200).map(function (a) { return { t: a.fecha, rut: a.rut || null, tipo: a.tipo, txt: a.texto }; });
    return { ok: true, empresas: empresas, meses: meses, actividad: act };
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

  'adm-bloqueo': function (q) { return conEmpresaAdm(q, function (e, s) { e.habilitada = !!q.habilitada; registrar(e.rut, 'seguridad', e.nombre + ': acceso ' + (e.habilitada ? 'habilitado' : 'bloqueado') + ' por ' + s.correo + '.'); }); },
  'adm-restablecer': function (q) { return conEmpresaAdm(q, function (e, s) { e.hash = ''; e.sal = ''; e.intentos = 0; e.pideReset = false; registrar(e.rut, 'seguridad', e.nombre + ': clave restablecida por ' + s.correo + '.'); }); },

  'adm-marcar': function (q) {
    return conEnvioAdm(q, function (x, e, s) {
      x.estado = q.estado === 'ingresada' ? 'ingresada' : q.estado === 'revisada' ? 'revisada' : 'enviado';
      registrar(e.rut, 'accion', e.nombre + ': marcada «' + x.estado + '» por ' + s.correo + '.');
    });
  },
  'adm-reabrir': function (q) {
    return conEnvioAdm(q, function (x, e, s) {
      x.estado = 'reabierto'; x.nota = String(q.nota || '').slice(0, 500);
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
