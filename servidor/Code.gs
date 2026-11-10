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
    var q = JSON.parse(e.postData.contents);
    var f = ACCIONES[q.accion];
    r = f ? f(q) : { error: 'Acción no válida.' };
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

function prop(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function libro() { return SpreadsheetApp.openById(prop('SHEET_ID')); }
function hoja(n) { return libro().getSheetByName(n); }
function filas(n) {
  var v = hoja(n).getDataRange().getValues(), cab = HOJAS[n], out = [];
  for (var i = 1; i < v.length; i++) { var o = { _fila: i + 1 }; cab.forEach(function (c, j) { o[c] = v[i][j]; }); out.push(o); }
  return out;
}
function escribir(n, o) { hoja(n).getRange(o._fila, 1, 1, HOJAS[n].length).setValues([HOJAS[n].map(function (c) { return o[c] === undefined ? '' : o[c]; })]); }
function agregar(n, o) { hoja(n).appendRow(HOJAS[n].map(function (c) { return o[c] === undefined ? '' : o[c]; })); }
function ahora() { return new Date().toISOString(); }
function hex(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function sha(t) { return hex(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)); }
function azar() { return sha(Utilities.getUuid() + Utilities.getUuid() + Date.now()).slice(0, 40); }
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
function correo(para, asunto, cuerpo) {
  if (!para || !RE_CORREO.test(para)) return false;
  MailApp.sendEmail({ to: para, subject: asunto, body: cuerpo + '\n\n—\nNovedades Contyme · mensaje automático, no respondas este correo.', name: 'Novedades Contyme' });
  return true;
}
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
      var ok = e.hash ? huella(clave, e.sal) === e.hash : (prop('CLAVE_INICIAL') !== '' && clave === prop('CLAVE_INICIAL'));
      if (!ok) {
        e.intentos = Number(e.intentos || 0) + 1; escribir('Empresas', e);
        if (e.intentos >= MAX_INTENTOS) { registrar(rut, 'seguridad', e.nombre + ': acceso bloqueado por ' + MAX_INTENTOS + ' intentos fallidos.'); return { error: BLOQUEO }; }
        var quedan = MAX_INTENTOS - e.intentos;
        return { campo: 'clave', error: GENERICO + (quedan <= 2 ? ' Te quedan ' + quedan + (quedan === 1 ? ' intento.' : ' intentos.') : '') };
      }
      e.intentos = 0; escribir('Empresas', e);
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
      e.sal = azar(); e.hash = huella(nueva, e.sal); e.contacto = contacto; e.intentos = 0; e.pideReset = false; e.actualizada = ahora();
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
        correo(para, 'Novedades · ' + e.nombre + ' pide restablecer su clave',
          e.nombre + ' (' + rut + ') pidió restablecer la clave de la app de Novedades.\nEntra a la app como administrador → Empresas → «Restablecer clave».');
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
      var raiz = DriveApp.getFolderById(prop('CARPETA_ID'));
      var fp = raiz.getFoldersByName(q.periodo), carPer = fp.hasNext() ? fp.next() : raiz.createFolder(q.periodo);
      var nomCar = s.rut + ' ' + e.nombre, fe = carPer.getFoldersByName(nomCar), car = fe.hasNext() ? fe.next() : carPer.createFolder(nomCar);
      if (deContyme(e.correoEncargada)) { try { car.addViewer(e.correoEncargada); } catch (err) {} }
      var urls = {};
      archivos.forEach(function (f) {
        var mime = /^(image\/[a-z0-9.+-]+|application\/pdf)$/i.test(f.mime || '') ? f.mime : 'application/octet-stream';
        var nombre = String(f.nombre || 'respaldo').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80);
        var arch = car.createFile(Utilities.newBlob(Utilities.base64Decode(String(f.base64 || '')), mime, nombre));
        urls[f.idx] = { nombre: nombre, url: arch.getUrl() };
      });
      // Si fue reabierto, se reemplazan las novedades anteriores
      if (previo) { var h = hoja('Novedades'), fl = filas('Novedades').filter(function (x) { return x.envioId === id; }).map(function (x) { return x._fila; }).sort(function (a, b) { return b - a; }); fl.forEach(function (f) { h.deleteRow(f); }); }
      var conocidos = trabajadoresDe(s.rut), lineas = [];
      lista.forEach(function (n, i) {
        var trab = String(n.trabajador || '').slice(0, 120), adj = urls[i] || null;
        agregar('Novedades', { envioId: id, periodo: q.periodo, rut: s.rut, trabajador: trab, tipo: String(n.tipo || '').slice(0, 40),
          detalle: String(n.detalle || '').slice(0, 300), comentario: String(n.comentario || '').slice(0, 500),
          respaldo: adj ? adj.nombre : (n.adjunto ? n.adjunto.nombre + ' (no llegó el archivo)' : ''), respaldoUrl: adj ? adj.url : '', datos: JSON.stringify(n.datos || {}).slice(0, 2000) });
        if (trab && n.trabajadorNuevo && conocidos.indexOf(trab) < 0) { agregar('Trabajadores', { rut: s.rut, nombre: trab }); conocidos.push(trab); }
        lineas.push('• ' + trab + ' — ' + n.tipo + ': ' + (n.detalle || '') + (n.comentario ? ' («' + n.comentario + '»)' : '') + (adj ? ' [respaldo adjunto en la carpeta]' : ''));
      });
      var env = { id: id, periodo: q.periodo, rut: s.rut, empresa: e.nombre, fecha: ahora(), cantidad: lista.length, sinNovedades: !lista.length, estado: 'enviado', nota: '', carpeta: car.getUrl() };
      if (previo) { env._fila = previo._fila; escribir('Envios', env); } else agregar('Envios', env);
      var cuantas = lista.length ? lista.length + (lista.length === 1 ? ' novedad' : ' novedades') : 'sin novedades';
      registrar(s.rut, 'accion', e.nombre + ' envió ' + cuantas + ' de ' + nombreMes(q.periodo) + '.');
      var para = deContyme(e.correoEncargada) ? e.correoEncargada : admins().join(',');
      correo(para, 'Novedades de ' + nombreMes(q.periodo) + ' · ' + e.nombre + ' (' + cuantas + ')',
        e.nombre + ' (' + s.rut + ') envió ' + cuantas + ' de ' + nombreMes(q.periodo) + (previo ? ' (corregidas)' : '') + '.\n\n' +
        (lineas.length ? lineas.join('\n') + '\n\n' : 'Avisó que este mes no tuvo novedades.\n\n') +
        (archivos.length ? 'Respaldos (carpeta privada en Drive): ' + car.getUrl() + '\n' : '') +
        'Revisa y marca el avance en la app de Novedades (administración).');
      if (e.contacto) correo(e.contacto, 'Recibimos tus novedades de ' + nombreMes(q.periodo),
        'Hola,\n\nContyme recibió las novedades de ' + nombreMes(q.periodo) + ' de ' + e.nombre + ' (' + cuantas + ').\nPuedes ver en qué van entrando a la app.');
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
      correo(c, 'Tu código para entrar a Novedades Contyme: ' + cod, 'Tu código es ' + cod + '. Vence en 10 minutos.\nSi no lo pediste, ignora este correo.');
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
      correo(e.contacto, 'Contyme te pidió revisar tus novedades de ' + nombreMes(x.periodo),
        'Hola,\n\nTu encargada en Contyme reabrió el envío de ' + nombreMes(x.periodo) + ' de ' + e.nombre + ' con este mensaje:\n\n«' + x.nota + '»\n\nEntra a la app, corrige y vuelve a enviar.');
    });
  },
  'adm-recordar': function (q) {
    var s = sesionDe(q, 'admin'); if (!s) return { error: 'Tu sesión venció. Vuelve a entrar.', salir: true };
    var e = empresa(normRut(q.rut)); if (!e || !periodoOk(q.periodo)) return { error: 'Empresa no encontrada.' };
    var x = envioDe(e.rut, q.periodo); if (x && x.estado !== 'reabierto') return { ok: true };
    var enviado = correo(e.contacto, 'Recuerda enviar tus novedades de ' + nombreMes(q.periodo),
      'Hola,\n\nTe recordamos enviar a Contyme las novedades de ' + nombreMes(q.periodo) + ' de ' + e.nombre + (q.plazo ? ' antes del ' + q.plazo : '') + ', o avisar que no hubo novedades, desde la app de Novedades.');
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
