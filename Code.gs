/**
 * COUTO GARAGE INVENTARIO — backend (Google Apps Script)
 * Va pegado en el editor de Apps Script de la hoja de cálculo del inventario.
 * 1) Ejecuta la función "instalar" una vez.  2) Rellena la hoja Usuarios.
 * 3) Añade GEMINI_API_KEY en Propiedades de la secuencia de comandos.
 * 4) Implementar > Nueva implementación > Aplicación web (Ejecutar como: yo; Acceso: cualquier usuario).
 */
var TABS = {
  Articulos: ['id', 'nombre', 'cantidad', 'estado', 'ubicacion', 'etiquetas', 'categoria', 'minimo', 'notas', 'precio', 'proveedor', 'fecha_compra', 'foto', 'creado', 'modificado', 'por', 'vehiculo', 'fuera', 'fuera_desde', 'imagen'],
  Cajas: ['id', 'tipo', 'ubicacion'],
  Palets: ['id', 'x', 'y', 'rot'],
  Posiciones: ['id', 'x', 'y', 'rot'],
  Usuarios: ['movil', 'pin', 'nombre', 'rol', 'activo'],
  Movimientos: ['fecha', 'usuario', 'accion', 'detalle'],
  Fotos: ['id', 'data'],
  Ubicaciones: ['codigo', 'zona', 'mueble', 'hueco'],
  Kits: ['id', 'nombre', 'articulos', 'cargados']
};
var ITEM_MAP = { name: 'nombre', qty: 'cantidad', st: 'estado', loc: 'ubicacion', tags: 'etiquetas', cat: 'categoria', min: 'minimo', notes: 'notas', price: 'precio', prov: 'proveedor', date: 'fecha_compra', veh: 'vehiculo', out: 'fuera', outd: 'fuera_desde', img: 'imagen' };
var ST_ES = { ok: 'hay', low: 'queda poco', empty: 'vacio' };
var ST_IN = { 'hay': 'ok', 'queda poco': 'low', 'vacio': 'empty' };
var CATS = ['Herramienta manual', 'Herramienta eléctrica', 'Tornillería', 'Recambios', 'Consumibles', 'Lubricantes y químicos', 'Electricidad', 'Ropa SC', 'Bebidas y comida', 'Menaje', 'Limpieza', 'Varios'];
var TZ = 'Europe/Madrid';

/* ---------- utilidades ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function str_(v) { if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd'); return v === null || v === undefined ? '' : String(v); }
function num_(v) { var n = Number(String(v === null || v === undefined ? '' : v).replace(',', '.')); return isFinite(n) ? n : 0; }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function fail_(msg, code) { var e = new Error(msg); e.code = code || 'error'; throw e; }
function props_() { return PropertiesService.getScriptProperties(); }

function Tab_(name) {
  this.name = name; this.cols = TABS[name]; this.sh = ss_().getSheetByName(name);
  if (!this.sh) fail_('Falta la hoja ' + name + '. Ejecuta "instalar".');
  var last = this.sh.getLastRow();
  this.rows = last > 1 ? this.sh.getRange(2, 1, last - 1, this.cols.length).getValues() : [];
  this.n0 = this.rows.length; this.dirty = false;
}
Tab_.prototype.col = function (c) { return this.cols.indexOf(c); };
Tab_.prototype.find = function (id) { var k = String(id); for (var i = 0; i < this.rows.length; i++) if (String(this.rows[i][0]) === k) return i; return -1; };
Tab_.prototype.blank = function () { var r = []; for (var i = 0; i < this.cols.length; i++) r.push(''); return r; };
Tab_.prototype.save = function () {
  if (!this.dirty) return;
  var need = this.rows.length + 1, max = this.sh.getMaxRows();
  if (max < need) this.sh.insertRowsAfter(max, need - max);
  if (this.n0 > 0) this.sh.getRange(2, 1, this.n0, this.cols.length).clearContent();
  if (this.rows.length) this.sh.getRange(2, 1, this.rows.length, this.cols.length).setValues(this.rows);
  this.n0 = this.rows.length; this.dirty = false;
};
function tabs_() { var cache = {}; return function (name) { return cache[name] || (cache[name] = new Tab_(name)); }; }

/* ---------- acceso: móvil + PIN ---------- */
function normMovil_(s) { var d = String(s === null || s === undefined ? '' : s).replace(/\D/g, ''); return d.length > 9 ? d.slice(-9) : d; }
function secret_() { var p = props_(), s = p.getProperty('SECRET'); if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); } return s; }
function sign_(payload) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(payload, secret_())); }
function findUser_(movil) {
  var t = new Tab_('Usuarios');
  for (var i = 0; i < t.rows.length; i++) {
    var r = t.rows[i];
    if (normMovil_(r[0]) === movil && movil && String(r[4]).toUpperCase().trim() !== 'NO') {
      return { movil: movil, pin: str_(r[1]).trim(), nombre: str_(r[2]) || movil, rol: String(r[3]).toLowerCase().trim() === 'admin' ? 'admin' : 'usuario' };
    }
  }
  return null;
}
function userFromToken_(token) {
  var p = String(token || '').split('.');
  if (p.length !== 3) return null;
  if (Number(p[1]) < Date.now()) return null;
  if (sign_(p[0] + '.' + p[1]) !== p[2]) return null;
  return findUser_(p[0]);
}
function login_(req) {
  var movil = normMovil_(req.movil), pin = String(req.pin || '').trim();
  var cache = CacheService.getScriptCache(), key = 'fallos_' + movil, fails = Number(cache.get(key) || 0);
  if (fails >= 5) fail_('Demasiados intentos. Espera 15 minutos.', 'bloqueado');
  var u = findUser_(movil);
  if (!u || !pin || u.pin !== pin) { cache.put(key, String(fails + 1), 900); fail_('Móvil o PIN incorrectos.', 'login'); }
  cache.remove(key);
  var exp = Date.now() + 60 * 24 * 3600 * 1000, payload = movil + '.' + exp;
  return { ok: true, token: payload + '.' + sign_(payload), user: { nombre: u.nombre, rol: u.rol }, data: getData_(tabs_()) };
}

/* ---------- lectura ---------- */
function getData_(T) {
  var a = T('Articulos'), c = a.col.bind(a);
  var items = a.rows.map(function (r) {
    return { id: str_(r[0]), name: str_(r[c('nombre')]), qty: num_(r[c('cantidad')]), st: ST_IN[String(r[c('estado')]).toLowerCase()] || 'ok', loc: str_(r[c('ubicacion')]),
      tags: str_(r[c('etiquetas')]), cat: str_(r[c('categoria')]), min: num_(r[c('minimo')]), notes: str_(r[c('notas')]), price: str_(r[c('precio')]),
      prov: str_(r[c('proveedor')]), date: str_(r[c('fecha_compra')]), foto: str_(r[c('foto')]) ? 1 : 0,
      veh: str_(r[c('vehiculo')]), out: str_(r[c('fuera')]), outd: str_(r[c('fuera_desde')]), img: str_(r[c('imagen')]) };
  }).filter(function (i) { return i.id; });
  var cajas = T('Cajas').rows.map(function (r) { return { id: str_(r[0]), t: str_(r[1]), loc: str_(r[2]) }; }).filter(function (x) { return x.id; });
  var pallets = T('Palets').rows.map(function (r) { return { id: str_(r[0]), x: num_(r[1]), y: num_(r[2]), rot: num_(r[3]) }; }).filter(function (x) { return x.id; });
  var pos = {};
  T('Posiciones').rows.forEach(function (r) { if (str_(r[0])) pos[str_(r[0])] = { x: num_(r[1]), y: num_(r[2]), rot: num_(r[3]) }; });
  var ids_ = function (v) { return str_(v).split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; }); };
  var kits = T('Kits').rows.map(function (r) { return { id: str_(r[0]), name: str_(r[1]), items: ids_(r[2]), done: ids_(r[3]) }; }).filter(function (x) { return x.id; });
  return { items: items, cajas: cajas, pallets: pallets, pos: pos, kits: kits };
}

/* ---------- escritura ---------- */
function itemVal_(k, v) {
  if (k === 'st') return ST_ES[v] || 'hay';
  if (k === 'qty' || k === 'min') return String(Math.max(0, num_(v)));
  if (k === 'img') return /^lotes\/[\w.\-\/]+\.(jpg|png|webp)$/.test(str_(v)) && str_(v).indexOf('..') < 0 ? str_(v) : '';
  return str_(v).slice(0, 2000);
}
function applyOp_(o, T, u) {
  var admin = u.rol === 'admin', t, i, r, k;
  switch (o.t) {
    case 'itemAdd':
      t = T('Articulos'); if (!o.item || !o.item.id || t.find(o.item.id) >= 0) return null;
      r = t.blank(); r[0] = String(o.item.id);
      for (k in ITEM_MAP) if (o.item[k] !== undefined) r[t.col(ITEM_MAP[k])] = itemVal_(k, o.item[k]);
      if (!r[t.col('estado')]) r[t.col('estado')] = 'hay';
      r[t.col('creado')] = now_(); r[t.col('modificado')] = now_(); r[t.col('por')] = u.nombre;
      t.rows.unshift(r); t.dirty = true; return 'Alta: ' + str_(o.item.name) + ' en ' + (str_(o.item.loc) || 'sin ubicar');
    case 'itemSet':
      t = T('Articulos'); i = t.find(o.id); if (i < 0 || !o.f) return null;
      var ch = [];
      for (k in o.f) if (ITEM_MAP[k]) { t.rows[i][t.col(ITEM_MAP[k])] = itemVal_(k, o.f[k]); ch.push(ITEM_MAP[k] + '=' + itemVal_(k, o.f[k])); }
      t.rows[i][t.col('modificado')] = now_(); t.rows[i][t.col('por')] = u.nombre; t.dirty = true;
      return str_(t.rows[i][1]) + ': ' + ch.join(', ');
    case 'itemDel':
      if (!admin) return null;
      t = T('Articulos'); i = t.find(o.id); if (i < 0) return null;
      var nm = str_(t.rows[i][1]); t.rows.splice(i, 1); t.dirty = true;
      var f = T('Fotos'), j = f.find(o.id); if (j >= 0) { f.rows.splice(j, 1); f.dirty = true; }
      return 'Borrado: ' + nm;
    case 'cajaSet':
      t = T('Cajas'); i = t.find(o.id); if (i < 0) return null;
      t.rows[i][2] = str_(o.loc); t.dirty = true; return o.id + ' a ' + (str_(o.loc) || 'sin ubicar');
    case 'cajaAdd':
      t = T('Cajas'); if (!o.id || t.find(o.id) >= 0) return null;
      t.rows.push([String(o.id), str_(o.tipo), '']); t.dirty = true; return 'Nueva caja ' + o.id;
    case 'palSet': case 'posSet':
      if (!admin) return null;
      t = T(o.t === 'palSet' ? 'Palets' : 'Posiciones'); i = t.find(o.id);
      r = [String(o.id), String(num_(o.x)), String(num_(o.y)), String(num_(o.rot))];
      if (i < 0) t.rows.push(r); else t.rows[i] = r;
      t.dirty = true; return 'Recolocado ' + o.id;
    case 'palDel':
      if (!admin) return null;
      t = T('Palets'); i = t.find(o.id); if (i < 0) return null;
      t.rows.splice(i, 1); t.dirty = true;
      var a = T('Articulos'), lc = a.col('ubicacion');
      a.rows.forEach(function (row) { if (str_(row[lc]) === String(o.id)) { row[lc] = ''; a.dirty = true; } });
      return 'Quitado palet ' + o.id;
    case 'kitSet':
      t = T('Kits'); if (!o.id) return null; i = t.find(o.id);
      r = [String(o.id), str_(o.name).slice(0, 120), (o.items || []).map(String).join(','), (o.done || []).map(String).join(',')];
      if (i < 0) t.rows.push(r); else t.rows[i] = r;
      t.dirty = true; return 'Kit ' + str_(o.name) + ': ' + (o.done || []).length + ' de ' + (o.items || []).length + ' cargados';
    case 'kitDel':
      if (!admin) return null;
      t = T('Kits'); i = t.find(o.id); if (i < 0) return null;
      var kn = str_(t.rows[i][1]); t.rows.splice(i, 1); t.dirty = true; return 'Borrado kit ' + kn;
    case 'posDel':
      if (!admin) return null;
      t = T('Posiciones'); i = t.find(o.id); if (i < 0) return null;
      t.rows.splice(i, 1); t.dirty = true; return 'Devuelto a su sitio ' + o.id;
  }
  return null;
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) fail_('La hoja está ocupada, se reintenta enseguida.', 'ocupado');
  try { return fn(); } finally { lock.releaseLock(); }
}
function ops_(u, ops) {
  return withLock_(function () {
    var T = tabs_(), logs = [];
    (ops || []).forEach(function (o) {
      try { var d = applyOp_(o, T, u); if (d) logs.push([now_(), u.nombre, o.t, d]); } catch (e) { logs.push([now_(), u.nombre, 'ERROR ' + o.t, String(e.message || e)]); }
    });
    ['Articulos', 'Cajas', 'Palets', 'Posiciones', 'Fotos', 'Kits'].forEach(function (n) { T(n).save(); });
    if (logs.length) { var m = ss_().getSheetByName('Movimientos'); m.getRange(m.getLastRow() + 1, 1, logs.length, 4).setValues(logs); }
    return { ok: true, data: getData_(T) };
  });
}

/* ---------- fotos de artículos ---------- */
function fotos_(req) {
  var t = new Tab_('Fotos'), out = {}, want = {};
  (req.ids || []).slice(0, 60).forEach(function (id) { want[String(id)] = 1; });
  t.rows.forEach(function (r) { if (want[str_(r[0])]) out[str_(r[0])] = str_(r[1]); });
  return { ok: true, fotos: out };
}
function fotoSet_(u, req) {
  var data = String(req.data || '');
  if (data.indexOf('data:image/') !== 0 || data.length > 49000) fail_('La foto es demasiado grande.', 'foto');
  return withLock_(function () {
    var T = tabs_(), f = T('Fotos'), a = T('Articulos'), i = a.find(req.id);
    if (i < 0) fail_('Ese artículo ya no existe.', 'foto');
    var j = f.find(req.id); if (j < 0) f.rows.push([String(req.id), data]); else f.rows[j][1] = data;
    f.dirty = true; a.rows[i][a.col('foto')] = 'SI'; a.dirty = true; f.save(); a.save();
    return { ok: true };
  });
}
function ubic_(u, req) {
  if (u.rol !== 'admin') fail_('Solo el administrador puede hacer esto.', 'permiso');
  return withLock_(function () {
    var t = new Tab_('Ubicaciones');
    t.rows = (req.rows || []).slice(0, 2000).map(function (r) { return [str_(r[0]), str_(r[1]), str_(r[2]), str_(r[3])]; });
    t.dirty = true; t.save(); return { ok: true, n: t.rows.length };
  });
}

/* ---------- Gemini ---------- */
function geminiKey_() { var k = props_().getProperty('GEMINI_API_KEY'); if (!k) fail_('Falta la clave de Gemini (GEMINI_API_KEY) en las propiedades del proyecto.', 'ia'); return k; }
function geminiModel_(key, refresh) {
  var p = props_(), fixed = p.getProperty('GEMINI_MODEL');
  if (fixed) return fixed.indexOf('models/') === 0 ? fixed : 'models/' + fixed;
  var auto = p.getProperty('GEMINI_MODEL_AUTO'); if (auto && !refresh) return auto;
  var res = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200&key=' + encodeURIComponent(key), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) fail_('Gemini no acepta la clave (' + res.getResponseCode() + ').', 'ia');
  var names = (JSON.parse(res.getContentText()).models || []).filter(function (m) {
    return (m.supportedGenerationMethods || []).indexOf('generateContent') >= 0 && /flash/.test(m.name) && !/image|tts|live|audio|embedding|thinking/.test(m.name);
  }).map(function (m) { return m.name; });
  if (!names.length) fail_('No hay ningún modelo Gemini Flash disponible para esta clave.', 'ia');
  var pick = names.indexOf('models/gemini-flash-latest') >= 0 ? 'models/gemini-flash-latest' : names.filter(function (n) { return !/lite|preview|exp/.test(n); }).sort().pop() || names.sort().pop();
  var alt = names.filter(function (n) { return n !== pick && !/preview|exp/.test(n); }).sort().pop() || '';
  p.setProperty('GEMINI_MODEL_AUTO', pick); p.setProperty('GEMINI_MODEL_ALT', alt); return pick;
}
function gemini_(parts) {
  var key = geminiKey_(), body = { contents: [{ role: 'user', parts: parts }], generationConfig: { temperature: 0.2, responseMimeType: 'application/json' } };
  var call = function (model) {
    return UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/' + model + ':generateContent?key=' + encodeURIComponent(key),
      { method: 'post', contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  };
  var model = geminiModel_(key, false), res = call(model);
  if (res.getResponseCode() === 404) { model = geminiModel_(key, true); res = call(model); }
  // 500 y 503 son saturación pasajera de Google: se reintenta y, si sigue, se prueba con el modelo de reserva.
  for (var n = 1; n <= 2 && (res.getResponseCode() === 503 || res.getResponseCode() === 500); n++) { Utilities.sleep(2500 * n); res = call(model); }
  if (res.getResponseCode() === 503 || res.getResponseCode() === 500) { var alt = props_().getProperty('GEMINI_MODEL_ALT'); if (alt && alt !== model) res = call(alt); }
  if (res.getResponseCode() === 503 || res.getResponseCode() === 500) fail_('Gemini está saturado ahora mismo. No es cosa de la app: vuelve a probar en un minuto.', 'ia');
  if (res.getResponseCode() === 429) fail_('Gemini ha llegado al límite gratuito por ahora. Prueba en un minuto.', 'ia');
  if (res.getResponseCode() !== 200) fail_('Gemini ha devuelto un error (' + res.getResponseCode() + ').', 'ia');
  var j = JSON.parse(res.getContentText()), txt = '';
  try { txt = j.candidates[0].content.parts.map(function (p) { return p.text || ''; }).join(''); } catch (e) { fail_('Gemini no ha devuelto nada utilizable.', 'ia'); }
  txt = txt.replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, '').trim();
  try { return JSON.parse(txt); } catch (e2) { fail_('No se ha podido leer la respuesta de Gemini.', 'ia'); }
}
function cleanItems_(arr) {
  if (!Array.isArray(arr)) arr = arr && Array.isArray(arr.articulos) ? arr.articulos : [];
  return arr.slice(0, 150).map(function (x) {
    return { name: str_(x.nombre).slice(0, 200), qty: Math.max(1, Math.round(num_(x.cantidad)) || 1), tags: str_(x.etiquetas).slice(0, 300),
      cat: CATS.indexOf(str_(x.categoria)) >= 0 ? str_(x.categoria) : 'Varios', price: str_(x.precio), prov: str_(x.proveedor), date: str_(x.fecha), dudoso: x.seguro === false ? 1 : 0,
      box: Array.isArray(x.caja) && x.caja.length === 4 && x.caja.every(function (n) { return isFinite(Number(n)); }) ? x.caja.map(function (n) { return Math.max(0, Math.min(1000, Number(n))); }) : null };
  }).filter(function (x) { return x.name; });
}
function aiFoto_(req) {
  var m = /^data:(image\/[a-z+]+);base64,(.+)$/.exec(String(req.img || ''));
  if (!m) fail_('No ha llegado ninguna imagen.', 'ia');
  var base = 'Responde solo con un array JSON. Cada elemento: {"nombre": string, "cantidad": number, "etiquetas": string, "categoria": string, "seguro": boolean}. ' +
    '"etiquetas" son sinónimos y términos de búsqueda en español separados por espacios. "categoria" es exactamente una de: ' + CATS.join(', ') + '. ';
  var prompt = req.modo === 'compra'
    ? 'Esta imagen es un ticket, factura o captura de un pedido. Extrae cada producto comprado. Añade además "precio" (número, precio unitario en euros), "proveedor" (tienda) y "fecha" (AAAA-MM-DD) si aparecen. No inventes datos que no se lean. ' + base
    : 'Eres un experto en herramientas y material de taller mecánico de competición. Identifica cada herramienta u objeto distinto que se vea en la foto, uno por uno, sin agrupar herramientas diferentes. ' +
      'Nombre concreto en español (por ejemplo "Llave combinada 13 mm", "Alicate de corte diagonal"). Pon la medida o la marca SOLO si se lee claramente en la foto; si no se lee, no la inventes y pon "seguro": false. ' +
      'Si hay varias unidades idénticas, una sola entrada con su cantidad. ' +
      'La foto puede estar girada o del revés. Si es una bandeja o maletín con el nombre o la medida impresos junto a cada hueco, lee esos rótulos y úsalos para nombrar cada pieza (por ejemplo "Vaso 1/2\" 17 mm", "Alargadera 3/8\" de 6 pulgadas"); en ese caso "seguro": true. Si un hueco rotulado está vacío, no lo incluyas. ' +
      'Añade a cada elemento "caja": [ymin, xmin, ymax, xmax], el recuadro que ocupa en la foto en escala de 0 a 1000 (si hay varias unidades, el de una sola). ' + base;
  return { ok: true, items: cleanItems_(gemini_([{ text: prompt }, { inline_data: { mime_type: m[1], data: m[2] } }])) };
}
function aiTags_(req) {
  var r = gemini_([{ text: 'Artículo de un garaje-taller: "' + str_(req.name).slice(0, 200) + '". Devuelve JSON {"etiquetas": string, "categoria": string}. "etiquetas": hasta 10 sinónimos, nombres coloquiales y términos relacionados en español (y gallego si es habitual), separados por espacios, sin repetir el nombre. "categoria": exactamente una de: ' + CATS.join(', ') + '.' }]);
  return { ok: true, tags: str_(r && r.etiquetas).slice(0, 300), cat: CATS.indexOf(str_(r && r.categoria)) >= 0 ? str_(r.categoria) : '' };
}
function aiBuscar_(req) {
  var a = new Tab_('Articulos'), c = a.col.bind(a);
  var lista = a.rows.slice(0, 3000).map(function (r) { return str_(r[0]) + '|' + str_(r[c('nombre')]) + '|' + str_(r[c('etiquetas')]).slice(0, 80) + '|' + str_(r[c('categoria')]); }).join('\n');
  var r = gemini_([{ text: 'Inventario de un garaje (id|nombre|etiquetas|categoría):\n' + lista + '\n\nBúsqueda del usuario: "' + str_(req.q).slice(0, 200) + '".\nDevuelve un array JSON con los id de los artículos que sirven para esa búsqueda (sinónimos, usos parecidos, alternativas), del más al menos adecuado, máximo 20. Si no hay ninguno, [].' }]);
  return { ok: true, ids: (Array.isArray(r) ? r : []).map(String) };
}
function aiDonde_(req) {
  var a = new Tab_('Articulos'), c = a.col.bind(a), cj = new Tab_('Cajas'), enCaja = {}, por = {};
  cj.rows.forEach(function (r) { enCaja[str_(r[0])] = str_(r[2]); });
  a.rows.forEach(function (r) { var l = str_(r[c('ubicacion')]); if (enCaja[l] !== undefined) l = enCaja[l]; if (!l) return; (por[l] = por[l] || []).push(str_(r[c('nombre')])); });
  var lista = (req.locs || []).slice(0, 400).map(function (l) { var k = str_(l.c); return k + ' | ' + str_(l.n) + ' | ' + (por[k] ? por[k].slice(0, 8).join('; ') : 'vacío'); }).join('\n');
  var r = gemini_([{ text: 'Huecos de un garaje-taller (código | descripción | qué hay ahora):\n' + lista + '\n\nNuevo artículo: "' + str_(req.name).slice(0, 200) + '".\nElige el mejor hueco: junto a cosas parecidas, lo pesado abajo, lo de uso frecuente a mano, comida y bebida en el living, trastos voluminosos en el foso. Devuelve JSON {"codigo": string, "motivo": string} con un código exacto de la lista y un motivo de una frase en español.' }]);
  return { ok: true, code: str_(r && r.codigo), why: str_(r && r.motivo).slice(0, 300) };
}

/* ---------- entrada web ---------- */
function prepararHojas_() {
  var p = props_(); if (p.getProperty('ESQUEMA') === '3') return;
  var ss = ss_();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name), cols = TABS[name];
    if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#111013').setFontColor('#ff2e9a');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), cols.length).setNumberFormat('@');
  });
  p.setProperty('ESQUEMA', '3');
}
function route_(req) {
  prepararHojas_();
  if (req.action === 'login') return login_(req);
  var u = userFromToken_(req.token); if (!u) fail_('La sesión ha caducado. Vuelve a entrar.', 'auth');
  switch (req.action) {
    case 'sync': return { ok: true, data: getData_(tabs_()), user: { nombre: u.nombre, rol: u.rol } };
    case 'ops': return ops_(u, req.ops);
    case 'fotos': return fotos_(req);
    case 'fotoSet': return fotoSet_(u, req);
    case 'ubic': return ubic_(u, req);
    case 'ai.foto': return aiFoto_(req);
    case 'ai.tags': return aiTags_(req);
    case 'ai.buscar': return aiBuscar_(req);
    case 'ai.donde': return aiDonde_(req);
  }
  fail_('Acción desconocida.', 'error');
}
function doPost(e) {
  var out;
  try { out = route_(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, error: String((err && err.message) || err), code: (err && err.code) || 'error' }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('Couto Garage Inventario: servicio activo.'); }

/* ---------- instalación y mantenimiento ---------- */
function instalar() {
  var ss = ss_();
  props_().deleteProperty('ESQUEMA'); prepararHojas_();
  var pal = ss.getSheetByName('Palets');
  if (pal.getLastRow() < 2) {
    var P = [[5.67, 3.28, -24], [6.75, 2.75, -24], [7.85, 2.2, -24], [1.75, 6.3, 90], [1.85, 7.35, 0], [3, 6.2, 0], [4.3, 6.2, 0], [5.75, 5.15, 0], [5.75, 6.2, 0], [2.35, 10.3, 90], [2.35, 11.6, 90], [3.05, 13.35, 0], [4.4, 13.35, 0], [6.65, 10.1, 90], [6.65, 11.45, 90], [6.65, 12.8, 90], [9.1, 10.3, 90], [9.1, 11.55, 90], [9.1, 12.8, 90]];
    pal.getRange(2, 1, P.length, 4).setValues(P.map(function (p, i) { return ['FOS-P' + (i < 9 ? '0' : '') + (i + 1), String(p[0]), String(p[1]), String(p[2])]; }));
  }
  var cj = ss.getSheetByName('Cajas');
  if (cj.getLastRow() < 2) {
    var C = [], add = function (pre, n, tipo) { for (var i = 1; i <= n; i++) C.push([pre + '-' + (i < 10 ? '0' : '') + i, tipo, '']); };
    add('C5', 12, 'Caja 5 L'); add('C11', 8, 'Caja 11 L'); add('ORG', 3, 'Organizador con tapa'); add('GAV', 24, 'Gaveta de tornillería');
    cj.getRange(2, 1, C.length, 3).setValues(C);
  }
  var us = ss.getSheetByName('Usuarios');
  if (us.getLastRow() < 2) us.getRange(2, 1, 3, 5).setValues([['', '', 'Iván', 'admin', 'SI'], ['', '', 'Padre', 'usuario', 'SI'], ['', '', 'Hermano', 'usuario', 'SI']]);
  var first = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (first && first.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(first);
  secret_();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'avisoSemanal'; });
  if (!has) ScriptApp.newTrigger('avisoSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  ss.toast('Instalado. Ahora rellena móvil y PIN en la hoja Usuarios.', 'Couto Garage', 10);
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Couto Garage').addItem('Instalar o reparar hojas', 'instalar').addItem('Probar Gemini', 'probarGemini').addItem('Enviar aviso ahora', 'avisoSemanal').addToUi();
}
function probarGemini() {
  var msg;
  try { var r = gemini_([{ text: 'Devuelve el JSON {"ok": true}' }]); msg = 'Gemini responde bien. Modelo: ' + geminiModel_(geminiKey_(), false) + '\n' + JSON.stringify(r); }
  catch (e) { msg = 'Fallo: ' + e.message; }
  try { SpreadsheetApp.getUi().alert(msg); } catch (e2) { Logger.log(msg); }
  return msg;
}
function avisoSemanal() {
  var a = new Tab_('Articulos'), c = a.col.bind(a), vac = [], poco = [];
  a.rows.forEach(function (r) {
    var e = String(r[c('estado')]).toLowerCase(), l = '- ' + str_(r[c('nombre')]) + ' (' + (str_(r[c('ubicacion')]) || 'sin ubicar') + ')';
    if (e === 'vacio') vac.push(l); else if (e === 'queda poco') poco.push(l + ', quedan ' + str_(r[c('cantidad')]));
  });
  if (!vac.length && !poco.length) return 'Nada que avisar';
  var body = 'Couto Garage, lista de la compra:\n\n' + (vac.length ? 'VACÍO\n' + vac.join('\n') + '\n\n' : '') + (poco.length ? 'QUEDA POCO\n' + poco.join('\n') + '\n' : '');
  MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Couto Garage: ' + (vac.length + poco.length) + ' cosas por reponer', body);
  return body;
}
