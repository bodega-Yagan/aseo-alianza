/**
 * APP PEDIDOS DE ASEO - ALIANZA
 * Planilla: Pedidos_Aseo_Alianza
 * Hojas: Productos, Personal, Pedidos, Detalle, Resumen, Config
 */

const SS_ID = '1EZ_kQuu8dKq3peSlQAtRSNXoS59BfEEdUi7l2QkaRow';
const TZ = 'America/Santiago';

function ss_() { return SpreadsheetApp.openById(SS_ID); }
function hoja_(nombre) { return ss_().getSheetByName(nombre); }
function txt_(v) { return String(v == null ? '' : v).trim(); }
function fmt_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) !== '[object Date]') return txt_(d);
  return Utilities.formatDate(d, TZ, 'dd-MM-yyyy HH:mm');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Pedidos Aseo Alianza')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------- Acceso desde la app instalable (GitHub Pages) ---------- */
function doPost(e) {
  const API = {
    getInicio: getInicio, login: login, crearPedido: crearPedido, misPedidos: misPedidos,
    adminLogin: adminLogin, getPendientes: getPendientes, getHistorial: getHistorial,
    registrarEntrega: registrarEntrega
  };
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = API[req.fn];
    if (!f) throw new Error('Función no válida');
    out = { ok: true, data: f.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: err.message || String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Datos base ---------- */

function getInicio() {
  const per = hoja_('Personal').getDataRange().getValues().slice(1)
    .filter(r => txt_(r[0]) && txt_(r[2]).toUpperCase() !== 'NO')
    .map(r => txt_(r[0]));
  return { personal: per, productos: productos_() };
}

function productos_() {
  return hoja_('Productos').getDataRange().getValues().slice(1)
    .filter(r => txt_(r[0]) && txt_(r[2]).toUpperCase() !== 'NO')
    .map(r => ({ nombre: txt_(r[0]), unidad: txt_(r[1]) || 'Unidad' }));
}

const UNIDADES = ['Unidad', 'Litros', 'CC', 'ML'];
// Unidad "base" con que se pide un producto: LT -> Litros, todo lo demás -> Unidad
function unidadBase_(u) { return /^(lt|lts|litro|litros)$/i.test(txt_(u)) ? 'Litros' : 'Unidad'; }

function config_(clave) {
  const filas = hoja_('Config').getDataRange().getValues();
  for (const r of filas) if (txt_(r[0]).toLowerCase() === clave.toLowerCase()) return txt_(r[1]);
  return '';
}

/* ---------- Personal de aseo ---------- */

function validarPersona_(nombre, pin) {
  const filas = hoja_('Personal').getDataRange().getValues().slice(1);
  const f = filas.find(r => txt_(r[0]) === txt_(nombre) && txt_(r[2]).toUpperCase() !== 'NO');
  if (!f) throw new Error('Nombre no encontrado o inactivo.');
  const pinHoja = txt_(f[1]);
  if (pinHoja && pinHoja !== txt_(pin)) throw new Error('PIN incorrecto.');
  return true;
}

function login(nombre, pin) {
  validarPersona_(nombre, pin);
  return { ok: true, nombre: nombre };
}

function crearPedido(nombre, pin, items, comentario) {
  validarPersona_(nombre, pin);
  const validos = productos_().map(p => p.nombre);
  items = (items || [])
    .map(i => ({ producto: txt_(i.producto), cantidad: Math.floor(Number(i.cantidad) || 0) }))
    .filter(i => i.cantidad > 0 && validos.indexOf(i.producto) >= 0);
  const uniProd = {};
  productos_().forEach(p => { uniProd[p.nombre] = unidadBase_(p.unidad); });
  if (!items.length) throw new Error('El pedido no tiene productos.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const shP = hoja_('Pedidos');
    const shD = hoja_('Detalle');
    const n = Math.max(shP.getLastRow(), 1); // fila 1 = encabezado
    const folio = 'ASEO-' + String(n).padStart(4, '0');
    const ahora = new Date();
    const total = items.reduce((s, i) => s + i.cantidad, 0);

    shP.appendRow([folio, ahora, nombre, txt_(comentario), 'PENDIENTE', '', '', '', total, '', '', '']);
    const filas = items.map(i => [folio, ahora, nombre, i.producto, i.cantidad, '', '', 'PENDIENTE', '', '', '', uniProd[i.producto] || 'Unidad', '']);
    shD.getRange(shD.getLastRow() + 1, 1, filas.length, 13).setValues(filas);
    return { ok: true, folio: folio };
  } finally {
    lock.releaseLock();
  }
}

function misPedidos(nombre, pin) {
  validarPersona_(nombre, pin);
  return pedidos_(p => p.solicitante === txt_(nombre), 30);
}

/* ---------- Lectura de pedidos ---------- */

function pedidos_(filtro, limite) {
  const cab = hoja_('Pedidos').getDataRange().getValues().slice(1);
  const det = hoja_('Detalle').getDataRange().getValues().slice(1);
  const porFolio = {};
  det.forEach(r => {
    const f = txt_(r[0]);
    if (!f) return;
    (porFolio[f] = porFolio[f] || []).push({
      producto: txt_(r[3]),
      solicitada: Number(r[4]) || 0,
      entregada: r[5] === '' ? null : Number(r[5]),
      diferencia: r[6] === '' ? null : Number(r[6]),
      estado: txt_(r[7]),
      obs: txt_(r[10]),
      unidadPed: txt_(r[11]) || 'Unidad',
      unidadEnt: txt_(r[12])
    });
  });
  let lista = cab.filter(r => txt_(r[0])).map(r => ({
    folio: txt_(r[0]),
    fecha: fmt_(r[1]),
    solicitante: txt_(r[2]),
    comentario: txt_(r[3]),
    estado: txt_(r[4]),
    fechaEntrega: fmt_(r[5]),
    entregadoPor: txt_(r[6]),
    obs: txt_(r[7]),
    totalSol: Number(r[8]) || 0,
    totalEnt: r[9] === '' ? null : Number(r[9]),
    nEntrega: txt_(r[10]),
    docErp: txt_(r[11]),
    lineas: porFolio[txt_(r[0])] || []
  })).filter(filtro);
  lista.reverse(); // más recientes primero
  return limite ? lista.slice(0, limite) : lista;
}

/* ---------- Encargado ---------- */

function validarAdmin_(clave) {
  if (txt_(clave) !== config_('Clave encargado')) throw new Error('Clave de encargado incorrecta.');
}

function adminLogin(clave) {
  validarAdmin_(clave);
  return { ok: true, nombre: config_('Nombre encargado') || 'Encargado' };
}

function getPendientes(clave) {
  validarAdmin_(clave);
  return pedidos_(p => p.estado === 'PENDIENTE', 0);
}

function getHistorial(clave) {
  validarAdmin_(clave);
  return pedidos_(p => p.estado !== 'PENDIENTE', 60);
}

/**
 * lineas: [{producto, entregada, unidad, obs}]
 */
function registrarEntrega(clave, folio, lineas, obs, docErp) {
  validarAdmin_(clave);
  const entregaPor = config_('Nombre encargado') || 'Encargado';
  const mapa = {}, mapaObs = {}, mapaUni = {};
  (lineas || []).forEach(l => {
    const p = txt_(l.producto);
    mapa[p] = Math.max(0, Number(String(l.entregada).replace(',', '.')) || 0);
    mapaObs[p] = txt_(l.obs);
    mapaUni[p] = UNIDADES.indexOf(txt_(l.unidad)) >= 0 ? txt_(l.unidad) : 'Unidad';
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const shP = hoja_('Pedidos');
    const shD = hoja_('Detalle');
    const cab = shP.getDataRange().getValues();
    let filaP = -1;
    for (let i = 1; i < cab.length; i++) if (txt_(cab[i][0]) === folio) { filaP = i + 1; break; }
    if (filaP < 0) throw new Error('No se encontró el pedido ' + folio);
    if (txt_(cab[filaP - 1][4]) !== 'PENDIENTE') throw new Error('Ese pedido ya fue registrado como entregado.');

    const ahora = new Date();
    // N° de documento de entrega correlativo (ENT-0001, ENT-0002...)
    let maxN = 0;
    for (let i = 1; i < cab.length; i++) {
      const m = /^ENT-(\d+)$/.exec(txt_(cab[i][10]));
      if (m) maxN = Math.max(maxN, Number(m[1]));
    }
    const nEntrega = 'ENT-' + String(maxN + 1).padStart(4, '0');
    const det = shD.getDataRange().getValues();
    let totSol = 0, totEnt = 0, nLin = 0, nComp = 0, nNada = 0;
    for (let i = 1; i < det.length; i++) {
      if (txt_(det[i][0]) !== folio) continue;
      const prod = txt_(det[i][3]);
      const sol = Number(det[i][4]) || 0;
      const ent = mapa.hasOwnProperty(prod) ? mapa[prod] : 0;
      const uPed = txt_(det[i][11]) || 'Unidad';
      const uEnt = mapaUni[prod] || uPed;
      const misma = uEnt === uPed;
      let dif, est;
      if (ent <= 0) { dif = sol; est = 'NO ENTREGADO'; }
      else if (!misma) { dif = ''; est = 'OTRA UNIDAD'; }
      else { dif = Math.max(0, sol - ent); est = ent >= sol ? 'COMPLETO' : 'PARCIAL'; }
      shD.getRange(i + 1, 6, 1, 8).setValues([[ent, dif, est, ahora, nEntrega, mapaObs[prod] || '', uPed, uEnt]]);
      nLin++; if (est === 'COMPLETO') nComp++; if (est === 'NO ENTREGADO') nNada++;
      totSol += sol; if (misma) totEnt += ent;
    }
    const estadoPed = nNada === nLin ? 'NO ENTREGADO' : (nComp === nLin ? 'ENTREGADO COMPLETO' : 'ENTREGADO PARCIAL');
    shP.getRange(filaP, 5, 1, 8).setValues([[estadoPed, ahora, entregaPor, txt_(obs), totSol, totEnt, nEntrega, txt_(docErp)]]);
    return { ok: true, estado: estadoPed, nEntrega: nEntrega };
  } finally {
    lock.releaseLock();
  }
}
