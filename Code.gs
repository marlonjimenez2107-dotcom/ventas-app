// Pega este archivo como Code.gs en Apps Script.
const VENTAS = ['Fecha','Código','Producto','Talla','Cant.','Precio','Desc. %','Total','Costo','Utilidad','Pago'];

// Ejecútala UNA vez: crea las hojas con datos de ejemplo.
function setup() {
  const ss = SpreadsheetApp.getActive();
  const inv = ss.getSheetByName('Inventario') || ss.insertSheet('Inventario');
  inv.clear();
  inv.getRange(1, 1, 1, 6).setValues([['Código','Producto','Talla','Costo','Precio','Stock']]);
  inv.getRange(2, 1, 4, 6).setValues([
    ['CAM-01','Camiseta básica','M',18000,39900,20],
    ['CAM-02','Camiseta básica','L',18000,39900,15],
    ['JEA-01','Jean slim','30',45000,99900,10],
    ['BUZ-01','Buzo con capucha','M',38000,89900,6]]);
  const v = ss.getSheetByName('Ventas') || ss.insertSheet('Ventas');
  v.clear();
  v.getRange(1, 1, 1, VENTAS.length).setValues([VENTAS]);
}

// Cambia este PIN por uno tuyo. Los vendedores lo escriben al entrar a la app.
const PIN = 'CAMBIA-ESTE-PIN';

function respuesta(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Ventas').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (p.pin !== PIN) return respuesta({ error: 'PIN incorrecto' });
  try {
    if (p.action === 'productos') return respuesta({ productos: productos() });
    if (p.action === 'vender') return respuesta(registrarVenta(p.id, p.q, p.d, p.pago));
    return respuesta({ error: 'Acción no válida' });
  } catch (err) {
    return respuesta({ error: err.message });
  }
}

function productos() {
  const d = SpreadsheetApp.getActive().getSheetByName('Inventario').getDataRange().getValues();
  d.shift();
  return d.map(r => ({ id: r[0], n: r[1], t: r[2], p: r[4], s: r[5] }));
}

function registrarVenta(id, q, d, pago) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = SpreadsheetApp.getActive();
    const inv = ss.getSheetByName('Inventario');
    const data = inv.getDataRange().getValues();
    const i = data.findIndex((r, k) => k > 0 && r[0] === id);
    if (i < 1) throw new Error('El producto no existe');
    const r = data[i];
    q = Math.max(1, Number(q) || 1);
    d = Math.min(100, Math.max(0, Number(d) || 0));
    if (q > r[5]) throw new Error('Stock insuficiente');
    const tot = r[4] * q * (1 - d / 100), co = r[3] * q;
    ss.getSheetByName('Ventas').appendRow([new Date(), id, r[1], r[2], q, r[4], d, tot, co, tot - co, pago]);
    inv.getRange(i + 1, 6).setValue(r[5] - q);
    return { total: tot };
  } finally { lock.releaseLock(); }
}
