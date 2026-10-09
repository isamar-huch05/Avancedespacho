/*  DESPACHO POR FLETE — servidor en Google Sheets (Apps Script)
    -------------------------------------------------------------
    Guarda el reporte del turno (hoja BD) y todas las lecturas de los celulares.
    Hojas que crea solo:
      · EVENTOS  → cada lectura, quitar, cierre y reapertura (con operario, hora y resultado)
      · BD_DATA  → el reporte publicado por la supervisora (uso interno, no editar)

    Cambia estas dos claves antes de publicar.
    CLAVE         → la misma que va en config.js (la usan todos los celulares)
    CLAVE_SUP     → solo la coordinadora (publicar el reporte y entrar a Coordinadora / Avance general)
*/
const CLAVE     = 'tailoy-despacho-2026';
const CLAVE_SUP = 'supervisora-2026';

const EVH = ['id','ts','hora','turno','tipo','flete','entrega','codigo','resultado','operario','via','dispositivo','ref','obs'];

function doGet() { return out({ok:true, app:'despacho-flete'}); }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ok:false, error:'json'}); }
  if (b.clave !== CLAVE) return out({ok:false, error:'clave'});
  try {
    switch (b.accion) {
      case 'estado':      return out(Object.assign({ok:true}, meta()));
      case 'bd_get':      return out(bdGet());
      case 'sup_check':   return out(b.clave_sup === CLAVE_SUP ? {ok:true} : {ok:false, error:'clave_sup'});
      case 'bd_set':      if (b.clave_sup !== CLAVE_SUP) return out({ok:false, error:'clave_sup'});
                          return out(bdSet(b));
      case 'eventos_get': return out(evGet(b));
      case 'eventos_add': return out(evAdd(b));
    }
    return out({ok:false, error:'accion'});
  } catch (err) {
    return out({ok:false, error:String(err)});
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function libro() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props() { return PropertiesService.getScriptProperties(); }

function hojaEventos() {
  let s = libro().getSheetByName('EVENTOS');
  if (!s) {
    s = libro().insertSheet('EVENTOS');
    s.getRange(1, 1, 1, EVH.length).setValues([EVH]).setFontWeight('bold');
    s.setFrozenRows(1);
    s.getRange('A:N').setNumberFormat('@');
  }
  return s;
}

function meta() {
  const p = props();
  return {
    version: p.getProperty('bd_version') || '',
    turno:   p.getProperty('turno') || '',
    name:    p.getProperty('bd_name') || '',
    fecha:   p.getProperty('bd_fecha') || ''
  };
}

function bdGet() {
  const m = meta();
  const s = libro().getSheetByName('BD_DATA');
  if (!s || !m.version || s.getLastRow() < 1) return Object.assign({ok:true, vacio:true}, m);
  const txt = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(r => r[0]).join('');
  return Object.assign({ok:true, data:JSON.parse(txt)}, m);
}

function bdSet(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let s = libro().getSheetByName('BD_DATA');
    if (!s) { s = libro().insertSheet('BD_DATA'); s.hideSheet(); }
    s.clear();
    const txt = JSON.stringify({cols:b.cols, rows:b.rows, cambios:b.cambios || []});
    const partes = [];
    for (let i = 0; i < txt.length; i += 40000) partes.push([txt.slice(i, i + 40000)]);
    s.getRange(1, 1, partes.length, 1).setNumberFormat('@').setValues(partes);
    const p = props(), v = String(Date.now());
    p.setProperty('bd_version', v);
    p.setProperty('bd_name', b.name || '');
    p.setProperty('bd_fecha', b.fecha || '');
    if (b.nuevo_turno || !p.getProperty('turno')) {
      p.setProperty('turno', v);
      p.setProperty('turno_fila', String(hojaEventos().getLastRow()));
    }
    return Object.assign({ok:true}, meta());
  } finally {
    lock.releaseLock();
  }
}

function evGet(b) {
  const m = meta();
  const s = hojaEventos();
  const last = s.getLastRow();
  const desde = Math.max(1, Number(b.desde) || 0, Number(props().getProperty('turno_fila')) || 0);
  let eventos = [];
  if (last > desde) {
    eventos = s.getRange(desde + 1, 1, last - desde, EVH.length).getValues()
      .filter(r => String(r[3]) === m.turno)
      .map(r => { const o = {}; EVH.forEach((k, i) => o[k] = String(r[i])); return o; });
  }
  return {ok:true, turno:m.turno, version:m.version, hasta:last, eventos:eventos};
}

function evAdd(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const m = meta();
    const s = hojaEventos();
    const filas = (b.eventos || []).map(e => EVH.map(k => k === 'turno' ? String(e.turno || m.turno) : (e[k] == null ? '' : String(e[k]))));
    if (filas.length) s.getRange(s.getLastRow() + 1, 1, filas.length, EVH.length).setNumberFormat('@').setValues(filas);
    return {ok:true, n:filas.length, hasta:s.getLastRow(), turno:m.turno};
  } finally {
    lock.releaseLock();
  }
}
