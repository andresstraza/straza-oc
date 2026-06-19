// ============================================================
// STRAZA — OC DATABASE · Google Apps Script
// Con soporte JSONP para evitar CORS
// ============================================================

function doGet(e) {
  var result;
  try {
    var action  = e.parameter.action;
    var payload = e.parameter.payload;
    var cb      = e.parameter.callback; // JSONP callback

    if (action === 'getOrdenes')      result = getOrdenes();
    else if (action === 'getProvs')   result = getProvs();
    else if (action === 'getCFG')     result = getCFG();
    else if (payload) {
      var body = JSON.parse(decodeURIComponent(payload));
      var act  = body.action;
      if      (act === 'saveOrden')   result = saveOrden(body.data);
      else if (act === 'anularOrden') result = anularOrden(body.id, body.motivo);
      else if (act === 'saveCFG')     result = saveCFG(body.data);
      else if (act === 'saveProvs')   result = saveProvs(body.data);
      else if (act === 'nextNum')     result = nextNum(body.tipo);
      else result = { error: 'Acción no reconocida' };
    }
    else result = { ok: true, ping: true };
  } catch(err) {
    result = { error: err.message };
  }

  var json = JSON.stringify(result);

  // Si viene parámetro callback → responder JSONP
  if (e.parameter.callback) {
    return ContentService
      .createTextOutput(e.parameter.callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;
    if      (action === 'saveOrden')   result = saveOrden(body.data);
    else if (action === 'anularOrden') result = anularOrden(body.id, body.motivo);
    else if (action === 'saveCFG')     result = saveCFG(body.data);
    else if (action === 'saveProvs')   result = saveProvs(body.data);
    else if (action === 'nextNum')     result = nextNum(body.tipo);
    else result = { error: 'Acción no reconocida' };
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function nextNum(tipo) {
  var sh = getSheet('CFG');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = sh.getDataRange().getValues();
    if (data.length === 0 || data[0][0] === '') {
      sh.clearContents();
      sh.appendRow(['TIPO','CONSECUTIVO']);
      ['COM','SER','TRA','LOG','OFI'].forEach(function(t){ sh.appendRow([t,15]); });
      data = sh.getDataRange().getValues();
    }
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === tipo) {
        var n = parseInt(data[i][1]) || 15;
        sh.getRange(i+1, 2).setValue(n + 1);
        return { num: tipo+'-'+String(n).padStart(3,'0'), ok: true };
      }
    }
    sh.appendRow([tipo, 16]);
    return { num: tipo+'-015', ok: true };
  } finally {
    lock.releaseLock();
  }
}

function getOrdenes() {
  var sh = getSheet('ORDENES');
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { ordenes: [] };
  var ordenes = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    try { ordenes.push(JSON.parse(data[i][1])); } catch(e) {}
  }
  return { ordenes: ordenes };
}

function saveOrden(oc) {
  var sh = getSheet('ORDENES');
  if (sh.getLastRow() === 0)
    sh.appendRow(['ID','JSON','NUM','TIPO','PROVEEDOR','TOTAL','FECHA','ESTADO']);
  var data = sh.getDataRange().getValues();
  var found = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === oc.id) { found = i+1; break; }
  }
  var row = [oc.id, JSON.stringify(oc), oc.num||'', oc.tipo||'',
             oc.prov||'', oc.total||0, oc.fecha||'', oc.anulada?'ANULADA':'ACTIVA'];
  if (found > 0) sh.getRange(found,1,1,8).setValues([row]);
  else sh.appendRow(row);
  return { ok: true };
}

function anularOrden(id, motivo) {
  var sh = getSheet('ORDENES');
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      var oc = JSON.parse(data[i][1]);
      oc.anulada = true;
      oc.motivoAnulacion = motivo;
      oc.fechaAnulacion = new Date().toISOString().slice(0,10);
      sh.getRange(i+1,2).setValue(JSON.stringify(oc));
      sh.getRange(i+1,8).setValue('ANULADA');
      return { ok: true };
    }
  }
  return { error: 'Orden no encontrada' };
}

function getCFG() {
  var sh = getSheet('CFG');
  var data = sh.getDataRange().getValues();
  var cfg = {COM:15,SER:15,TRA:15,LOG:15,OFI:15};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) cfg[String(data[i][0])] = parseInt(data[i][1])||15;
  }
  return { cfg: cfg, ok: true };
}

function saveCFG(cfg) {
  var sh = getSheet('CFG');
  sh.clearContents();
  sh.appendRow(['TIPO','CONSECUTIVO']);
  ['COM','SER','TRA','LOG','OFI'].forEach(function(t){ sh.appendRow([t, cfg[t]||15]); });
  return { ok: true };
}

function getProvs() {
  var sh = getSheet('PROVEEDORES');
  var data = sh.getDataRange().getValues();
  if (data.length <= 1) return { provs: [] };
  var provs = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    try { provs.push(JSON.parse(data[i][1])); } catch(e) {}
  }
  return { provs: provs };
}

function saveProvs(list) {
  var sh = getSheet('PROVEEDORES');
  sh.clearContents();
  sh.appendRow(['NOMBRE','JSON']);
  list.forEach(function(p){ sh.appendRow([p.nombre, JSON.stringify(p)]); });
  return { ok: true };
}
