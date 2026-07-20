// FarmaPos — Apps Script Backend
// Pega este código en script.google.com y despliega como Web App

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── LICENCIA PRO: se valida en el Panel Maestro de Omnia (Control Central),
// NO en este script. Cada cliente ya no necesita su propia hoja "licencias":
// la app llama directo a la URL /exec del Panel Maestro (action=verify).
// Este backend solo guarda los datos del negocio (productos, ventas, etc.).

function doGet(e) {
  const action = e.parameter.action || '';

  if (action === 'ping') {
    return json({ok:true, msg:'pong', ts:new Date().toISOString()});
  }

  const sheet = getSheet(e.parameter.sheet || 'productos');
  if (action === 'get') {
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data = rows.slice(1).map(r => Object.fromEntries(headers.map((h,i)=>[h,r[i]])));
    return json({ok:true, data});
  }
  return json({ok:false, error:'Unknown action'});
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const {action, sheet: sheetName, data, id} = body;

    const sheet = getSheet(sheetName || 'productos');
    const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

    if (action === 'upsert') {
      const rows = sheet.getDataRange().getValues();
      // La hoja "config" usa 'key' como identificador único, no 'id'
      const keyCol = sheetName === 'config' ? headers.indexOf('key') : headers.indexOf('id');
      const keyVal = sheetName === 'config' ? data.key : data.id;
      const existing = rows.findIndex((r,i)=> i>0 && keyCol>=0 && r[keyCol]===keyVal);
      const row = headers.map(h => data[h] ?? '');
      if (existing > 0) sheet.getRange(existing+1,1,1,row.length).setValues([row]);
      else sheet.appendRow(row);
      return json({ok:true});
    }
    if (action === 'delete') {
      const rows = sheet.getDataRange().getValues();
      const idCol = headers.indexOf('id');
      const idx = rows.findIndex((r,i)=>i>0 && r[idCol]===id);
      if (idx > 0) sheet.deleteRow(idx+1);
      return json({ok:true});
    }
    // 'sync' = reemplazo completo de UNA hoja
    if (action === 'sync') {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow-1);
      (data||[]).forEach(d => sheet.appendRow(headers.map(h=>d[h]??'')));
      return json({ok:true});
    }
    // 'sync_all' = varias hojas a la vez, una por cada clave del objeto data
    if (action === 'sync_all') {
      Object.keys(data||{}).forEach(sheetKey => {
        const s = getSheet(sheetKey);
        const h = s.getRange(1,1,1,s.getLastColumn()).getValues()[0];
        const lastRow = s.getLastRow();
        if (lastRow > 1) s.deleteRows(2, lastRow-1);
        (data[sheetKey]||[]).forEach(d => s.appendRow(h.map(col=>d[col]??'')));
      });
      return json({ok:true});
    }
    return json({ok:false, error:'Unknown action'});
  } catch(err) {
    return json({ok:false, error:err.toString()});
  }
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    const headers = {
      productos:   ['id','barcode','nombre','cat','proveedor','precio','costo','stock','stockmin','vence','unidad','desc'],
      ventas:      ['id','folio','fecha','items','subtotal','descuento','total','recibido','cambio','metodo'],
      movimientos: ['id','fecha','tipo','monto','concepto'],
      cortes:      ['apertura','cierre','fondo','ingresos','egresos','saldoFinal'],
      config:      ['key','value']
    };
    if (headers[name]) s.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
  }
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
