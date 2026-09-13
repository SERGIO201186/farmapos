// FarmaPos — Apps Script Backend
// Pega este código en script.google.com y despliega como Web App

const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// ── LICENCIA PRO: se valida en el Panel Maestro de Omnia (Control Central),
// NO en este script. Cada cliente ya no necesita su propia hoja "licencias":
// la app llama directo a la URL /exec del Panel Maestro (action=verify).
// Este backend solo guarda los datos del negocio (productos, ventas, etc.).

function doGet(e) {
  const action = e.parameter.action || '';

  // Verificación del webhook de WhatsApp (Meta Cloud API la llama por GET
  // una vez, al registrar la URL /exec, para confirmar que somos dueños de ella).
  if (e.parameter['hub.mode'] === 'subscribe') {
    const verifyToken = waProp('verifyToken');
    if (verifyToken && e.parameter['hub.verify_token'] === verifyToken) {
      return ContentService.createTextOutput(e.parameter['hub.challenge']);
    }
    return ContentService.createTextOutput('Verificación fallida');
  }

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

    // Mensaje entrante de WhatsApp (Meta Cloud API) — se distingue de las
    // peticiones normales de la app porque trae este "object" fijo.
    if (body.object === 'whatsapp_business_account') {
      manejarMensajeWhatsApp(body);
      return json({ok:true});
    }

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
      config:      ['key','value'],
      pacientes:   ['id','nombre','fechaNacimiento','sexo','telefono','direccion','alergias','cronicas','notas','creado'],
      recetas:     ['id','folio','fecha','pacienteId','pacienteNombre','pacienteEdad','pacienteSexo','diagnostico','notas','medicamentos','vigenciaDias','medicoNombre','medicoCedula','medicoCedulaEsp','medicoEspecialidad'],
      citas:       ['id','fecha','hora','paciente','telefono','motivo','estado','creado','origen']
    };
    if (headers[name]) s.getRange(1,1,1,headers[name].length).setValues([headers[name]]);
  }
  return s;
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── ASISTENTE DE WHATSAPP CON IA PARA AGENDAR CITAS ──────────────────────
// El bot recibe mensajes de WhatsApp, conversa con el paciente para reunir
// nombre / fecha / hora / motivo, y al completarlos guarda la cita en la
// hoja "citas". El doctor solo consulta esa hoja desde el panel de citas
// de la app — no participa en la conversación.
//
// INSTALACIÓN EN UN CLIENTE NUEVO (una sola vez, desde el editor de Apps
// Script, con las credenciales propias de ese cliente):
//
//   configurarAsistenteWhatsApp({
//     whatsappToken:   'EAAxxxxx...',           // token permanente, Meta for Developers > WhatsApp > API Setup
//     phoneNumberId:   '123456789012345',       // Phone Number ID de ese mismo panel
//     verifyToken:     'un-secreto-a-elegir',   // cualquier cadena; debe repetirse al registrar el webhook
//     anthropicApiKey: 'sk-ant-xxxxx...',       // console.anthropic.com
//     doctorTelefono:  '521XXXXXXXXXX',         // opcional, en formato E.164 sin "+": para avisarle cada cita nueva
//     nombreClinica:   'Nombre de la clínica'   // se usa en el prompt del asistente
//   });
//
// Después, en Meta for Developers > WhatsApp > Configuration, registra la
// URL /exec de este Web App como Webhook y usa el mismo verifyToken.

function configurarAsistenteWhatsApp(cfg) {
  const props = PropertiesService.getScriptProperties();
  Object.keys(cfg || {}).forEach(k => props.setProperty('WA_' + k, String(cfg[k])));
  return 'Configuración del asistente de WhatsApp guardada.';
}

function waProp(key, def) {
  return PropertiesService.getScriptProperties().getProperty('WA_' + key) || def || '';
}

function manejarMensajeWhatsApp(body) {
  try {
    const entry = (body.entry || [])[0];
    const change = entry && (entry.changes || [])[0];
    const value = change && change.value;
    const msg = value && (value.messages || [])[0];
    if (!msg || msg.type !== 'text') return; // ignora confirmaciones de entrega, multimedia, etc.

    const telefono = msg.from;
    const texto = msg.text.body;

    const respuesta = procesarMensajeIA(telefono, texto);
    enviarWhatsApp(telefono, respuesta.reply);

    if (respuesta.citaLista && respuesta.cita) {
      guardarCita(respuesta.cita, telefono);
      const doctorTel = waProp('doctorTelefono');
      if (doctorTel) {
        enviarWhatsApp(doctorTel,
          '📅 Nueva cita agendada\nPaciente: ' + respuesta.cita.paciente +
          '\nFecha: ' + respuesta.cita.fecha + ' ' + respuesta.cita.hora +
          '\nMotivo: ' + respuesta.cita.motivo);
      }
    }
  } catch (err) {
    // Nunca dejar que un error rompa el webhook: si no respondemos 200,
    // Meta reintenta la misma entrega varias veces.
    console.error('Error en manejarMensajeWhatsApp: ' + err);
  }
}

function procesarMensajeIA(telefono, texto) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'wa_chat_' + telefono;
  let historial = [];
  try { historial = JSON.parse(cache.get(cacheKey) || '[]'); } catch (e) { historial = []; }

  historial.push({ role: 'user', content: texto });

  const zona = Session.getScriptTimeZone() || 'America/Mexico_City';
  const hoy = Utilities.formatDate(new Date(), zona, "yyyy-MM-dd (EEEE)");
  const nombreClinica = waProp('nombreClinica', 'la clínica');

  const systemPrompt =
    'Eres el asistente de WhatsApp de ' + nombreClinica + '. Tu único trabajo es ' +
    'ayudar a agendar citas médicas por chat, de forma breve y amable.\n\n' +
    'Hoy es ' + hoy + '.\n\n' +
    'Necesitas reunir estos datos antes de agendar:\n' +
    '- Nombre completo del paciente\n' +
    '- Fecha deseada (conviértela siempre a formato AAAA-MM-DD)\n' +
    '- Hora deseada (formato HH:mm, 24 horas)\n' +
    '- Motivo breve de la consulta\n\n' +
    'Si falta algún dato, pregúntalo de forma natural, uno o dos a la vez, sin listarlos ' +
    'como un formulario. Si el paciente da una fecha/hora ambigua o ya pasada, pide que la ' +
    'aclare o corrija. No inventes datos ni marques la cita como lista sin tener los 4.\n\n' +
    'Responde SIEMPRE con un único objeto JSON válido, sin texto fuera del JSON, con esta forma exacta:\n' +
    '{"reply": "<mensaje para el paciente>", "citaLista": <true o false>, ' +
    '"cita": {"paciente": "...", "fecha": "AAAA-MM-DD", "hora": "HH:mm", "motivo": "..."}}\n\n' +
    '"cita" solo va completo cuando citaLista es true. Mientras falten datos, citaLista debe ser false.';

  const aiRespuesta = llamarClaude(systemPrompt, historial);

  historial.push({ role: 'assistant', content: JSON.stringify(aiRespuesta) });
  if (historial.length > 20) historial = historial.slice(-20); // recorta el contexto
  cache.put(cacheKey, JSON.stringify(historial), 1800); // 30 min de contexto por número

  return aiRespuesta;
}

function llamarClaude(systemPrompt, historial) {
  const apiKey = waProp('anthropicApiKey');
  if (!apiKey) {
    // Sin IA configurada, no se rompe el flujo: se avisa al paciente que
    // alguien del equipo lo contactará, en vez de fallar en silencio.
    return { reply: 'Gracias por tu mensaje. En breve una persona del equipo te contactará para agendar tu cita.', citaLista: false };
  }

  const payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 400,
    system: systemPrompt,
    messages: historial.map(m => ({ role: m.role, content: m.content }))
  };

  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  let data;
  try { data = JSON.parse(resp.getContentText()); } catch (e) { data = {}; }
  const texto = (data.content && data.content[0] && data.content[0].text) || '';

  try {
    return JSON.parse(texto);
  } catch (e) {
    // La IA no devolvió JSON puro (raro, pero posible): se manda su texto tal cual.
    return { reply: texto || 'Disculpa, ¿me repites tu mensaje?', citaLista: false };
  }
}

function enviarWhatsApp(telefono, texto) {
  const token = waProp('whatsappToken');
  const phoneId = waProp('phoneNumberId');
  if (!token || !phoneId) return;

  UrlFetchApp.fetch('https://graph.facebook.com/v20.0/' + phoneId + '/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({
      messaging_product: 'whatsapp',
      to: telefono,
      type: 'text',
      text: { body: texto }
    }),
    muteHttpExceptions: true
  });
}

function guardarCita(cita, telefono) {
  const sheet = getSheet('citas');
  sheet.appendRow([
    Utilities.getUuid(),
    cita.fecha,
    cita.hora,
    cita.paciente,
    telefono,
    cita.motivo,
    'pendiente',
    new Date().toISOString(),
    'whatsapp'
  ]);
}
