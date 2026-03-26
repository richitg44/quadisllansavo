// ═══════════════════════════════════════════════════════════════════
//  NISSAN VO MANAGER — Google Apps Script v4
//  Hoja Clientes unificada: manuales + CRM4YOU
//  + Lectura de Sheet externo (stock del jefe)
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = "1AVH8yxQ5GRAa3XLEoRYU7xielB8g-dZF4TqMd2aa4_A";

const HOJAS = {
  clientes:    "Clientes",
  stock:       "Stock",
  stockPasado: "Stock Pasado",
  ubicaciones: "Ubicaciones",
};

// ── Columnas Clientes (compatible hacia atrás) ────────────────────
const COLS_CLIENTES = [
  "id","nombre","telefono","email","estado",
  "pMin","pMax","modelos","comb","cambio","kmMax","añoMin","etiq",
  "notas","notasCrm",
  "fuente","crmId","gestor","origen","vehiculoInteres",
  "historico","misNotas","descartado","followups","ultimaImport",
];

const COLS_STOCK = [
  "id","matricula","marca","modelo","version",
  "año","km","precio","color","combustible","cambio","etiqueta","foto","isNew",
  "precioContado","procedencia","estado","tipo","fecMat","impFinanciar"
];

// ── ENTRY POINT ───────────────────────────────────────────────────
function doGet(e) {
  let action, body = {};
  try {
    if (e.parameter && e.parameter.d) {
      const parsed = JSON.parse(decodeURIComponent(e.parameter.d));
      action = parsed.action; body = parsed;
    } else { action = e.parameter && e.parameter.action; }
  } catch(err) { action = e.parameter && e.parameter.action; }

  return respond(e, dispatch(action, body));
}

function doPost(e) {
  let action, body = {};
  try { const p = JSON.parse(e.postData.contents); action = p.action; body = p; }
  catch(err) { return json({error:"Error POST: "+err.message}); }
  return json(dispatch(action, body));
}

function dispatch(action, body) {
  try {
    if      (action === "getAll")             return getAll();
    else if (action === "saveCliente")        return saveCliente(body.data);
    else if (action === "deleteCliente")      return deleteRow(HOJAS.clientes, body.id);
    else if (action === "saveStock")          return saveStock(body.data);
    else if (action === "replaceStock")       return replaceStock(body.data);
    else if (action === "getPasado")          return getPasado();
    else if (action === "saveUbicacion")      return saveUbicacion(body.data);
    else if (action === "enviarCartel")       return enviarCartel(body.data);
    else if (action === "replaceLeadsCrm")    return replaceLeadsCrm(body.data);
    else if (action === "updateClienteMeta")  return updateClienteMeta(body.data);
    else if (action === "getExternalStock")   return getExternalStock(body.sheetId);
    else return { error: "Acción desconocida: " + action };
  } catch(err) { return { error: err.message }; }
}

function respond(e, result) {
  const cb = e.parameter && e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb+"("+JSON.stringify(result)+")")
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return json(result);
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LEER TODO ──────────────────────────────────────────────────────
function getAll() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return {
    clientes:    sheetToObjects(ss.getSheetByName(HOJAS.clientes),    COLS_CLIENTES),
    stock:       sheetToObjects(ss.getSheetByName(HOJAS.stock),       COLS_STOCK),
    ubicaciones: sheetToObjects(ss.getSheetByName(HOJAS.ubicaciones), ['id','zona']),
  };
}

function getPasado() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return { stock: sheetToObjects(ss.getSheetByName(HOJAS.stockPasado), COLS_STOCK) };
}

// ── CLIENTES MANUALES ─────────────────────────────────────────────
function saveCliente(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(HOJAS.clientes);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === data.id) {
      sheet.getRange(i+1, 1, 1, COLS_CLIENTES.length).setValues([clienteToRow(data)]);
      return { ok:true, action:"updated" };
    }
  }
  sheet.appendRow(clienteToRow(data));
  return { ok:true, action:"inserted" };
}

function clienteToRow(c) {
  return [
    c.id, c.nombre, c.telefono, c.email, c.estado,
    c.pref.pMin, c.pref.pMax,
    JSON.stringify(c.pref.modelos), JSON.stringify(c.pref.comb),
    c.pref.cambio, c.pref.kmMax, c.pref.añoMin,
    JSON.stringify(c.pref.etiq),
    c.pref.notas||"",
    "",
    c.fuente||"manual",
    c.crmId||"", c.gestor||"", c.origen||"", c.vehiculoInteres||"",
    JSON.stringify(c.historico||[]), JSON.stringify(c.misNotas||[]),
    c.descartado ? 1 : 0,
    JSON.stringify(c.followups||[]),
    c.ultimaImport||"",
  ];
}

// ── LEADS CRM4YOU → Clientes ───────────────────────────────────────
function replaceLeadsCrm(newLeads) {
  if (!Array.isArray(newLeads)) return { error:"Se esperaba un array" };

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ensureClientesSheet(ss);
  const rows  = sheet.getDataRange().getValues();
  const now   = new Date().toISOString().slice(0,10);

  const existingByCrmId = {};
  const existingByName  = {};
  for (let i = 1; i < rows.length; i++) {
    const fuente = String(rows[i][15]||'').trim();
    if (fuente !== 'crm4you') continue;
    const crmId = String(rows[i][16]||'').trim();
    const nombre = String(rows[i][1]||'').trim();
    if (crmId)  existingByCrmId[crmId]  = { row: rows[i], idx: i+1 };
    if (nombre) existingByName[nombre]   = { row: rows[i], idx: i+1 };
  }

  let inserted = 0, updated = 0;

  for (let li = 0; li < newLeads.length; li++) {
    const lead = newLeads[li];
    const key  = String(lead.crmId||'').trim();
    const nom  = String(lead.nombre||'').trim();

    const pref = inferirPreferencias(lead);
    const existing = (key && existingByCrmId[key]) || existingByName[nom];

    const primerMsgCrm = (lead.historico||[])
      .filter(function(n){ return n.texto && n.texto.trim(); })
      .slice(-1).map(function(n){ return n.texto; })[0] || '';

    const newRow = [
      existing ? existing.row[0] : 'CRM-' + (key || nom.replace(/\s/g, '-')),
      nom,
      lead.telefono||'',
      lead.email||'',
      lead.estado||'Nuevo',
      existing ? existing.row[5]  : '',
      existing ? existing.row[6]  : '',
      existing ? existing.row[7]  : '[]',
      existing ? existing.row[8]  : '[]',
      existing ? existing.row[9]  : 'Cualquiera',
      existing ? existing.row[10] : '',
      existing ? existing.row[11] : '',
      existing ? existing.row[12] : '[]',
      existing ? existing.row[13] : '',
      primerMsgCrm,
      'crm4you',
      key,
      lead.gestor||'',
      lead.origen||'',
      lead.vehiculo||'',
      JSON.stringify(lead.historico||[]),
      existing ? (existing.row[21]||'[]') : '[]',
      existing ? (existing.row[22]||0)    : 0,
      existing ? (existing.row[23]||'[]') : '[]',
      now,
    ];

    if (existing) {
      sheet.getRange(existing.idx, 1, 1, COLS_CLIENTES.length).setValues([newRow]);
      updated++;
    } else {
      sheet.appendRow(newRow);
      inserted++;
    }
  }

  sheet.getRange("A1").setNote("Última importación CRM4YOU: " + new Date().toLocaleString("es-ES"));
  return { ok:true, inserted, updated };
}

function inferirPreferencias(lead) {
  const veh     = (lead.vehiculo||'').toLowerCase();
  const notas   = (lead.historico||[]).map(n=>n.texto||'').join(' ').toLowerCase();
  const fuente  = veh + ' ' + notas;

  const modelosNissan = ['qashqai','juke','x-trail','xtrail','micra','leaf','ariya',
    'navara','nv200','nv300','nv400','kicks','townstar','primastar'];
  const modelos = modelosNissan.filter(m => fuente.includes(m))
    .map(m => m === 'xtrail' ? 'x-trail' : m.charAt(0).toUpperCase()+m.slice(1));

  const comb = [];
  if (fuente.match(/\béléctri/))  comb.push('Eléctrico');
  if (fuente.match(/\bhíbrido|hybrid/)) comb.push('Híbrido');
  if (fuente.match(/\bgasoil|diésel|diesel/)) comb.push('Diésel');
  if (fuente.match(/\bgasolina/)) comb.push('Gasolina');

  let pMax = 50000;
  const presMatch = fuente.match(/(\d[\d.]{3,})\s*€?|\b(\d{2})[\s.]?000\b|hasta\s+(\d+)k/i);
  if (presMatch) {
    const raw = (presMatch[1]||presMatch[2]+'000'||presMatch[3]+'000').replace(/\./g,'');
    const val = parseInt(raw);
    if (val > 3000 && val < 150000) pMax = val;
  }

  let añoMin = 2019;
  const añoMatch = fuente.match(/\b(202\d|201[89])\b/);
  if (añoMatch) añoMin = parseInt(añoMatch[1]);

  let kmMax = 100000;
  const kmMatch = fuente.match(/(\d+)\s*km/i);
  if (kmMatch) {
    const km = parseInt(kmMatch[1]);
    if (km > 1000 && km < 300000) kmMax = km;
  }

  return {
    pMin: Math.max(0, pMax - 10000),
    pMax,
    modelos,
    comb: comb.length ? comb : [],
    cambio: fuente.includes('automático') ? 'Automático' : 'Cualquiera',
    kmMax,
    añoMin,
    etiq: fuente.match(/\beco\b/) ? ['ECO'] : [],
  };
}

function updateClienteMeta(data) {
  if (!data) return {error:"data requerido"};
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(HOJAS.clientes);
  if (!sheet) return {error:"Hoja Clientes no existe"};

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const rowId    = String(rows[i][0]||'').trim();
    const rowCrmId = String(rows[i][16]||'').trim();
    const rowNom   = String(rows[i][1]||'').trim();
    const match    = (data.id && rowId === String(data.id)) ||
                     (data.crmId && rowCrmId === String(data.crmId)) ||
                     (data.nombre && rowNom === data.nombre);
    if (!match) continue;
    if (data.misNotas   !== undefined) sheet.getRange(i+1,22).setValue(JSON.stringify(data.misNotas));
    if (data.descartado !== undefined) sheet.getRange(i+1,23).setValue(data.descartado ? 1 : 0);
    if (data.followups  !== undefined) sheet.getRange(i+1,24).setValue(JSON.stringify(data.followups));
    if (data.pMin    !== undefined) sheet.getRange(i+1, 6).setValue(data.pMin);
    if (data.pMax    !== undefined) sheet.getRange(i+1, 7).setValue(data.pMax);
    if (data.modelos !== undefined) sheet.getRange(i+1, 8).setValue(JSON.stringify(data.modelos));
    if (data.comb    !== undefined) sheet.getRange(i+1, 9).setValue(JSON.stringify(data.comb));
    if (data.cambio  !== undefined) sheet.getRange(i+1,10).setValue(data.cambio);
    if (data.kmMax   !== undefined) sheet.getRange(i+1,11).setValue(data.kmMax);
    if (data.añoMin  !== undefined) sheet.getRange(i+1,12).setValue(data.añoMin);
    if (data.etiq    !== undefined) sheet.getRange(i+1,13).setValue(JSON.stringify(data.etiq));
    return {ok:true};
  }
  return {ok:false, error:"Cliente no encontrado"};
}

// ── STOCK ─────────────────────────────────────────────────────────
function saveStock(data) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(HOJAS.stock);
  const rows  = sheet.getDataRange().getValues();
  const mat   = (data.matricula||"").toUpperCase().replace(/[^A-Z0-9]/g,"");
  for (let i = 1; i < rows.length; i++) {
    const rowMat = (rows[i][1]||"").toUpperCase().replace(/[^A-Z0-9]/g,"");
    if (rowMat === mat) {
      sheet.getRange(i+1,1,1,COLS_STOCK.length).setValues([stockToRow(data)]);
      return {ok:true, action:"updated"};
    }
  }
  sheet.appendRow(stockToRow(data));
  return {ok:true, action:"inserted"};
}

function clearSheetData(sheet) {
  const lr = sheet.getLastRow(), lc = sheet.getLastColumn();
  if (lr > 1 && lc > 0) sheet.getRange(2,1,lr-1,lc).clearContent();
}

function replaceStock(rows) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheetA = ss.getSheetByName(HOJAS.stock);
  const sheetP = ss.getSheetByName(HOJAS.stockPasado);
  const ad = sheetA.getDataRange().getValues();
  clearSheetData(sheetP);
  if (ad.length > 1) sheetP.getRange(2,1,ad.length-1,COLS_STOCK.length).setValues(ad.slice(1));
  clearSheetData(sheetA);
  if (rows.length > 0) sheetA.getRange(2,1,rows.length,COLS_STOCK.length).setValues(rows.map(stockToRow));
  sheetP.getRange("A1").setNote("Importado el: "+new Date().toLocaleString("es-ES"));
  return {ok:true, replaced:rows.length};
}

function limpiarFecha(v) {
  if (!v) return '';
  var s = String(v).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})T/); if(m) return m[3]+'-'+m[2]+'-'+m[1];
  m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);      if(m) return m[3]+'-'+m[2]+'-'+m[1];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(m) return ('0'+m[1]).slice(-2)+'-'+('0'+m[2]).slice(-2)+'-'+m[3];
  return s;
}

function stockToRow(v) {
  return [
    v.id||"", v.matricula||"", v.marca||"Nissan", v.modelo||"", v.version||"",
    v.año||"", v.km||0, v.precio||0, v.color||"",
    v.combustible||"", v.cambio||"", v.etiqueta||"", v.foto||"🚗",
    v.isNew ? "SI" : "",
    v.precioContado||0, v.procedencia||"",
    v.estado||"", v.tipo||"", limpiarFecha(v.fecMat), v.impFinanciar||0
  ];
}

function deleteRow(sheetName, id) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const rows  = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === id) { sheet.deleteRow(i+1); return {ok:true}; }
  }
  return {ok:false, error:"No encontrado"};
}

function sheetToObjects(sheet, cols) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];
  return rows.slice(1).map(row => {
    const obj = {};
    cols.forEach((col, i) => { obj[col] = row[i] ?? ""; });
    return obj;
  });
}

function saveUbicacion(data) {
  if (!data || !data.id) return {error:"id requerido"};
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(HOJAS.ubicaciones);
  if (!sh) { sh = ss.insertSheet(HOJAS.ubicaciones); sh.appendRow(['id','zona']); }
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === data.id) {
      if (!data.zona) sh.deleteRow(i+1); else sh.getRange(i+1,2).setValue(data.zona);
      return {ok:true};
    }
  }
  if (data.zona) sh.appendRow([data.id, data.zona]);
  return {ok:true};
}

function enviarCartel(data) {
  if (!data || !data.email) return {error:"email requerido"};
  const pdfBlob = Utilities.newBlob(
    Utilities.base64Decode(data.pdfBase64), 'application/pdf',
    'cartel_'+data.matricula.replace(/\s/g,'_')+'.pdf'
  );
  GmailApp.sendEmail(data.email, 'Cartel precio: '+data.matricula+' · '+data.modelo,
    'Adjunto el cartel de precio para '+data.modelo+' ('+data.matricula+').',
    {name:'Nissan VO Manager', attachments:[pdfBlob]});
  return {ok:true};
}

// ── LEER STOCK DESDE SHEET EXTERNO (el del jefe) ─────────────────
function getExternalStock(sheetId) {
  try {
    var ss = SpreadsheetApp.openById(sheetId);
    var result = [];

    // Leer hoja "Stock VO"
    var wsVO = null;
    ss.getSheets().forEach(function(s) {
      if (s.getName().trim().toLowerCase() === 'stock vo') wsVO = s;
    });

    if (wsVO) {
      var dataVO = wsVO.getDataRange().getValues();
      var headersVO = dataVO[0].map(function(h) { return String(h).trim(); });

      var colMap = {};
      headersVO.forEach(function(h, i) {
        var hl = h.toLowerCase();
        if (hl === 'estado') colMap.estado = i;
        if (hl === 'matricula') colMap.matricula = i;
        if (hl === 'combustible') colMap.combustible = i;
        if (hl === 'dist medioamb') colMap.etiqueta = i;
        if (hl === 'pvp f') colMap.precio = i;
        if (hl === 'precio') colMap.precioBase = i;
        if (hl === 'kms') colMap.km = i;
        if (hl === 'fecha matricula') colMap.fecMat = i;
        if (hl.match(/a[ñn]o\s*fact/i)) colMap.año = i;
        if (hl === 'marca') colMap.marca = i;
        if (hl === 'modelo') colMap.modelo = i;
        if (hl.match(/versi[oó]n/i)) colMap.version = i;
        if (hl === 'color') colMap.color = i;
        if (hl.match(/transmi/i)) colMap.cambio = i;
        if (hl === 'procedencia') colMap.procedencia = i;
        if (hl.includes('financ')) colMap.impFinanciar = i;
      });

      for (var r = 1; r < dataVO.length; r++) {
        var row = dataVO[r];
        var estado = String(row[colMap.estado] || '').trim();
        if (!estado || estado === 'Vendido' || estado === 'IAC') continue;
        var mat = String(row[colMap.matricula] || '').toUpperCase().trim();
        if (!mat || mat.length < 6) continue;
        var precio = parseInt(String(row[colMap.precio] || '0').replace(/[^0-9]/g, '')) || 0;
        if (!precio) precio = parseInt(String(row[colMap.precioBase] || '0').replace(/[^0-9]/g, '')) || 0;
        if (!precio) continue;

        var fecMat = '';
        if (colMap.fecMat !== undefined && row[colMap.fecMat]) {
          var d = row[colMap.fecMat];
          if (d instanceof Date) {
            fecMat = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
          } else {
            fecMat = String(d);
          }
        }

        result.push({
          matricula: mat,
          marca: String(row[colMap.marca] || 'Nissan').trim(),
          modelo: String(row[colMap.modelo] || '').trim(),
          version: colMap.version !== undefined ? String(row[colMap.version] || '').trim() : '',
          año: colMap.año !== undefined ? parseInt(row[colMap.año]) || 0 : 0,
          km: colMap.km !== undefined ? parseInt(String(row[colMap.km] || '0').replace(/[^0-9]/g, '')) || 0 : 0,
          precio: precio,
          precioBase: colMap.precioBase !== undefined ? parseInt(String(row[colMap.precioBase] || '0').replace(/[^0-9]/g, '')) || 0 : precio,
          fecMat: fecMat,
          color: colMap.color !== undefined ? String(row[colMap.color] || '').trim() : '',
          combustible: colMap.combustible !== undefined ? String(row[colMap.combustible] || '').trim() : '',
          cambio: colMap.cambio !== undefined ? String(row[colMap.cambio] || '').trim() : '',
          etiqueta: colMap.etiqueta !== undefined ? String(row[colMap.etiqueta] || '').trim() : '',
          procedencia: colMap.procedencia !== undefined ? String(row[colMap.procedencia] || '').toUpperCase().trim() : 'VO',
          impFinanciar: colMap.impFinanciar !== undefined ? parseInt(String(row[colMap.impFinanciar] || '0').replace(/[^0-9]/g, '')) || 0 : 0,
          estado: estado,
          tipo: 'VO'
        });
      }
    }

    // Leer hoja "Stock DEMO"
    var wsDemo = null;
    ss.getSheets().forEach(function(s) {
      if (s.getName().trim().toLowerCase() === 'stock demo') wsDemo = s;
    });

    if (wsDemo) {
      var dataDemo = wsDemo.getDataRange().getValues();
      var headerRow = 0;
      for (var h = 0; h < Math.min(dataDemo.length, 5); h++) {
        if (dataDemo[h].some(function(c) { return String(c).trim().toLowerCase() === 'matricula'; })) {
          headerRow = h;
          break;
        }
      }
      var headersDemo = dataDemo[headerRow].map(function(h) { return String(h).trim(); });

      var colMapD = {};
      headersDemo.forEach(function(h, i) {
        var hl = h.toLowerCase();
        if (hl === 'estado') colMapD.estado = i;
        if (hl === 'matricula') colMapD.matricula = i;
        if (hl === 'combustible') colMapD.combustible = i;
        if (hl === 'pvp f') colMapD.precio = i;
        if (hl === 'precio') colMapD.precioBase = i;
        if (hl === 'kms') colMapD.km = i;
        if (hl === 'f. mat' || hl === 'fecha matricula') colMapD.fecMat = i;
        if (hl.match(/a[ñn]o\s*fact/i)) colMapD.año = i;
        if (hl === 'modelo') colMapD.modelo = i;
        if (hl.match(/versi[oó]n/i)) colMapD.version = i;
        if (hl === 'color') colMapD.color = i;
        if (hl.match(/transmi/i)) colMapD.cambio = i;
        if (hl === 'procedencia') colMapD.procedencia = i;
        if (hl.includes('financ')) colMapD.impFinanciar = i;
      });

      for (var r = headerRow + 1; r < dataDemo.length; r++) {
        var row = dataDemo[r];
        var estado = String(row[colMapD.estado] || '').trim();
        if (!estado || estado === 'Vendido' || estado === 'IAC') continue;
        var mat = String(row[colMapD.matricula] || '').toUpperCase().trim();
        if (!mat || mat.length < 6) continue;
        var precio = parseInt(String(row[colMapD.precio] || '0').replace(/[^0-9]/g, '')) || 0;
        if (!precio) precio = parseInt(String(row[colMapD.precioBase] || '0').replace(/[^0-9]/g, '')) || 0;
        if (!precio) continue;

        var fecMat = '';
        if (colMapD.fecMat !== undefined && row[colMapD.fecMat]) {
          var d = row[colMapD.fecMat];
          if (d instanceof Date) {
            fecMat = Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MM-yyyy');
          } else {
            fecMat = String(d);
          }
        }

        result.push({
          matricula: mat,
          marca: 'Nissan',
          modelo: colMapD.modelo !== undefined ? String(row[colMapD.modelo] || '').trim() : '',
          version: colMapD.version !== undefined ? String(row[colMapD.version] || '').trim() : '',
          año: colMapD.año !== undefined ? parseInt(row[colMapD.año]) || 0 : 0,
          km: colMapD.km !== undefined ? parseInt(String(row[colMapD.km] || '0').replace(/[^0-9]/g, '')) || 0 : 0,
          precio: precio,
          precioBase: colMapD.precioBase !== undefined ? parseInt(String(row[colMapD.precioBase] || '0').replace(/[^0-9]/g, '')) || 0 : precio,
          fecMat: fecMat,
          color: colMapD.color !== undefined ? String(row[colMapD.color] || '').trim() : '',
          combustible: colMapD.combustible !== undefined ? String(row[colMapD.combustible] || '').trim() : '',
          cambio: colMapD.cambio !== undefined ? String(row[colMapD.cambio] || '').trim() : '',
          etiqueta: '',
          procedencia: colMapD.procedencia !== undefined ? String(row[colMapD.procedencia] || '').toUpperCase().trim() : 'DEMO',
          impFinanciar: colMapD.impFinanciar !== undefined ? parseInt(String(row[colMapD.impFinanciar] || '0').replace(/[^0-9]/g, '')) || 0 : 0,
          estado: estado,
          tipo: 'Demo'
        });
      }
    }

    return { ok: true, stock: result };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ── ENSURE CLIENTES SHEET (migración automática) ──────────────────
function ensureClientesSheet(ss) {
  let sheet = ss.getSheetByName(HOJAS.clientes);
  if (!sheet) {
    sheet = ss.insertSheet(HOJAS.clientes);
    sheet.appendRow(COLS_CLIENTES);
    formatHeader(sheet, COLS_CLIENTES.length);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const currentCols = sheet.getLastColumn();
  if (currentCols < COLS_CLIENTES.length) {
    const header = sheet.getRange(1, 1, 1, currentCols).getValues()[0];
    const newCols = COLS_CLIENTES.slice(currentCols);
    sheet.getRange(1, currentCols+1, 1, newCols.length).setValues([newCols]);
    sheet.getRange(1, currentCols+1, 1, newCols.length)
         .setBackground("#C3002F").setFontColor("white").setFontWeight("bold");
    Logger.log("Migradas "+newCols.length+" columnas nuevas a Clientes: "+newCols.join(", "));
  }
  return sheet;
}

function formatHeader(sheet, numCols) {
  sheet.getRange(1,1,1,numCols).setBackground("#C3002F").setFontColor("white").setFontWeight("bold");
}

// ── SETUP ─────────────────────────────────────────────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  function ensureSheet(name, cols) {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    if (s.getLastRow() === 0) {
      s.appendRow(cols);
      formatHeader(s, cols.length);
    }
    s.setFrozenRows(1);
    return s;
  }

  ensureClientesSheet(ss);
  ensureSheet(HOJAS.stock,       COLS_STOCK);
  ensureSheet(HOJAS.stockPasado, COLS_STOCK);

  Logger.log("✅ Hojas listas. Clientes tiene "+COLS_CLIENTES.length+" columnas.");
}
