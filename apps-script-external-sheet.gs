/**
 * INSTRUCCIONES:
 * Copia SOLO la función getExternalStock() de abajo y pégala
 * en tu Google Apps Script existente (nissan-vo-script.gs).
 *
 * Luego haz una NUEVA IMPLEMENTACIÓN del script para que los cambios se apliquen.
 * (Implementar > Nueva implementación > Aplicación web)
 *
 * IMPORTANTE: El Apps Script debe ejecutarse con TU cuenta de Google,
 * y tu cuenta debe tener acceso de lectura al Sheet de tu jefe.
 */

// ======= AÑADIR ESTO A TU APPS SCRIPT EXISTENTE =======

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

      // Mapear columnas
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
      // La hoja DEMO a veces tiene 3 filas de cabecera, buscar la fila con "Matricula"
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


// ======= MODIFICAR TU doPost EXISTENTE =======
// Añade este case dentro del switch/if de acciones:
//
//   if (action === 'getExternalStock') {
//     var sheetId = payload.sheetId;
//     return send(getExternalStock(sheetId));
//   }
//
// Donde "send" es tu función que devuelve ContentService.createTextOutput(JSON.stringify(data))
