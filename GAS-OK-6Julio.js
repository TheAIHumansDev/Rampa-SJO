// ============================================================
//  TURNAROUND CHECK — Apps Script v2.0
//  NUEVAS FUNCIONES:
//    • Fotos en Google Drive (carpeta por vuelo)
//    • Hoja FACTURABLE  — servicios adicionales por vuelo
//    • Hoja INCIDENTES  — incidentes operacionales
//    • Hoja SEGURIDAD   — reportes SMS / Seguridad Operacional
//    • Hoja CANCELADOS  — vuelos cancelados con facturación
//    • Hoja CHECKPOINTS_DETALLE — fila por cada checkpoint
//    • Hojas TURN / RON_Llegada / RON_Salida (existentes, mejoradas)
//
//  INSTRUCCIONES DE DEPLOY:
//  1. Apps Script → selecciona todo → pega este archivo
//  2. Ajusta SPREADSHEET_ID y DRIVE_FOLDER_ID (raíz de fotos)
//  3. Deploy → New deployment → Web App
//       Execute as: Me
//       Who has access: Anyone (o Anyone with Google account)
//  4. Copia la URL del deployment → ponla en SCRIPT_URL del HTML
// ============================================================

// ── CONFIGURACIÓN ────────────────────────────────────────────
var SPREADSHEET_ID  = '1GgS964BGGVOD_4jgftFpAMcwZ9O4N8w2csyxzNG2pUc';
var DRIVE_FOLDER_ID = '1S0UV94TZRwuqsY0yCJZgTNBCHlrU4-O-'; // Carpeta raíz "Turnaround Fotos"

// ── CABECERAS POR HOJA ───────────────────────────────────────

var HEADERS_TURN = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo Origen','Vuelo Destino',
  'Vuelo','Tipo','Fecha/Hora','Inspector','Rol','Posición','Estatus',
  'Aeropuerto','Tiempo Total (min)','Items Completados','Items Totales',
  'Tasa Completitud (%)',
  // checkpoints
  'Briefing','Inicio FOD','Beacon OFF',
  '360° Inicio','Planta Eléctrica Conecta','Planta Eléctrica Desconecta',
  'Aire Acondicionado','Banda','Escalera',
  'Inicia Descarga Carga','Finaliza Descarga Carga',
  'Inicia Carga Carga','Finaliza Carga Carga',
  'Waste Residuales','Water Potable',
  'Inicia Descarga Equipaje','Finaliza Descarga Equipaje',
  'Inicia Carga Equipaje','Finaliza Carga Equipaje',
  'Limpieza Sube','Limpieza Baja',
  'Limpieza Baño Delantero','Limpieza Baño Trasero',
  'Limpieza Pasillos','Limpieza Cocina Delantera','Limpieza Cocina Trasera',
  'Limpieza Basureros','Limpieza Bolsas Asiento','Limpieza Compartimentos',
  '1er FOD','2do FOD','3er FOD',
  'Último Viaje Carreta','Towbar',
  'Bin Delantero','Bin Trasero',
  '360° Fin','Push Back'
];

var HEADERS_RON_LLEGADA = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo Origen','Vuelo Destino',
  'Vuelo','Tipo','Fecha/Hora','Inspector','Rol','Posición','Estatus',
  'Aeropuerto','Tiempo Total (min)','Items Completados','Items Totales',
  'Tasa Completitud (%)',
  'Briefing','FOD','Beacon OFF',
  '360° Inicio','Planta Eléctrica Conecta','Planta Eléctrica Desconecta',
  'Aire Acondicionado','Banda','Escalera',
  'Inicia Descarga Equipaje','Finaliza Descarga Equipaje',
  'Waste Residuales','Water Potable',
  'Limpieza Sube','Limpieza Baja',
  'Limpieza Baño Delantero','Limpieza Baño Trasero',
  'Limpieza Pasillos','Limpieza Cocina Delantera','Limpieza Cocina Trasera',
  'Limpieza Basureros','Limpieza Bolsas Asiento','Limpieza Compartimentos',
  '1er FOD','2do FOD',
  'Último Viaje Carreta',
  'Bin Delantero','Bin Trasero',
  '360° Fin','Towbar'
];

var HEADERS_RON_SALIDA = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo Origen','Vuelo Destino',
  'Vuelo','Tipo','Fecha/Hora','Inspector','Rol','Posición','Estatus',
  'Aeropuerto','Tiempo Total (min)','Items Completados','Items Totales',
  'Tasa Completitud (%)',
  'Briefing','FOD','Avión en Posición',
  'Planta Eléctrica Conecta','Planta Eléctrica Desconecta',
  'Aire Acondicionado','Banda','Escalera',
  'Inicia Carga Equipaje','Finaliza Carga Equipaje',
  'Inicia Carga Carga','Finaliza Carga Carga',
  'Waste Residuales','Water Potable',
  '1er FOD','2do FOD',
  'Towbar',
  'Bin Delantero','Bin Trasero',
  '360° Fin','Push Back'
];

var HEADERS_FACTURABLE = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo','Tipo','Fecha/Hora',
  'Inspector','Posición',
  'Servicio','Icono','Detalle','Cantidad','Unidad','Tarifa','Monto Total',
  'Proveedor','Notas','Timestamp Registro'
];

var HEADERS_INCIDENTES = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo','Tipo','Fecha/Hora',
  'Inspector','Posición',
  'Tipo Incidente','Severidad','Descripción','Timestamp'
];

var HEADERS_SEGURIDAD = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo','Tipo','Fecha/Hora',
  'Inspector','Posición',
  'Categoría','Nivel Riesgo','Personas Involucradas',
  'Descripción','Acción Tomada','Requiere Seguimiento','Timestamp'
];

var HEADERS_CANCELADOS = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo','Tipo','Fecha/Hora',
  'Inspector','Posición',
  'Motivo Cancelación','Hora Notificación','Personal Disponible',
  'Equipo Desplegado','Tiempo en Posición',
  'Tarifa Disponibilidad','Horas Personal','Tarifa Hora','Cargos Equipo',
  'Total Facturar','Observaciones','Timestamp'
];

var HEADERS_CHECKPOINTS = [
  'Flight Key','Código Aerolínea','Aerolínea','Vuelo','Tipo','Fecha/Hora',
  'Inspector','Posición',
  'Step #','Checkpoint','Emoji','Sub-Item',
  'Timestamp Registro','Foto URL','Drive File ID'
];

// ── MAPA: item label → columna header ────────────────────────
// Permite que saveInspection ubique la columna correcta en cada hoja
var ITEM_TO_HEADER = {
  // Briefing
  'Briefing':                    'Briefing',
  // FOD
  'Inicio FOD':                  'Inicio FOD',
  'FOD':                         'FOD',
  '1er FOD':                     '1er FOD',
  '2do FOD':                     '2do FOD',
  '3er FOD':                     '3er FOD',
  // Beacon
  'Beacon OFF':                  'Beacon OFF',
  // 360
  '360° Aeronave — Inicio':      '360° Inicio',
  '360° Aeronave — Fin':         '360° Fin',
  // Planta
  'Planta Eléctrica — conecta':  'Planta Eléctrica Conecta',
  'Planta Eléctrica — desconecta':'Planta Eléctrica Desconecta',
  'Planta Eléctrica — Desconecta (auto)':'Planta Eléctrica Desconecta',
  // Otros servicios
  'Aire Acondicionado':          'Aire Acondicionado',
  'Banda':                       'Banda',
  'Escalera':                    'Escalera',
  // Carga
  'Inicia Descarga Carga':       'Inicia Descarga Carga',
  'Finaliza Descarga Carga':     'Finaliza Descarga Carga',
  'Inicia Carga Carga':          'Inicia Carga Carga',
  'Finaliza Carga Carga':        'Finaliza Carga Carga',
  // Waste & Water
  '🗑️ Waste (Aguas Residuales)': 'Waste Residuales',
  '💧 Water (Agua Potable)':     'Water Potable',
  // Equipaje
  'Inicia Descarga Equipaje':    'Inicia Descarga Equipaje',
  'Finaliza Descarga Equipaje':  'Finaliza Descarga Equipaje',
  'Inicia Carga Equipaje':       'Inicia Carga Equipaje',
  'Finaliza Carga Equipaje':     'Finaliza Carga Equipaje',
  // Limpieza
  '🧹 Sube Limpieza':            'Limpieza Sube',
  '🧹 Baja Limpieza':            'Limpieza Baja',
  '🚻 Baño Delantero':           'Limpieza Baño Delantero',
  '🚻 Baño Trasero':             'Limpieza Baño Trasero',
  '🛤️ Pasillos':                 'Limpieza Pasillos',
  '🍽️ Cocina Delantera':        'Limpieza Cocina Delantera',
  '🍽️ Cocina Trasera':          'Limpieza Cocina Trasera',
  '🗑️ Basureros':               'Limpieza Basureros',
  '📄 Bolsas de Asiento':        'Limpieza Bolsas Asiento',
  '🗄️ Compartimentos Superiores':'Limpieza Compartimentos',
  // Otros
  'Último Viaje Carreta':        'Último Viaje Carreta',
  'Towbar':                      'Towbar',
  '🚪 Bin Delantero':            'Bin Delantero',
  '🚪 Bin Trasero':              'Bin Trasero',
  'Push Back':                   'Push Back',
  'Avión en Posición':           'Avión en Posición',
  // RON sub-checkpoints
  'Inicia Descarga':             'Inicia Descarga Equipaje',
  'Finaliza Descarga':           'Finaliza Descarga Equipaje',
  'Inicia Carga':                'Inicia Carga Carga',
  'Finaliza Carga':              'Finaliza Carga Carga',
  // Reporte Seguridad (no va a col de checkpoint sino a hoja Seguridad)
};


// ============================================================
//  doGet
// ============================================================
function doGet(e) {
  var params   = e.parameter;
  var callback = params.callback || '';
  var action   = params.action   || '';
  var result;

  try {
    if (action === 'getHistory') {
      result = getHistory(params.inspector);
    } else if (action === 'debugInspector') {
      result = debugInspector(params.inspector);
    } else if (params.username && params.password) {
      result = authenticateUser(params.username, params.password);
    } else {
      result = { success: false, message: 'Acción no reconocida' };
    }
  } catch (err) {
    result = { success: false, message: 'Error: ' + err.message };
  }

  var json = JSON.stringify(result);
  return callback
    ? ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
//  doPost — enrutador de acciones POST
// ============================================================
function doPost(e) {
  try {
    var data   = JSON.parse(e.postData.contents);
    var action = data.action || '';
    var result;

    if      (action === 'saveInspection')  result = saveInspection(data);
    else if (action === 'updateStatus')    result = updateStatus(data);
    else if (action === 'saveFacturable')  result = saveFacturable(data);
    else if (action === 'saveIncidente')   result = saveIncidente(data);
    else if (action === 'saveSeguridad')   result = saveSeguridad(data);
    else if (action === 'saveCancelado')   result = saveCancelado(data);
    else result = { success: false, message: 'Acción no reconocida: ' + action };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log('doPost error: ' + err.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, message: 'Error POST: ' + err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
//  saveInspection — guarda checkpoint + foto en Drive
// ============================================================
function saveInspection(data) {
  try {
    var ss         = SpreadsheetApp.openById(SPREADSHEET_ID);
    var flightType = data.flightType || 'TURN';
    var sheetName  = getSheetName(flightType);
    var headers    = getHeadersForType(flightType);
    var sheet      = getOrCreateSheet(ss, sheetName, headers);

    var flightKey = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp);
    var targetRow = findOrCreateFlightRow(sheet, flightKey, data, headers);

    // ── Determinar columna del checkpoint ────────────────
    var itemLabel   = String(data.item || '').trim();
    var headerLabel = ITEM_TO_HEADER[itemLabel] || itemLabel;
    var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var itemColIdx   = findColIdx(sheetHeaders, headerLabel);

    var photoUrl    = '';
    var driveFileId = '';

    // ── Subir foto a Drive si viene base64 ───────────────
    if (data.photoBase64 && data.photoBase64.length > 100) {
      try {
        var driveResult = savePhotoToDrive(
          data.photoBase64,
          flightKey,
          itemLabel,
          data.timestamp
        );
        photoUrl    = driveResult.url;
        driveFileId = driveResult.fileId;
      } catch (photoErr) {
        Logger.log('Error foto Drive: ' + photoErr.toString());
      }
    }

    // ── Escribir timestamp en la columna del checkpoint ──
    if (itemColIdx >= 0) {
      var cellVal = data.timestamp || new Date().toISOString();
      // Si ya hay URL de Drive, concatenar al timestamp
      if (photoUrl) {
        cellVal = cellVal + ' | ' + photoUrl;
      }
      sheet.getRange(targetRow, itemColIdx + 1).setValue(cellVal);
    }

    // ── Guardar fila detallada en CHECKPOINTS_DETALLE ───
    saveCheckpointDetail(ss, data, flightKey, itemLabel, photoUrl, driveFileId);

    // ── Actualizar métricas ──────────────────────────────
    updateMetrics(sheet, targetRow, sheetHeaders);

    // ── Guardar seguridad si viene en el payload ─────────
    if (data.seguridad) {
      var segPayload = {
        airline:          data.airline,
        flightType:       data.flightType,
        flightNumber:     data.flightNumber,
        flightTimestamp:  data.flightTimestamp,
        inspector:        data.inspector,
        posicion:         data.posicion,
        seguridad:        data.seguridad
      };
      saveSeguridad(segPayload);
    }

    return { success: true, message: 'Guardado', photoUrl: photoUrl, driveFileId: driveFileId };

  } catch (err) {
    Logger.log('saveInspection error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  savePhotoToDrive — convierte base64 → archivo JPEG en Drive
// ============================================================
function savePhotoToDrive(base64Data, flightKey, checkpointLabel, timestamp) {
  // Limpiar prefijo data:image/jpeg;base64,
  var cleanBase64 = base64Data.replace(/^data:image\/\w+;base64,/, '');
  var blob        = Utilities.newBlob(
    Utilities.base64Decode(cleanBase64),
    'image/jpeg',
    buildPhotoFileName(flightKey, checkpointLabel, timestamp)
  );

  // Carpeta raíz
  var rootFolder;
  try {
    rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  } catch (e) {
    // Si no existe la carpeta raíz, crearla en My Drive
    rootFolder = DriveApp.createFolder('Turnaround Fotos');
  }

  // Subcarpeta por vuelo (Flight Key)
  var flightFolder = getOrCreateSubfolder(rootFolder, flightKey);

  // Crear archivo
  var file    = flightFolder.createFile(blob);
  var fileId  = file.getId();
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var fileUrl = 'https://drive.google.com/file/d/' + fileId + '/view';

  return { fileId: fileId, url: fileUrl };
}

function buildPhotoFileName(flightKey, label, timestamp) {
  var ts   = timestamp ? new Date(timestamp) : new Date();
  var time = Utilities.formatDate(ts, Session.getScriptTimeZone(), 'HHmmss');
  var clean = String(label).replace(/[^a-zA-Z0-9_\-]/g, '_').substring(0, 40);
  return flightKey + '_' + clean + '_' + time + '.jpg';
}

function getOrCreateSubfolder(parent, name) {
  var cleanName = String(name).replace(/[\/\\:*?"<>|]/g, '-');
  var it = parent.getFoldersByName(cleanName);
  if (it.hasNext()) return it.next();
  return parent.createFolder(cleanName);
}


// ============================================================
//  saveCheckpointDetail — hoja CHECKPOINTS_DETALLE
// ============================================================
function saveCheckpointDetail(ss, data, flightKey, itemLabel, photoUrl, driveFileId) {
  try {
    var sheet = getOrCreateSheet(ss, 'CHECKPOINTS_DETALLE', HEADERS_CHECKPOINTS);
    var row = [
      flightKey,
      data.airline          || '',
      getAirlineName(data.airline),
      data.flightNumber     || '',
      data.flightType       || '',
      data.flightTimestamp  || new Date().toISOString(),
      data.inspector        || '',
      data.posicion         || '',
      data.step             || '',
      itemLabel,
      data.emoji            || '',
      data.sub              || '',
      data.timestamp        || new Date().toISOString(),
      photoUrl              || '',
      driveFileId           || ''
    ];
    sheet.appendRow(row);
  } catch (e) {
    Logger.log('saveCheckpointDetail error: ' + e.toString());
  }
}


// ============================================================
//  saveFacturable — hoja FACTURABLE
// ============================================================
function saveFacturable(data) {
  try {
    var ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet   = getOrCreateSheet(ss, 'FACTURABLE', HEADERS_FACTURABLE);
    var flightKey = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp);
    var items   = data.items || [];

    // Borrar entradas previas de este flight key para evitar duplicados
    deleteRowsByFlightKey(sheet, flightKey);

    items.forEach(function(item) {
      var row = [
        flightKey,
        data.airline         || '',
        getAirlineName(data.airline),
        data.flightNumber    || '',
        data.flightType      || '',
        data.flightTimestamp || new Date().toISOString(),
        data.inspector       || '',
        data.posicion        || '',
        item.name            || '',
        item.icon            || '',
        item.detail          || '',
        item.cantidad        || '',
        item.unidad          || '',
        item.tarifa          || '',
        item.amount          || 0,
        item.proveedor       || '',
        item.notes           || '',
        item.timestamp       || new Date().toISOString()
      ];
      sheet.appendRow(row);
    });

    applySheetFormatting(sheet, 'FACTURABLE');
    return { success: true, message: 'Servicios facturables guardados: ' + items.length };

  } catch (err) {
    Logger.log('saveFacturable error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  saveIncidente — hoja INCIDENTES
// ============================================================
function saveIncidente(data) {
  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet     = getOrCreateSheet(ss, 'INCIDENTES', HEADERS_INCIDENTES);
    var flightKey = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp);
    var inc       = data.incident || {};

    var row = [
      flightKey,
      data.airline         || '',
      getAirlineName(data.airline),
      data.flightNumber    || '',
      data.flightType      || '',
      data.flightTimestamp || new Date().toISOString(),
      data.inspector       || '',
      data.posicion        || '',
      inc.type             || '',
      inc.severity         || '',
      inc.description      || '',
      inc.timestamp        || new Date().toISOString()
    ];
    sheet.appendRow(row);

    applySheetFormatting(sheet, 'INCIDENTES');
    return { success: true, message: 'Incidente guardado' };

  } catch (err) {
    Logger.log('saveIncidente error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  saveSeguridad — hoja SEGURIDAD
// ============================================================
function saveSeguridad(data) {
  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet     = getOrCreateSheet(ss, 'SEGURIDAD', HEADERS_SEGURIDAD);
    var flightKey = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp);
    var seg       = data.seguridad || {};

    var row = [
      flightKey,
      data.airline         || '',
      getAirlineName(data.airline),
      data.flightNumber    || '',
      data.flightType      || '',
      data.flightTimestamp || new Date().toISOString(),
      data.inspector       || '',
      data.posicion        || '',
      seg.categoria        || '',
      seg.riesgo           || '',
      seg.personas         || '',
      seg.descripcion      || '',
      seg.accion           || '',
      seg.seguimiento ? 'Sí' : 'No',
      seg.timestamp        || new Date().toISOString()
    ];
    sheet.appendRow(row);

    applySheetFormatting(sheet, 'SEGURIDAD');
    return { success: true, message: 'Reporte de seguridad guardado' };

  } catch (err) {
    Logger.log('saveSeguridad error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  saveCancelado — hoja CANCELADOS
// ============================================================
function saveCancelado(data) {
  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet     = getOrCreateSheet(ss, 'CANCELADOS', HEADERS_CANCELADOS);
    var flightKey = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp);
    var cd        = data.cancelacion || {};
    var fac       = cd.facturacion || {};

    var row = [
      flightKey,
      data.airline         || '',
      getAirlineName(data.airline),
      data.flightNumber    || '',
      data.flightType      || '',
      data.flightTimestamp || new Date().toISOString(),
      data.inspector       || '',
      data.posicion        || '',
      cd.motivo            || '',
      cd.horaNot           || '',
      cd.personal          || '',
      cd.equipo            || '',
      cd.tiempoPos         || '',
      fac.disponibilidad   || 0,
      fac.horasPersonal    || 0,
      fac.tarifaHora       || 0,
      fac.cargosEquipo     || 0,
      fac.total            || 0,
      cd.observaciones     || '',
      new Date().toISOString()
    ];
    sheet.appendRow(row);

    // También guardar servicios adicionales si existen
    if (data.billing && data.billing.length > 0) {
      saveFacturable({
        airline:          data.airline,
        flightNumber:     data.flightNumber,
        flightType:       data.flightType,
        flightTimestamp:  data.flightTimestamp,
        inspector:        data.inspector,
        posicion:         data.posicion,
        items:            data.billing
      });
    }

    applySheetFormatting(sheet, 'CANCELADOS');
    return { success: true, message: 'Vuelo cancelado guardado' };

  } catch (err) {
    Logger.log('saveCancelado error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  updateStatus — cierra inspección (Completado / Cancelado)
// ============================================================
function updateStatus(data) {
  try {
    var ss        = SpreadsheetApp.openById(SPREADSHEET_ID);
    var flightType = data.flightType || 'TURN';
    var sheetName  = getSheetName(flightType);
    var sheet      = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Hoja no encontrada: ' + sheetName };

    var flightKey    = buildFlightKey(data.airline, data.flightNumber, data.flightTimestamp || data.timestamp);
    var sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var allData      = sheet.getDataRange().getValues();
    var colFKey      = findColIdx(sheetHeaders, 'Flight Key');
    var colEstatus   = findColIdx(sheetHeaders, 'Estatus');

    for (var i = 1; i < allData.length; i++) {
      if (String(allData[i][colFKey]).trim() === flightKey) {
        if (colEstatus >= 0) {
          sheet.getRange(i + 1, colEstatus + 1).setValue(data.status || 'Completado');
        }
        // Guardar observaciones finales si vienen
        var colObs = findColIdx(sheetHeaders, 'Observaciones Finales');
        if (colObs >= 0 && data.observaciones) {
          sheet.getRange(i + 1, colObs + 1).setValue(data.observaciones);
        }
        break;
      }
    }

    // Guardar incidente si viene
    if (data.incident) {
      saveIncidente({
        airline:         data.airline,
        flightNumber:    data.flightNumber,
        flightType:      data.flightType,
        flightTimestamp: data.flightTimestamp,
        inspector:       data.inspector,
        posicion:        data.posicion,
        incident:        data.incident
      });
    }

    // Guardar cancelación si aplica
    if (data.cancelacion) {
      saveCancelado({
        airline:         data.airline,
        flightNumber:    data.flightNumber,
        flightType:      data.flightType,
        flightTimestamp: data.flightTimestamp,
        inspector:       data.inspector,
        posicion:        data.posicion,
        cancelacion:     data.cancelacion,
        billing:         data.billing || []
      });
    }

    return { success: true, message: 'Estado actualizado a ' + (data.status || 'Completado') };

  } catch (err) {
    Logger.log('updateStatus error: ' + err.toString());
    return { success: false, message: err.message };
  }
}


// ============================================================
//  getHistory — historial por inspector (todas las hojas vuelo)
// ============================================================
function getHistory(inspector) {
  if (!inspector) return { success: false, message: 'Inspector requerido' };

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var SHEET_DEFS = [
    { name: 'TURN',         type: 'TURN'          },
    { name: 'RON_Llegada',  type: 'RON-ARRIVAL'   },
    { name: 'RON_Salida',   type: 'RON-DEPARTURE' },
  ];

  var reports = [];

  SHEET_DEFS.forEach(function(def) {
    var sheet = ss.getSheetByName(def.name);
    if (!sheet || sheet.getLastRow() < 2) return;

    var lastCol = sheet.getLastColumn();
    var data    = sheet.getRange(1, 1, sheet.getLastRow(), lastCol).getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });

    function ci(name) { return findColIdx(headers, name); }

    data.slice(1).forEach(function(row) {
      var rowInsp = ci('Inspector') >= 0 ? String(row[ci('Inspector')] || '').trim() : '';
      if (!rowInsp || rowInsp.toLowerCase() !== inspector.toLowerCase()) return;

      var flightKey   = ci('Flight Key')         >= 0 ? String(row[ci('Flight Key')]   || '').trim() : '';
      var codigo      = ci('Código Aerolínea')   >= 0 ? String(row[ci('Código Aerolínea')] || '').trim() : '';
      var aerolinea   = ci('Aerolínea')          >= 0 ? String(row[ci('Aerolínea')]    || '').trim() : '';
      var vuelo       = ci('Vuelo')              >= 0 ? String(row[ci('Vuelo')]         || '').trim() : '';
      var tipo        = ci('Tipo')               >= 0 ? String(row[ci('Tipo')]          || '').trim() : def.type;
      var fechaRaw    = ci('Fecha/Hora')         >= 0 ? row[ci('Fecha/Hora')]           : '';
      var estatusRaw  = ci('Estatus')            >= 0 ? String(row[ci('Estatus')]       || '').trim() : '';
      var posicion    = ci('Posición')           >= 0 ? String(row[ci('Posición')]      || '').trim() : '';
      var itemsComp   = ci('Items Completados')  >= 0 ? (Number(row[ci('Items Completados')]) || 0) : 0;
      var itemsTotal  = ci('Items Totales')      >= 0 ? (Number(row[ci('Items Totales')])      || 0) : 23;
      var tasa        = ci('Tasa Completitud (%)') >= 0 ? (Number(row[ci('Tasa Completitud (%)')]) || 0) : 0;

      // Fecha ISO
      var fechaISO = '';
      if (fechaRaw instanceof Date) {
        fechaISO = fechaRaw.toISOString();
      } else if (fechaRaw) {
        var p = new Date(String(fechaRaw).trim());
        fechaISO = isNaN(p.getTime()) ? String(fechaRaw).trim() : p.toISOString();
      }

      var estatus = estatusRaw || (tasa >= 100 ? 'Completado' : 'En Proceso');
      var progress = estatus === 'Completado' ? 100
        : Math.min(100, itemsTotal > 0 ? Math.round((itemsComp / itemsTotal) * 100) : 0);

      // Checkpoints completados (columnas después de metadatos)
      var META = ['flight key','código aerolínea','aerolinea','aerolínea','vuelo origen',
        'vuelo destino','vuelo','tipo','fecha/hora','inspector','rol','posición',
        'aeropuerto','tiempo total (min)','items completados','items totales',
        'tasa completitud (%)','estatus','observaciones finales'];
      var checkpoints = [];
      headers.forEach(function(h, idx) {
        var normH = norm(h);
        if (!normH) return;
        if (META.some(function(m) { return norm(m) === normH; })) return;
        var val = row[idx];
        if (!val || val === '' || val === false || val === 0) return;
        var ts = val instanceof Date ? val.toISOString() : String(val).split(' | ')[0].trim();
        var photoUrl = '';
        if (String(val).indexOf('drive.google.com') >= 0) {
          photoUrl = String(val).split(' | ')[1] || '';
        }
        checkpoints.push({ label: h, timestamp: ts, photoUrl: photoUrl });
      });

      reports.push({
        flightKey:       flightKey,
        airline:         codigo,
        airlineName:     aerolinea,
        flightNumber:    vuelo,
        flightType:      tipo,
        flightTimestamp: fechaISO,
        inspector:       rowInsp,
        role:            ci('Rol') >= 0 ? String(row[ci('Rol')] || '').trim() : '',
        status:          estatus,
        posicion:        posicion,
        checkpoints:     checkpoints,
        photoCount:      checkpoints.filter(function(c) { return c.photoUrl; }).length,
        itemsCompletados:itemsComp,
        itemsTotales:    itemsTotal,
        progress:        progress,
        _sheet:          def.name,
      });
    });
  });

  reports.sort(function(a, b) {
    return (b.flightTimestamp ? new Date(b.flightTimestamp) : 0) -
           (a.flightTimestamp ? new Date(a.flightTimestamp) : 0);
  });

  return { success: true, reports: reports, total: reports.length };
}


// ============================================================
//  authenticateUser
// ============================================================
function authenticateUser(username, password) {
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Usuarios') || ss.getSheetByName('Users') || ss.getSheetByName('usuarios');
    if (!sheet) return { success: false, message: 'Hoja de usuarios no encontrada' };

    var data    = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: 'No hay usuarios registrados' };

    var headers = data[0].map(function(h) { return norm(String(h)); });
    function colIdx(names) {
      for (var i = 0; i < names.length; i++) {
        var idx = headers.indexOf(norm(names[i]));
        if (idx >= 0) return idx;
      }
      return -1;
    }

    var cUser = colIdx(['usuario','user','username','email','correo']);
    var cPass = colIdx(['contraseña','password','contrasena','pass','clave']);
    var cRol  = colIdx(['rol','role','cargo','perfil']);

    if (cUser < 0 || cPass < 0) return { success: false, message: 'Estructura de hoja de usuarios incorrecta' };

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][cUser] || '').trim().toLowerCase() === username.toLowerCase().trim() &&
          String(data[i][cPass] || '').trim() === password.trim()) {
        return {
          success:  true,
          message:  'Login exitoso',
          username: String(data[i][cUser]).trim(),
          role:     cRol >= 0 ? String(data[i][cRol] || '').trim() : 'Inspector'
        };
      }
    }
    return { success: false, message: 'Usuario o contraseña incorrectos' };

  } catch (err) {
    return { success: false, message: 'Error de autenticación: ' + err.message };
  }
}


// ============================================================
//  debugInspector (mantenida para compatibilidad)
// ============================================================
function debugInspector(inspector) {
  if (!inspector) return { success: false, message: 'Inspector requerido' };
  try {
    var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('TURN');
    if (!sheet) return { success: false, message: 'Hoja TURN no encontrada' };
    var data    = sheet.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).trim(); });
    var ci = findColIdx(headers, 'Inspector');
    var matched = [];
    data.slice(1).forEach(function(row, idx) {
      var insp = String(row[ci] || '').trim();
      if (insp.toLowerCase() === inspector.toLowerCase()) {
        matched.push({ row: idx + 2, inspector: insp, flightKey: row[0] });
      }
    });
    return { success: true, headers: headers, matched: matched, matchedCount: matched.length };
  } catch (err) {
    return { success: false, message: err.message };
  }
}


// ============================================================
//  Helpers internos
// ============================================================

function getSheetName(flightType) {
  if (flightType === 'RON-ARRIVAL')   return 'RON_Llegada';
  if (flightType === 'RON-DEPARTURE') return 'RON_Salida';
  return 'TURN';
}

function getHeadersForType(flightType) {
  if (flightType === 'RON-ARRIVAL')   return HEADERS_RON_LLEGADA;
  if (flightType === 'RON-DEPARTURE') return HEADERS_RON_SALIDA;
  return HEADERS_TURN;
}

function buildFlightKey(airline, flightNumber, timestamp) {
  var date    = timestamp ? new Date(timestamp) : new Date();
  var dateStr = date.getFullYear() + '-' +
    String(date.getMonth() + 1).padStart(2, '0') + '-' +
    String(date.getDate()).padStart(2, '0');
  return (airline || '') + '-' + (flightNumber || '') + '-' + dateStr;
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    // Escribir headers con formato
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#1e3a5f');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setFontSize(10);
    sheet.setFrozenRows(1);
    // Autosize primeras columnas
    sheet.setColumnWidth(1, 220); // Flight Key
    sheet.setColumnWidth(2, 80);
    sheet.setColumnWidth(3, 130);
    sheet.setColumnWidth(4, 90);
  } else {
    // Verificar que los headers existan; si no, reescribir
    if (sheet.getLastRow() === 0) {
      var hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setValues([headers]);
      hr.setBackground('#1e3a5f');
      hr.setFontColor('#ffffff');
      hr.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function findOrCreateFlightRow(sheet, flightKey, data, headers) {
  var sheetHeaders = sheet.getLastRow() > 0
    ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    : headers;

  var colFKey  = findColIdx(sheetHeaders, 'Flight Key');
  var lastRow  = sheet.getLastRow();

  // Buscar fila existente
  if (colFKey >= 0 && lastRow > 1) {
    var keys = sheet.getRange(2, colFKey + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === flightKey) return i + 2;
    }
  }

  // Crear fila nueva
  var newRow = lastRow + 1;
  var rowData = new Array(sheetHeaders.length).fill('');

  function setCol(name, value) {
    var idx = findColIdx(sheetHeaders, name);
    if (idx >= 0) rowData[idx] = value;
  }

  setCol('Flight Key',         flightKey);
  setCol('Código Aerolínea',   data.airline          || '');
  setCol('Aerolínea',          getAirlineName(data.airline));
  setCol('Vuelo Origen',       data.flightOrigen      || '');
  setCol('Vuelo Destino',      data.flightDestino     || '');
  setCol('Vuelo',              data.flightNumber      || '');
  setCol('Tipo',               data.flightType        || '');
  setCol('Fecha/Hora',         data.flightTimestamp   || new Date().toISOString());
  setCol('Inspector',          data.inspector         || '');
  setCol('Rol',                data.role              || '');
  setCol('Posición',           data.posicion          || '');
  setCol('Estatus',            'En Proceso');
  setCol('Items Totales',      getTotalItems(data.flightType));
  setCol('Items Completados',  0);
  setCol('Tasa Completitud (%)', 0);

  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);

  // Alternar color de fila
  var bgColor = (newRow % 2 === 0) ? '#0d1b30' : '#0a1425';
  sheet.getRange(newRow, 1, 1, sheetHeaders.length).setBackground(bgColor);
  sheet.getRange(newRow, 1, 1, sheetHeaders.length).setFontColor('#e2e8f0');

  return newRow;
}

function updateMetrics(sheet, row, headers) {
  try {
    var lastCol  = sheet.getLastColumn();
    var rowData  = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
    var META_SET = new Set([
      'flight key','código aerolínea','aerolinea','aerolínea','vuelo origen',
      'vuelo destino','vuelo','tipo','fecha/hora','inspector','rol','posición',
      'aeropuerto','tiempo total (min)','items completados','items totales',
      'tasa completitud (%)','estatus','observaciones finales'
    ]);

    var completed = 0, total = 0;
    headers.forEach(function(h, i) {
      if (!h || META_SET.has(norm(h))) return;
      total++;
      var val = rowData[i];
      if (val !== null && val !== undefined && val !== '' && val !== false && val !== 0) completed++;
    });

    var tasa = total > 0 ? Math.round((completed / total) * 100) : 0;
    function sc(name, value) {
      var idx = findColIdx(headers, name);
      if (idx >= 0) sheet.getRange(row, idx + 1).setValue(value);
    }
    sc('Items Completados', completed);
    sc('Items Totales',     total);
    sc('Tasa Completitud (%)', tasa);
    if (tasa >= 100) sc('Estatus', 'Completado');

  } catch (e) {
    Logger.log('updateMetrics error: ' + e.toString());
  }
}

function deleteRowsByFlightKey(sheet, flightKey) {
  if (sheet.getLastRow() < 2) return;
  var headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colFKey  = findColIdx(headers, 'Flight Key');
  if (colFKey < 0) return;
  var data     = sheet.getRange(2, colFKey + 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (String(data[i][0]).trim() === flightKey) {
      sheet.deleteRow(i + 2);
    }
  }
}

function applySheetFormatting(sheet, type) {
  // Congelar headers si aún no está hecho
  if (sheet.getFrozenRows() === 0) sheet.setFrozenRows(1);
  // Color header
  var lastCol = sheet.getLastColumn();
  if (lastCol > 0) {
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground('#1e3a5f')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
}

// ── Utilidades básicas ───────────────────────────────────────

function findColIdx(headers, name) {
  var normName = norm(name);
  for (var i = 0; i < headers.length; i++) {
    if (norm(String(headers[i])) === normName) return i;
  }
  // partial match
  for (var j = 0; j < headers.length; j++) {
    if (norm(String(headers[j])).indexOf(normName) >= 0) return j;
  }
  return -1;
}

function norm(s) {
  return String(s).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^\w\s%\/()]/g, '')
    .trim();
}

function getAirlineName(code) {
  var names = {
    'DL':'Delta','WN':'Southwest','AV':'Avianca','UA':'United',
    'AA':'American Airlines','WS':'Westjet','E9':'Iberojet','PO':'Porter',
    'NK':'Spirit','IB':'Iberia','TS':'Air Transat','B6':'Jetblue',
    'D0':'DHL','CG':'Cargo'
  };
  return names[String(code || '').toUpperCase()] || code || '';
}

function getTotalItems(flightType) {
  if (flightType === 'TURN')          return 39;
  if (flightType === 'RON-ARRIVAL')   return 28;
  if (flightType === 'RON-DEPARTURE') return 22;
  return 39;
}

// ============================================================
//  INSTRUCCIONES DE INTEGRACIÓN CON EL HTML
//  ─────────────────────────────────────────
//  En el HTML, reemplaza los calls a post() por la versión
//  extendida de abajo (o añade estas funciones nuevas):
//
//  1. saveInspection:  action:'saveInspection' — ya existe,
//     ahora añade campo "posicion: sPosicion" al payload
//
//  2. Al Finalizar (confirmFinalReport):
//     post({ action:'saveFacturable', airline:sA, flightNumber:sF,
//            flightType:sT, flightTimestamp:fTS, inspector:U.username,
//            posicion:sPosicion, items: billingItems })
//
//  3. saveInc() → post({ action:'saveIncidente', airline:sA,
//            flightNumber:sF, flightType:sT, flightTimestamp:fTS,
//            inspector:U.username, posicion:sPosicion, incident:incData })
//
//  4. saveSeg() → post({ action:'saveSeguridad', airline:sA,
//            flightNumber:sF, flightType:sT, flightTimestamp:fTS,
//            inspector:U.username, posicion:sPosicion, seguridad:segData })
//
//  5. closeCancelReport() → post({ action:'saveCancelado', ... })
//
//  El campo DRIVE_FOLDER_ID debe ser el ID de la carpeta de Drive
//  donde quieres guardar las fotos (visible en la URL de Drive).
// ============================================================
