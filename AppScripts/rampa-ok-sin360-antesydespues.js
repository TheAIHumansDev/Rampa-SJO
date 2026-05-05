// ============================================
// APPS SCRIPT - TURNAROUND CHECK CON FOTOS
// VERSIÓN: Una fila por vuelo completo
// ============================================

// CONFIGURACIÓN - ACTUALIZAR ESTOS IDs
const SPREADSHEET_ID = '1dLDXOTzhy8K23Kljg8Ikioazmyj-vMqHRTgS98LjqAE';
const DRIVE_FOLDER_ID = '17QalYwu1mXAavWmvDh2LyVK3NxgfH0jQ';

// Nombres de las hojas según tipo de vuelo
const SHEET_NAMES = {
  'TURN': 'TURN',
  'RON-ARRIVAL': 'RON_Llegada',
  'RON-DEPARTURE': 'RON_Salida'
};

// Nombres de aerolíneas
const AIRLINE_NAMES = {
  'DL': 'Delta',
  'WN': 'Southwest',
  'E9': 'Iberojet',
  'AV': 'Avianca',
  'UA': 'United',
  'AA': 'American Airlines'
};

// Checkpoints por tipo de vuelo
const CHECKPOINTS = {
  'TURN': [
    'briefing', 'inicio_fod', 'beacon_off', 'planta_electrica', 'aire_acondicionado',
    'banda', 'escalera', 'inicia_descarga_carga', 'waste_water', 'descarga_equipaje',
    'inicia_equipaje', 'finaliza_equipaje', 'sube_limpieza', 'baja_limpieza',
    'segundo_fod', 'ultimo_viaje_carreta', 'inicia_carga_maleta', 'ultima_maleta',
    'towbar', 'cierre_bines', 'push_back', 'tercer_fod', 'en_aire'
  ],
  'RON-ARRIVAL': [
    'briefing', 'fod', 'beacon_off', 'planta_electrica', 'aire_acondicionado',
    'banda', 'escalera', 'inicia_descarga', 'finaliza_descarga', 'inicia_carga',
    'finaliza_carga', 'waste_water', 'sube_limpieza', 'baja_limpieza',
    'segundo_fod', 'ultimo_viaje_carreta', 'cierre_bines'
  ],
  'RON-DEPARTURE': [
    'briefing', 'fod', 'avion_posicion', 'planta_electrica', 'aire_acondicionado',
    'banda', 'escalera', 'inicia_carga', 'waste_water', 'segundo_fod',
    'finaliza_carga', 'inicia_carga_carga', 'finaliza_carga_carga', 'towbar',
    'cierre_bines', 'tercer_fod'
  ]
};

// Labels amigables para cada checkpoint
const CHECKPOINT_LABELS = {
  'briefing': 'Briefing',
  'inicio_fod': 'Inicio FOD',
  'beacon_off': 'Beacon OFF',
  'planta_electrica': 'Planta Eléctrica',
  'aire_acondicionado': 'Aire Acondicionado',
  'banda': 'Banda',
  'escalera': 'Escalera',
  'inicia_descarga_carga': 'Inicia Descarga Carga',
  'waste_water': 'Waste & Water',
  'descarga_equipaje': 'Finaliza Descarga Carga',
  'inicia_equipaje': 'Inicia Descarga Equipaje',
  'finaliza_equipaje': 'Finaliza Descarga Equipaje',
  'sube_limpieza': 'Sube Limpieza',
  'baja_limpieza': 'Baja Limpieza',
  'segundo_fod': '2do FOD',
  'ultimo_viaje_carreta': 'Último Viaje Carreta',
  'inicia_carga_maleta': 'Inicia Carga Maleta',
  'ultima_maleta': 'Última Maleta',
  'towbar': 'Towbar',
  'cierre_bines': 'Cierre Bines',
  'push_back': 'Push Back',
  'tercer_fod': '3er FOD',
  'en_aire': 'En el Aire',
  'fod': 'FOD',
  'avion_posicion': 'Avión en Posición',
  'inicia_descarga': 'Inicia Descarga Equipaje',
  'finaliza_descarga': 'Finaliza Descarga Equipaje',
  'inicia_carga': 'Inicia Carga',
  'finaliza_carga': 'Finaliza Carga',
  'inicia_carga_carga': 'Inicia Carga de Carga',
  'finaliza_carga_carga': 'Finaliza Carga de Carga'
};

// Cache de datos de vuelos en proceso
const flightDataCache = {};

// ============================================
// FUNCIÓN PRINCIPAL - doPost
// ============================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    Logger.log('📥 Datos recibidos:', data.action);
    
    if (data.action === 'saveInspection') {
      return saveInspection(data);
    } else if (data.action === 'updateStatus') {
      return updateStatus(data);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Acción no reconocida'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('❌ Error en doPost:', error);
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// GUARDAR INSPECCIÓN (Actualizar fila del vuelo)
// ============================================
function saveInspection(data) {
  try {
    // Validar configuración
    if (SPREADSHEET_ID === 'TU_SPREADSHEET_ID_AQUI') {
      throw new Error('ERROR: Actualizar SPREADSHEET_ID en línea 6');
    }
    
    Logger.log('📊 Abriendo spreadsheet...');
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = SHEET_NAMES[data.flightType] || 'TURN';
    let sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log('⚠️ Creando hoja: ' + sheetName);
      sheet = spreadsheet.insertSheet(sheetName);
      createHeaders(sheet, data.flightType);
    }
    
    if (sheet.getLastRow() === 0) {
      createHeaders(sheet, data.flightType);
    }
    
    const flightKey = `${data.airline}_${data.flightNumber}_${data.flightTimestamp}`;
    Logger.log('🔑 Flight Key: ' + flightKey);
    
    // Buscar o crear fila del vuelo
    let rowNumber = findFlightRow(sheet, flightKey);
    let isNewFlight = false;
    
    if (!rowNumber) {
      // Nuevo vuelo - crear fila
      Logger.log('✨ Nuevo vuelo - creando fila');
      rowNumber = sheet.getLastRow() + 1;
      isNewFlight = true;
      
      // Inicializar fila con datos básicos
      const basicData = [
        flightKey,                                    // A: Flight Key
        data.airline,                                  // B: Código Aerolínea
        AIRLINE_NAMES[data.airline] || data.airline,  // C: Aerolínea
        data.flightNumber,                             // D: Número Vuelo
        data.flightType,                               // E: Tipo Vuelo
        new Date(data.flightTimestamp),                // F: Fecha/Hora Inicio
        data.inspector,                                // G: Inspector
        data.role                                      // H: Rol
      ];
      
      // Agregar fila vacía
      sheet.appendRow(basicData);
    }
    
    Logger.log('📍 Fila del vuelo: ' + rowNumber);
    
    // Guardar foto en Drive si existe
    let photoUrl = '';
    if (data.photoBase64) {
      Logger.log('📸 Guardando foto...');
      photoUrl = savePhotoToDrive(
        data.photoBase64,
        data.airline,
        data.flightNumber,
        data.item,
        data.timestamp
      );
    }
    
    // Actualizar columna del checkpoint
    updateCheckpointColumn(sheet, rowNumber, data, photoUrl);
    
    // Si es Beacon OFF, guardar tiempo de referencia
    if (data.item === 'Beacon OFF') {
      const beaconColIndex = getColumnIndexForCheckpoint('beacon_off', data.flightType);
      if (beaconColIndex > 0) {
        sheet.getRange(rowNumber, beaconColIndex).setNote('REFERENCIA: ' + data.timestamp);
      }
    }
    
    // Actualizar métricas del vuelo
    updateFlightMetrics(sheet, rowNumber, data.flightType);
    
    Logger.log('✅ Checkpoint guardado: ' + data.item);
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Checkpoint guardado',
      photoUrl: photoUrl
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('❌ Error:', error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// CREAR HEADERS SEGÚN TIPO DE VUELO
// ============================================
function createHeaders(sheet, flightType) {
  const checkpoints = CHECKPOINTS[flightType] || CHECKPOINTS['TURN'];
  
  const headers = [
    'Flight Key',        // A
    'Código Aerolínea',  // B
    'Aerolínea',         // C
    'Vuelo',             // D
    'Tipo',              // E
    'Fecha/Hora',        // F
    'Inspector',         // G
    'Rol'                // H
  ];
  
  // Agregar columnas para cada checkpoint
  checkpoints.forEach(checkpoint => {
    headers.push(CHECKPOINT_LABELS[checkpoint] || checkpoint);
  });
  
  // Agregar columnas de métricas finales
  headers.push('Tiempo Total (min)');
  headers.push('Items Completados');
  headers.push('Items Totales');
  headers.push('Tasa Completitud (%)');
  headers.push('Estatus');
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatear headers
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a237e');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setWrap(true);
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, Math.min(headers.length, 26));
}

// ============================================
// BUSCAR FILA DEL VUELO
// ============================================
function findFlightRow(sheet, flightKey) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === flightKey) {
      return i + 1; // +1 porque las filas empiezan en 1
    }
  }
  return null;
}

// ============================================
// OBTENER ÍNDICE DE COLUMNA PARA CHECKPOINT
// ============================================
function getColumnIndexForCheckpoint(checkpointKey, flightType) {
  const checkpoints = CHECKPOINTS[flightType] || CHECKPOINTS['TURN'];
  const index = checkpoints.indexOf(checkpointKey);
  if (index === -1) return -1;
  
  // 8 columnas base + índice del checkpoint
  return 9 + index;
}

// ============================================
// ACTUALIZAR COLUMNA DEL CHECKPOINT
// ============================================
function updateCheckpointColumn(sheet, rowNumber, data, photoUrl) {
  const checkpointKey = convertItemToKey(data.item);
  const colIndex = getColumnIndexForCheckpoint(checkpointKey, data.flightType);
  
  if (colIndex < 0) {
    Logger.log('⚠️ Checkpoint no encontrado: ' + checkpointKey);
    return;
  }
  
  const timestamp = new Date(data.timestamp);
  const timeStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'HH:mm:ss');
  
  // Obtener valor actual (puede tener múltiples timestamps)
  const currentValue = sheet.getRange(rowNumber, colIndex).getValue();
  let newValue = timeStr;
  
  if (currentValue && currentValue.toString().trim() !== '') {
    // Ya hay timestamps, agregar nuevo
    newValue = currentValue + '\n' + timeStr;
  }
  
  // Si es briefing, agregar temas y notas
  if (checkpointKey === 'briefing' && (data.briefingTopics || data.briefingNotes)) {
    const topics = data.briefingTopics ? data.briefingTopics.join(', ') : '';
    const notes = data.briefingNotes || '';
    newValue = timeStr + '\nTemas: ' + topics + '\n' + notes;
  }
  
  const cell = sheet.getRange(rowNumber, colIndex);
  cell.setValue(newValue);
  
  // Agregar nota con URL de foto si existe
  if (photoUrl) {
    cell.setNote('Foto: ' + photoUrl);
    cell.setFontColor('#1565c0');
  }
  
  // Calcular diferencia vs Beacon OFF
  const beaconOffTime = getBeaconOffTimeForFlight(sheet, rowNumber, data.flightType);
  if (beaconOffTime && checkpointKey !== 'beacon_off') {
    const diffMs = timestamp - beaconOffTime;
    const diffMins = Math.floor(diffMs / 60000);
    const diffSecs = Math.floor((diffMs % 60000) / 1000);
    const diffStr = `${diffMins >= 0 ? '+' : ''}${diffMins}m ${diffSecs}s`;
    
    // Agregar diferencia como comentario
    const currentNote = cell.getNote() || '';
    cell.setNote(currentNote + '\nDiff vs Beacon OFF: ' + diffStr);
    
    // Colorear según diferencia
    if (diffMins < 0) {
      cell.setBackground('#ffcdd2'); // Rojo (antes)
    } else if (diffMins <= 30) {
      cell.setBackground('#c8e6c9'); // Verde (óptimo)
    } else if (diffMins <= 60) {
      cell.setBackground('#fff9c4'); // Amarillo (aceptable)
    } else {
      cell.setBackground('#ffcdd2'); // Rojo (tardío)
    }
  } else if (checkpointKey === 'beacon_off') {
    cell.setBackground('#ffd700'); // Dorado (referencia)
    cell.setFontWeight('bold');
  }
}

// ============================================
// OBTENER BEACON OFF TIME PARA EL VUELO
// ============================================
function getBeaconOffTimeForFlight(sheet, rowNumber, flightType) {
  const beaconColIndex = getColumnIndexForCheckpoint('beacon_off', flightType);
  if (beaconColIndex < 0) return null;
  
  const beaconValue = sheet.getRange(rowNumber, beaconColIndex).getValue();
  if (!beaconValue) return null;
  
  // Extraer primer timestamp (formato HH:mm:ss)
  const timeStr = beaconValue.toString().split('\n')[0];
  
  // Obtener fecha del vuelo
  const flightDate = sheet.getRange(rowNumber, 6).getValue(); // Columna F
  
  // Combinar fecha con hora
  const [hours, minutes, seconds] = timeStr.split(':').map(Number);
  const beaconTime = new Date(flightDate);
  beaconTime.setHours(hours, minutes, seconds, 0);
  
  return beaconTime;
}

// ============================================
// ACTUALIZAR MÉTRICAS DEL VUELO
// ============================================
function updateFlightMetrics(sheet, rowNumber, flightType) {
  const checkpoints = CHECKPOINTS[flightType] || CHECKPOINTS['TURN'];
  const totalCheckpoints = checkpoints.length;
  
  let completedCount = 0;
  let firstTime = null;
  let lastTime = null;
  
  // Contar checkpoints completados y encontrar tiempos
  checkpoints.forEach(checkpoint => {
    const colIndex = getColumnIndexForCheckpoint(checkpoint, flightType);
    if (colIndex > 0) {
      const value = sheet.getRange(rowNumber, colIndex).getValue();
      if (value && value.toString().trim() !== '') {
        completedCount++;
        
        // Extraer timestamp
        const timeStr = value.toString().split('\n')[0];
        const flightDate = sheet.getRange(rowNumber, 6).getValue();
        const [hours, minutes, seconds] = timeStr.split(':').map(Number);
        const checkpointTime = new Date(flightDate);
        checkpointTime.setHours(hours, minutes, seconds, 0);
        
        if (!firstTime || checkpointTime < firstTime) {
          firstTime = checkpointTime;
        }
        if (!lastTime || checkpointTime > lastTime) {
          lastTime = checkpointTime;
        }
      }
    }
  });
  
  // Calcular tiempo total
  let totalMinutes = 0;
  if (firstTime && lastTime) {
    totalMinutes = Math.round((lastTime - firstTime) / 60000);
  }
  
  // Calcular tasa de completitud
  const completionRate = totalCheckpoints > 0 ? 
    Math.round((completedCount / totalCheckpoints) * 100) : 0;
  
  // Actualizar columnas de métricas
  const metricsStartCol = 9 + checkpoints.length; // Después de todos los checkpoints
  
  sheet.getRange(rowNumber, metricsStartCol).setValue(totalMinutes);     // Tiempo Total
  sheet.getRange(rowNumber, metricsStartCol + 1).setValue(completedCount); // Items Completados
  sheet.getRange(rowNumber, metricsStartCol + 2).setValue(totalCheckpoints); // Items Totales
  sheet.getRange(rowNumber, metricsStartCol + 3).setValue(completionRate); // Tasa Completitud
  
  // Formatear porcentaje
  sheet.getRange(rowNumber, metricsStartCol + 3).setNumberFormat('0"%"');
}

// ============================================
// ACTUALIZAR ESTATUS FINAL
// ============================================
function updateStatus(data) {
  try {
    const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = SHEET_NAMES[data.flightType] || 'TURN';
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: 'Hoja no encontrada'
    })).setMimeType(ContentService.MimeType.JSON);
    
    const flightKey = `${data.airline}_${data.flightNumber}_${data.timestamp}`;
    const rowNumber = findFlightRow(sheet, flightKey);
    
    if (rowNumber) {
      const checkpoints = CHECKPOINTS[data.flightType] || CHECKPOINTS['TURN'];
      const statusCol = 9 + checkpoints.length + 4; // Última columna de métricas + 1
      sheet.getRange(rowNumber, statusCol).setValue(data.status);
      
      Logger.log('✅ Estatus actualizado a: ' + data.status);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      success: true,
      message: 'Estatus actualizado'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('❌ Error:', error.toString());
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// CONVERTIR ITEM A KEY
// ============================================
function convertItemToKey(itemName) {
  // Convertir nombre de item a key del checkpoint
  for (const [key, label] of Object.entries(CHECKPOINT_LABELS)) {
    if (label === itemName) {
      return key;
    }
  }
  return itemName.toLowerCase().replace(/ /g, '_');
}

// ============================================
// GUARDAR FOTO EN DRIVE
// ============================================
function savePhotoToDrive(base64Data, airlineCode, flightNumber, checkpointName, timestamp) {
  try {
    if (DRIVE_FOLDER_ID === 'TU_FOLDER_ID_REPORTES_RAMPA_AQUI') {
      Logger.log('⚠️ DRIVE_FOLDER_ID no configurado');
      return '';
    }
    
    const rootFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const airlineName = AIRLINE_NAMES[airlineCode] || airlineCode;
    
    let airlineFolder = getFolderByName(rootFolder, airlineName);
    if (!airlineFolder) {
      airlineFolder = rootFolder.createFolder(airlineName);
    }
    
    let flightFolder = getFolderByName(airlineFolder, flightNumber);
    if (!flightFolder) {
      flightFolder = airlineFolder.createFolder(flightNumber);
    }
    
    const base64Clean = base64Data.split(',')[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Clean),
      'image/jpeg',
      sanitizeFileName(checkpointName) + '_' + formatTimestampForFile(timestamp) + '.jpg'
    );
    
    const file = flightFolder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
    
  } catch (error) {
    Logger.log('❌ Error guardando foto:', error);
    return 'Error: ' + error.toString();
  }
}

// ============================================
// FUNCIONES AUXILIARES
// ============================================
function getFolderByName(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : null;
}

function sanitizeFileName(name) {
  return name.replace(/[^a-z0-9_\-]/gi, '_');
}

function formatTimestampForFile(timestamp) {
  const date = new Date(timestamp);
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
}

// ============================================
// FUNCIÓN DE DIAGNÓSTICO
// ============================================
function diagnosticarConfiguracion() {
  Logger.log('=== DIAGNÓSTICO ===');
  
  if (SPREADSHEET_ID === 'TU_SPREADSHEET_ID_AQUI') {
    Logger.log('❌ Actualizar SPREADSHEET_ID');
    return;
  }
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log('✅ Spreadsheet: ' + ss.getName());
    
    ['TURN', 'RON_Llegada', 'RON_Salida'].forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (!sheet) {
        sheet = ss.insertSheet(name);
        const type = name === 'TURN' ? 'TURN' : name === 'RON_Llegada' ? 'RON-ARRIVAL' : 'RON-DEPARTURE';
        createHeaders(sheet, type);
        Logger.log('✅ Creada: ' + name);
      } else {
        Logger.log('✅ Existe: ' + name);
      }
    });
    
    if (DRIVE_FOLDER_ID !== 'TU_FOLDER_ID_REPORTES_RAMPA_AQUI') {
      const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
      Logger.log('✅ Drive: ' + folder.getName());
    }
    
    Logger.log('=== ✅ SISTEMA LISTO ===');
  } catch (e) {
    Logger.log('❌ Error: ' + e.toString());
  }
}
