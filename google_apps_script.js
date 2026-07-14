/**
 * JUSTIFALTAS - GOOGLE APPS SCRIPT BACKEND
 * 
 * Este script actúa como API para la aplicación Justifaltas.
 * Permite leer/escribir datos en Google Sheets y subir archivos a Google Drive.
 * 
 * INSTRUCCIONES DE INSTALACIÓN:
 * 1. Abre tu hoja de cálculo de Google Sheets.
 * 2. Ve a "Extensiones" -> "Apps Script".
 * 3. Borra cualquier código existente y pega todo este código.
 * 4. Haz clic en "Guardar" (icono de disquete).
 * 5. Haz clic en "Implementar" -> "Nueva implementación".
 * 6. Configura:
 *    - Tipo de implementación: Aplicación web
 *    - Ejecutar como: Yo (tu correo)
 *    - Quién tiene acceso: Cualquiera (esto es clave para que la app pueda comunicarse sin autenticación compleja)
 * 7. Haz clic en "Implementar", autoriza los accesos de Google Drive y Sheets, y copia la "URL de la aplicación web".
 * 8. Pega esa URL en el archivo `app.js` de tu aplicación en la constante `GOOGLE_SCRIPT_URL`.
 */

// Nombre de la carpeta en Google Drive donde se guardarán las evidencias.
// El script la creará automáticamente si no existe.
const DRIVE_FOLDER_NAME = "Justifaltas_Evidencias";

// Configuración de Hojas y Columnas (Esquema)
const SCHEMAS = {
  'Usuarios': ['ID_Usuario', 'Correo_Electronico', 'Nombre_Completo', 'Rol', 'Contrasena', 'Fecha_Registro', 'Activo'],
  'Justificantes': ['ID_Justificante', 'ID_Alumno', 'Fecha_Falta', 'Periodo_Tetra', 'Parcial', 'Motivo', 'Descripcion', 'Estado', 'Fecha_Registro', 'ID_Coordinador_Revisor', 'Fecha_Revision'],
  'Archivos Adjuntos': ['ID_Archivo', 'ID_Justificante', 'Nombre_Archivo', 'Ruta_Archivo', 'Tipo_Archivo', 'Fecha_Carga', 'Tamano_KB'],
  'Justificante_Maestro': ['ID', 'ID_Justificante', 'ID_Maestro', 'Estado_Maestro', 'Fecha_Notificacion', 'Fecha_Justificacion'],
  'Observaciones': ['ID_Observacion', 'ID_Justificante', 'ID_Usuario', 'Comentario', 'Tipo', 'Fecha'],
  'Notificaciones': ['ID_Notificacion', 'ID_Usuario', 'ID_Justificante', 'Mensaje', 'Leida', 'Fecha']
};

/**
 * Maneja las peticiones GET (Lectura de la Base de Datos).
 */
function doGet(e) {
  try {
    ensureSheetsAndSchemas();
    const data = readAllData();
    return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Maneja las peticiones POST (Escritura y Subida de Archivos).
 */
function doPost(e) {
  // Configuración de CORS
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
  
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Cuerpo de la petición vacío' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    
    if (action === 'upload_file') {
      const fileUrl = saveFileToDrive(payload.fileName, payload.base64Data);
      return ContentService.createTextOutput(JSON.stringify({ success: true, url: fileUrl }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === 'save_db') {
      ensureSheetsAndSchemas();
      saveAllData(payload.data, payload.callerRole, payload.callerId);
      const data = readAllData();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else if (action === 'delete_user') {
      ensureSheetsAndSchemas();
      deleteUser(payload.userId);
      const data = readAllData();
      return ContentService.createTextOutput(JSON.stringify({ success: true, data: data }))
        .setMimeType(ContentService.MimeType.JSON);
        
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Acción no reconocida: ' + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Asegura que existan todas las hojas requeridas y que tengan las columnas correctas en la primera fila.
 * Si faltan columnas críticas en Usuarios (como Contrasena), las agrega con valores por defecto.
 */
function ensureSheetsAndSchemas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  for (let sheetName in SCHEMAS) {
    let sheet = ss.getSheetByName(sheetName);
    const expectedHeaders = SCHEMAS[sheetName];
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(expectedHeaders);
      
      // Rellenar usuarios por defecto si la hoja Usuarios es recién creada
      if (sheetName === 'Usuarios') {
        const defaultUsers = [
          ['sim_user_epy1yk2', 'Prueba@uin.edu.mx', 'Usuario de Pruebas UIN', 'alumno', '12345', new Date().toISOString(), 1],
          ['Uguzman', 'ulisega03s24a@uinteramericana.edu.mx', 'Ulise Guzman Alvaro', 'alumno', '12345', new Date().toISOString(), 1],
          ['Syanet', 'syanethh03s24a@uinteramericana.edu.mx', 'Sandra Yanet Hernández Hernández', 'alumno', '12345', new Date().toISOString(), 1],
          ['Ihernandez', 'irenehh03s24a@uinteramericana.edu.mx', 'Irene Hernández Hernández ', 'alumno', '12345', new Date().toISOString(), 1],
          ['Yulila', 'iyuleidals03s24a@uinteramericana.edu.mx', 'Itzel Yuleida Lila Sanchez ', 'alumno', '12345', new Date().toISOString(), 1],
          ['Dantonio', 'dantonioie03s24a@uinteramericana.edu.mx', 'Diego Antonio Iracheta Escareño', 'alumno', '12345', new Date().toISOString(), 1],
          ['Mangel', 'mdavila0311@uinteramericana.edu.mx', 'Miguel Angel Davila Calzada', 'maestro', '12345', new Date().toISOString(), 1],
          ['Aeloisa', 'ana.hernandez03ea26@uinteramericana.edu.mx', 'Ana Eloisa Hernández Jimenez', 'maestro', '12345', new Date().toISOString(), 1],
          ['Oalavez', 'oscar.alavez03ea26p@uinteramericana.edu.mx', 'Oscar Alavez Reyes', 'maestro', '12345', new Date().toISOString(), 1],
          ['Asimon', 'andres.simon03ea26@uinteramericana.edu.mx', 'Andres Simón Treviño', 'maestro', '12345', new Date().toISOString(), 1],
          ['SanSilva', 'direccionsc_sd25o@uinteramericana.edu.mx', 'Lic. Sandra Silva', 'coordinacion', '12345', new Date().toISOString(), 1],
          ['EscSanta', 'escolarsc@universidadinteramericana.edu.mx', 'Escolar Santa Catarina', 'coordinacion', '12345', new Date().toISOString(), 1]
        ];
        sheet.getRange(2, 1, defaultUsers.length, expectedHeaders.length).setValues(defaultUsers);
      }
      continue;
    }
    
    // Si la hoja ya existe, validar y complementar columnas
    const range = sheet.getDataRange();
    const values = range.getValues();
    let headers = [];
    
    if (values.length > 0) {
      headers = values[0].map(h => h.toString().trim());
    } else {
      sheet.appendRow(expectedHeaders);
      continue;
    }
    
    // Rellenar usuarios por defecto si la hoja Usuarios ya existía pero está vacía
    if (sheetName === 'Usuarios' && sheet.getLastRow() === 1) {
      const defaultUsers = [
        ['sim_user_epy1yk2', 'Prueba@uin.edu.mx', 'Usuario de Pruebas UIN', 'alumno', '12345', new Date().toISOString(), 1],
        ['Uguzman', 'ulisega03s24a@uinteramericana.edu.mx', 'Ulise Guzman Alvaro', 'alumno', '12345', new Date().toISOString(), 1],
        ['Syanet', 'syanethh03s24a@uinteramericana.edu.mx', 'Sandra Yanet Hernández Hernández', 'alumno', '12345', new Date().toISOString(), 1],
        ['Ihernandez', 'irenehh03s24a@uinteramericana.edu.mx', 'Irene Hernández Hernández ', 'alumno', '12345', new Date().toISOString(), 1],
        ['Yulila', 'iyuleidals03s24a@uinteramericana.edu.mx', 'Itzel Yuleida Lila Sanchez ', 'alumno', '12345', new Date().toISOString(), 1],
        ['Dantonio', 'dantonioie03s24a@uinteramericana.edu.mx', 'Diego Antonio Iracheta Escareño', 'alumno', '12345', new Date().toISOString(), 1],
        ['Mangel', 'mdavila0311@uinteramericana.edu.mx', 'Miguel Angel Davila Calzada', 'maestro', '12345', new Date().toISOString(), 1],
        ['Aeloisa', 'ana.hernandez03ea26@uinteramericana.edu.mx', 'Ana Eloisa Hernández Jimenez', 'maestro', '12345', new Date().toISOString(), 1],
        ['Oalavez', 'oscar.alavez03ea26p@uinteramericana.edu.mx', 'Oscar Alavez Reyes', 'maestro', '12345', new Date().toISOString(), 1],
        ['Asimon', 'andres.simon03ea26@uinteramericana.edu.mx', 'Andres Simón Treviño', 'maestro', '12345', new Date().toISOString(), 1],
        ['SanSilva', 'direccionsc_sd25o@uinteramericana.edu.mx', 'Lic. Sandra Silva', 'coordinacion', '12345', new Date().toISOString(), 1],
        ['EscSanta', 'escolarsc@universidadinteramericana.edu.mx', 'Escolar Santa Catarina', 'coordinacion', '12345', new Date().toISOString(), 1]
      ];
      sheet.getRange(2, 1, defaultUsers.length, expectedHeaders.length).setValues(defaultUsers);
    }
    
    // Buscar si faltan columnas esperadas
    let modified = false;
    const missingHeaders = expectedHeaders.filter(h => !headers.includes(h));
    
    if (missingHeaders.length > 0) {
      // Agregar las columnas faltantes al final de la primera fila
      const lastCol = sheet.getLastColumn();
      sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
      
      // Si agregamos columnas a Usuarios, rellenar valores por defecto para filas existentes
      if (sheetName === 'Usuarios') {
        const updatedRange = sheet.getDataRange();
        const updatedValues = updatedRange.getValues();
        const newHeaders = updatedValues[0].map(h => h.toString().trim());
        
        for (let i = 1; i < updatedValues.length; i++) {
          if (missingHeaders.includes('Contrasena')) {
            const colIndex = newHeaders.indexOf('Contrasena') + 1;
            sheet.getRange(i + 1, colIndex).setValue('12345'); // Contraseña por defecto
          }
          if (missingHeaders.includes('Activo')) {
            const colIndex = newHeaders.indexOf('Activo') + 1;
            sheet.getRange(i + 1, colIndex).setValue(1); // Activo por defecto
          }
        }
      }
    }
  }
}

/**
 * Lee todas las hojas y las transforma en el formato de base de datos JSON de Justifaltas.
 */
function readAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const db = {};
  
  const sheetMapping = {
    'usuarios': 'Usuarios',
    'justificaciones': 'Justificantes',
    'archivos_adjuntos': 'Archivos Adjuntos',
    'justificante_maestro': 'Justificante_Maestro',
    'observaciones': 'Observaciones',
    'notificaciones': 'Notificaciones'
  };
  
  for (let apiKey in sheetMapping) {
    const sheetName = sheetMapping[apiKey];
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      db[apiKey] = [];
      continue;
    }
    
    const range = sheet.getDataRange();
    const rows = range.getValues();
    
    if (rows.length <= 1) {
      db[apiKey] = [];
      continue;
    }
    
    const headers = rows[0].map(h => h.toString().trim());
    const data = [];
    
    for (let i = 1; i < rows.length; i++) {
      const obj = {};
      let hasData = false;
      
      for (let j = 0; j < headers.length; j++) {
        let val = rows[i][j];
        
        if (val !== undefined && val !== null && val !== "") {
          hasData = true;
        }
        
        if (val instanceof Date) {
          obj[headers[j]] = val.toISOString();
        } else if (val === "") {
          obj[headers[j]] = null;
        } else {
          obj[headers[j]] = val;
        }
      }
      
      if (hasData) {
        data.push(obj);
      }
    }
    db[apiKey] = data;
  }
  
  return db;
}

/**
 * Sobrescribe las hojas con los nuevos datos guardados por la aplicación.
 */
function saveAllData(db, callerRole, callerId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  callerRole = callerRole || 'guest';
  callerId = callerId || '';
  
  const sheetMapping = {
    'usuarios': 'Usuarios',
    'justificaciones': 'Justificantes',
    'archivos_adjuntos': 'Archivos Adjuntos',
    'justificante_maestro': 'Justificante_Maestro',
    'observaciones': 'Observaciones',
    'notificaciones': 'Notificaciones'
  };
  
  const primaryKeys = {
    'Usuarios': 'ID_Usuario',
    'Justificantes': 'ID_Justificante',
    'Archivos Adjuntos': 'ID_Archivo',
    'Justificante_Maestro': 'ID',
    'Observaciones': 'ID_Observacion',
    'Notificaciones': 'ID_Notificacion'
  };
  
  for (let apiKey in sheetMapping) {
    const sheetName = sheetMapping[apiKey];
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    const expectedHeaders = SCHEMAS[sheetName];
    const pkName = primaryKeys[sheetName];
    
    // Garantizar que la hoja esté inicializada con cabeceras correctas
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(expectedHeaders);
    }
    
    const range = sheet.getDataRange();
    let rows = range.getValues();
    let headers = rows[0].map(h => h.toString().trim());
    
    // Validar cabeceras, si no coinciden inicializar
    const match = expectedHeaders.every(h => headers.includes(h));
    if (!match) {
      sheet.clear();
      sheet.appendRow(expectedHeaders);
      rows = [expectedHeaders];
      headers = expectedHeaders;
    }
    
    const pkIndex = headers.indexOf(pkName);
    if (pkIndex === -1) {
      sheet.clear();
      sheet.appendRow(expectedHeaders);
      rows = [expectedHeaders];
      headers = expectedHeaders;
    }
    
    // Mapear filas existentes
    const existingRows = rows.slice(1);
    const existingMap = {};
    existingRows.forEach((row, idx) => {
      const pkVal = row[headers.indexOf(pkName)];
      if (pkVal !== undefined && pkVal !== null && pkVal !== '') {
        existingMap[pkVal.toString()] = { rowIndex: idx + 2, data: row };
      }
    });
    
    const clientItems = db[apiKey] || [];
    const clientItemIds = new Set();
    
    clientItems.forEach(item => {
      const pkVal = item[pkName];
      if (pkVal === undefined || pkVal === null || pkVal === '') return;
      
      const pkValStr = pkVal.toString();
      clientItemIds.add(pkValStr);
      
      // Construir fila en orden de cabeceras
      const rowData = headers.map(header => {
        let val = item[header];
        return (val === undefined || val === null) ? '' : val;
      });
      
      if (existingMap[pkValStr]) {
        // Reglas de negocio para proteger estados críticos frente a sobrescrituras concurrentes:
        const existingData = existingMap[pkValStr].data;
        const rowIndex = existingMap[pkValStr].rowIndex;
        
        // 1. En Justificantes, no regresar a "Pendiente" si ya está "Aprobada" o "Rechazada"
        if (sheetName === 'Justificantes') {
          const stateCol = headers.indexOf('Estado');
          const existingState = existingData[stateCol];
          const incomingState = item['Estado'];
          if ((existingState === 'Aprobada' || existingState === 'Rechazada') && incomingState === 'Pendiente') {
            rowData[stateCol] = existingState; // Conservar el estado de la hoja
          }
        }
        
        // 2. En Justificante_Maestro, no regresar a "Pendiente" si ya está "Enterada por Maestro"
        if (sheetName === 'Justificante_Maestro') {
          const stateCol = headers.indexOf('Estado_Maestro');
          const existingState = existingData[stateCol];
          const incomingState = item['Estado_Maestro'];
          if (existingState === 'Enterada por Maestro' && incomingState === 'Pendiente') {
            rowData[stateCol] = existingState;
          }
        }
        
        // 3. En Usuarios, solo permitir actualizar la contraseña si es el propio usuario logueado (callerId)
        // o si la contraseña en la hoja está vacía.
        if (sheetName === 'Usuarios') {
          const pwdCol = headers.indexOf('Contrasena');
          if (pwdCol !== -1) {
            const existingPwd = existingData[pwdCol];
            const incomingPwd = item['Contrasena'];
            if (existingPwd && existingPwd !== '' && item['ID_Usuario'] !== callerId) {
              rowData[pwdCol] = existingPwd; // Conservar la contraseña de la hoja
            }
          }
        }
        
        // Actualizar fila
        sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
      } else {
        // Insertar nueva fila
        sheet.appendRow(rowData);
      }
    });
    
  }
}

/**
 * Elimina un usuario por su ID de la hoja Usuarios y limpia sus relaciones en otras hojas.
 */
function deleteUser(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!userId) return;
  
  // 1. Eliminar de Usuarios
  const userSheet = ss.getSheetByName('Usuarios');
  if (userSheet) {
    const range = userSheet.getDataRange();
    const rows = range.getValues();
    const headers = rows[0].map(h => h.toString().trim());
    const pkIndex = headers.indexOf('ID_Usuario');
    if (pkIndex !== -1) {
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][pkIndex].toString() === userId.toString()) {
          userSheet.deleteRow(i + 1);
          break;
        }
      }
    }
  }
  
  // 2. Eliminar de Justificante_Maestro (limpiar relaciones con docente)
  const mappingSheet = ss.getSheetByName('Justificante_Maestro');
  if (mappingSheet) {
    const range = mappingSheet.getDataRange();
    const rows = range.getValues();
    const headers = rows[0].map(h => h.toString().trim());
    const maestroIndex = headers.indexOf('ID_Maestro');
    if (maestroIndex !== -1) {
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][maestroIndex].toString() === userId.toString()) {
          mappingSheet.deleteRow(i + 1);
        }
      }
    }
  }
  
  // 3. Eliminar de Notificaciones (limpiar notificaciones del docente)
  const notifSheet = ss.getSheetByName('Notificaciones');
  if (notifSheet) {
    const range = notifSheet.getDataRange();
    const rows = range.getValues();
    const headers = rows[0].map(h => h.toString().trim());
    const userIndex = headers.indexOf('ID_Usuario');
    if (userIndex !== -1) {
      for (let i = rows.length - 1; i >= 1; i--) {
        if (rows[i][userIndex].toString() === userId.toString()) {
          notifSheet.deleteRow(i + 1);
        }
      }
    }
  }
}

/**
 * Guarda un archivo base64 en Google Drive, le da permisos de visor público y retorna la URL.
 */
function saveFileToDrive(fileName, base64Data) {
  let base64Content = base64Data;
  if (base64Data.indexOf(';base64,') !== -1) {
    base64Content = base64Data.split(';base64,')[1];
  }
  
  let contentType = 'application/octet-stream';
  if (base64Data.indexOf('data:') === 0 && base64Data.indexOf(';base64,') !== -1) {
    contentType = base64Data.substring(5, base64Data.indexOf(';base64,'));
  }
  
  const decoded = Utilities.base64Decode(base64Content);
  const blob = Utilities.newBlob(decoded, contentType, fileName);
  
  // Buscar o crear la carpeta contenedora en Google Drive
  let folder;
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  const file = folder.createFile(blob);
  
  // Compartir archivo para que cualquiera con el enlace lo pueda ver (necesario para el visor del navegador)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  return file.getUrl();
}
