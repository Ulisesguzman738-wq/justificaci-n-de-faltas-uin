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
      saveAllData(payload.data);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
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
function saveAllData(db) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
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
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    sheet.clear();
    
    const headers = SCHEMAS[sheetName];
    sheet.appendRow(headers);
    
    const items = db[apiKey] || [];
    if (items.length === 0) continue;
    
    const values = [];
    for (let item of items) {
      const row = [];
      for (let header of headers) {
        let val = item[header];
        if (val === undefined || val === null) {
          row.push('');
        } else {
          row.push(val);
        }
      }
      values.push(row);
    }
    
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
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
