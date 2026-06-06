/**
 * GESTORPRO - API DE CONEXIÓN DE GOOGLE SHEETS
 * 
 * Este script actúa como API (Backend) para la aplicación web GestorPRO.
 * Debe pegarse en Google Sheets: Extensión -> Apps Script.
 * Luego implementarse como "Aplicación web" con acceso para "Cualquier persona".
 */

// CLAVE SECRETA DE ACCESO - CAMBIA ESTO POR UN PIN O CONTRASEÑA MÁS SEGURA
const SECRET_TOKEN = "gestorpro123";

// Nombre de la carpeta en Google Drive donde se guardarán los comprobantes
const DRIVE_FOLDER_NAME = "GestorPRO_Comprobantes";

/**
 * Función auxiliar para verificar la contraseña/token de acceso
 */
function checkAuth(e) {
  var token = "";
  if (e.parameter && e.parameter.token) {
    token = e.parameter.token;
  }
  return token === SECRET_TOKEN;
}

/**
 * Retorna JSON con formato estándar CORS
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Convierte un valor de celda a número de forma segura.
 * Si la celda tiene un error de fórmula (#ERROR!, #REF!, etc.) o un valor no numérico, retorna 0.
 */
function safeNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

/**
 * ENDPOINT GET: Obtiene toda la información de la planilla
 */
function doGet(e) {
  if (!checkAuth(e)) {
    return jsonResponse({ status: "error", message: "No autorizado. Token incorrecto." });
  }

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    
    // Auto-repair formulas if needed to prevent locale syntax errors
    ensureCorrectFormulas(doc);
    
    // 1. LEER PROVEEDORES
    const sheetProv = doc.getSheetByName("Proveedores");
    ensureExtraHeaders(sheetProv);
    const provRange = sheetProv.getRange(6, 1, sheetProv.getLastRow() - 5, 15);
    const provValues = provRange.getValues();
    const proveedores = [];
    
    for (let i = 0; i < provValues.length; i++) {
      const row = provValues[i];
      if (row[0]) { // Verificar que exista alias
        let obsVal = row[10] || '';
        let cat = row[11] || '';
        let pri = row[12] || 'normal';
        let fechaInicio = row[13] || '';
        let condiciones = row[14] || '';
        
        // Decodificar META legacy si existe en las observaciones
        if (obsVal.startsWith('[META:')) {
          const decoded = decodeMetaObs(obsVal);
          cat = decoded.meta.cat || cat;
          pri = decoded.meta.pri || pri;
          fechaInicio = decoded.meta.fechaInicio || fechaInicio;
          condiciones = decoded.meta.cond || condiciones;
          obsVal = decoded.text;
        }
        
        // Asegurar que la fecha sea string YYYY-MM-DD
        if (fechaInicio instanceof Date) {
          fechaInicio = fechaInicio.toISOString().slice(0, 10);
        } else if (fechaInicio) {
          fechaInicio = String(fechaInicio).slice(0, 10);
        }
        
        proveedores.push({
          alias: row[0],
          nombre: row[1],
          cbu: row[2],
          telefono: row[3],
          email: row[4],
          deudaTotal: safeNum(row[5]),
          pagado: safeNum(row[6]),
          pendiente: safeNum(row[7]),
          pctPagado: safeNum(row[8]),
          estado: row[9],
          obs: obsVal,
          categoria: cat,
          prioridad: pri,
          fechaInicio: fechaInicio,
          condiciones: condiciones
        });
      }
    }

    // 2. LEER MOVIMIENTOS
    const sheetMov = doc.getSheetByName("Movimientos");
    const movimientos = [];
    if (sheetMov.getLastRow() >= 5) {
      const movRange = sheetMov.getRange(5, 1, sheetMov.getLastRow() - 4, 8);
      const movValues = movRange.getValues();
      for (let i = 0; i < movValues.length; i++) {
        const row = movValues[i];
        if (row[1]) { // Alias del proveedor
          movimientos.push({
            fecha: row[0] ? (row[0] instanceof Date ? row[0].toISOString().slice(0,10) : String(row[0]).slice(0,10)) : "",
            provAlias: row[1],
            provNombre: row[2],
            tipo: row[3] === "Pago total" ? "total" : "parcial",
            monto: safeNum(row[4]),
            pendiente: safeNum(row[5]),
            comprobante: row[6],
            obs: row[7]
          });
        }
      }
    }

    // 3. LEER COMPROBANTES
    const sheetComp = doc.getSheetByName("Comprobantes");
    const comprobantes = [];
    if (sheetComp.getLastRow() >= 5) {
      const compRange = sheetComp.getRange(5, 1, sheetComp.getLastRow() - 4, 7);
      const compValues = compRange.getValues();
      for (let i = 0; i < compValues.length; i++) {
        const row = compValues[i];
        if (row[1]) { // Alias del proveedor
          comprobantes.push({
            fecha: row[0] ? (row[0] instanceof Date ? row[0].toISOString().slice(0,10) : String(row[0]).slice(0,10)) : "",
            provAlias: row[1],
            monto: safeNum(row[2]),
            tipo: row[3],
            name: row[4],
            url: row[5],
            obs: row[6]
          });
        }
      }
    }

    return jsonResponse({
      status: "success",
      proveedores: proveedores,
      movimientos: movimientos,
      comprobantes: comprobantes
    });

  } catch (error) {
    return jsonResponse({ status: "error", message: "Error al leer planilla: " + error.toString() });
  }
}

/**
 * ENDPOINT POST: Recibe y procesa acciones de edición / registro
 */
function doPost(e) {
  if (!checkAuth(e)) {
    return jsonResponse({ status: "error", message: "No autorizado. Token incorrecto." });
  }

  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "add_provider") {
      const sheet = doc.getSheetByName("Proveedores");
      const nextRow = sheet.getLastRow() + 1;
      const sep = getFormulaSeparator();
      
      ensureExtraHeaders(sheet);
      
      // Creamos la fila
      sheet.getRange(nextRow, 1).setValue(data.alias); // A
      sheet.getRange(nextRow, 2).setValue(data.nombre || ""); // B
      sheet.getRange(nextRow, 3).setValue(data.cbu || ""); // C
      sheet.getRange(nextRow, 4).setValue(data.telefono || ""); // D
      sheet.getRange(nextRow, 5).setValue(data.email || ""); // E
      sheet.getRange(nextRow, 6).setValue(Number(data.deudaTotal || 0)); // F
      
      // Fórmulas automáticas de Excel (usando el separador de la planilla)
      sheet.getRange(nextRow, 7).setFormula(`=IFERROR(SUMIF(Movimientos!$B$5:$B$10000${sep} A${nextRow}${sep} Movimientos!$E$5:$E$10000)${sep} 0)`); // G (Pagado)
      sheet.getRange(nextRow, 8).setFormula(`=IFERROR(F${nextRow}-G${nextRow}${sep} 0)`); // H (Pendiente)
      sheet.getRange(nextRow, 9).setFormula(`=IFERROR(IF(F${nextRow}>0${sep} G${nextRow}/F${nextRow}${sep} 1)${sep} 1)`); // % Pagado
      sheet.getRange(nextRow, 10).setFormula(`=IFERROR(IF(H${nextRow}<=0${sep} "✅ Al día"${sep} IF(G${nextRow}>0${sep} "🟡 Deuda parcial"${sep} "🔴 Con deuda"))${sep} "")`); // J (Estado)
      sheet.getRange(nextRow, 11).setValue(data.obs || ""); // K
      sheet.getRange(nextRow, 12).setValue(data.categoria || ""); // L
      sheet.getRange(nextRow, 13).setValue(data.prioridad || "normal"); // M
      sheet.getRange(nextRow, 14).setValue(data.fechaInicio || ""); // N
      sheet.getRange(nextRow, 15).setValue(data.condiciones || ""); // O
      
      return jsonResponse({ status: "success", message: "Proveedor creado con éxito." });
    }
    
    else if (action === "edit_provider") {
      const sheet = doc.getSheetByName("Proveedores");
      const lastRow = sheet.getLastRow();
      const aliases = sheet.getRange(6, 1, lastRow - 5, 1).getValues();
      let rowIdx = -1;
      
      for (let i = 0; i < aliases.length; i++) {
        if (aliases[i][0] === data.oldAlias) {
          rowIdx = i + 6;
          break;
        }
      }
      
      if (rowIdx === -1) {
        return jsonResponse({ status: "error", message: "Proveedor no encontrado." });
      }
      
      ensureExtraHeaders(sheet);
      
      // Actualizamos datos
      sheet.getRange(rowIdx, 1).setValue(data.alias);
      sheet.getRange(rowIdx, 2).setValue(data.nombre || "");
      sheet.getRange(rowIdx, 3).setValue(data.cbu || "");
      sheet.getRange(rowIdx, 4).setValue(data.telefono || "");
      sheet.getRange(rowIdx, 5).setValue(data.email || "");
      sheet.getRange(rowIdx, 6).setValue(Number(data.deudaTotal || 0));
      sheet.getRange(rowIdx, 11).setValue(data.obs || "");
      sheet.getRange(rowIdx, 12).setValue(data.categoria || "");
      sheet.getRange(rowIdx, 13).setValue(data.prioridad || "normal");
      sheet.getRange(rowIdx, 14).setValue(data.fechaInicio || "");
      sheet.getRange(rowIdx, 15).setValue(data.condiciones || "");
      
      // Si el alias cambió, actualizamos los movimientos asociados en la hoja Movimientos
      if (data.oldAlias !== data.alias) {
        const sheetMov = doc.getSheetByName("Movimientos");
        const lastMovRow = sheetMov.getLastRow();
        if (lastMovRow >= 5) {
          const movAliases = sheetMov.getRange(5, 2, lastMovRow - 4, 1);
          const movVals = movAliases.getValues();
          for (let m = 0; m < movVals.length; m++) {
            if (movVals[m][0] === data.oldAlias) {
              sheetMov.getRange(m + 5, 2).setValue(data.alias);
            }
          }
        }
      }
      
      return jsonResponse({ status: "success", message: "Proveedor modificado con éxito." });
    }
    
    else if (action === "delete_provider") {
      const sheet = doc.getSheetByName("Proveedores");
      const lastRow = sheet.getLastRow();
      const aliases = sheet.getRange(6, 1, lastRow - 5, 1).getValues();
      let rowIdx = -1;
      
      for (let i = 0; i < aliases.length; i++) {
        if (aliases[i][0] === data.alias) {
          rowIdx = i + 6;
          break;
        }
      }
      
      if (rowIdx === -1) {
        return jsonResponse({ status: "error", message: "Proveedor no encontrado." });
      }
      
      sheet.deleteRow(rowIdx);
      return jsonResponse({ status: "success", message: "Proveedor eliminado con éxito de la hoja." });
    }
    
    else if (action === "adjust_debt") {
      const sheet = doc.getSheetByName("Proveedores");
      const lastRow = sheet.getLastRow();
      const aliases = sheet.getRange(6, 1, lastRow - 5, 1).getValues();
      let rowIdx = -1;
      
      for (let i = 0; i < aliases.length; i++) {
        if (aliases[i][0] === data.alias) {
          rowIdx = i + 6;
          break;
        }
      }
      
      if (rowIdx === -1) {
        return jsonResponse({ status: "error", message: "Proveedor no encontrado." });
      }
      
      sheet.getRange(rowIdx, 6).setValue(Number(data.deudaTotal || 0));
      if (data.obs) {
        const currentObs = sheet.getRange(rowIdx, 11).getValue();
        sheet.getRange(rowIdx, 11).setValue((currentObs ? currentObs + "\n" : "") + data.obs);
      }
      
      return jsonResponse({ status: "success", message: "Deuda ajustada con éxito." });
    }
    
    else if (action === "add_payment") {
      const sheetMov = doc.getSheetByName("Movimientos");
      const nextMovRow = sheetMov.getLastRow() + 1;
      
      // Buscar el nombre completo del proveedor en la hoja de Proveedores
      const sheetProv = doc.getSheetByName("Proveedores");
      const aliases = sheetProv.getRange(6, 1, sheetProv.getLastRow() - 5, 2).getValues();
      let nombreProv = "";
      for (let i = 0; i < aliases.length; i++) {
        if (aliases[i][0] === data.provAlias) {
          nombreProv = aliases[i][1];
          break;
        }
      }
      
      // Subir archivos a Google Drive si vienen adjuntos
      let comprobanteLinks = [];
      if (data.files && data.files.length > 0) {
        const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
        for (let f = 0; f < data.files.length; f++) {
          const fileData = data.files[f];
          const fileUrl = uploadFileToDrive(fileData.base64, fileData.name, fileData.mimeType, folder);
          comprobanteLinks.push(fileUrl);
          
          // Registrar en la hoja de Comprobantes
          const sheetComp = doc.getSheetByName("Comprobantes");
          const nextCompRow = sheetComp.getLastRow() + 1;
          sheetComp.getRange(nextCompRow, 1).setValue(data.fecha);
          sheetComp.getRange(nextCompRow, 2).setValue(data.provAlias);
          sheetComp.getRange(nextCompRow, 3).setValue(Number(data.monto));
          sheetComp.getRange(nextCompRow, 4).setValue(fileData.mimeType.startsWith("image/") ? "Captura/Foto" : "Documento PDF");
          sheetComp.getRange(nextCompRow, 5).setValue(fileData.name);
          sheetComp.getRange(nextCompRow, 6).setValue(fileUrl);
          sheetComp.getRange(nextCompRow, 7).setValue(data.obs || "");
        }
      }
      
      // Registrar en la hoja Movimientos
      const sep = getFormulaSeparator();
      sheetMov.getRange(nextMovRow, 1).setValue(data.fecha); // A (Fecha)
      sheetMov.getRange(nextMovRow, 2).setValue(data.provAlias); // B (Alias)
      sheetMov.getRange(nextMovRow, 3).setValue(nombreProv); // C (Nombre)
      sheetMov.getRange(nextMovRow, 4).setValue(data.tipo === "total" ? "Pago total" : "Pago parcial"); // D (Tipo)
      sheetMov.getRange(nextMovRow, 5).setValue(Number(data.monto)); // E (Monto)
      sheetMov.getRange(nextMovRow, 6).setFormula(`=IFERROR(VLOOKUP(B${nextMovRow}${sep} Proveedores!$A$6:$H$1000${sep} 8${sep} FALSE)${sep} "")`); // F (Pendiente después del pago)
      sheetMov.getRange(nextMovRow, 7).setValue(comprobanteLinks.length > 0 ? comprobanteLinks.join(" | ") : "Sin comprobante"); // G (Link Drive)
      sheetMov.getRange(nextMovRow, 8).setValue(data.obs || ""); // H
      
      return jsonResponse({ status: "success", message: "Pago registrado y comprobantes guardados." });
    }
    
    else if (action === "add_document") {
      const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
      const fileUrls = [];
      
      if (data.files && data.files.length > 0) {
        const sheetComp = doc.getSheetByName("Comprobantes");
        for (let f = 0; f < data.files.length; f++) {
          const fileData = data.files[f];
          const fileUrl = uploadFileToDrive(fileData.base64, fileData.name, fileData.mimeType, folder);
          fileUrls.push(fileUrl);
          
          const nextCompRow = sheetComp.getLastRow() + 1;
          sheetComp.getRange(nextCompRow, 1).setValue(data.fecha); // A
          sheetComp.getRange(nextCompRow, 2).setValue(data.provAlias); // B
          sheetComp.getRange(nextCompRow, 3).setValue(Number(data.monto || 0)); // C
          sheetComp.getRange(nextCompRow, 4).setValue(fileData.mimeType.startsWith("image/") ? "Captura/Foto" : "Documento PDF"); // D
          sheetComp.getRange(nextCompRow, 5).setValue(fileData.name); // E
          sheetComp.getRange(nextCompRow, 6).setValue(fileUrl); // F
          sheetComp.getRange(nextCompRow, 7).setValue(data.obs || ""); // G
        }
      }
      
      return jsonResponse({ status: "success", message: "Documento subido y registrado con éxito.", urls: fileUrls });
    }
    
    else if (action === "sync_queue") {
      // Sincroniza múltiples operaciones a la vez (cola offline)
      const results = [];
      const queue = data.queue;
      
      for (let i = 0; i < queue.length; i++) {
        const item = queue[i];
        // Simulamos una llamada recursiva simulada
        const mockE = {
          postData: { contents: JSON.stringify({ action: item.action, data: item.data }) },
          parameter: { token: SECRET_TOKEN }
        };
        const res = doPost(mockE);
        results.push({ index: i, response: JSON.parse(res.getContent()) });
      }
      
      return jsonResponse({ status: "success", results: results });
    }
    
    return jsonResponse({ status: "error", message: "Acción no reconocida." });

  } catch (error) {
    return jsonResponse({ status: "error", message: "Error en procesamiento POST: " + error.toString() });
  }
}

/**
 * Obtiene o crea la carpeta especificada en Google Drive
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    const folder = DriveApp.createFolder(folderName);
    // Cambiar permisos para que cualquiera con el link pueda ver (necesario para descargar desde la web app)
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  }
}

/**
 * Sube un archivo en Base64 a Drive y retorna su URL
 */
function uploadFileToDrive(base64Data, filename, mimeType, folder) {
  // Extraer el contenido base64 puro (removiendo el header data:*/*;base64,)
  const base64Clean = base64Data.split(",")[1] || base64Data;
  const decoded = Utilities.base64Decode(base64Clean);
  const blob = Utilities.newBlob(decoded, mimeType, filename);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

/**
 * Detecta el separador de argumentos de fórmulas de acuerdo a la configuración regional de la planilla.
 * Las planillas en regiones que usan coma como separador decimal (como Argentina o España) usan punto y coma (;).
 */
function getFormulaSeparator() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const locale = ss.getSpreadsheetLocale();
    const lang = locale.split('_')[0];
    const semicolonLanguages = ['de', 'fr', 'it', 'es', 'pt', 'tr', 'ru', 'pl', 'nl', 'sv', 'no', 'da', 'fi'];
    if (semicolonLanguages.indexOf(lang) !== -1) {
      return ';';
    }
  } catch (e) {
    // fallback
  }
  return ',';
}

/**
 * Corrige y regenera las fórmulas de la hoja en caso de que contengan errores o separadores incorrectos.
 * Esto se ejecuta automáticamente al iniciar doGet para sanear la planilla.
 */
function ensureCorrectFormulas(doc) {
  try {
    const sheetProv = doc.getSheetByName("Proveedores");
    const lastRowProv = sheetProv.getLastRow();
    if (lastRowProv < 6) return;
    
    const sep = getFormulaSeparator();
    
    // Comprobamos si la primera fila tiene error o fórmulas rotas
    const rangeG = sheetProv.getRange(6, 7, lastRowProv - 5, 1);
    const valuesG = rangeG.getValues();
    
    let needsFix = false;
    for (let i = 0; i < valuesG.length; i++) {
      const valStr = String(valuesG[i][0]);
      if (valStr.indexOf("#ERROR") !== -1 || valStr === "NaN") {
        needsFix = true;
        break;
      }
    }
    
    if (needsFix) {
      // Regeneramos las fórmulas con el separador regional correcto
      for (let r = 6; r <= lastRowProv; r++) {
        sheetProv.getRange(r, 7).setFormula(`=IFERROR(SUMIF(Movimientos!$B$5:$B$10000${sep} A${r}${sep} Movimientos!$E$5:$E$10000)${sep} 0)`);
        sheetProv.getRange(r, 8).setFormula(`=IFERROR(F${r}-G${r}${sep} 0)`);
        sheetProv.getRange(r, 9).setFormula(`=IFERROR(IF(F${r}>0${sep} G${r}/F${r}${sep} 1)${sep} 1)`);
        sheetProv.getRange(r, 10).setFormula(`=IFERROR(IF(H${r}<=0${sep} "✅ Al día"${sep} IF(G${r}>0${sep} "🟡 Deuda parcial"${sep} "🔴 Con deuda"))${sep} "")`);
      }
    }
    
    // Corregimos la hoja Movimientos
    const sheetMov = doc.getSheetByName("Movimientos");
    const lastRowMov = sheetMov.getLastRow();
    if (lastRowMov >= 5) {
      const rangeF = sheetMov.getRange(5, 6, lastRowMov - 4, 1);
      const valuesF = rangeF.getValues();
      let needsMovFix = false;
      for (let i = 0; i < valuesF.length; i++) {
        const valStr = String(valuesF[i][0]);
        if (valStr.indexOf("#ERROR") !== -1 || valStr === "NaN") {
          needsMovFix = true;
          break;
        }
      }
      if (needsMovFix) {
        for (let r = 5; r <= lastRowMov; r++) {
          sheetMov.getRange(r, 6).setFormula(`=IFERROR(VLOOKUP(B${r}${sep} Proveedores!$A$6:$H$1000${sep} 8${sep} FALSE)${sep} "")`);
        }
      }
    }
  } catch (e) {
    // Evitar que un error en la corrección bloquee toda la lectura
  }
}

/**
 * Asegura que existan los encabezados para las nuevas columnas L, M, N, O en la hoja de Proveedores
 */
function ensureExtraHeaders(sheet) {
  try {
    const maxCols = sheet.getMaxColumns();
    if (maxCols < 15) {
      sheet.insertColumnsAfter(maxCols, 15 - maxCols);
    }
    const headerRow = 5;
    if (sheet.getRange(headerRow, 12).getValue() !== "Categoría") {
      sheet.getRange(headerRow, 12).setValue("Categoría");
    }
    if (sheet.getRange(headerRow, 13).getValue() !== "Prioridad") {
      sheet.getRange(headerRow, 13).setValue("Prioridad");
    }
    if (sheet.getRange(headerRow, 14).getValue() !== "Fecha inicio") {
      sheet.getRange(headerRow, 14).setValue("Fecha inicio");
    }
    if (sheet.getRange(headerRow, 15).getValue() !== "Condiciones") {
      sheet.getRange(headerRow, 15).setValue("Condiciones");
    }
  } catch (e) {
    // Evitar que falle el proceso principal si hay problemas con los encabezados
  }
}

/**
 * Decodifica el bloque META heredado de las observaciones
 */
function decodeMetaObs(obs) {
  if (!obs) return { meta: {}, text: '' };
  const match = obs.match(/^\[META:(.*?)\]\n?([\s\S]*)$/);
  if (match) {
    try { return { meta: JSON.parse(match[1]), text: (match[2] || '').trim() }; }
    catch (e) { return { meta: {}, text: obs }; }
  }
  return { meta: {}, text: obs };
}

