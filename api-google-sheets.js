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
 * ENDPOINT GET: Obtiene toda la información de la planilla
 */
function doGet(e) {
  if (!checkAuth(e)) {
    return jsonResponse({ status: "error", message: "No autorizado. Token incorrecto." });
  }

  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. LEER PROVEEDORES
    const sheetProv = doc.getSheetByName("Proveedores");
    const provRange = sheetProv.getRange(6, 1, sheetProv.getLastRow() - 5, 11);
    const provValues = provRange.getValues();
    const proveedores = [];
    
    for (let i = 0; i < provValues.length; i++) {
      const row = provValues[i];
      if (row[0]) { // Verificar que exista alias
        proveedores.push({
          alias: row[0],
          nombre: row[1],
          cbu: row[2],
          telefono: row[3],
          email: row[4],
          deudaTotal: Number(row[5] || 0),
          pagado: Number(row[6] || 0),
          pendiente: Number(row[7] || 0),
          pctPagado: Number(row[8] || 0),
          estado: row[9],
          obs: row[10]
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
            monto: Number(row[4] || 0),
            pendiente: Number(row[5] || 0),
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
            monto: Number(row[2] || 0),
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
      
      // Creamos la fila
      sheet.getRange(nextRow, 1).setValue(data.alias); // A
      sheet.getRange(nextRow, 2).setValue(data.nombre || ""); // B
      sheet.getRange(nextRow, 3).setValue(data.cbu || ""); // C
      sheet.getRange(nextRow, 4).setValue(data.telefono || ""); // D
      sheet.getRange(nextRow, 5).setValue(data.email || ""); // E
      sheet.getRange(nextRow, 6).setValue(Number(data.deudaTotal || 0)); // F
      
      // Fórmulas automáticas de Excel
      sheet.getRange(nextRow, 7).setFormula(`=IFERROR(SUMIF(Movimientos!$B$5:$B$10000, A${nextRow}, Movimientos!$E$5:$E$10000), 0)`); // G (Pagado)
      sheet.getRange(nextRow, 8).setFormula(`=IFERROR(F${nextRow}-G${nextRow}, 0)`); // H (Pendiente)
      sheet.getRange(nextRow, 9).setFormula(`=IFERROR(IF(F${nextRow}>0, G${nextRow}/F${nextRow}, 1), 1)`); // % Pagado
      sheet.getRange(nextRow, 10).setFormula(`=IFERROR(IF(H${nextRow}<=0, "✅ Al día", IF(G${nextRow}>0, "🟡 Deuda parcial", "🔴 Con deuda")), "")`); // J (Estado)
      sheet.getRange(nextRow, 11).setValue(data.obs || ""); // K
      
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
      
      // Actualizamos datos
      sheet.getRange(rowIdx, 1).setValue(data.alias);
      sheet.getRange(rowIdx, 2).setValue(data.nombre || "");
      sheet.getRange(rowIdx, 3).setValue(data.cbu || "");
      sheet.getRange(rowIdx, 4).setValue(data.telefono || "");
      sheet.getRange(rowIdx, 5).setValue(data.email || "");
      sheet.getRange(rowIdx, 6).setValue(Number(data.deudaTotal || 0));
      sheet.getRange(rowIdx, 11).setValue(data.obs || "");
      
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
      sheetMov.getRange(nextMovRow, 1).setValue(data.fecha); // A (Fecha)
      sheetMov.getRange(nextMovRow, 2).setValue(data.provAlias); // B (Alias)
      sheetMov.getRange(nextMovRow, 3).setValue(nombreProv); // C (Nombre)
      sheetMov.getRange(nextMovRow, 4).setValue(data.tipo === "total" ? "Pago total" : "Pago parcial"); // D (Tipo)
      sheetMov.getRange(nextMovRow, 5).setValue(Number(data.monto)); // E (Monto)
      sheetMov.getRange(nextMovRow, 6).setFormula(`=IFERROR(VLOOKUP(B${nextMovRow}, Proveedores!$A$6:$H$1000, 8, FALSE), "")`); // F (Pendiente después del pago)
      sheetMov.getRange(nextMovRow, 7).setValue(comprobanteLinks.length > 0 ? comprobanteLinks.join(" | ") : "Sin comprobante"); // G (Link Drive)
      sheetMov.getRange(nextMovRow, 8).setValue(data.obs || ""); // H
      
      return jsonResponse({ status: "success", message: "Pago registrado y comprobantes guardados." });
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
