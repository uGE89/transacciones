// --- CONFIGURACIÓN GENERAL ---

const ENDPOINT_ALEGRA = "https://api.alegra.com/api/v1/payments";

// --- CORRECCIÓN 1: El token ya no incluye la palabra "Basic " ---
const TOKEN_ALEGRA = "Y2FybG9zdWdlMzAwODg5QHlhaG9vLmNvbTphNTRlZWY4ODQxN2YyMTg4ZWE4MA=="; 
const MODO_TEST_API_ALEGRA = false; // Cambiar a false para modo producción

// --- FUNCIÓN PRINCIPAL ---
/**
 * Registra pagos en Alegra. Opera en modo "lote" o "específico".
 * @param {string|number|null} idEspecifico - (Opcional) El ID de la única transacción a procesar.
 * @returns {object|undefined} Si se procesa un ID específico, devuelve el resultado.
 */
function registrarPagosDesdeSheet(idEspecifico = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_TRANSACCIONES);
  if (!hoja) throw new Error(`No se encontró la hoja '${HOJA_TRANSACCIONES}'.`);

  const datos = hoja.getDataRange().getValues();
  const encabezados = datos[0].map(h => String(h).toLowerCase().trim());

  const campo = (nombre) => encabezados.indexOf(nombre.toLowerCase());

  // Definición de todas las columnas necesarias
  const colId = campo("Id");
  const colIdAlegra = campo("ID_Alegra");
  const colEstadoInterno = campo("Estado");
  const colRegistradoAlegra = campo("Registrado");
  const colDate = campo("date");
  const colType = campo("type");
  const colQuantity = campo("quantity");
  const colBank = campo("bankaccountid");
  const colClient = campo("client");
  const colCategory = campo("category");
  const colAnotation = campo("anotation");
  const colObservations = campo("observations");

  const filasProcesadas = [];
  const errores = [];
  const anotacionesProcesadas = new Set();

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const filaNum = i + 1;
    const idActual = fila[colId];

    // --- Lógica de filtrado ---
    if (idEspecifico && String(idActual) !== String(idEspecifico)) {
      continue;
    }

    // --- Reglas para el modo lote ---
    if (!idEspecifico) {
      if (fila[colRegistradoAlegra] && fila[colRegistradoAlegra].toString().toLowerCase() === "ok") continue;
      if (fila[colEstadoInterno].toLowerCase() !== "aprobada") continue; // Solo procesa las aprobadas
    }
    
    // --- Lógica de validación ---
    const anotation = fila[colAnotation] || "";
    if (!idEspecifico && anotation && anotacionesProcesadas.has(anotation)) {
      hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue("Duplicado");
      continue;
    }
    if (anotation) anotacionesProcesadas.add(anotation);

    const cantidad = parseFloat(fila[colQuantity]);
    const bankAccountId = parseInt(fila[colBank]);
    const clientId = parseInt(fila[colClient]);
    const categoryId = parseInt(fila[colCategory]);

    if (!fila[colDate] || isNaN(bankAccountId) || isNaN(clientId) || isNaN(categoryId) || isNaN(cantidad)) {
      hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue("Datos inválidos");
      const errorMsg = "Campos requeridos vacíos o inválidos";
      if (idEspecifico) return { exito: false, mensaje: errorMsg };
      errores.push({ fila: filaNum, status: "Validación", msg: errorMsg });
      continue;
    }
    
    const payload = {
        date: Utilities.formatDate(new Date(fila[colDate]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        type: fila[colType],
        paymentMethod: "transfer",
        bankAccount: { id: bankAccountId },
        client: { id: clientId },
        anotation: anotation,
        observations: fila[colObservations] || "",
        categories: [{ id: categoryId, quantity: cantidad, price: 1 }]
    };
    
    if (MODO_TEST_API_ALEGRA) {
      console.log(`TEST - Enviar fila ${filaNum}:`, JSON.stringify(payload, null, 2));
      hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue("TEST");
      if (idEspecifico) return { exito: true, mensaje: "Ejecución en modo TEST." };
      filasProcesadas.push(filaNum);
      continue;
    }

    try {
      // --- CORRECCIÓN 2: La llamada a UrlFetchApp se construye correctamente ---
      const options = {
        'method': 'post',
        'headers': {
          'Authorization': 'Basic ' + TOKEN_ALEGRA,
          'Content-Type': 'application/json'
        },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };
      const response = UrlFetchApp.fetch(ENDPOINT_ALEGRA, options);
      const resCode = response.getResponseCode();
      const resText = response.getContentText();

      if (resCode === 200 || resCode === 201) {
        const json = JSON.parse(resText);
        hoja.getRange(filaNum, colIdAlegra + 1).setValue(json.id || "");
        hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue("OK");
        if (idEspecifico) return { exito: true, idAlegra: json.id };
        filasProcesadas.push(filaNum);
      } else {
        hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue(`Error ${resCode}`);
        if (idEspecifico) return { exito: false, mensaje: `Error de Alegra (${resCode}): ${resText}` };
        errores.push({ fila: filaNum, status: resCode, msg: resText });
      }
    } catch (err) {
      hoja.getRange(filaNum, colRegistradoAlegra + 1).setValue("Excepción");
      if (idEspecifico) return { exito: false, mensaje: `Excepción al llamar a Alegra: ${err.message}` };
      errores.push({ fila: filaNum, status: "Exception", msg: err.message });
    }
    
    if (idEspecifico) break;
  }

  // --- Resumen final para el modo lote ---
  if (!idEspecifico) {
    console.log("\n=== RESUMEN DE PROCESAMIENTO EN LOTE ===");
    console.log(`Filas procesadas exitosamente: ${filasProcesadas.length}`);
    if (errores.length > 0) {
      console.log(`Errores encontrados: ${errores.length}`);
      errores.forEach(err => console.log(` - Fila ${err.fila}: [${err.status}] ${err.msg}`));
    } else {
      console.log("Sin errores.");
    }
  }

  if (idEspecifico) {
    return { exito: false, mensaje: "No se encontró la transacción para procesar." };
  }
}

/**
 * Función de ayuda para ejecutar el proceso en lote desde el editor de Apps Script.
 */
function test_registrarPagosEnLote() {
  registrarPagosDesdeSheet();
}
