// --- CONFIGURACIÓN GLOBAL ---
// Asegúrate de que estas constantes estén disponibles globalmente (p. ej., en un archivo Global.gs)

// MODO_TEST_Copia: Si es 'true', el script escribirá en hojas con el sufijo "_Test" 
// (ej. "Transacciones_Test") para no afectar los datos de producción.
const MODO_TEST_Copia = false; 

// MAX_RETRIES: Número máximo de veces que se reintentará una llamada fallida a la API de OpenAI.


// --- PROCESO PRINCIPAL ---
function procesarImagenesNuevas() {
  // --- SECCIÓN DE DEPURACIÓN INICIAL ---
  Logger.log("--- INICIANDO procesarImagenesNuevas ---");
  let contadorNuevas = 0; // Contador para saber cuántas imágenes se procesaron
  let contadorSaltadas = 0; // Contador para las ya revisadas

  try {
    const sufijoTest = MODO_TEST_Copia ? '_Test' : '';
    const hojaTransacciones = SpreadsheetApp.getActive().getSheetByName(HOJA_TRANSACCIONES + sufijoTest);
    const hojaRevisadas = SpreadsheetApp.getActive().getSheetByName(HOJA_REVISADAS + sufijoTest);
    const CONFIG = cargarConfiguracion();
    
    Logger.log(`Modo Test: ${MODO_TEST_Copia}. Usando sufijo: '${sufijoTest}'`);

    const folder = DriveApp.getFolderById(getConfig('FOLDER_ID'));
    const processedFolder = DriveApp.getFolderById(getConfig('PROCESSED_FOLDER_ID'));
    const archivos = folder.getFiles();

    // Carga de datos existentes
    const nombresRevisados = hojaRevisadas.getLastRow() > 1
      ? hojaRevisadas.getRange(2, 1, hojaRevisadas.getLastRow() - 1).getValues().flat()
      : [];
    Logger.log(`Se encontraron ${nombresRevisados.length} archivos ya revisados.`);

    const anotacionesRegistradas = hojaTransacciones.getLastRow() > 1
      ? hojaTransacciones.getRange(2, 9, hojaTransacciones.getLastRow() - 1, 1).getValues().flat().map(val => String(val).trim())
      : [];
    Logger.log(`Se encontraron ${anotacionesRegistradas.length} anotaciones ya registradas.`);
    
    // Cálculo del correlativo
    let proximoCorrelativo = 1;
    const ultimaFila = hojaTransacciones.getLastRow();
    if (ultimaFila >= 2) {
      const datos = hojaTransacciones.getRange(2, 16, ultimaFila - 1, 1).getValues().flat().map(Number).filter(n => Number.isInteger(n) && n > 0);
      if (datos.length) proximoCorrelativo = Math.max(...datos) + 1;
    }
    Logger.log(`El próximo correlativo a usar es: ${proximoCorrelativo}`);

    if (!archivos.hasNext()) {
        Logger.log("No se encontraron archivos en la carpeta de Drive.");
        return "No se encontraron imágenes en la carpeta."; // <-- Devolvemos un mensaje claro
    }

    // --- BUCLE PRINCIPAL DE PROCESAMIENTO ---
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();
      const urlImagenRef = `https://drive.google.com/uc?id=${archivo.getId()}`;
      const fechaRevision = new Date();

      if (nombresRevisados.includes(nombreArchivo)) {
        // Logger.log(`SALTADO (ya revisado): ${nombreArchivo}`); // Descomenta si quieres ver todos los saltados
        contadorSaltadas++;
        continue;
      }

      Logger.log(`--- Procesando archivo nuevo: ${nombreArchivo} ---`);

      try {
        // 1) Extraer y formatear la línea OCR
        Logger.log("Paso 1: Llamando a OpenAI para extraer texto...");
        let line = extraerDesdeOpenAI(archivo);
        Logger.log(`Respuesta cruda de OpenAI: "${line}"`);

        if (!line) {
          logError(nombreArchivo, fechaRevision, 'Respuesta vacía o malformada de OpenAI', 'Null');
          continue;
        }
        
        line = line.replace(/^\|+|\|+$/g, '');
        const parts = line.split('|').map(p => p.trim());
        
        if (parts.length < 7) {
          logError(nombreArchivo, fechaRevision, `Línea incompleta de OpenAI (${parts.length} partes)`, line);
          continue;
        }
        let [date, type, qtyStr, currency, bankIdStr, anotation, observations] = parts;
        Logger.log(`Paso 2: Texto parseado correctamente. Anotación: ${anotation}`);

        // 2) Convertir y ajustar campos
        const rawAmount = parseFloat(qtyStr) || 0;
        let quantity = rawAmount;
        if (/^USD|U\$/i.test(currency)) {
          const tipoCambio = Number(CONFIG.tipoCambio) || 1;
          quantity = parseFloat((rawAmount * tipoCambio).toFixed(2));
          Logger.log(`Conversión de moneda: ${rawAmount} ${currency} -> ${quantity} (TC: ${tipoCambio})`);
        }
        const bankAccountId = Number(bankIdStr) || null;

        // 3) Duplicados
        if (anotation && anotacionesRegistradas.includes(anotation)) {
          hojaRevisadas.appendRow([nombreArchivo, fechaRevision, 'duplicado (anotación)']);
          Logger.log(`SALTADO (anotación duplicada): ${nombreArchivo} - ${anotation}`);
          continue;
        }

        // 4) Insertar fila
        const fila = [
          nombreArchivo, date, type, 'transfer', quantity, bankAccountId,
          null, null, anotation, observations, 'Pendiente', urlImagenRef,
          null, null, null, proximoCorrelativo
        ];
        Logger.log("Paso 3: Fila preparada para insertar: " + JSON.stringify(fila));
        
        hojaTransacciones.appendRow(fila);
        hojaRevisadas.appendRow([nombreArchivo, fechaRevision, 'Pendiente']);
        // Mover el archivo a la carpeta de procesados
        processedFolder.addFile(archivo);
        folder.removeFile(archivo);

        if (anotation) anotacionesRegistradas.push(anotation);

        Logger.log(`✅ ÉXITO: Archivo "${nombreArchivo}" procesado e insertado.`);
        contadorNuevas++;
        proximoCorrelativo++;
      } catch (err) {
        Logger.log(`‼️ ERROR procesando "${nombreArchivo}": ${err.message}`);
        logError(nombreArchivo, fechaRevision, 'Error general', err.message || err);
      }
    }
    
    // --- SECCIÓN DE DEPURACIÓN FINAL ---
    Logger.log("--- Bucle de archivos finalizado ---");
    Logger.log(`Resumen: ${contadorNuevas} imágenes nuevas procesadas, ${contadorSaltadas} ya estaban revisadas.`);

    // Devolvemos un resumen claro a la función principal
    return `${contadorNuevas} imágenes nuevas fueron procesadas.`;

  } catch(e) {
    Logger.log(`‼️ ERROR FATAL en procesarImagenesNuevas: ${e.stack}`);
    return `Error fatal en el proceso: ${e.message}`; // Devolvemos el error
  }
}



// --- FUNCIONES DE PROCESAMIENTO (ACTUALIZADA) ---
function logError(nombreArchivo, fecha, mensaje, detalleError) {
  // Se determina el sufijo de prueba dentro de la función también.
  const sufijoTest = MODO_TEST_Copia ? '_Test' : '';
  const hojaErrores = SpreadsheetApp.getActive().getSheetByName(HOJA_ERRORES + sufijoTest);
  const hojaRevisadas = SpreadsheetApp.getActive().getSheetByName(HOJA_REVISADAS + sufijoTest);

  // Si el detalle es un objeto, lo convierte a JSON. Si es una cadena, lo usa directamente.
  const detalleParaLog = (typeof detalleError === 'object' && detalleError !== null) ? JSON.stringify(detalleError) : detalleError;

  hojaErrores.appendRow([nombreArchivo, fecha, mensaje, detalleParaLog]);
  hojaRevisadas.appendRow([nombreArchivo, fecha, "falló"]);
  Logger.log(`ERROR: ${nombreArchivo} - ${mensaje} - Detalle: ${detalleParaLog}`);
}


// --- UTILIDADES ---
function cargarConfiguracion() {
  // Se asume que la hoja de configuración no tiene una versión de prueba
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_CONFIG);
  const config = {};
  if (!hoja || hoja.getLastRow() < 2) return config; 
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
  datos.forEach(([clave, valor]) => { 
    if (clave) { 
      config[clave.toString().trim()] = valor; 
    }
  });
  return config;
}

function getConfig(clave) {
  return PropertiesService.getScriptProperties().getProperty(clave);
}

function setConfig(clave, valor) {
  PropertiesService.getScriptProperties().setProperty(clave, valor);
}


// --- EXTRAER DESDE OPENAI (VERSIÓN ROBUSTA CON REINTENTOS) ---
function extraerDesdeOpenAI(archivo) {
  const OPENAI_API_KEY = getConfig("OPENAI_API_KEY");
  if (!OPENAI_API_KEY) {
    Logger.log("ERROR: OPENAI_API_KEY no está configurada.");
    return null;
  }

  const blob = archivo.getBlob();
  if (!blob.getContentType().startsWith("image/")) {
    Logger.log("ERROR: El archivo no es una imagen.");
    return null;
  }
  const imageBase64 = Utilities.base64Encode(blob.getBytes());
  const mimeType = blob.getContentType();



  const payload = {
    model: "gpt-4o",
    messages: [
      { role: "system", content: SYSTEM_PROMPT_OCR },
      {
        role: "user",
        content: [
          { type: "image_url", image_url: { url: `data:${mimeType};base64,${imageBase64}`, detail: "auto" } }
        ]
      }
    ],
    temperature: 0,
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // Muy importante para manejar los errores manualmente
  };

  // --- INICIO DE LA LÓGICA DE REINTENTOS ---
  for (let i = 0; i < MAX_RETRIES; i++) {
    const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
    const code = res.getResponseCode();
    const txt = res.getContentText();

    if (code === 200) {
      // Éxito: procesamos la respuesta y salimos de la función
      const response = JSON.parse(txt);
      const line = response.choices?.[0]?.message?.content?.trim();
      if (!line) {
        Logger.log("❌ No se obtuvo texto en response.content aunque el código fue 200. Respuesta: " + txt);
        return null;
      }
      Logger.log("✅ Línea OCR formateada para " + archivo.getName() + ": " + line);
      return line;
    }
    
    if (code === 429) {
      // Error de Límite de Tasa: esperamos y reintentamos
      Logger.log(`⚠️ Rate Limit (429) en intento ${i + 1} para el archivo ${archivo.getName()}. Esperando para reintentar...`);
      const errorResponse = JSON.parse(txt);
      const errorMessage = errorResponse.error.message || "";
      
      const waitTimeMatch = errorMessage.match(/try again in ([\d.]+)s/);
      let waitMilliseconds = 0;

      if (waitTimeMatch && waitTimeMatch[1]) {
        // Usamos el tiempo sugerido por la API + un pequeño búfer (500ms)
        waitMilliseconds = parseFloat(waitTimeMatch[1]) * 1000 + 500;
        Logger.log(`   La API sugiere esperar ${waitMilliseconds / 1000} segundos.`);
      } else {
        // Si no, usamos un backoff exponencial
        waitMilliseconds = Math.pow(2, i + 1) * 1000; // 2s, 4s, 8s...
        Logger.log(`   La API no sugirió tiempo. Usando backoff exponencial: ${waitMilliseconds / 1000} segundos.`);
      }
      Utilities.sleep(waitMilliseconds);

    } else {
      // Otro tipo de error (ej. 500, 400, etc.)
      Logger.log(`❌ OpenAI API error ${code} (intento ${i + 1}/${MAX_RETRIES}) para ${archivo.getName()}: ${txt}`);
      Utilities.sleep(Math.pow(2, i) * 1000); // 1s, 2s, 4s...
    }
  }

  Logger.log(`❌ FALLO PERMANENTE para el archivo ${archivo.getName()} después de ${MAX_RETRIES} intentos.`);
  return null;
}


// ---- FUNCIÓN DE PRUEBA ----
function testBase64Encoding() {
  try {
    const TEST_FILE_ID = "1D0ti_Nt5QeHwvNvgQjXruEfz9tIgOKr-"; 
    const file = DriveApp.getFileById(TEST_FILE_ID);
    Logger.log("Iniciando prueba de codificación Base64 para el archivo: " + file.getName());
    const imageBytes = file.getBlob().getBytes();
    const imageBase64 = Utilities.base64Encode(imageBytes); 
    Logger.log("✅ ¡ÉXITO! La codificación Base64 funcionó correctamente.");
    Logger.log("Los primeros 50 caracteres del resultado son: " + imageBase64.substring(0, 50));
  } catch (e) {
    Logger.log("❌ ERROR DURANTE LA PRUEBA: " + e.toString());
    Logger.log("Stack del error: " + e.stack);
  }
}
