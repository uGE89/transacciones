// =================================================================
// CONFIGURACIÓN GLOBAL
// =================================================================
// =================================================================
// FUNCIÓN PRINCIPAL - SERVIR LA APP
// =================================================================

/**
 * Sirve la interfaz de usuario principal (el visor de transacciones).
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Visor y Procesador de Transacciones')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// =================================================================
// FUNCIONES PARA OBTENER DATOS (LLAMADAS DESDE EL FRONTEND)
// =================================================================

/**
 * Obtiene y formatea todas las transacciones para el frontend.
 * Esta función es el corazón de la carga de datos.
 */
function obtenerTransacciones() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRANSACCIONES);
    if (!hoja) {
      return { error: `La hoja "${HOJA_TRANSACCIONES}" no existe.` };
    }
    
    let datosCompletos = hoja.getDataRange().getValues();
    if (datosCompletos.length <= 1) {
      return { encabezados: datosCompletos[0] || [], filas: [] };
    }
    
    // Extrae y normaliza los encabezados de la hoja (a minúsculas para evitar errores)
    const todosLosEncabezadosOriginales = datosCompletos.shift();
    const todosLosEncabezados = todosLosEncabezadosOriginales.map(h => String(h).trim().toLowerCase());

    // Mapeo de los nombres que necesita el frontend (claves) a los nombres REALES de tu hoja (valores).
    // Las claves aquí son las que el frontend usará para crear sus índices.
    const mapeoColumnas = {
      // Frontend Key (lo que el JS usará) : "Header en la Hoja de Cálculo"
      "origen": "origen",
      "id": "id",
      "date": "date",
      "type": "type",
      "quantity": "quantity",
      "anotation": "anotation",
      "observations": "observations",
      "bankAccountId": "bankaccountid",
      "clientId": "client",
      "categoryId": "category",
      "ID_Alegra": "id_alegra",
      "Estado": "estado",
      "DirecciónImagen": "dirección"
    };

    const encabezadosParaEnviar = Object.keys(mapeoColumnas);

    // Encuentra el índice de cada columna requerida en la hoja real
    const indicesEnHoja = encabezadosParaEnviar.map(key => {
      const nombreRealEnHoja = mapeoColumnas[key];
const indice = todosLosEncabezados.indexOf(nombreRealEnHoja.toLowerCase());
      if (indice === -1) {
        Logger.log(`Advertencia: La columna requerida "${nombreRealEnHoja}" (mapeada desde "${key}") no se encontró en la hoja.`);
      }
      return indice;
    });

    const indiceColumnaFecha = todosLosEncabezados.indexOf("date");
    
    // Formatea los datos, asegurándose de que cada fila tenga el orden correcto
    const datosFormateados = datosCompletos.map(filaEnteraOriginal => {
      return indicesEnHoja.map((indiceOriginal) => {
        if (indiceOriginal === -1) return null; // Si una columna no se encontró, devuelve null

        let celda = filaEnteraOriginal[indiceOriginal];
        
        // Formateo especial para fechas
        if (indiceOriginal === indiceColumnaFecha && celda instanceof Date) {
          return Utilities.formatDate(celda, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        return celda;
      });
    });
    
    return { encabezados: encabezadosParaEnviar, filas: datosFormateados };

  } catch (error) {
    Logger.log('Error en obtenerTransacciones: ' + error.stack);
    return { error: 'Error general en el servidor: ' + error.message };
  }
}


/**
 * Obtiene la lista de bancos para los menús desplegables.
 */
/**
 * Obtiene la lista de bancos para los menús desplegables.
 * Ahora también obtiene el color de la columna E.
 */
function getBancosLista() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_BANCOS);
    if (!hoja) throw new Error(`No se encontró la hoja '${HOJA_BANCOS}'.`);
    
    // Lee hasta la columna E para incluir el color
    const rango = hoja.getRange('A2:E' + hoja.getLastRow());
    const valores = rango.getValues();

    return valores.map(fila => {
        // fila[0] = ID, fila[1] = Nombre, fila[4] = Color
        if (fila[0] && fila[1]) {
          return {
            bankAccountId: fila[0],
            name: fila[1],
            color: fila[4] || '#78909c' // Usa el color de la columna E o un gris por defecto
          };
        }
        return null;
      })
      .filter(item => item !== null);

  } catch (e) {
    console.error('Error en getBancosLista:', e);
    return [];
  }
}

/**
 * Obtiene la lista de clientes.
 */
function getClientesLista() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CLIENTES); 
    if (!hoja) return [];
    const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues();
    return datos.map(fila => ({ id: fila[0], nombre: fila[1] })).filter(item => item.id && item.nombre);
  } catch (e) {
    Logger.log("Error en getClientesLista: " + e);
    return [];
  }
}

/**
 * Obtiene la lista de categorías.
 */
function getCategoriasLista() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CATEGORIAS);
    if (!hoja) return [];
    const ultimaFila = hoja.getLastRow();
    if (ultimaFila <= 1) return [];
    
    const datos = hoja.getRange(2, 1, ultimaFila - 1, 2).getValues();
    return datos.map(fila => ({ id: fila[0], nombre: fila[1] })).filter(item => item.id && item.nombre); 
  } catch (e) {
    Logger.log("Error en getCategoriasLista: " + e.stack);
    return [];
  }
}


// =================================================================
// FUNCIONES DE MODIFICACIÓN DE DATOS
// =================================================================

/**
 * Orquesta los procesos de actualización (correos, imágenes).
 */
function ejecutarActualizacionDepositos() {
  Logger.log("--- INICIANDO PROCESO GLOBAL DE ACTUALIZACIÓN ---");
  try {
    // CORREGIDO: Se llaman las funciones de procesamiento reales.
    // Asegúrate de que estas funciones existan en tu proyecto de Apps Script.
    const resultadoCorreos = procesarTransferenciasTextoLafise(); 
    const resultadoImagenes = procesarImagenesNuevas(); 

    const mensajeFinal = `Resumen de actualización:\n- Correos: ${resultadoCorreos}\n- Imágenes: ${resultadoImagenes}`;
    Logger.log("Proceso de actualización finalizado. " + mensajeFinal.replace(/\n/g, ' '));
    
    return mensajeFinal;
  } catch (err) {
    Logger.log("‼️ ERROR FATAL en el orquestador principal: " + err.stack);
    return "Error durante la actualización: " + err.message;
  }
}



/**
 * Actualiza el cliente y/o categoría de una transacción existente.
 */
function actualizarClienteCategoria(idTransaccion, nuevoClienteId, nuevaCategoriaId) {
  Logger.log(`Actualizando Tx ID: ${idTransaccion} -> Cliente ID: ${nuevoClienteId}, Categoría ID: ${nuevaCategoriaId}`);
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRANSACCIONES);
    if (!hoja) throw new Error(`La hoja "${HOJA_TRANSACCIONES}" no existe.`);

    // Normaliza los encabezados a minúsculas para una búsqueda robusta
    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());

    const indiceColId = encabezados.indexOf("id");
    const indiceColCliente = encabezados.indexOf("client");
    const indiceColCategoria = encabezados.indexOf("category");

    if (indiceColId === -1) throw new Error("No se encontró la columna 'Id'.");
    if (indiceColCliente === -1) throw new Error("No se encontró la columna 'client'.");
    if (indiceColCategoria === -1) throw new Error("No se encontró la columna 'category'.");

    const idsEnHoja = hoja.getRange(2, indiceColId + 1, hoja.getLastRow() - 1, 1).getValues();
    let filaEncontrada = -1;

    for (let i = 0; i < idsEnHoja.length; i++) {
      if (String(idsEnHoja[i][0]) === String(idTransaccion)) {
        filaEncontrada = i + 2;
        break;
      }
    }

    if (filaEncontrada !== -1) {
      Logger.log(`Fila encontrada en la hoja: ${filaEncontrada}. Actualizando...`);
      hoja.getRange(filaEncontrada, indiceColCliente + 1).setValue(nuevoClienteId || null);
      hoja.getRange(filaEncontrada, indiceColCategoria + 1).setValue(nuevaCategoriaId || null);
      SpreadsheetApp.flush();
      return "Cliente y Categoría actualizados con éxito.";
    } else {
      Logger.log(`Error: No se encontró la transacción con ID ${idTransaccion}.`);
      return `Error: No se encontró la transacción con ID ${idTransaccion}.`;
    }
  } catch (error) {
    Logger.log('Error en actualizarClienteCategoria: ' + error.stack);
    return 'Error en el servidor: ' + error.message;
  }
}

/**
 * Agrega una nueva fila de transacción a la hoja de cálculo.
 */
/**
 * Escribe un objeto de datos en una hoja de cálculo, mapeando las claves del objeto
 * a los encabezados de la hoja para asegurar el orden correcto de las columnas.
 * @param {Sheet} sheet El objeto de la hoja de cálculo donde se escribirá la fila.
 * @param {Object} datosFila Un objeto donde las claves coinciden (sin importar mayúsculas) con los encabezados de la hoja.
 */

function escribirFilaMapeada(sheet, datosFila) {
  const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headerMap = {};
  encabezados.forEach((header, index) => {
    if (header) {
      headerMap[header.toLowerCase().trim()] = index;
    }
  });

  const newRow = new Array(encabezados.length).fill('');
  for (const key in datosFila) {
    const lowerKey = key.toLowerCase().trim();
    if (headerMap.hasOwnProperty(lowerKey)) {
      const index = headerMap[lowerKey];
      newRow[index] = datosFila[key];
    }
  }
  
  sheet.appendRow(newRow);
}
// Reemplaza tu vieja función 'agregarNuevaTransaccion' por esta.
/**
 * Procesa y guarda una nueva transacción, detectando el tipo de operación 
 * (Simple, Transferencia, Préstamo) y actuando en consecuencia.
 * Reemplaza a la antigua 'agregarNuevaTransaccion'.
 * @param {Object} datosFormulario Un objeto con todos los datos del formulario del frontend.
 * @returns {string} Un mensaje de éxito o lanza un error.
 */
/**
 * Procesa y guarda una nueva transacción, detectando el tipo de operación 
 * (Simple, Transferencia, Préstamo) y actuando en consecuencia.
 * @param {Object} datosFormulario Un objeto con todos los datos del formulario del frontend.
 * @returns {string} Un mensaje de éxito o lanza un error.
 */

function generarIdTransaccion() {
  return 'TX-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
}


function procesarNuevaTransaccion(datosFormulario) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transaccionesSheet = ss.getSheetByName(HOJA_TRANSACCIONES);
    const controlPrestamosSheet = ss.getSheetByName(HOJA_CONTROL_PRESTAMOS);

    // Verificaciones de existencia de hojas...
    if (!transaccionesSheet) { throw new Error(`Error Crítico: No se pudo encontrar la hoja llamada '${HOJA_TRANSACCIONES}'.`); }
    if (datosFormulario.tipoOperacion === 'PRESTAMO' && !controlPrestamosSheet) { throw new Error(`Error Crítico: No se pudo encontrar la hoja llamada '${HOJA_CONTROL_PRESTAMOS}'.`); }

    let mensajeExito = "";

    switch (datosFormulario.tipoOperacion) {
      case 'TRANSFERENCIA':
        const idTransferencia = "TXF-" + new Date().getTime();
        let idSalida = obtenerSiguienteId(transaccionesSheet, 'Id');
        const datosSalida = {
          'Id': idSalida,
          'Date': datosFormulario.fecha,
          'Type': 'out',
          'Quantity': datosFormulario.monto,
          'Observations': `Transferencia enviada a ${datosFormulario.cuentaDestinoNombre}`,
          'BankAccountId': datosFormulario.cuentaId,
          'Client': datosFormulario.clienteId,
          'Anotation': datosFormulario.anotacion,
          'ID_Vinculado': idTransferencia,
          'Category': datosFormulario.categoriaId,
          'Origen': 'Registro manual',
          'Estado': 'Pendiente',
          'paymentMethod' : 'cash'
        };
        escribirFilaMapeada(transaccionesSheet, datosSalida);

        let idEntrada = obtenerSiguienteId(transaccionesSheet, 'Id');
        const datosEntrada = {
          'Id': idEntrada,
          'Date': datosFormulario.fecha,
          'Type': 'in',
          'paymentMethod' : 'cash',
          'Quantity': datosFormulario.monto,
          'Observations': `Transferencia recibida de ${datosFormulario.cuentaOrigenNombre}`,
          'BankAccountId': datosFormulario.cuentaDestinoId,
          'Client': datosFormulario.clienteId,
          'Anotation': datosFormulario.anotacion,
          'ID_Vinculado': idTransferencia,
          'Category': datosFormulario.categoriaId,
          'Origen': 'Registro manual',
          'Estado': 'Pendiente'
        };
        escribirFilaMapeada(transaccionesSheet, datosEntrada);

        mensajeExito = "Transferencia registrada con éxito.";
        break;

      case 'PRESTAMO':
        const idPrestamo = "PREST-" + new Date().getTime();
        let idDesembolso = obtenerSiguienteId(transaccionesSheet, 'Id');
        const montoTotal = parseFloat(datosFormulario.montoTotal);

        const datosDesembolso = {
          'Id': idDesembolso,
          'Date': datosFormulario.fecha,
          'Type': 'out',
          'paymentMethod' : 'cash',
          'Quantity': montoTotal,
          'Observations': `Otorgamiento de préstamo a ${datosFormulario.deudorNombre}`,
          'BankAccountId': datosFormulario.cuentaId,
          'Client': datosFormulario.deudorId,
          'Category': datosFormulario.categoriaId,
          'Anotation': datosFormulario.anotacion,
          'ID_Vinculado': idPrestamo,
          'Origen': 'Registro manual',
          'Estado': 'Pendiente'
        };
        escribirFilaMapeada(transaccionesSheet, datosDesembolso);

        const numCuotas = parseInt(datosFormulario.numCuotas, 10);
        const montoCuota = montoTotal / numCuotas;
        const primerPago = new Date(datosFormulario.fechaPrimerPago);
        for (let i = 0; i < numCuotas; i++) {
          const fechaVencimiento = new Date(primerPago);
          fechaVencimiento.setUTCMonth(fechaVencimiento.getUTCMonth() + i);
          const datosCuota = {
            'ID_Prestamo': idPrestamo,
            'ID_Cuota': `${idPrestamo}-${i + 1}`,
            'ID_Deudor': datosFormulario.deudorId,
            'Nombre_Deudor': datosFormulario.deudorNombre,
            'Monto_Total_Prestamo': montoTotal,
            'Cuota_Num': `${i + 1} / ${numCuotas}`,
            'Monto_Cuota': montoCuota.toFixed(2),
            'Fecha_Vencimiento': fechaVencimiento,
            'Estado': 'Pendiente',
            'Origen': 'Registro manual',
            'paymentMethod' : 'cash',
          };
          escribirFilaMapeada(controlPrestamosSheet, datosCuota);
        }

        mensajeExito = `Préstamo de ${numCuotas} cuotas registrado exitosamente.`;
        break;

      default: // Caso 'SIMPLE' (incluye abonos a préstamos)
        let idSimple = obtenerSiguienteId(transaccionesSheet, 'Id');
        const datosSimples = {
          'Id': idSimple,
          'Date': datosFormulario.fecha,
          'Type': datosFormulario.tipo,
          'Quantity': datosFormulario.monto,
          'Observations': datosFormulario.observaciones,
          'BankAccountId': datosFormulario.cuentaId,
          'Client': datosFormulario.clienteId,
          'Category': datosFormulario.categoriaId,
          'Anotation': datosFormulario.anotacion,
          'Origen': 'Registro manual',
          'Estado': 'Pendiente',
          'PaymentMethod': 'cash'
        };
        escribirFilaMapeada(transaccionesSheet, datosSimples);

        // --- BLOQUE CORREGIDO Y MOVIDO AQUÍ ---
        // Si este registro simple incluye un pago de cuotas, las marcamos como pagadas.
        if (datosFormulario.cuotasAbonar && datosFormulario.cuotasAbonar.length > 0) {
          marcarCuotasComoPagadas(datosFormulario.cuotasAbonar, idSimple);
          mensajeExito = "Abono a préstamo registrado y cuotas actualizadas.";
        } else {
          mensajeExito = "Transacción simple registrada.";
        }
        break;
    }

    SpreadsheetApp.flush();
    return mensajeExito;

  } catch (e) {
    Logger.log('Error en procesarNuevaTransaccion: ' + e.stack);
    throw new Error('Error en el servidor al procesar la operación: ' + e.message);
  }
}


/**
 * Rellena la columna "Id" con números secuenciales para mantener el orden.
 */
function numerarRegistros() {
  try {
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRANSACCIONES);
    if (!hoja) return;

    const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
    const indiceColId = encabezados.map(h => String(h).trim().toLowerCase()).indexOf("id");

    if (indiceColId === -1) {
      Logger.log("No se encontró la columna 'Id' para numerar.");
      return;
    }

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return; // No hay datos que numerar

    const rangoIds = hoja.getRange(2, indiceColId + 1, ultimaFila - 1, 1);
    const numeros = [];
    for (let i = 0; i < ultimaFila - 1; i++) {
      numeros.push([i + 1]); // Crea un array 2D para el método setValues
    }
    rangoIds.setValues(numeros);
    Logger.log(`Se numeraron ${numeros.length} registros en la columna 'Id'.`);
  } catch (e) {
    Logger.log('Error en numerarRegistros: ' + e.stack);
  }
}
/**
 * Calcula el siguiente ID correlativo para una hoja de cálculo,
 * basándose en el valor de la última fila.
 * @param {Sheet} sheet La hoja de la cual obtener el último ID.
 * @param {string} nombreColumnaId El nombre exacto del encabezado de la columna ID.
 * @returns {number} El siguiente ID secuencial.
 */
function obtenerSiguienteId(sheet, nombreColumnaId) {
  const lastRow = sheet.getLastRow();
  // Si no hay datos (solo encabezados), el primer ID es 1.
  if (lastRow < 2) {
    return 1;
  }
  
  // Encontrar el índice de la columna ID
  const encabezados = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = encabezados.map(h => h.toLowerCase().trim()).indexOf(nombreColumnaId.toLowerCase().trim());

  if (colIndex === -1) {
    throw new Error(`No se encontró la columna ID llamada '${nombreColumnaId}'.`);
  }

  // Obtener el valor de la última fila en la columna ID
  const ultimoId = sheet.getRange(lastRow, colIndex + 1).getValue();
  
  // Convertir a número y sumar 1. Si no es un número, empezar de 1.
  return !isNaN(ultimoId) && ultimoId > 0 ? parseInt(ultimoId) + 1 : 1;
}
/**
 * Aprueba una transacción validando un PIN.
 */


/**
 * Obtiene una imagen desde una URL y la devuelve como un objeto Base64.
 * Esto evita problemas de CORS y permisos en el navegador del cliente.
 * @param {string} imageUrl La URL de la imagen a obtener.
 * @return {object} Un objeto con los datos base64 y el tipo de contenido, o null si falla.
 */
function getImageAsBase64(imageUrl) {
  // Maneja URLs vacías o nulas.
  if (!imageUrl) {
    return null;
  }
  
  try {
    // Usamos el token de autorización del script para acceder a archivos de Drive si es necesario.
    const response = UrlFetchApp.fetch(imageUrl, {
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      // Ignora respuestas inválidas del certificado SSL, útil para algunas URLs.
      validateHttpsCertificates: false 
    });

    const imageBlob = response.getBlob();
    const contentType = imageBlob.getContentType();
    const base64Data = Utilities.base64Encode(imageBlob.getBytes());

    // Devuelve un objeto con todo lo que el cliente necesita.
    return {
      base64Data: base64Data,
      contentType: contentType
    };
  } catch (e) {
    console.error("Error al obtener la imagen como Base64 para la URL: " + imageUrl + ". Error: " + e.toString());
    return null; // Devuelve null en caso de error.
  }
}

/**
 * Lee la hoja de cálculo "Plantillas" y devuelve una lista de objetos, 
 * donde cada objeto representa una plantilla de transacción.
 * @returns {Array<Object>} Un array con las plantillas disponibles.
 */
function obtenerPlantillas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Plantillas");
    if (!sheet) {
      throw new Error("La hoja 'Plantillas' no fue encontrada.");
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data.shift(); // Saca la primera fila (encabezados)
    
    // Convertir las filas en objetos para un manejo más fácil
    const plantillas = data.map(row => {
      const plantillaObj = {};
      headers.forEach((header, index) => {
        plantillaObj[header] = row[index];
      });
      return plantillaObj;
    });
    
    return plantillas;
  } catch (e) {
    console.error("Error en obtenerPlantillas: " + e.toString());
    return []; // Devuelve un array vacío en caso de error
  }
}

function obtenerCuotasPorVencer(fecha) {
  // Lógica para leer la hoja "Control de Préstamos", filtrar por fecha y estado "Pendiente", y devolver los resultados.
}

function registrarPagoAgrupado(listaDeIDsDeCuota, datosDelPago) {
  // Lógica para:
  // 1. Calcular el monto total sumando las cuotas.
  // 2. Crear UN SOLO registro de ingreso en la hoja "Transacciones".
  // 3. Actualizar el estado a "Pagada" en TODAS las cuotas correspondientes en la hoja "Control de Préstamos".
}

// ESTA ES LA VERSIÓN COMPLETA Y RECOMENDADA
function aprobarTransaccionConPin(idTransaccion, pin, anotationGenerado) {
    // Se recomienda usar PropertiesService para el PIN
    const PIN_SUPERVISOR = PropertiesService.getScriptProperties().getProperty('SUPERVISOR_PIN') || "1511";

    if (pin !== PIN_SUPERVISOR) {
        return { exito: false, mensaje: "PIN incorrecto." };
    }

    try {
        const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_TRANSACCIONES);
        if (!hoja) throw new Error("No se encontró la hoja de transacciones.");

        const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
        
        // --- CORRECCIÓN 1: Añadido el índice para 'origen' ---
        const indiceColId = encabezados.indexOf("id");
        const indiceColAnotation = encabezados.indexOf("anotation");
        const indiceColEstado = encabezados.indexOf("estado");
        const indiceColOrigen = encabezados.indexOf("origen");

        // Verificación de que todas las columnas necesarias existen
        if ([indiceColId, indiceColAnotation, indiceColEstado, indiceColOrigen].includes(-1)) {
            throw new Error("Columnas 'id', 'anotation', 'estado' u 'origen' no encontradas.");
        }

        const idsEnHoja = hoja.getRange(2, indiceColId + 1, hoja.getLastRow() - 1, 1).getValues();
        let filaEncontrada = -1;
        for (let i = 0; i < idsEnHoja.length; i++) {
            if (String(idsEnHoja[i][0]) === String(idTransaccion)) {
                filaEncontrada = i + 2;
                break;
            }
        }

        if (filaEncontrada !== -1) {
            // --- CORRECCIÓN 2 y 3: Lógica consolidada y llamada a Alegra movida aquí ---

            // Lógica condicional para la anotación
            const origenValor = hoja.getRange(filaEncontrada, indiceColOrigen + 1).getValue();
            if (String(origenValor).trim().toLowerCase() === 'registro manual') {
                hoja.getRange(filaEncontrada, indiceColAnotation + 1).setValue(anotationGenerado);
            }

            // El estado se actualiza SIEMPRE
            hoja.getRange(filaEncontrada, indiceColEstado + 1).setValue("Aprobada");
            
            // Forzamos que los cambios se guarden en la hoja ANTES de continuar
            SpreadsheetApp.flush();

            // LLAMAMOS A LA FUNCIÓN DE REGISTRO, PASÁNDOLE EL ID
            const resultadoAlegra = registrarPagosDesdeSheet(idTransaccion);

            // Devolvemos un mensaje completo basado en el resultado de Alegra
            if (resultadoAlegra && resultadoAlegra.exito) {
                return { exito: true, mensaje: `Tx ${idTransaccion} aprobada y registrada en Alegra (ID: ${resultadoAlegra.idAlegra}).` };
            } else {
                // La transacción SÍ se aprobó en la hoja, pero falló en Alegra.
                return { exito: true, mensaje: `Tx ${idTransaccion} aprobada, pero falló al registrar en Alegra: ${resultadoAlegra ? resultadoAlegra.mensaje : 'Error desconocido.'}` };
            }

        } else {
            return { exito: false, mensaje: "No se encontró la transacción." };
        }

    } catch (e) {
        Logger.log('Error en aprobarTransaccionConPin: ' + e.stack);
        throw new Error('Error en el servidor al aprobar: ' + e.message);
    }
}


/**
 * Verifica si un cliente tiene abono en proceso.
 * Si no está bloqueado, lo bloquea para este usuario.
 * Retorna {ok: true/false, mensaje: "..."}
 */
function verificarYMarcarBloqueoAbono(clienteId, usuarioEmail) {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName('AbonosEnProceso') || ss.insertSheet('AbonosEnProceso');
  const now = new Date();
  const BLOQUEO_MINUTOS = 15;
  let data = hoja.getDataRange().getValues();
  if (data.length === 0) data = [['ClienteID','Usuario','FechaHoraInicio']];
  
  // Limpiar bloqueos viejos (más de 15 min)
  let nuevos = [data[0]];
  let bloqueado = false;
  let bloqueadoPor = '';
  for (let i = 1; i < data.length; i++) {
    const [cid, usuario, inicio] = data[i];
    if (!cid || !usuario || !inicio) continue;
    const fecha = new Date(inicio);
    const minutos = (now - fecha) / 60000;
    if (minutos > BLOQUEO_MINUTOS) continue; // Expirado, no copiar
    if (cid == clienteId && usuario != usuarioEmail) {
      bloqueado = true;
      bloqueadoPor = usuario;
      nuevos.push(data[i]);
    } else if (!(cid == clienteId && usuario == usuarioEmail)) {
      nuevos.push(data[i]);
    }
  }
  // Si no estaba bloqueado, lo agregamos
  if (!bloqueado) nuevos.push([clienteId, usuarioEmail, now]);
  hoja.clearContents();
  hoja.getRange(1,1,nuevos.length,3).setValues(nuevos);

  if (bloqueado) {
    return {ok: false, mensaje: 'Otro usuario (' + bloqueadoPor + ') ya está abonando a este cliente.'};
  }
  return {ok: true, mensaje: 'Listo para abonar.'};
}

/**
 * Libera el bloqueo para ese cliente y usuario.
 */
function liberarBloqueoAbono(clienteId, usuarioEmail) {
  const ss = SpreadsheetApp.getActive();
  const hoja = ss.getSheetByName('AbonosEnProceso');
  if (!hoja) return;
  let data = hoja.getDataRange().getValues();
  let nuevos = [data[0]];
  for (let i = 1; i < data.length; i++) {
    const [cid, usuario] = data[i];
    if (cid == clienteId && usuario == usuarioEmail) continue;
    nuevos.push(data[i]);
  }
  hoja.clearContents();
  hoja.getRange(1,1,nuevos.length,3).setValues(nuevos);
}

/**
 * Devuelve cuotas pendientes de un cliente (por orden de vencimiento)
 */
/**
 * Devuelve cuotas pendientes de un cliente (por orden de vencimiento).
 * VERSIÓN CORREGIDA Y ROBUSTA.
 */
/**
 * Devuelve cuotas pendientes de un cliente (por orden de vencimiento).
 * VERSIÓN FINAL CON CORRECCIÓN DE FECHAS.
 */
function obtenerCuotasPendientes(clienteId) {
    try {
        const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONTROL_PRESTAMOS);
        if (!hoja || hoja.getLastRow() < 2) return [];

        const data = hoja.getDataRange().getValues();
        const encabezados = data.shift().map(h => String(h).trim().toLowerCase());

        const idxDeudor = encabezados.indexOf('id_deudor');
        const idxEstado = encabezados.indexOf('estado');
        const idxCuota = encabezados.indexOf('id_cuota');
        const idxMonto = encabezados.indexOf('monto_cuota');
        const idxVenc = encabezados.indexOf('fecha_vencimiento');
        
        if (idxDeudor === -1 || idxEstado === -1) {
            Logger.log("Error: No se encontraron las columnas 'id_deudor' o 'estado'.");
            return [];
        }

        const idClienteLimpio = String(clienteId).trim();
        const pendientes = [];

        for (const fila of data) {
            const deudorEnFila = String(fila[idxDeudor]).trim();
            const estadoEnFila = String(fila[idxEstado]).trim().toLowerCase();
            
            if (deudorEnFila === idClienteLimpio && estadoEnFila === 'pendiente') {
                const fechaCruda = fila[idxVenc];
                pendientes.push({
                    ID_Cuota: fila[idxCuota],
                    Monto: fila[idxMonto],
                    // --- CAMBIO CLAVE Y DEFINITIVO ---
                    // Convertimos la fecha a un string ISO, que sí puede viajar por la red.
                    Fecha_Vencimiento: fechaCruda instanceof Date ? fechaCruda.toISOString() : fechaCruda,
                    Estado: fila[encabezados.indexOf('estado')]
                });
            }
        }

        // El ordenamiento sigue funcionando porque los strings ISO se pueden comparar.
        pendientes.sort((a, b) => new Date(a.Fecha_Vencimiento) - new Date(b.Fecha_Vencimiento));
        return pendientes;

    } catch (e) {
        Logger.log(`Error en obtenerCuotasPendientes: ${e.stack}`);
        return [];
    }
}

/**
 * Marca como pagadas las cuotas seleccionadas y les asigna el ID de la transacción de abono
 */
function marcarCuotasComoPagadas(listaIdCuotas, idTransaccionAbono) {
  const ss = SpreadsheetApp.getActive();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CONTROL_PRESTAMOS);
  const data = hoja.getDataRange().getValues();
  const encabezados = data[0];
  const idxCuota = encabezados.indexOf('ID_Cuota');
  const idxEstado = encabezados.indexOf('Estado');
  const idxIdAbono = encabezados.indexOf('ID_Transaccion_Abono');
  
  for (let i = 1; i < data.length; i++) {
    if (listaIdCuotas.includes(data[i][idxCuota])) {
      data[i][idxEstado] = 'Pagado';
      data[i][idxIdAbono] = idTransaccionAbono;
    }
  }
  hoja.getRange(1,1,data.length,encabezados.length).setValues(data);
}



function test_bloqueoCliente() {
  // Cambia los valores por un cliente y correo real de prueba
  let resultado = verificarYMarcarBloqueoAbono("77", "uge@flores.com");
  Logger.log(resultado);
}

function test_liberarBloqueo() {
  liberarBloqueoAbono("77", "uge@flores.com");
}

function test_obtenerCuotas() {
  let pendientes = obtenerCuotasPendientes("351");
  Logger.log(pendientes);
}

function test_marcarCuotasPagadas() {
  // Simula pago de las dos primeras cuotas
  marcarCuotasComoPagadas(["PREST-1749750821585-1","PREST-1749750821585-2"], "PREST-1749750821585");
}

/**
 * Función de prueba para verificar la lógica de obtenerCuotasPendientes
 * para un cliente específico.
 */
function test_ObtenerCuotasDeUnCliente() {
    // --- CONFIGURACIÓN DE LA PRUEBA ---
    // Simula el ID del cliente que se selecciona en el formulario.
    // Para la prueba de "Idelma Lilly Flores", el ID es '351'.
    const idClientePrueba = '351'; 
    
    Logger.log(`--- INICIANDO PRUEBA: Buscando cuotas para el cliente ID: ${idClientePrueba} ---`);

    try {
        // --- EJECUCIÓN ---
        // Llama a la función que queremos probar con el ID simulado.
        const cuotasEncontradas = obtenerCuotasPendientes(idClientePrueba);

        // --- VERIFICACIÓN DE RESULTADOS ---
        if (Array.isArray(cuotasEncontradas)) {
            if (cuotasEncontradas.length > 0) {
                Logger.log(`✅ ÉXITO: Se encontraron ${cuotasEncontradas.length} cuotas pendientes.`);
                Logger.log("Datos de las cuotas encontradas:");
                // Imprime los resultados en un formato fácil de leer.
                Logger.log(JSON.stringify(cuotasEncontradas, null, 2));
            } else {
                Logger.log(`⚠️ FALLO: La función se ejecutó, pero no encontró ninguna cuota para el ID '${idClientePrueba}'.`);
                Logger.log("Posibles causas: 1) El ID del cliente es incorrecto. 2) No hay cuotas con estado 'Pendiente' para ese ID. 3) Hay un problema de formato de datos en la hoja que la función no pudo limpiar.");
            }
        } else {
            Logger.log("❌ ERROR CRÍTICO: La función obtenerCuotasPendientes no devolvió un array. Devolvió:");
            Logger.log(cuotasEncontradas);
        }

    } catch (error) {
        Logger.log(`❌ ERROR FATAL: La prueba falló con una excepción: ${error.message}`);
        Logger.log(error.stack);
    }

    Logger.log("--- PRUEBA FINALIZADA ---");
}
