// --- CONSTANTES GLOBALES Y DE CONFIGURACIÓN ---
//const HOJA_SCRIPT_ID = "1WX497JkSJkJS2kaNo38MJVHBa1nzE4OraOKE07JNkPg";
//const HOJA_TRANSACCIONES = "Transacciones";
//const HOJA_CONFIG = "Configuración";
//const REMITENTE_LAFISE = "bancanet@notificaciones.lafise.com";
//const FECHA_MINIMA_STR = "2025-04-15"; // Ajusta la fecha mínima según necesites. Para probar hoy, usa una fecha de ayer o antes.

// --- VARIABLES GLOBALES ---
// let CONFIG_GLOBAL = null;
// let CUENTAS_PROPIAS_OBJ = null;
//let CUENTAS_PROPIAS_LISTA = null;

// Asegúrate de que esta función esté en tu script y sea la misma que funcionó antes.
function cargarConfiguracionGlobal() {
  if (CONFIG_GLOBAL && Object.keys(CONFIG_GLOBAL).length > 0 && !CONFIG_GLOBAL.errorCarga) {
    // Logger.log("Usando configuración global previamente cargada.");
    return CONFIG_GLOBAL;
  }
  Logger.log("Iniciando carga de configuración global...");
  try {
    const spreadsheet = SpreadsheetApp.openById(HOJA_SCRIPT_ID);
    const hojaConfig = spreadsheet.getSheetByName(HOJA_CONFIG);
    CONFIG_GLOBAL = {};
    CUENTAS_PROPIAS_OBJ = {};
    const tempCuentasPropiasLista = [];

    if (!hojaConfig) {
      Logger.log(`ERROR CRÍTICO: Hoja de Configuración "${HOJA_CONFIG}" no encontrada.`);
      CONFIG_GLOBAL.tipoCambio = 36.60; // Default muy básico
      CONFIG_GLOBAL.errorCarga = true;
      return CONFIG_GLOBAL;
    }

    const ultimaFilaConfig = hojaConfig.getLastRow();
    if (ultimaFilaConfig < 2) {
      Logger.log(`Hoja de Configuración "${HOJA_CONFIG}" está vacía o no tiene datos.`);
      CONFIG_GLOBAL.tipoCambio = 36.60; // Default
      // No marcar como error de carga necesariamente, podría ser intencional
      return CONFIG_GLOBAL;
    }

    const datos = hojaConfig.getRange(2, 1, ultimaFilaConfig - 1, 2).getValues();
    let tipoCambioEncontrado = false;

    datos.forEach(([clave, valor]) => {
      if (clave && clave.toString().trim() !== "") {
        const claveStr = clave.toString().trim();
        CONFIG_GLOBAL[claveStr] = valor;
        if (claveStr.toLowerCase() === 'tipocambio') {
          const tcNum = parseFloat(valor);
          if (!isNaN(tcNum) && isFinite(tcNum)) {
            CONFIG_GLOBAL.tipoCambio = tcNum;
            tipoCambioEncontrado = true;
          } else {
            Logger.log(`ADVERTENCIA: 'tipoCambio' ('${valor}') no es un número válido.`);
          }
        } else if (claveStr.startsWith('cuenta_')) {
          const numeroCuenta = claveStr.replace('cuenta_', '').trim();
          if (numeroCuenta) {
            const idCuenta = Number(valor);
            if (!isNaN(idCuenta) && isFinite(idCuenta)) {
               CUENTAS_PROPIAS_OBJ[numeroCuenta] = { id: idCuenta, esPropia: true, numero: numeroCuenta };
               tempCuentasPropiasLista.push(numeroCuenta);
            } else {
               Logger.log(`ADVERTENCIA: ID para ${claveStr} ('${valor}') no es numérico. Se omite.`);
            }
          }
        }
      }
    });

    if (!tipoCambioEncontrado && !CONFIG_GLOBAL.hasOwnProperty('tipoCambio')) {
      Logger.log("ADVERTENCIA: 'tipoCambio' no encontrado en Configuración. Usando valor por defecto: 36.60");
      CONFIG_GLOBAL.tipoCambio = 36.60;
    }
    
    CUENTAS_PROPIAS_LISTA = tempCuentasPropiasLista;
    Logger.log(`Configuración cargada. Tipo de cambio: ${CONFIG_GLOBAL.tipoCambio}`);
    Logger.log(`Cuentas propias (${CUENTAS_PROPIAS_LISTA.length}): ${CUENTAS_PROPIAS_LISTA.join(", ")}`);
    if (CUENTAS_PROPIAS_LISTA.length === 0) {
        Logger.log("ADVERTENCIA: No se cargaron cuentas propias. Verifica la hoja de 'Configuración'.");
    }

  } catch (error) {
    Logger.log(`ERROR CRÍTICO al cargar la configuración: ${error.message} Stack: ${error.stack}`);
    CONFIG_GLOBAL = { tipoCambio: 36.60, errorCarga: true }; // Configuración mínima de emergencia
    CUENTAS_PROPIAS_OBJ = {};
    CUENTAS_PROPIAS_LISTA = [];
  }
  return CONFIG_GLOBAL;
}

/**
 * VERSIÓN DE DIAGNÓSTICO COMPLETA
 * Procesa correos de Gmail de transferencias Lafise y los registra en una hoja de cálculo.
 */
function procesarTransferenciasTextoLafise() {
  Logger.log("********************************************************************************");
  Logger.log("*** EJECUTANDO 'procesarTransferenciasTextoLafise' (VERSIÓN DIAGNÓSTICO COMPLETA) ***");
  Logger.log("********************************************************************************");
  
  try {
    cargarConfiguracionGlobal();

    if (!CONFIG_GLOBAL || CONFIG_GLOBAL.errorCarga) {
      Logger.log("ERROR: No se pudo cargar la configuración. Abortando `procesarTransferenciasTextoLafise`.");
      return;
    }
    if (typeof CONFIG_GLOBAL.tipoCambio !== 'number' || isNaN(CONFIG_GLOBAL.tipoCambio)) {
      Logger.log(`ERROR: El tipo de cambio (${CONFIG_GLOBAL.tipoCambio}) no es un número válido. Abortando.`);
      return;
    }
     if (!CUENTAS_PROPIAS_LISTA) { // Comprobación adicional
      Logger.log("ERROR CRÍTICO: CUENTAS_PROPIAS_LISTA no está inicializada. Revise `cargarConfiguracionGlobal`.");
      return;
    }
    
    const FECHA_MINIMA = new Date(FECHA_MINIMA_STR);
    if (isNaN(FECHA_MINIMA.getTime())) {
        Logger.log(`ERROR: FECHA_MINIMA_STR ('${FECHA_MINIMA_STR}') no es una fecha válida. Abortando.`);
        return;
    }
    Logger.log(`Procesando correos con fecha de transacción a partir de: ${FECHA_MINIMA.toISOString()}`);

    const spreadsheet = SpreadsheetApp.openById(HOJA_SCRIPT_ID);
    const hoja = spreadsheet.getSheetByName(HOJA_TRANSACCIONES);

    if (!hoja) {
      Logger.log(`ERROR CRÍTICO: Hoja de Transacciones "${HOJA_TRANSACCIONES}" no encontrada. Abortando.`);
      return;
    }

    // --- LÓGICA PARA NÚMERO CORRELATIVO ---
    let proximoCorrelativo = 1;
    const ultimaFilaTransacciones = hoja.getLastRow();
    if (ultimaFilaTransacciones >= 2) { // Asumimos que la fila 1 es de encabezados
      const rangoColumnaP = hoja.getRange(2, 16, ultimaFilaTransacciones - 1, 1); // Columna P es la 16
      const datosColumnaP = rangoColumnaP.getValues()
        .flat()
        .map(val => {
          const num = Number(val);
          return Number.isInteger(num) && isFinite(num) && num > 0 ? num : null;
        })
        .filter(val => val !== null);
      if (datosColumnaP.length > 0) proximoCorrelativo = Math.max(...datosColumnaP) + 1;
    }
    Logger.log("Próximo número correlativo a usar: " + proximoCorrelativo);
    
    // --- ANOTACIONES REGISTRADAS (Columna I es la 9) ---
    const anotacionesRegistradas = ultimaFilaTransacciones > 1
      ? hoja.getRange(2, 9, ultimaFilaTransacciones - 1, 1).getValues().flat().map(r => r.toString().trim()).filter(r => r !== "")
      : [];
    Logger.log(`Anotaciones (referencias) ya registradas en la hoja: ${anotacionesRegistradas.length}`);

    // --- PASO 1: CONSTRUIR Y EJECUTAR LA CONSULTA DE GMAIL ---
    // Cambiado a `newer_than:3d` (últimos 3 días) para ser más específico y cubrir el día de hoy y un margen.
    // Ajusta `3d` (3 días) según necesites. Si ejecutas esto una vez al día, `2d` o `3d` es seguro.
    let query = `from:"${REMITENTE_LAFISE}" newer_than:3d (` +
      'subject:"Aviso de transferencia Envío Veloz" OR ' +
      'subject:"Notificación de pago de servicio" OR ' +
      'subject:"Aviso de transferencia a terceros LAFISE mismo país")';
    
    query = `${query} label:inbox`; // Busca solo en la bandeja de entrada.
    Logger.log(`Consulta de Gmail a ejecutar: ${query}`);

    const threads = GmailApp.search(query);
    Logger.log(`\n>>> HILOS DE CORREO ENCONTRADOS POR GMAIL: ${threads.length} <<<\n`);

    if (threads.length === 0) {
      Logger.log("--- No se encontraron hilos de correo que coincidan con la consulta. ---");
      Logger.log("CONSEJOS DE DEPURACIÓN SI ESPERABAS CORREOS:");
      Logger.log(`1. REMITENTE: Verifica que el remitente del correo sea exactamente: "${REMITENTE_LAFISE}"`);
      Logger.log("2. ASUNTO: Verifica que los asuntos de tus correos nuevos coincidan EXACTAMENTE (mayúsculas, acentos, etc.) con alguno de los buscados.");
      Logger.log("3. FECHA DEL CORREO: La consulta busca correos recibidos en los últimos 3 días (`newer_than:3d`).");
      Logger.log("4. UBICACIÓN: Revisa si los correos están en la 'Bandeja de entrada'. Si fueron archivados o están en Spam/Papelera, `label:inbox` no los encontrará.");
      Logger.log("   -> Para probar sin el filtro de bandeja de entrada, comenta la línea `query = `${query} label:inbox`;`");
      Logger.log("--- FIN DEL PROCESO DE TRANSFERENCIAS ---");
      return;
    }

    // --- PASO 2: ITERAR Y PROCESAR CADA MENSAJE ---
    let correosProcesadosExitosamente = 0;
    threads.forEach((thread, threadIndex) => {
      const messages = thread.getMessages(); // Obtener todos los mensajes del hilo
      messages.forEach((message, messageIndex) => {
        const subject = message.getSubject();
        const body = message.getPlainBody(); // Usar PlainBody para regex más simples
        const messageId = message.getId();
        const messageDate = message.getDate(); // Fecha en que se recibió el correo

        Logger.log(`\n--- [INICIO] Procesando Mensaje ${threadIndex + 1}-${messageIndex + 1} (ID Gmail: ${messageId}) ---`);
        Logger.log(`Asunto del correo: "${subject}"`);
        Logger.log(`Fecha de recepción del correo: ${messageDate.toISOString()}`);

        const isPagoServicio = subject.includes("Notificación de pago de servicio");
        if (isPagoServicio) Logger.log("Este correo parece ser una 'Notificación de pago de servicio'.");

        // PASO 2.1: FILTRO INICIAL DEL CUERPO
        Logger.log("Aplicando filtro inicial del cuerpo del correo...");
        const PasaFiltroExitoso = body.includes("Estado: Exitoso");
        const PasaFiltroReferencia = body.includes("Referencia:");
        const PasaFiltroMonto = body.includes("Monto:");
        const ContieneAsteriscos = body.includes("*****"); // Para transferencias enmascaradas/fallidas

        Logger.log(`  - ¿Contiene 'Estado: Exitoso'? : ${PasaFiltroExitoso}`);
        Logger.log(`  - ¿Contiene 'Referencia:'?     : ${PasaFiltroReferencia}`);
        Logger.log(`  - ¿Contiene 'Monto:'?          : ${PasaFiltroMonto}`);
        Logger.log(`  - ¿Contiene '*****'? (debe ser falso para procesar): ${ContieneAsteriscos}`);

        if (!PasaFiltroExitoso || !PasaFiltroReferencia || !PasaFiltroMonto || ContieneAsteriscos) {
          Logger.log(`--- [FIN-DESCARTADO POR FILTRO INICIAL] Mensaje (ID: ${messageId}). ---`);
          return; // Siguiente mensaje
        }
        Logger.log("  => Mensaje PASÓ el filtro inicial.");

        // PASO 2.2: EXTRACCIÓN Y VALIDACIÓN DE LA REFERENCIA
        Logger.log("Extrayendo número de referencia...");
        const referenciaMatch = body.match(/Referencia:\s*(\d+)/);
        const referencia = referenciaMatch ? referenciaMatch[1] : null;
        Logger.log(`  - Referencia extraída: ${referencia}`);

        if (!referencia) {
          Logger.log(`--- [FIN-DESCARTADO POR FALTA DE REFERENCIA] Mensaje (ID: ${messageId}). ---`);
          return;
        }
        if (anotacionesRegistradas.includes(referencia)) {
          Logger.log(`--- [FIN-DESCARTADO POR REFERENCIA DUPLICADA] Mensaje (ID: ${messageId}) con Referencia ${referencia} ya fue registrado. ---`);
          return;
        }
        Logger.log(`  => Referencia ${referencia} es nueva y válida.`);

        // PASO 2.3: EXTRACCIÓN Y VALIDACIÓN DE FECHA DE TRANSACCIÓN
        Logger.log("Extrayendo y formateando la fecha de transacción del cuerpo del correo...");
        const fechaTextoMatch = body.match(/Fecha:\s*(.+?)(?:\r?\n|$)/); // Captura hasta el fin de línea
        const fechaTexto = fechaTextoMatch ? fechaTextoMatch[1].trim() : null;
        const fechaFormateada = formatearFecha(fechaTexto); // Usa tu función formatearFecha
        Logger.log(`  - Texto de fecha encontrado en correo: "${fechaTexto}"`);
        Logger.log(`  - Fecha formateada para la hoja: "${fechaFormateada}"`);

        if (!fechaFormateada) {
          Logger.log(`--- [FIN-DESCARTADO POR FECHA INVÁLIDA/NO FORMATEABLE] Mensaje (ID: ${messageId}, Ref: ${referencia}). ---`);
          return;
        }
        const fechaTransaccion = new Date(fechaFormateada); // Convertir a objeto Date para comparar
        if (fechaTransaccion < FECHA_MINIMA) {
          Logger.log(`--- [FIN-DESCARTADO POR FECHA ANTERIOR A MÍNIMA] Mensaje (ID: ${messageId}, Ref: ${referencia}). Fecha transacción ${fechaFormateada} es anterior a ${FECHA_MINIMA_STR}. ---`);
          return;
        }
        Logger.log(`  => Fecha de transacción ${fechaFormateada} es válida y posterior o igual a la mínima.`);

        // PASO 2.4: EXTRACCIÓN DE CUENTAS
        Logger.log("Extrayendo cuentas Origen y Destino...");
        const cuentaOrigenRegex = /Cuenta\s*(?:de\s*)?Origen[\s\S]+?N(?:ú|u)mero de cuenta:\s*([A-Z0-9]+)/i;
        const cuentaDestinoRegex = /Cuenta\s*Destino[\s\S]+?N(?:ú|u)mero de cuenta:\s*([A-Z0-9]+)/i;
        
        const cuentaOrigenExtraida = body.match(cuentaOrigenRegex)?.[1] || "";
        const cuentaDestinoExtraida = body.match(cuentaDestinoRegex)?.[1] || "";
        const cuentaOrigen = cuentaOrigenExtraida.trim();
        const cuentaDestino = cuentaDestinoExtraida.trim();
        Logger.log(`  - Cuenta Origen extraída: '${cuentaOrigen}'`);
        Logger.log(`  - Cuenta Destino extraída: '${cuentaDestino}'`);

        if (!cuentaOrigen && !cuentaDestino && isPagoServicio) {
            Logger.log("  ADVERTENCIA: Para Pago de Servicio, no se extrajo ni origen ni destino. Se asumirá origen si se configura una sola cuenta propia o se requiere configuración específica.");
            // Para pagos de servicio, la cuenta destino puede no ser relevante o no estar presente.
            // La cuenta origen es la importante. Si no se extrae, puede ser un problema.
        } else if (!cuentaOrigen && !cuentaDestino) {
            Logger.log(`--- [FIN-DESCARTADO POR FALTA DE CUENTAS] Mensaje (ID: ${messageId}, Ref: ${referencia}). No se extrajo ni origen ni destino. ---`);
            return;
        }

        // PASO 2.5: DETERMINAR TIPO DE MOVIMIENTO Y BANK ACCOUNT ID
        Logger.log("Determinando tipo de movimiento y Bank Account ID...");
        const esOrigenPropio = cuentaOrigen && CUENTAS_PROPIAS_LISTA.includes(cuentaOrigen);
        const esDestinoPropio = cuentaDestino && CUENTAS_PROPIAS_LISTA.includes(cuentaDestino);
        Logger.log(`  - ¿Cuenta Origen es propia? (${cuentaOrigen}): ${esOrigenPropio}`);
        Logger.log(`  - ¿Cuenta Destino es propia? (${cuentaDestino}): ${esDestinoPropio}`);

        let tipoMovimiento = "unknown";
        let bankAccountId = 0; // Default

        if (isPagoServicio) {
            if (esOrigenPropio) {
                tipoMovimiento = "out";
                bankAccountId = CUENTAS_PROPIAS_OBJ[cuentaOrigen]?.id || 0;
                Logger.log("    Pago de servicio desde cuenta propia.");
            } else {
                Logger.log(`  ADVERTENCIA: Pago de servicio desde cuenta origen NO PROPIA ('${cuentaOrigen}'). Se descartará o requiere lógica especial.`);
                Logger.log(`--- [FIN-DESCARTADO PAGO DE SERVICIO NO PROPIO] Mensaje (ID: ${messageId}, Ref: ${referencia}). ---`);
                return;
            }
        } else { // Lógica para transferencias
            if (esOrigenPropio && !esDestinoPropio) {
                tipoMovimiento = "out";
                bankAccountId = CUENTAS_PROPIAS_OBJ[cuentaOrigen]?.id || 0;
            } else if (!esOrigenPropio && esDestinoPropio) {
                tipoMovimiento = "in";
                bankAccountId = CUENTAS_PROPIAS_OBJ[cuentaDestino]?.id || 0;
            } else if (esOrigenPropio && esDestinoPropio) {
                // Transferencia entre cuentas propias. Registrar como 'out' de la origen.
                tipoMovimiento = "out"; 
                bankAccountId = CUENTAS_PROPIAS_OBJ[cuentaOrigen]?.id || 0;
                Logger.log("    Transferencia interna entre cuentas propias.");
            } else {
                Logger.log("    Transferencia no involucra una cuenta propia conocida como origen o destino principal.");
                Logger.log(`--- [FIN-DESCARTADO TRANSFERENCIA NO PROPIA] Mensaje (ID: ${messageId}, Ref: ${referencia}). ---`);
                return;
            }
        }
        Logger.log(`  => Tipo Movimiento: ${tipoMovimiento}, Bank Account ID: ${bankAccountId}`);
        if (bankAccountId === 0 && tipoMovimiento !== "unknown") {
            Logger.log(`  ADVERTENCIA: Bank Account ID es 0 para un movimiento ${tipoMovimiento}. Cuenta ${tipoMovimiento === 'out' ? cuentaOrigen : cuentaDestino} podría no estar bien mapeada en Configuración.`);
        }


        // PASO 2.6: EXTRACCIÓN DE MONEDA Y MONTO
        Logger.log("Extrayendo Moneda y Monto...");
        const monedaMatch = body.match(/Moneda:\s*(USD|NIO)/i);
        const moneda = monedaMatch ? monedaMatch[1].toUpperCase() : "NIO"; // Default NIO
        const montoRegex = /Monto:\s*(?:USD|NIO)?\s*([\d,]+\.?\d*)/i;
        const montoCapturado = body.match(montoRegex)?.[1];
        let monto = parseFloat(montoCapturado ? montoCapturado.replace(/,/g, "") : "0");

        Logger.log(`  - Moneda extraída: ${moneda}, Monto capturado: '${montoCapturado}', Monto parseado inicial: ${monto}`);
        if (isNaN(monto) || !isFinite(monto)) {
            Logger.log(`  ERROR: Monto parseado no es un número válido. Se usará 0.`);
            monto = 0;
        }

        let montoOriginal = monto;
        let monedaOriginal = moneda;
        if (moneda === "USD") {
          monto = parseFloat((monto * CONFIG_GLOBAL.tipoCambio).toFixed(2));
          Logger.log(`  Convertido de USD a NIO: ${montoOriginal} USD * ${CONFIG_GLOBAL.tipoCambio} = ${monto} NIO`);
        } else {
          monto = parseFloat(monto.toFixed(2)); // Asegurar dos decimales
        }
        if (isNaN(monto) || !isFinite(monto)) {
            Logger.log(`  ERROR: Monto final (post-conversión/toFixed) no es válido. Se usará 0.`);
            monto = 0;
        }
        Logger.log(`  => Monto final (NIO): ${monto}`);

        // PASO 2.7: EXTRACCIÓN DE CONCEPTO Y OBSERVACIONES
        Logger.log("Extrayendo Concepto y construyendo Observaciones...");
        const concepto =
          body.match(/Concepto de la\s*transacci(?:ó|o)n:\s*(.+?)(?:\r?\n|$)/i)?.[1]?.trim() ||
          body.match(/Servicio pagado:\s*(.+?)(?:\r?\n|$)/i)?.[1]?.trim() ||
          body.match(/Concepto:\s*(.+?)(?:\r?\n|$)/i)?.[1]?.trim() ||
          (isPagoServicio ? `Pago de Servicio Ref ${referencia}` : `Transferencia Ref ${referencia}`);
        Logger.log(`  - Concepto extraído/generado: "${concepto}"`);

        const titularOrigenMatch = body.match(/Cuenta Origen[\s\S]+?Titular:\s*(.+?)(?:\r?\n|$)/i);
        const titularOrigen = titularOrigenMatch ? titularOrigenMatch[1].trim() : "";
        const operadorMatch = body.match(/Datos de la operaci(?:ó|o)n[\s\S]+?Operador:\s*(.+?)(?:\r?\n|$)/i);
        const operador = operadorMatch ? operadorMatch[1].trim() : titularOrigen; // Usar titular origen si operador no está

        let observaciones = concepto;
        if (tipoMovimiento === "in" && operador) {
          observaciones = `${concepto} - Recibido de: ${operador}`;
        } else if (tipoMovimiento === "out") {
            const titularDestinoMatch = body.match(/Cuenta Destino[\s\S]+?Titular:\s*(.+?)(?:\r?\n|$)/i);
            const titularDestino = titularDestinoMatch ? titularDestinoMatch[1].trim() : "";
            if (esDestinoPropio) {
                 observaciones = `${concepto} - Transf. interna a cta: ${cuentaDestino}`;
            } else if (titularDestino) {
                observaciones = `${concepto} - Enviado a: ${titularDestino} (Cta: ${cuentaDestino})`;
            } else if (cuentaDestino) { // Si no hay titular destino pero sí número de cuenta
                observaciones = `${concepto} - Enviado a cta: ${cuentaDestino}`;
            }
        }
        Logger.log(`  => Observaciones finales: "${observaciones.substring(0,499)}"`);

        // PASO 2.8: REGISTRAR EN LA HOJA
        Logger.log("Preparando para registrar en la hoja de cálculo...");
        const numeroCorrelativoActual = proximoCorrelativo;
        
        try {
          hoja.appendRow([
            `Correo ${messageDate.toISOString().substring(0,10)} ID ${messageId}`, // Col A
            fechaFormateada,                               // Col B
            tipoMovimiento,                                // Col C
            "transfer",       // Col D
            monto,                                         // Col E
            bankAccountId,                                 // Col F
            "",                                            // Col G: Client ID
            "",                                            // Col H: Category ID
            referencia,                                    // Col I
            observaciones.substring(0, 499),               // Col J (limitado a 500 chars)
            "Pendiente",                                   // Col K
            "",                                            // Col L: URL Imagen
            null, null, null,                              // Col M, N, O
            numeroCorrelativoActual                        // Col P
          ]);
          Logger.log(`✅ REGISTRADO EXITOSAMENTE! Ref ${referencia} | Correlativo: ${numeroCorrelativoActual} | Tipo: ${tipoMovimiento} | Monto: ${monto}`);
          anotacionesRegistradas.push(referencia); // Añadir a la lista en memoria para esta ejecución
          proximoCorrelativo++; 
          correosProcesadosExitosamente++;
        } catch (e) {
          Logger.log(`❌ ERROR AL INTENTAR ESCRIBIR EN LA HOJA para Ref ${referencia}: ${e.message} ${e.stack}`);
        }
        Logger.log(`--- [FIN] Procesamiento Mensaje (ID: ${messageId}, Ref: ${referencia}) ---`);
      }); // Fin forEach message
    }); // Fin forEach thread

    Logger.log(`\n--- Resumen del Proceso ---`);
    Logger.log(`Total de hilos de correo encontrados por Gmail: ${threads.length}`);
    Logger.log(`Total de correos procesados exitosamente y registrados: ${correosProcesadosExitosamente}`);
    Logger.log("--- FIN DEL PROCESO DE TRANSFERENCIAS ---");

    return `${correosProcesadosExitosamente} nuevos registrados.`;

  } catch (errorGlobal) {
    Logger.log(`Error FATAL en 'procesarTransferenciasTextoLafise': ${errorGlobal.message} - Stack: ${errorGlobal.stack}`);
  }
}

// La función formatearFecha() se mantiene igual.
function formatearFecha(fechaTexto) {
  if (!fechaTexto || typeof fechaTexto !== 'string') {
    // Logger.log("formatearFecha: entrada no válida o vacía."); // Puede ser muy verboso
    return "";
  }
  const fechaParte = fechaTexto.split(" ")[0]; 
  const meses = { ENE: "01", FEB: "02", MAR: "03", ABR: "04", MAY: "05", JUN: "06", JUL: "07", AGO: "08", SEP: "09", OCT: "10", NOV: "11", DIC: "12" };
  const partes = fechaParte.match(/^(\d{1,2})\/([A-Z]{3})\/(\d{4})$/i); 
  if (!partes) {
    // Logger.log(`formatearFecha: No se pudo parsear: "${fechaParte}" (Original: "${fechaTexto}")`);
    return "";
  }
  let [, dia, mesTexto, anio] = partes;
  dia = dia.padStart(2, '0'); 
  const mesNumero = meses[mesTexto.toUpperCase()];
  if (!mesNumero) {
    // Logger.log(`formatearFecha: Mes no reconocido "${mesTexto}"`);
    return "";
  }
  return `${anio}-${mesNumero}-${dia}`;
}

// Para ejecutar manualmente desde el editor:
// 1. Selecciona 'testProcesarTransferencias' en el desplegable.
// 2. Haz clic en 'Ejecutar'.
function testProcesarTransferencias() {
  procesarTransferenciasTextoLafise();
}
