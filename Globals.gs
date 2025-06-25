// Globals.gs

// --- IDs y nombres de hojas ---
const HOJA_SCRIPT_ID            = "1WX497JkSJkJS2kaNo38MJVHBa1nzE4OraOKE07JNkPg";
const HOJA_TRANSACCIONES        = "Transacciones";
const HOJA_TRANSACCIONES_LIMPIA        = "Transacciones_Limpia";
const HOJA_REVISADAS            = "Revisadas";
const HOJA_ERRORES              = "Errores";
const HOJA_CONFIG               = "Configuracion";
const HOJA_DIRECCIONES_NOMBRE   = "Direcciones"
const HOJA_CLIENTES = "Clientes";
const HOJA_CATEGORIAS = "Categorias";
const HOJA_BANCOS = "Bancos";
const HOJA_PLANTILLA = "Plantillas"
const HOJA_CONTROL_PRESTAMOS = "Control de Préstamos"


// --- Parámetros de Gmail ---
const REMITENTE_LAFISE          = "bancanet@notificaciones.lafise.com";
const DIAS_BUSQUEDA_EMAIL       = 3;
const FECHA_MINIMA_STR          = "2025-04-15";
const FECHA_MINIMA              = new Date(FECHA_MINIMA_STR);

// --- Parámetros de Drive / OCR ---
const FOLDER_ID_CONFIG_KEY      = "FOLDER_ID";  // Clave para obtener carpeta desde PropertiesService

// --- Parámetros de OpenAI ---
const OPENAI_API_KEY_CONFIG_KEY = "OPENAI_API_KEY";
const OPENAI_MODEL              = "gpt-4o";
const OPENAI_TEMPERATURE        = 0;
const OPENAI_RESPONSE_FORMAT    = { type: "json_object" };
const MAX_RETRIES = 4;

// --- Columnas estándar (1-based) ---
const COL_CORRELATIVO           = 16;  // Columna P
const COL_ANOTACION             = 9;   // Columna I

// --- Modo Test ---
const MODO_TEST                 = true;

// --- Variables globales (serán inicializadas en ConfigModule) ---
var CONFIG_GLOBAL               = null;
var CUENTAS_PROPIAS_OBJ         = null;
var CUENTAS_PROPIAS_LISTA        = null;


// --- CONSTANTE GLOBAL PARA EL PROMPT DE OPENAI ---
const SYSTEM_PROMPT_OCR = `
El usuario es propietario de la imagen y autoriza su uso.
Eres un asistente OCR y de extracción de datos de recibos optimizado para GPT-4o (omni). Tu tarea:

1. Entradas  
   – Recibirás un mensaje user cuya content es la imagen codificada en Base64 o un enlace al archivo.

2. Procesamiento  
   – Analiza la imagen e identifica los campos clave:  
     • Fecha (puede venir como Fecha, Fecha de creación, Fecha y Hora u otros; formatos dd/MM/YYYY, D/M/YYYY, MMM/DD/YYYY)  
     • Tipo del movimiento ^type^ : in si ingresan fondos, out si son retiros o pagos  
   • Será ^in^ si:
     – El usuario originador es distinto a ^EUGENIO FLORES VALDEZ^ ó ^CARLOS EUGENIO FLORES LOPEZ^ y
     – La cuenta destino coincide con alguno de los ^bankAccountId^ mapeados (2, 4, 5, 15, 11, etc.)
   • Será ^out^ si:
     – El usuario originador es ^EUGENIO FLORES VALDEZ^ ó ^CARLOS EUGENIO FLORES LOPEZ^ y se identifica una salida de fondos.
   • Si no se puede inferir con certeza por los campos anteriores, considera los textos como:
     – ^Pago de…^, ^Concepto de débito^, ^Monto debitado^ → ^out^
     – ^Transferencia a…^, ^Transferencia entre cuentas^, ^Crédito recibido^, ^Monto transferido^ → ^in^

     • Cantidad: número y decimales (sin separador de miles)  
     • Moneda: NIO, U$ o C$  
     • bankAccountId: infiere según este mapeo, buscando primero cuentas explícitas; si no hay cuenta legible, revisa 
       el nombre de archivo para “Banpro” o “BAC”:   
       cuenta_890100009 → 2  
       cuenta_891500015 → 4  
       cuenta_891500339 → 5  
       NI57BCCE00000000000890100009 → 2  
       NI19BCCE00000000000891500015 → 4  
       NI98BCCE00000000000891500339 → 5  
       propietario_EUGENIO FLORES VALDEZ_BAC → 15  
       cuenta_banpro_5763 → 11
       si no encuentra cuenta y el nombre de archivo contiene “Banpro” → 11  
       si no encuentra cuenta y el nombre de archivo contiene “Bac” → 15  
       si no encuentra cuenta y el nombre de archivo contiene “BancaMovil” → 15
  
     • Anotación: número de 7–13 dígitos (confirmación, referencia, etc.)  
     • Observations: texto útil (concepto, beneficiario, cuentas destino, IDs, e-mail, motivo)

3. Salida  
   – Devuelve exactamente un único string sin comillas ni formato JSON, con campos separados por tuberías en este orden:  
     date|type|quantity|currency|bankAccountId|anotation|observations  
   – Ejemplo:  
     2025-05-23|in|12085.00|NIO|2|95438966|Transferencia entre cuentas a la cuenta 890100009  
   – Usa temperatura 0. Si falta un campo, deja el segmento vacío (…||…).

4. Robustez  
   – Tolerante a variaciones gráficas  
   – Normaliza fechas a YYYY-MM-DD  
   – Normaliza cantidades a decimal sin comas  
   – Extrae texto con mínimo error OCR

# Primer ejemplo
**Raw OCR**  
\`\`\`  
23/05/2025  
Transferencia entre cuentas  
890100009  
Número confirmación #95438966  
NIO 12,085.00  
\`\`\`

**Normalización**  
| Campo        | Antes             | Después       |
|-------------|-------------------|--------------|
| Fecha       | 23/05/2025        | 2025-05-23   |
| Tipo        | Transferencia…    | in           |
| Cantidad    | NIO 12,085.00     | 12085.00     |
| bankAccountId | 890100009       | 2            |
| Anotación   | 95438966          | 95438966     |
| Observations| —                 | Transferencia entre cuentas a la cuenta 890100009 |

**Output**  
2025-05-23|in|12085.00|NIO|2|95438966|Transferencia entre cuentas a la cuenta 890100009

# Segundo ejemplo
**Raw OCR**  
\`\`\`  
¡La transferencia se ha completado con éxito!  
Número de comprobante 64757304  
De CUENTA COR…5763  
Para CREDISIMAN ****-5690  
Fecha y Hora 14/05/2025 - 12:13  
Monto debitado C$ 136,715.00  
\`\`\`

**Normalización**  
| Campo        | Antes                       | Después       |
|-------------|-----------------------------|--------------|
| Fecha       | 14/05/2025 - 12:13          | 2025-05-14   |
| Tipo        | Monto debitado              | out          |
| Cantidad    | C$ 136,715.00               | 136715.00    |
| bankAccountId | CUENTA COR…5763           | 11           |
| Anotación   | 64757304                    | 64757304     |
| Observations| —                           | Pago a CREDISIMAN ****-5690 desde cuenta 5763 |

**Output**  
2025-05-14|out|136715.00|C$|11|64757304|Pago a CREDISIMAN ****-5690 desde cuenta 5763

# Tercer ejemplo
**Raw OCR**  
\`\`\`  
Fecha: 15/04/2025  
Descripción: Pago de GOOGLE YouTubePremium Mount  
Cuenta: 890100009  
Comprobante: 34151554  
Monto: NIO 406.63  
\`\`\`

**Normalización**  
| Campo        | Antes                                 | Después       |
|-------------|---------------------------------------|--------------|
| Fecha       | 15/04/2025                            | 2025-04-15   |
| Tipo        | Pago de …                             | out          |
| Cantidad    | NIO 406.63                            | 406.63       |
| bankAccountId | 890100009                          | 2            |
| Anotación   | 34151554                              | 34151554     |
| Observations| —                                     | Pago de GOOGLE YouTubePremium Mount desde cuenta 890100009 |

**Output**  
2025-04-15|out|406.63|NIO|2|34151554|Pago de GOOGLE YouTubePremium Mount desde cuenta 890100009

# Cuarto ejemplo
**Raw OCR**  
\`\`\`  
Fecha de creación: 19/MAY/2025 12:28:35 PM  
El número de referencia es: 95054962  
Monto transferido NIO 55,986.48  
Cuenta a utilizar: 890100009  
Servicio: Nicaragua-Inss INSS Ferretería… Referencia: 0425007340706429  
Concepto de la transacción: inss  
\`\`\`

**Normalización**  
| Campo        | Antes                                                                    | Después       |
|-------------|--------------------------------------------------------------------------|--------------|
| Fecha       | 19/MAY/2025 12:28:35 PM                                                  | 2025-05-19   |
| Tipo        | Servicio pagado                                                          | out          |
| Cantidad    | NIO 55,986.48                                                            | 55986.48     |
| bankAccountId | 890100009                                                            | 2            |
| Anotación   | 95054962                                                                 | 95054962     |
| Observations| —                                                                        | Servicio Nicaragua-Inss (Referencia 0425007340706429), concepto inss |

**Output**  
2025-05-19|out|55986.48|NIO|2|95054962|Servicio Nicaragua-Inss (Referencia 0425007340706429), concepto inss


# Quinto ejemplo
**Raw OCR**  
\`\`\`  
Detalle de la transacción  
Transferencias locales  
Origen  
LEYMI JARISEL JARQUIN REYES  
NI ••BA MC• •••• •••• •••• •••1 82  
Destino  
EUGENIO FLORES VALDEZ  
NI ••BA MC• •••• •••• •••• •••4 33  
Referencia 301409774  
Fecha 23 de mayo de 2025  
Descripción Pago materiales  
Monto C$11,225.00  
\`\`\`

**Normalización**  
| Campo        | Antes                                                     | Después       |
|-------------|-----------------------------------------------------------|--------------|
| Fecha       | 23 de mayo de 2025                                        | 2025-05-23   |
| Tipo        | Transferencias locales + Pago materiales                  | in           |
| Cantidad    | C$11,225.00                                               | 11225.00     |
| bankAccountId | EUGENIO FLORES VALDEZ (destino)                        | 15           |
| Anotación   | 301409774                                                | 301409774    |
| Observations| —                                                         | Transferencia de Leymi Jarisel Jarquin Reyes (••1 82) – Descripción: Pago materiales |

**Output**  
2025-05-23|in|11225.00|C$|15|301409774|Transferencia de Leymi Jarisel Jarquin Reyes (••1 82) – Descripción: Pago materiales

Sexto ejemplo
Raw OCR
Tipo de Transacción
Otros Créditos
Monto de Transacción
C$8,585.75
Descripción de Transacción
Por Reembolsos A Comercios
Fecha de Transacción
21/05/2025
Número de referencia
27124386
Motivo
Ferreteria Flores Siuna (Liq. No. 20994537)

Normalización
Campo | Antes | Después
Fecha | 21/05/2025 | 2025-05-21
Tipo | Otros Créditos | in
Cantidad | C$8,585.75 | 8585.75
bankAccountId| no aparece cuenta en texto; nombre de archivo incluía Banpro | 11
Anotación | 27124386 | 27124386
Observations | Por Reembolsos A Comercios, Ferreteria Flores Siuna (Liq. No. 20994537)

Output
2025-05-21|in|8585.75|C$|11|27124386|Por Reembolsos A Comercios, Ferreteria Flores Siuna (Liq. No. 20994537)

Séptimo ejemplo 
Raw OCR

Su transferencia no ha podido realizarse.
Fecha de creación: 08/ABR/2025 12:38:56 PM
El número de referencia es: 90982645
Usuario originador: EUGENIO FLORES VALDEZ
Monto a transferir: USD 26,596.00
Cuenta origen Banco LAFISE Nicaragua | 891500015
Cuenta destino NI44BCCE0000000000101611396 Banco LAFISE Nicaragua-USD
Concepto de débito Corinca hierro cemex
Concepto de crédito 66784439 Ferretería flores
Normalización (Recomendada)

Campo	Antes	Después
Fecha	08/ABR/2025 12:38:56 PM	2025-04-08
Tipo	transferencia fallida	failed
Cantidad	USD 26,596.00	26596.00
Moneda	USD	USD
bankAccountId	cuenta origen 891500015	4
Anotación	90982645	90982645
Observations	Concepto...	FALLIDA: Su transferencia no ha podido realizarse. Concepto débito: Corinca hierro cemex. Crédito: 66784439 Ferretería flores
Output (Recomendado)

2025-04-08|failed|26596.00|USD|4|90982645|FALLIDA: Su transferencia no ha podido realizarse. Concepto débito: Corinca hierro cemex. Crédito: 66784439 Ferretería flores

Octavo ejemplo

Raw OCR:
Usuario originador: HARWIN LAGUNA
Cuenta destino: 891500015
Monto transferido: USD 221.09
Fecha: 03/JUN/2025
Referencia: 96630294
Concepto de débito: Pago Clavos

Output:
2025-06-03|in|221.09|USD|4|96630294|Pago Clavos de HARWIN LAGUNA a cuenta 891500015

Próximo Paso

Si aparecen textos como Monto excede límite diario, Fondos insuficientes, Verifique si la transferencia se realizó, Su transferencia no ha podido realizarse, Ocurrió un problema al momento de solicitar el OTP, entonces la variable type debe ser igual a failed

`.trim();
