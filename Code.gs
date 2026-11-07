/**
 * VARIABLES GLOBALES
 */
var PARAMS_SHEET_ID = '1UJPefwcIOhF65ujBDE5FGEwJaHr-1-hddgC5BLWdGfE';
var BD_SHEET_ID = '1IcQA1aWWz6aOZijOP8Nwpmr3Tp6bk0O6n4VBhCXWRQ0';
//var BD_SHEET_ID = '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';
/**
 * BD independiente para presupuesto / proyección.
 * Dentro de este Spreadsheet debe existir una hoja llamada BD.
 *
 * Headers esperados:
 * idx;fecha_ejecucion;tipo;grupo;categoria;tercero;valor;descripcion;estado
 */
var PRESUPUESTO_SHEET_ID = '19qER6hdcfHqtH9MHxkE8yqX8I15ntuJ1SM0x8KMk4ag';

/**
 * NEPTUNO FZ - CONTROL DE MIGRACION
 *
 * false: conserva las rutas productivas actuales.
 * true: habilita los motores nuevos.
 *
 * NO ACTIVAR hasta completar frontend,
 * preparacion de BD y despliegue controlado.
 */
var TRANSACTION_V3_ENABLED = true;

/**
 * =========================
 * HELPERS DE RESPUESTA API
 * =========================
 */
function apiSuccess(data, message) {
  return {
    success: true,
    data: data || {},
    message: message || ''
  };
}

function apiError(message, data) {
  return {
    success: false,
    data: data || {},
    message: message || 'Ocurrió un error'
  };
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseBoolean(value) {
  return String(value).toLowerCase() === 'true';
}

/**
 * =========================
 * ENTRADA API
 * =========================
 */
function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : '';

    switch (action) {
      case 'getFormData':
        return jsonOutput(getFormData());

      case 'getListData':
        return jsonOutput(
          getListData({
            onlyCurrentYear: parseBoolean(e.parameter.onlyCurrentYear)
          })
        );
      
      case 'getProjectionData':
        return jsonOutput(
          getProjectionData({
            onlyCurrentYear: parseBoolean(e.parameter.onlyCurrentYear)
          })
        );

      default:
        return jsonOutput(apiError('Acción GET no válida', { action: action }));
    }
  } catch (error) {
    return jsonOutput(apiError(error.toString()));
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput(apiError('No se recibió body en la petición'));
    }

    var body = JSON.parse(e.postData.contents);
    var action = body.action || '';
    var payload = body.payload || {};

    switch (action) {
      case 'validateLogin':
        return jsonOutput(validateLogin(payload));

      
      case 'saveTransaction':
        return jsonOutput(
          TRANSACTION_V3_ENABLED
            ? saveTransactionV2_(
                BD_SHEET_ID,
                payload,
                true
              )
            : saveTransaction(payload)
        );

      case 'updateTransaction':
        if (!TRANSACTION_V3_ENABLED) {
          return jsonOutput(
            apiError(
              'Actualizacion no habilitada'
            )
          );
        }

        return jsonOutput(
          updateTransactionOnSheet_(
            payload,
            BD_SHEET_ID,
            false
          )
        );

      
      case 'softDeleteTransaction':
        return jsonOutput(
          TRANSACTION_V3_ENABLED
            ? softDeleteTransactionV2_(
                BD_SHEET_ID,
                payload,
                true
              )
            : softDeleteTransaction(payload)
        );


      case 'saveProjectionTasks':
        return jsonOutput(saveProjectionTasks(payload));

      case 'logAccess':
        return jsonOutput(logAccess(payload));

      default:
        return jsonOutput(apiError('Acción POST no válida', { action: action }));
    }
  } catch (error) {
    return jsonOutput(apiError(error.toString()));
  }
}


/**
 * =========================
 * UTILIDADES
 * =========================
 */
function normalizeText(text) {
  if (!text) return '';
  return String(text)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function getColIdx(headers, target) {
  var normalizedTarget = normalizeText(target);
  for (var i = 0; i < headers.length; i++) {
    if (normalizeText(headers[i]) === normalizedTarget) return i;
  }
  return -1;
}

function mapRowToObject(headers, row) {
  var obj = {};
  headers.forEach(function(header, index) {
    obj[header] = row[index] || '';
  });
  return obj;
}

function getRequiredColumnIndex(headers, columnName) {
  var idx = getColIdx(headers, columnName);
  if (idx === -1) {
    throw new Error('No se encontró la columna requerida: ' + columnName);
  }
  return idx;
}

/**
 * =========================
 * AUTENTICACIÓN
 * =========================
 */
function checkInitialUser() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) return apiError('No email');

    var data = Sheets.Spreadsheets.Values.get(PARAMS_SHEET_ID, 'Users!A:Z').values;
    if (!data || data.length < 2) return apiError('No hay usuarios configurados');

    var headers = data[0];
    var mailIdx = getRequiredColumnIndex(headers, 'mail');
    var nameIdx = getRequiredColumnIndex(headers, 'user_name');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][mailIdx]) === String(email)) {
        return apiSuccess({
          user: {
            email: email,
            name: data[i][nameIdx] || ''
          }
        });
      }
    }

    return apiError('Usuario no encontrado');
  } catch (e) {
    return apiError(e.toString());
  }
}

function validateLogin(payload) {
  try {
    var username = payload.username || '';
    var password = payload.password || '';

    if (!username || !password) {
      return apiError('Username y password son requeridos');
    }

    var data = Sheets.Spreadsheets.Values.get(PARAMS_SHEET_ID, 'Users!A:Z').values;
    if (!data || data.length < 2) {
      return apiError('No hay usuarios configurados');
    }

    var headers = data[0];
    var nameIdx = getRequiredColumnIndex(headers, 'user_name');
    var pswIdx = getRequiredColumnIndex(headers, 'psw');

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][nameIdx]) === String(username) &&
        String(data[i][pswIdx]) === String(password)
      ) {
        return apiSuccess({
          user: {
            name: data[i][nameIdx] || ''
          }
        }, 'Login correcto');
      }
    }

    return apiError('Credenciales inválidas');
  } catch (e) {
    return apiError(e.toString());
  }
}

/**
 * =========================
 * PARÁMETROS / MAESTROS
 * =========================
 */
function getFormData() {
  try {
    var ranges = ['terceros!A:A', 'forma_pago!A:Z', 'tabla_pg!A:Z', 'flash!A:Z'];
    var response = Sheets.Spreadsheets.Values.batchGet(PARAMS_SHEET_ID, { ranges: ranges });

    var res = {
      terceros: [],
      formasPago: [],
      tablaPG: [],
      flashData: []
    };

    // 1. Terceros
    var vTerceros = response.valueRanges[0].values;
    if (vTerceros && vTerceros.length > 1) {
      res.terceros = vTerceros
        .slice(1)
        .map(function(r) { return r[0]; })
        .filter(String);
    }

    // 2. Formas de pago
    var vFP = response.valueRanges[1].values;
    if (vFP && vFP.length > 0) {
      var hFP = vFP[0];
      var iM = getRequiredColumnIndex(hFP, 'medio');
      var iC = getRequiredColumnIndex(hFP, 'clase');
      var iT = getRequiredColumnIndex(hFP, 'tipo_fuente');
      var iE = getRequiredColumnIndex(hFP, 'estado_fuente');
      var iMon = getColIdx(hFP, 'moneda');
      var iGP = getColIdx(hFP, 'genera_pasivo');
      var iTFuente = getColIdx(hFP, 'tercero_fuente');

      res.formasPago = vFP.slice(1).map(function(r) {
        return {
          medio: r[iM] || '',
          clase: r[iC] || '',
          tipo_fuente: r[iT] || '',
          estado_fuente: r[iE] || 'h',
          moneda: iMon >= 0 ? (r[iMon] || 'COP') : 'COP',
          genera_pasivo: iGP >= 0 ? (r[iGP] || 'n') : 'n',
          tercero_fuente: iTFuente >= 0 ? (r[iTFuente] || '') : ''
        };
      }).filter(function(item) {
        return item.medio;
      });
    }

    // 3. Tabla P&G
    var vPG = response.valueRanges[2].values;
    if (vPG && vPG.length > 0) {
      var hPG = vPG[0];
      var idxTipo = getRequiredColumnIndex(hPG, 'tipo');
      var idxGrupo = getRequiredColumnIndex(hPG, 'grupo');
      var idxCat = getRequiredColumnIndex(hPG, 'categoria');

      res.tablaPG = vPG.slice(1).map(function(r) {
        return {
          tipo: r[idxTipo] || '',
          grupo: r[idxGrupo] || '',
          categoria: r[idxCat] || ''
        };
      }).filter(function(item) {
        return item.tipo;
      });
    }

    // 4. Flash
    var vFlash = response.valueRanges[3].values;
    if (vFlash && vFlash.length > 0) {
      var hF = vFlash[0];

      var idxFTipo = getRequiredColumnIndex(hF, 'tipo');
      var idxFGrupo = getRequiredColumnIndex(hF, 'grupo');
      var idxFCategoria = getRequiredColumnIndex(hF, 'categoria');
      var idxFMedio = getRequiredColumnIndex(hF, 'medio');
      var idxFTercero = getRequiredColumnIndex(hF, 'tercero');
      var idxFDescripcion = getRequiredColumnIndex(hF, 'descripcion');

      res.flashData = vFlash.slice(1).map(function(r) {
        return {
          tipo: r[idxFTipo] || '',
          grupo: r[idxFGrupo] || '',
          categoria: r[idxFCategoria] || '',
          medio: r[idxFMedio] || '',
          tercero: r[idxFTercero] || '',
          descripcion: r[idxFDescripcion] || '',
          valor: '',
          isSaving: false,
          checked: false
        };
      }).filter(function(item) {
        return item.tipo;
      });
    }

    return apiSuccess(res, 'Datos de formulario cargados');
  } catch (e) {
    return apiError('Error en getFormData: ' + e.toString());
  }
}

function appendTransactionRow(sheetHeaders, record, newId) {

  return sheetHeaders.map(function(header, index) {

    if (index === 0) return newId;

    return record.hasOwnProperty(header)
      ? record[header]
      : '';

  });

}

function buildPasivoRecord(originalRecord) {

  return {
    fecha_registro: originalRecord.fecha_registro,
    fecha_pago: originalRecord.fecha_pago,

    tipo: 'Pasivo',

    grupo: originalRecord.tipo_fuente,

    categoria: originalRecord.medio,

    medio: 'Bridge',

    clase: originalRecord.clase,

    tipo_fuente: originalRecord.tipo_fuente,

    tercero: originalRecord.tercero_fuente,

    valor: originalRecord.valor,

    descripcion: originalRecord.descripcion,

    estado: 'h'
  };

}

/**
 * =========================
 * TRANSACCIONES
 * =========================
 *
 * payload esperado:
 * {
 *   record: {
 *     campo1: valor1,
 *     campo2: valor2,
 *     ...
 *   }
 * }
 */


/**
 * IDEMPOTENCIA V3 - HELPERS
 * Solo utilizados por saveTransactionV2_.
 */

function hashRequestV3_(record) {
  var campos = [
    'fecha_registro',
    'fecha_pago',
    'tipo',
    'grupo',
    'categoria',
    'medio',
    'tercero',
    'valor',
    'descripcion'
  ];

  var canonical = campos.map(function(campo) {
    var value = record[campo];

    if (campo === 'valor') {
      return Number(value);
    }

    return String(value == null ? '' : value);
  });

  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(canonical),
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(b) {
    return ('0' + ((b + 256) % 256).toString(16))
      .slice(-2);
  }).join('');
}


function readIdempotencyLogV3_(spreadsheetId) {

  var response = Sheets.Spreadsheets.Values.get(
    spreadsheetId,
    "'Idempotency_Log'!A:G",
    { valueRenderOption: 'UNFORMATTED_VALUE' }
  );

  var rows = response.values || [];

  var expected = [
    'request_id',
    'payload_hash',
    'status',
    'principal_id',
    'pasivo_id',
    'created_at',
    'result_json'
  ];

  if (!rows.length ||
      expected.some(function(h, i) {
        return rows[0][i] !== h;
      })) {
    throw new Error(
      'Estructura inválida de Idempotency_Log'
    );
  }

  return rows.slice(1).filter(function(row) {
    return String(row[0] || '').trim() !== '';
  }).map(function(row) {
    return {
      requestId: String(row[0]),
      hash: String(row[1] || ''),
      status: String(row[2] || ''),
      principalId: String(row[3] || ''),
      pasivoId: String(row[4] || ''),
      createdAt: String(row[5] || ''),
      resultJson: String(row[6] || '')
    };
  });
}


function verifyIdempotencyRowsV3_(
  existingData, requestId, principalId, pasivoId
) {

  var headers = existingData[0] || [];

  var requestCol = getRequiredColumnIndex(
    headers, 'request_id'
  );

  var idCol = getRequiredColumnIndex(
    headers, 'idx'
  );

  var matches = existingData.slice(1)
    .filter(function(row) {
      return String(row[requestCol] || '') === requestId;
    })
    .map(function(row) {
      return String(row[idCol] || '');
    });

  var expectedIds = [String(principalId)];

  if (pasivoId) {
    expectedIds.push(String(pasivoId));
  }

  return matches.length === expectedIds.length &&
    expectedIds.every(function(id) {
      return matches.filter(function(x) {
        return x === id;
      }).length === 1;
    });
}


/**
 * NEPTUNO FZ - GUARDADO SEGURO V2
 * Primera etapa: cálculo de ID y planificación.
 * No realiza escrituras.
 */
function planSaveTransactionV2_(spreadsheetId, record) {
  if (!spreadsheetId) {
    throw new Error('Spreadsheet ID requerido');
  }

  if (!record || typeof record !== 'object' ||
      Array.isArray(record) || !Object.keys(record).length) {
    throw new Error('Registro inválido');
  }

  var response = Sheets.Spreadsheets.Values.get(
    spreadsheetId,
    'BD!A:Z',
    { valueRenderOption: 'FORMATTED_VALUE' }
  );

  var data = response.values || [];

  if (!data.length) {
    throw new Error('BD sin encabezados');
  }

  var headers = data[0];
  var idxCol = getRequiredColumnIndex(headers, 'idx');

  var maxId = 0;
  var ids = {};
  var duplicados = [];
  var filasConDatos = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i] || [];

    if (!row.some(function(v) {
      return String(v == null ? '' : v).trim() !== '';
    })) {
      continue;
    }

    filasConDatos++;

    var rawId = String(
      row[idxCol] == null ? '' : row[idxCol]
    ).trim();

    if (!/^[1-9]\d*$/.test(rawId) ||
        !Number.isSafeInteger(Number(rawId))) {
      throw new Error(
        'ID inválido en fila física ' + (i + 1)
      );
    }

    var id = Number(rawId);

    if (ids[id]) {
      duplicados.push(id);
    }

    ids[id] = true;
    maxId = Math.max(maxId, id);
  }

  if (duplicados.length) {
    throw new Error(
      'IDs duplicados detectados: ' +
      [...new Set(duplicados)].join(', ')
    );
  }

  var generaPasivo =
    String(record.genera_pasivo || 'n')
      .toLowerCase() === 's';

  var nextId = maxId + 1;

  if (!Number.isSafeInteger(nextId + (generaPasivo ? 1 : 0))) {
    throw new Error('Límite de IDs alcanzado');
  }

  var principal = appendTransactionRow(
    headers,
    record,
    nextId
  );

  var rows = [principal];

  if (generaPasivo) {
    var pasivo = buildPasivoRecord(record);
    pasivo.request_id = record.request_id;

    rows.push(
      appendTransactionRow(
        headers,
        pasivo,
        nextId + 1
      )
    );
  }

  return {
    headers: headers,
    rows: rows,
    nextId: nextId,
    pasivoId: generaPasivo ? nextId + 1 : null,
    maxIdAnterior: maxId,
    filasConDatos: filasConDatos,
    registrosAInsertar: rows.length,
    totalFilasLeidas: data.length
  };
}


/**
 * NEPTUNO FZ - GUARDADO SEGURO V2
 *
 * Uso exclusivo interno durante las pruebas.
 * No conectado a doPost().
 *
 * Inserta filas y valores mediante una única
 * llamada a Sheets.Spreadsheets.batchUpdate().
 */


function isValidISODate_(value) {
  if (typeof value !== 'string' ||
      !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
    return false;
  }

  var parts = value.split('-').map(Number);

  var date = new Date(Date.UTC(
    parts[0], parts[1] - 1, parts[2]
  ));

  return date.getUTCFullYear() === parts[0] &&
    date.getUTCMonth() === parts[1] - 1 &&
    date.getUTCDate() === parts[2];
}


function resolvePaymentSourceV2_(record) {

  var response = getFormData();

  if (!response.success) {
    throw new Error(
      'No se pudo consultar el catálogo de medios'
    );
  }

  var medio = String(record.medio || '').trim();

  var matches = (response.data.formasPago || [])
    .filter(function(source) {
      return String(source.medio || '').trim() === medio;
    });

  if (matches.length !== 1) {
    throw new Error(
      'Medio inexistente o ambiguo: ' + medio
    );
  }

  var source = matches[0];

  if (String(source.estado_fuente || '').toLowerCase() !== 'h') {
    throw new Error('El medio de pago no está activo');
  }

  var genera = String(
    source.genera_pasivo || 'n'
  ).trim().toLowerCase();

  if (genera !== 's' && genera !== 'n') {
    throw new Error(
      'Configuración genera_pasivo inválida'
    );
  }

  if (genera === 's' &&
      (!source.tipo_fuente || !source.tercero_fuente)) {
    throw new Error(
      'Fuente generadora de Pasivo incompleta'
    );
  }

  return {
    medio: source.medio,
    clase: source.clase,
    tipo_fuente: source.tipo_fuente,
    tercero_fuente: source.tercero_fuente,
    genera_pasivo: genera
  };
}




function saveTransactionV2_(
  spreadsheetId,
  payload,
  allowProduction
) {

  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    payload = payload || {};

    
    if (!spreadsheetId) {
      return apiError('Spreadsheet ID requerido');
    }

    if (spreadsheetId === BD_SHEET_ID &&
        !(allowProduction === true &&
          TRANSACTION_V3_ENABLED === true)) {
      return apiError(
        'Guardado V3 productivo no habilitado'
      );
    }


    var requestId = String(
      payload.requestId == null ? '' : payload.requestId
    ).trim();

    if (!/^[A-Za-z0-9_-]{16,100}$/.test(requestId)) {
      return apiError('requestId inválido o ausente');
    }

    
    var record = payload.record || {};


    if (!record || typeof record !== 'object' ||
        Array.isArray(record) || !Object.keys(record).length) {
      return apiError('Registro inválido');
    }

    // Evitar escritura con datos de prueba incompletos.
    var required = [
      'fecha_registro', 'fecha_pago', 'tipo',
      'grupo', 'categoria', 'medio',
      'tercero', 'valor', 'descripcion'
    ];

    for (var r = 0; r < required.length; r++) {
      var key = required[r];

      if (record[key] === null ||
          record[key] === undefined ||
          String(record[key]).trim() === '') {
        return apiError('Campo obligatorio: ' + key);
      }
    }

    
    if (
      !isValidISODate_(record.fecha_registro) ||
      !isValidISODate_(record.fecha_pago)
    ) {
      return apiError('Fecha inválida o inexistente');
    }


    if (typeof record.valor === 'boolean' ||
        !Number.isFinite(Number(record.valor))) {
      return apiError('Valor no numérico');
    }

    
    
    // Normalizar los datos de la solicitud.
    // No consultar todavía el catálogo.
    record = Object.assign({}, record, {
      valor: Number(record.valor),
      estado: 'h'
    });

    // Firma inmutable de los nueve campos
    // financieros originales.
    var firmaActual = hashRequestV3_(record);

    // Bloquear antes de consultar solicitudes previas.
    lock.waitLock(10000);
    locked = true;


    
    // Detectar solicitudes previamente procesadas.
    var existingData = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z',
      { valueRenderOption: 'UNFORMATTED_VALUE' }
    ).values || [];

    if (!existingData.length) {
      return apiError('BD sin encabezados');
    }

    var existingHeaders = existingData[0];
    var requestCol = getRequiredColumnIndex(
      existingHeaders, 'request_id'
    );
    var existingIdCol = getRequiredColumnIndex(
      existingHeaders, 'idx'
    );

    var previous = [];

    for (var q = 1; q < existingData.length; q++) {
      var row = existingData[q] || [];

      if (String(row[requestCol] || '') === requestId) {
        previous.push({
          id: row[existingIdCol],
          row: q + 1
        });
      }
    }

    
    
    
    // Consultar registro técnico dentro del bloqueo.
    var logEntries = readIdempotencyLogV3_(
      spreadsheetId
    );

    var logMatches = logEntries.filter(function(entry) {
      return entry.requestId === requestId;
    });

    if (logMatches.length > 1) {
      return apiError(
        'requestId duplicado en registro técnico. ' +
        'Revisión manual requerida.',
        { requestId: requestId, integrityError: true }
      );
    }

    if (logMatches.length === 1) {

      var entry = logMatches[0];

      if (entry.hash !== firmaActual) {
        return apiError(
          'Conflicto: requestId reutilizado con datos diferentes.',
          {
            requestId: requestId,
            alreadyProcessed: true,
            samePayload: false
          }
        );
      }

      if (entry.status !== 'COMMITTED') {
        return apiError(
          'Solicitud con estado técnico no confirmado.',
          {
            requestId: requestId,
            integrityError: true
          }
        );
      }

      var integrity = verifyIdempotencyRowsV3_(
        existingData,
        requestId,
        entry.principalId,
        entry.pasivoId
      );

      if (!integrity) {
        return apiError(
          'Solicitud registrada, pero movimientos inconsistentes. ' +
          'Revisión manual requerida.',
          {
            requestId: requestId,
            integrityError: true
          }
        );
      }

      var originalResult;

      try {
        originalResult = JSON.parse(entry.resultJson);
      } catch (parseError) {
        return apiError(
          'Resultado original ilegible. Revisión manual requerida.',
          {
            requestId: requestId,
            integrityError: true
          }
        );
      }

      if (!originalResult ||
          String(originalResult.newId) !== entry.principalId ||
          String(originalResult.pasivoId == null
            ? '' : originalResult.pasivoId) !== entry.pasivoId) {
        return apiError(
          'Resultado original inconsistente.',
          {
            requestId: requestId,
            integrityError: true
          }
        );
      }

      return apiSuccess(
        Object.assign({}, originalResult, {
          requestId: requestId,
          alreadyProcessed: true,
          replayed: true,
          samePayload: true
        }),
        'Solicitud previamente registrada. Resultado recuperado.'
      );
    }

    // Solicitudes históricas anteriores al registro técnico.
    // No reconstruir firmas a partir de movimientos editables.
    if (previous.length > 0) {
      return apiError(
        'Solicitud histórica sin firma inmutable. ' +
        'No se permite repetir automáticamente.',
        {
          requestId: requestId,
          alreadyProcessed: true,
          legacyRequest: true
        }
      );
    }




    
    // ==========================================
    // SOLICITUD NUEVA - VALIDAR CATÁLOGO
    // ==========================================

    // Este bloque solo se ejecuta después
    // de descartar solicitudes ya procesadas.

    var source;

    try {
      source = resolvePaymentSourceV2_(record);
    } catch (sourceError) {
      return apiError(sourceError.message);
    }

    // El catálogo sigue siendo la fuente
    // autorizada para clase, tipo_fuente,
    // tercero_fuente y genera_pasivo.
    record = Object.assign({}, record, source, {
      valor: Number(record.valor),
      estado: 'h',
      request_id: requestId
    });

    // Planificar dentro del bloqueo.
    var plan = planSaveTransactionV2_(
      spreadsheetId,
      record
    );


    var metadata = Sheets.Spreadsheets.get(
      spreadsheetId,
      {
        fields: 'sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)),basicFilter)'
      }
    );

    var sheet = (metadata.sheets || []).find(
      function(s) {
        return s.properties.title === 'BD';
      }
    );

    if (!sheet) {
      return apiError('No existe la hoja BD');
    }

    var sheetId = sheet.properties.sheetId;
    var grid = sheet.properties.gridProperties;

    
    var logSheet = (metadata.sheets || []).find(
      function(s) {
        return s.properties.title === 'Idempotency_Log';
      }
    );

    if (!logSheet) {
      return apiError(
        'No existe la hoja técnica Idempotency_Log'
      );
    }

    if (logSheet.properties.gridProperties.rowCount < 2 ||
        logSheet.properties.gridProperties.columnCount < 7) {
      return apiError(
        'Idempotency_Log no tiene dimensiones suficientes'
      );
    }


    if (plan.headers.length > grid.columnCount) {
      return apiError(
        'La hoja no tiene suficientes columnas'
      );
    }

    // Determinar última fila física con contenido.
    var read = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z'
    );

    var allRows = read.values || [];
    var lastOccupied = 0;

    for (var i = 0; i < allRows.length; i++) {
      if ((allRows[i] || []).some(function(v) {
        return String(v == null ? '' : v).trim() !== '';
      })) {
        lastOccupied = i + 1;
      }
    }

    if (lastOccupied < 1) {
      return apiError('Encabezados no encontrados');
    }

    // Posición de inserción, índice base cero.
    var insertionIndex = lastOccupied;
    var count = plan.rows.length;

    // Construir valores para updateCells.
    var rowData = plan.rows.map(function(row) {
      return {
        values: row.map(function(value) {
          if (value === null || value === undefined) {
            return { userEnteredValue: {
              stringValue: ''
            }};
          }

          if (typeof value === 'number') {
            return { userEnteredValue: {
              numberValue: value
            }};
          }

          if (typeof value === 'boolean') {
            return { userEnteredValue: {
              boolValue: value
            }};
          }

          return { userEnteredValue: {
            stringValue: String(value)
          }};
        })
      };
    });

    // Inserción y escritura en una operación.
    var requests = [
      {
        insertDimension: {
          range: {
            sheetId: sheetId,
            dimension: 'ROWS',
            startIndex: insertionIndex,
            endIndex: insertionIndex + count
          },
          inheritFromBefore: true
        }
      },
      {
        updateCells: {
          start: {
            sheetId: sheetId,
            rowIndex: insertionIndex,
            columnIndex: 0
          },
          rows: rowData,
          fields: 'userEnteredValue'
        }
      }
    ];

    
    // ==========================================
    // V2 - AMPLIAR FILTRO BÁSICO SI ES NECESARIO
    // ==========================================

    var filtroActual = sheet.basicFilter;

    if (filtroActual && filtroActual.range) {

      var rangoActual = filtroActual.range;

      // Índice final exclusivo, base cero.
      var nuevoFin = insertionIndex + count;

      var finActual = rangoActual.endRowIndex;

      if (finActual != null && finActual < nuevoFin) {

        var nuevoFiltro = JSON.parse(
          JSON.stringify(filtroActual)
        );

        nuevoFiltro.range.endRowIndex = nuevoFin;

        // Preservar el rango horizontal y la configuración
        // del filtro. Solo ampliamos su límite inferior.
        requests.push({
          setBasicFilter: {
            filter: nuevoFiltro
          }
        });
      }
    }

    
    // ==========================================
    // V3 - REGISTRO TÉCNICO INMUTABLE
    // ==========================================

    var originalResult = {
      newId: plan.nextId,
      record: plan.rows[0],
      pasivoId: plan.pasivoId,
      insertedRows: count,
      firstRow: insertionIndex + 1,
      lastRow: insertionIndex + count,
      verified: true
    };

    var createdAt = Utilities.formatDate(
      new Date(),
      'UTC',
      "yyyy-MM-dd'T'HH:mm:ss'Z'"
    );

    var logValues = [
      requestId,
      firmaActual,
      'COMMITTED',
      plan.nextId,
      plan.pasivoId == null ? '' : plan.pasivoId,
      createdAt,
      JSON.stringify(originalResult)
    ];

    var logRowData = {
      values: logValues.map(function(value) {
        if (typeof value === 'number') {
          return {
            userEnteredValue: { numberValue: value }
          };
        }

        return {
          userEnteredValue: {
            stringValue: String(value)
          }
        };
      })
    };

    // Insertar siempre debajo del encabezado.
    // Las entradas anteriores se desplazan sin sobrescribirse.
    requests.push({
      insertDimension: {
        range: {
          sheetId: logSheet.properties.sheetId,
          dimension: 'ROWS',
          startIndex: 1,
          endIndex: 2
        },
        inheritFromBefore: true
      }
    });

    requests.push({
      updateCells: {
        start: {
          sheetId: logSheet.properties.sheetId,
          rowIndex: 1,
          columnIndex: 0
        },
        rows: [logRowData],
        fields: 'userEnteredValue'
      }
    });


    Sheets.Spreadsheets.batchUpdate(
      { requests: requests },
      spreadsheetId
    );

    // Verificación posterior de las filas insertadas.
    var firstRow = insertionIndex + 1;
    var lastRow = insertionIndex + count;

    var range = 'BD!A' + firstRow + ':' +
      columnToLetter(plan.headers.length) +
      lastRow;

    var verification = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      range,
      { valueRenderOption: 'UNFORMATTED_VALUE' }
    );

    var savedRows = verification.values || [];
    var idxCol = getRequiredColumnIndex(
      plan.headers, 'idx'
    );

    var verified = savedRows.length === count;

    if (verified) {
      for (var j = 0; j < count; j++) {
        var expectedRow = plan.rows[j];
        var actualRow = savedRows[j] || [];

        if (String(actualRow[idxCol]) !==
            String(expectedRow[idxCol])) {
          verified = false;
          break;
        }

        for (var c = 0; c < plan.headers.length; c++) {
          var expectedValue = expectedRow[c] == null
            ? '' : String(expectedRow[c]);

          var actualValue = actualRow[c] == null
            ? '' : String(actualRow[c]);

          if (expectedValue !== actualValue) {
            verified = false;
            break;
          }
        }

        if (!verified) break;
      }
    }

    if (!verified) {
      return apiError(
        'La escritura fue enviada, pero la verificación falló. ' +
        'Revisar BD antes de reintentar.',
        {
          newId: plan.nextId,
          firstRow: firstRow,
          lastRow: lastRow,
          writeAttempted: true
        }
      );
    }

    
    // Confirmar que el registro técnico existe
    // y conserva la firma y los IDs originales.
    var logCheck = readIdempotencyLogV3_(
      spreadsheetId
    ).filter(function(entry) {
      return entry.requestId === requestId;
    });

    var logVerified =
      logCheck.length === 1 &&
      logCheck[0].hash === firmaActual &&
      logCheck[0].status === 'COMMITTED' &&
      logCheck[0].principalId === String(plan.nextId) &&
      logCheck[0].pasivoId === String(
        plan.pasivoId == null ? '' : plan.pasivoId
      ) &&
      logCheck[0].resultJson ===
        JSON.stringify(originalResult);

    if (!logVerified) {
      return apiError(
        'Escritura enviada, pero registro técnico no verificado. ' +
        'No reintentar con una clave nueva.',
        {
          requestId: requestId,
          newId: plan.nextId,
          writeAttempted: true,
          integrityError: true
        }
      );
    }


    
    return apiSuccess(
      Object.assign({}, originalResult, {
        requestId: requestId,
        alreadyProcessed: false,
        replayed: false
      }),
      'Transacción guardada correctamente'
    );


  } catch (error) {
    return apiError(
      'Error en saveTransactionV2_: ' + error.toString()
    );
  } finally {
    if (locked) lock.releaseLock();
  }
}



function saveTransaction(payload) {
  var lock = LockService.getScriptLock();

  try {

    var record = payload.record || {};


    if (!record || Object.keys(record).length === 0) {
      return apiError('No se recibió el objeto record');
    }

    lock.waitLock(10000);

    var sheetData = Sheets.Spreadsheets.Values.get(BD_SHEET_ID, 'BD!A:Z').values;
    if (!sheetData || sheetData.length === 0) {
      return apiError('La hoja BD no tiene encabezados');
    }

    var headers = sheetData[0];
    var currentData = sheetData.slice(1);

    var newId = 1;
    if (currentData.length > 0) {
      var lastRow = currentData[currentData.length - 1];
      var lastId = lastRow[0];
      newId = isNaN(parseInt(lastId, 10)) ? (currentData.length + 1) : (parseInt(lastId, 10) + 1);
    }

    var rowToSave = appendTransactionRow(
      headers,
      record,
      newId
    );

    var valueRange = Sheets.newValueRange();
    valueRange.values = [rowToSave];

    Sheets.Spreadsheets.Values.append(
      valueRange,
      BD_SHEET_ID,
      'BD!A1',
      { valueInputOption: 'USER_ENTERED' }
    );

    // =======================================
    // GENERAR PASIVO AUTOMÁTICO
    // =======================================

    if (String(record.genera_pasivo || 'n') === 's') {

      var pasivoRecord = buildPasivoRecord(record);

      var pasivoId = newId + 1;

      var pasivoRow = appendTransactionRow(
        headers,
        pasivoRecord,
        pasivoId
      );

      var pasivoRange = Sheets.newValueRange();
      pasivoRange.values = [pasivoRow];

      Sheets.Spreadsheets.Values.append(
        pasivoRange,
        BD_SHEET_ID,
        'BD!A1',
        { valueInputOption: 'USER_ENTERED' }
      );

    }
    return apiSuccess(
      { newId: newId, record: rowToSave },
      'Transacción guardada correctamente'
    );
  } catch (e) {
    return apiError(e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * payload opcional:
 * {
 *   onlyCurrentYear: true
 * }
 */
function getListData(payload) {
  try {
    payload = payload || {};
    var onlyCurrentYear = !!payload.onlyCurrentYear;

    var data = Sheets.Spreadsheets.Values.get(BD_SHEET_ID, 'BD!A:Z').values;
    if (!data || data.length <= 1) {
      return apiSuccess({ items: [] }, 'Sin datos');
    }

    var headers = data[0];
    var estadoIdx = getRequiredColumnIndex(headers, 'estado');
    var fechaRegIdx = getRequiredColumnIndex(headers, 'fecha_registro');

    var idIdx = 0;
    var idxIdx = getColIdx(headers, 'idx');
    var currentYear = new Date().getFullYear();

    var result = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      if ((row[estadoIdx] || '') !== 'h') continue;

      var rowDateStr = row[fechaRegIdx] || '';
      var rowYear = parseInt(String(rowDateStr).split('-')[0], 10);

      if (onlyCurrentYear && rowYear !== currentYear) continue;

      result.push(mapRowToObject(headers, row));
    }

    result.sort(function(a, b) {
      var aOrder = idxIdx >= 0 ? parseInt(a[headers[idxIdx]], 10) || 0 : parseInt(a[headers[idIdx]], 10) || 0;
      var bOrder = idxIdx >= 0 ? parseInt(b[headers[idxIdx]], 10) || 0 : parseInt(b[headers[idIdx]], 10) || 0;
      return bOrder - aOrder;
    });

    return apiSuccess({ items: result }, 'Listado cargado correctamente');
  } catch (e) {
    return apiError(e.toString());
  }
}



/**
 * NEPTUNO FZ - ACTUALIZAR MOVIMIENTO
 *
 * payload:
 * {
 *   id: 123,
 *   record: { descripcion: "Nuevo texto" },
 *   expected: { descripcion: "Texto anterior" },
 *   dryRun: true
 * }
 *
 * dryRun=true: valida sin escribir.
 * Para escribir: expected es obligatorio.
 */
function updateTransaction(payload) {
  return updateTransactionOnSheet_(
    payload,
    BD_SHEET_ID,
    false
  );
}

function updateTransactionOnSheet_(payload, spreadsheetId, allowManualBridgeTest) {
  var lock = LockService.getScriptLock();
  var locked = false;

  try {
    payload = payload || {};

    var id = String(payload.id == null ? '' : payload.id).trim();
    var patch = payload.record;
    var expected = payload.expected;
    var dryRun = payload.dryRun === true;

    if (!/^[1-9]\d*$/.test(id)) {
      return apiError('ID inválido');
    }

    if (!patch || typeof patch !== 'object' ||
        Array.isArray(patch) || !Object.keys(patch).length) {
      return apiError('No se recibieron campos para actualizar');
    }

    var allowed = [
      'fecha_registro', 'fecha_pago',
      'tipo', 'grupo', 'categoria',
      'medio', 'tercero', 'valor', 'descripcion'
    ];

    var keys = Object.keys(patch);

    for (var k = 0; k < keys.length; k++) {
      if (allowed.indexOf(keys[k]) === -1) {
        return apiError('Campo no autorizado: ' + keys[k]);
      }
    }

    if (!dryRun && (!expected || typeof expected !== 'object' ||
        Array.isArray(expected))) {
      return apiError(
        'Se requieren los valores originales para actualizar'
      );
    }

    lock.waitLock(10000);
    locked = true;

    
    var response = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z',
      { valueRenderOption: 'FORMATTED_VALUE' }
    );


    var data = response.values || [];

    if (data.length < 2) {
      return apiError('BD sin registros');
    }

    var headers = data[0];
    var idCol = getRequiredColumnIndex(headers, 'idx');
    var estadoCol = getRequiredColumnIndex(headers, 'estado');

    // Buscar ID real, incluyendo filas filtradas u ocultas.
    var matches = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol] == null
          ? '' : data[i][idCol]).trim() === id) {
        matches.push(i);
      }
    }

    if (matches.length === 0) {
      return apiError('ID no encontrado', { id: id });
    }

    if (matches.length !== 1) {
      return apiError(
        'ID duplicado. Actualización bloqueada',
        { id: id, coincidencias: matches.length }
      );
    }

    var index = matches[0];
    var rowNumber = index + 1;
    var currentRow = data[index];

    if (String(currentRow[estadoCol] || '').trim() !== 'h') {
      return apiError('El movimiento no está activo');
    }

    var current = {};
    headers.forEach(function(h, j) {
      current[h] = currentRow[j] == null ? '' : currentRow[j];
    });

    // Protección frente a pasivos automáticos.
    // Consultar catálogo de fuentes de pago.
    var formResponse = getFormData();

    if (!formResponse.success) {
      return apiError('No se pudo validar el catálogo de medios');
    }

    var sources = formResponse.data.formasPago || [];

    function generatesLiability(medio) {
      return sources.some(function(s) {
        return String(s.medio) === String(medio) &&
          String(s.genera_pasivo).toLowerCase() === 's';
      });
    }

    var nextMedio = Object.prototype.hasOwnProperty.call(
      patch, 'medio'
    ) ? patch.medio : current.medio;

    var possibleAutomaticLiability =
      String(current.tipo).toLowerCase() === 'pasivo' &&
      String(current.medio).toLowerCase() === 'bridge';

    
    var isAuthorizedTest =
      allowManualBridgeTest === true &&
      spreadsheetId === '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0' &&
      id === '3662' &&
      keys.length === 1 &&
      keys[0] === 'descripcion' &&
      String(current.descripcion) === 'Test' &&
      String(current.tipo).toLowerCase() === 'pasivo' &&
      String(current.medio).toLowerCase() === 'bridge' &&
      String(current.estado) === 'h';

    if (!isAuthorizedTest && (
      generatesLiability(current.medio) ||
      generatesLiability(nextMedio) ||
      possibleAutomaticLiability ||
      String(current.genera_pasivo || '').toLowerCase() === 's'
    )) {
      return apiError(
        'Movimiento con posible pasivo automático. ' +
        'Edición protegida hasta implementar actualización vinculada.'
      );
    }


    // Construir estado final sin modificar el original.
    var merged = Object.assign({}, current, patch);

    var required = [
      'fecha_registro', 'fecha_pago',
      'tipo', 'grupo', 'categoria',
      'medio', 'tercero', 'descripcion'
    ];

    for (var r = 0; r < required.length; r++) {
      var field = required[r];

      if (!String(merged[field] == null
          ? '' : merged[field]).trim()) {
        return apiError('Campo obligatorio: ' + field);
      }
    }

    var datePattern = /^\d{4}-\d{2}-\d{2}$/;

    ['fecha_registro', 'fecha_pago'].forEach(function(field) {
      var value = String(merged[field]);
      if (!datePattern.test(value)) {
        throw new Error('Fecha inválida: ' + field);
      }

      var parts = value.split('-').map(Number);
      var date = new Date(Date.UTC(
        parts[0], parts[1] - 1, parts[2]
      ));

      if (
        date.getUTCFullYear() !== parts[0] ||
        date.getUTCMonth() !== parts[1] - 1 ||
        date.getUTCDate() !== parts[2]
      ) {
        throw new Error('Fecha inexistente: ' + field);
      }
    });

    var amount = merged.valor;

    if (amount === '' || amount == null ||
        typeof amount === 'boolean' ||
        !Number.isFinite(Number(amount))) {
      return apiError('El valor debe ser numérico');
    }

    merged.valor = Number(amount);

    // Comprobar que el usuario editó la versión esperada.
    if (!dryRun) {
      var expectedKeys = Object.keys(expected);

      if (!expectedKeys.length) {
        return apiError('Valores originales vacíos');
      }

      // Exigir los nueve campos originales.
      for (var e = 0; e < allowed.length; e++) {
        if (!Object.prototype.hasOwnProperty.call(
          expected, allowed[e]
        )) {
          return apiError(
            'Falta valor original: ' + allowed[e]
          );
        }
      }

      for (var x = 0; x < allowed.length; x++) {
        var key = allowed[x];

        if (String(current[key] == null ? '' : current[key]) !==
            String(expected[key] == null ? '' : expected[key])) {
          return apiError(
            'El movimiento cambió desde su lectura. Recarga antes de editar.',
            { id: id, campo: key }
          );
        }
      }
    }

    // Solo actualizar celdas realmente modificadas.
    var changes = [];

    keys.forEach(function(key) {
      var col = getRequiredColumnIndex(headers, key);
      var nextValue = merged[key];
      var oldValue = current[key];

      if (String(nextValue) !== String(oldValue)) {
        changes.push({
          field: key,
          column: col + 1,
          value: nextValue
        });
      }
    });

    var result = {
      id: id,
      row: rowNumber,
      dryRun: dryRun,
      changedFields: changes.map(function(c) {
        return c.field;
      }),
      updated: false
    };

    if (dryRun) {
      return apiSuccess(
        result,
        'Validación correcta. No se escribió en BD.'
      );
    }

    if (!changes.length) {
      return apiSuccess(result, 'Sin cambios que actualizar');
    }

    var requests = changes.map(function(change) {
      return {
        range: 'BD!' +
          columnToLetter(change.column) + rowNumber,
        values: [[change.value]]
      };
    });

    
    Sheets.Spreadsheets.Values.batchUpdate(
      {
        valueInputOption: 'RAW',
        data: requests
      },
      spreadsheetId
    );


    
    // Verificar nuevamente por ID, no por fila.
    // No realizar una segunda escritura.
    var checkResponse = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z',
      { valueRenderOption: 'FORMATTED_VALUE' }
    );

    var checkData = checkResponse.values || [];
    var checkHeaders = checkData[0] || [];

    var checkIdCol = getRequiredColumnIndex(
      checkHeaders, 'idx'
    );

    var checkMatches = [];

    for (var v = 1; v < checkData.length; v++) {
      if (String(checkData[v][checkIdCol] == null
          ? '' : checkData[v][checkIdCol]).trim() === id) {
        checkMatches.push({
          row: v + 1,
          values: checkData[v]
        });
      }
    }

    if (checkMatches.length !== 1) {
      return apiError(
        'Actualización enviada, pero ID no verificable. ' +
        'Revisar BD antes de reintentar.',
        {
          id: id,
          writeAttempted: true,
          coincidencias: checkMatches.length
        }
      );
    }

    var verifiedRow = checkMatches[0].values;

    // Comparar todos los campos, no solamente
    // los modificados.
    var verified = headers.every(function(header, c) {

      var expectedValue = Object.prototype.hasOwnProperty.call(
        merged, header
      ) ? merged[header] : currentRow[c];

      var actualValue = verifiedRow[c];

      return String(expectedValue == null
        ? '' : expectedValue) ===
        String(actualValue == null
          ? '' : actualValue);
    });

    if (!verified) {
      return apiError(
        'Actualización enviada, pero los valores ' +
        'no coinciden. Revisar BD antes de reintentar.',
        {
          id: id,
          writeAttempted: true,
          verified: false
        }
      );
    }

    result.updated = true;
    result.verified = true;
    result.row = checkMatches[0].row;

    return apiSuccess(
      result,
      'Transacción actualizada correctamente'
    );


  } catch (error) {
    return apiError(
      'Error en updateTransaction: ' + error.toString()
    );
  } finally {
    if (locked) lock.releaseLock();
  }
}

/**
 * =========================
 * PROYECCIÓN / PRESUPUESTO
 * =========================
 *
 * Hoja presupuesto:
 * Spreadsheet: PRESUPUESTO_SHEET_ID
 * Sheet: BD
 *
 * Headers:
 * idx;fecha_ejecucion;tipo;grupo;categoria;tercero;valor;descripcion;estado
 */

/**
 * GET getProjectionData
 *
 * payload opcional:
 * {
 *   onlyCurrentYear: true
 * }
 *
 * Retorna:
 * {
 *   budget: [],
 *   execution: []
 * }
 */
function getProjectionData(payload) {
  try {
    payload = payload || {};

    var budget = getProjectionBudgetData(payload);

    // La ejecución real sale de la BD actual de movimientos.
    // Reutilizamos getListData para mantener la misma lógica de estado y año.
    var executionResponse = getListData(payload);

    if (!executionResponse.success) {
      return apiError(
        executionResponse.message || 'No se pudo cargar la ejecución real',
        { executionResponse: executionResponse }
      );
    }

    var executionItems = executionResponse.data.items || [];

    // Para Proyección excluimos movimientos operativos de ajuste.
    // Esto NO afecta Movimientos, Tesorería ni Dashboard.
    executionItems = executionItems.filter(function(item) {
      return !isProjectionAdjustmentRecord(item);
    });

    return apiSuccess({
      budget: budget,
      execution: executionItems
    }, 'Datos de proyección cargados correctamente');

  } catch (e) {
    return apiError('Error en getProjectionData: ' + e.toString());
  }
}

/**
 * Lee la BD de presupuesto.
 */
function getProjectionBudgetData(payload) {
  payload = payload || {};

  var onlyCurrentYear = !!payload.onlyCurrentYear;

  var data = Sheets.Spreadsheets.Values.get(PRESUPUESTO_SHEET_ID, 'BD!A:Z').values;

  if (!data || data.length <= 1) {
    return [];
  }

  var headers = data[0];

  validateProjectionHeaders(headers);

  var estadoIdx = getRequiredColumnIndex(headers, 'estado');
  var fechaIdx = getRequiredColumnIndex(headers, 'fecha_ejecucion');
  var idxIdx = getRequiredColumnIndex(headers, 'idx');

  var currentYear = new Date().getFullYear();

  var result = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    if ((row[estadoIdx] || '') !== 'h') continue;

    var rowDateStr = normalizeProjectionDateValue(row[fechaIdx]);
    var rowYear = parseInt(String(rowDateStr).split('-')[0], 10);

    if (onlyCurrentYear && rowYear !== currentYear) continue;

    var obj = mapRowToObject(headers, row);

    // Forzamos fecha ISO para que el front siempre reciba yyyy-mm-dd
    obj.fecha_ejecucion = rowDateStr;

    result.push(obj);
  }

  result.sort(function(a, b) {
    var aOrder = parseInt(a[headers[idxIdx]], 10) || 0;
    var bOrder = parseInt(b[headers[idxIdx]], 10) || 0;
    return bOrder - aOrder;
  });

  return result;
}

/**
 * POST saveProjectionTasks
 *
 * payload esperado:
 * {
 *   tasks: [
 *     {
 *       action: 'create' | 'update' | 'delete',
 *       id: 123,
 *       record: {
 *         fecha_ejecucion: '2026-06-01',
 *         tipo: 'Gasto',
 *         grupo: 'Servicios',
 *         categoria: 'Internet',
 *         tercero: 'Claro',
 *         valor: 140000,
 *         descripcion: 'Internet hogar',
 *         estado: 'h'
 *       }
 *     }
 *   ]
 * }
 */
function saveProjectionTasks(payload) {
  var lock = LockService.getScriptLock();

  try {
    payload = payload || {};

    var tasks = payload.tasks || [];

    if (!Array.isArray(tasks) || tasks.length === 0) {
      return apiError('No se recibieron tareas de presupuesto');
    }

    lock.waitLock(10000);

    var data = Sheets.Spreadsheets.Values.get(PRESUPUESTO_SHEET_ID, 'BD!A:Z').values;

    if (!data || data.length === 0) {
      return apiError('La hoja BD de presupuesto no tiene encabezados');
    }

    var headers = data[0];

    validateProjectionHeaders(headers);

    var idxIdx = getRequiredColumnIndex(headers, 'idx');
    var estadoIdx = getRequiredColumnIndex(headers, 'estado');

    var nextId = getNextProjectionId(data, idxIdx);

    var rowsToAppend = [];

    var result = {
      created: 0,
      updated: 0,
      deleted: 0,
      errors: []
    };

    for (var i = 0; i < tasks.length; i++) {
      var task = tasks[i] || {};
      var action = String(task.action || '').toLowerCase();

      try {
        if (action === 'create') {
          var createRecord = sanitizeProjectionRecord(task.record || {});
          var createErrors = validateProjectionRecord(createRecord);

          if (createErrors.length > 0) {
            result.errors.push({
              index: i,
              action: action,
              message: createErrors.join(', ')
            });
            continue;
          }

          createRecord.estado = createRecord.estado || 'h';

          var newId = nextId;
          nextId++;

          var newRow = buildProjectionRow(headers, createRecord, newId);
          rowsToAppend.push(newRow);

          result.created++;
        }

        else if (action === 'update') {
          var updateId = task.id;

          if (!updateId) {
            result.errors.push({
              index: i,
              action: action,
              message: 'La tarea update no tiene id'
            });
            continue;
          }

          var updateDataIndex = findProjectionDataIndexById(data, headers, updateId);

          if (updateDataIndex === -1) {
            result.errors.push({
              index: i,
              action: action,
              id: updateId,
              message: 'No se encontró el presupuesto a actualizar'
            });
            continue;
          }

          var currentObj = mapRowToObject(headers, data[updateDataIndex]);
          var updateRecord = sanitizeProjectionRecord(task.record || {});

          var mergedRecord = {};
          headers.forEach(function(header) {
            mergedRecord[header] = currentObj[header] || '';
          });

          Object.keys(updateRecord).forEach(function(key) {
            mergedRecord[key] = updateRecord[key];
          });

          mergedRecord.estado = mergedRecord.estado || 'h';

          var updateErrors = validateProjectionRecord(mergedRecord);

          if (updateErrors.length > 0) {
            result.errors.push({
              index: i,
              action: action,
              id: updateId,
              message: updateErrors.join(', ')
            });
            continue;
          }

          var updatedRow = buildProjectionRow(headers, mergedRecord, updateId);
          var updateRowNumber = updateDataIndex + 1;
          var updateRange = 'BD!A' + updateRowNumber + ':' + columnToLetter(headers.length) + updateRowNumber;

          var updateValueRange = Sheets.newValueRange();
          updateValueRange.values = [updatedRow];

          Sheets.Spreadsheets.Values.update(
            updateValueRange,
            PRESUPUESTO_SHEET_ID,
            updateRange,
            { valueInputOption: 'RAW' }
          );

          data[updateDataIndex] = updatedRow;

          result.updated++;
        }

        else if (action === 'delete') {
          var deleteId = task.id;

          if (!deleteId) {
            result.errors.push({
              index: i,
              action: action,
              message: 'La tarea delete no tiene id'
            });
            continue;
          }

          var deleteDataIndex = findProjectionDataIndexById(data, headers, deleteId);

          if (deleteDataIndex === -1) {
            result.errors.push({
              index: i,
              action: action,
              id: deleteId,
              message: 'No se encontró el presupuesto a eliminar'
            });
            continue;
          }

          var deleteRowNumber = deleteDataIndex + 1;
          var estadoColLetter = columnToLetter(estadoIdx + 1);
          var deleteRange = 'BD!' + estadoColLetter + deleteRowNumber;

          var deleteValueRange = Sheets.newValueRange();
          deleteValueRange.values = [['n']];

          Sheets.Spreadsheets.Values.update(
            deleteValueRange,
            PRESUPUESTO_SHEET_ID,
            deleteRange,
            { valueInputOption: 'RAW' }
          );

          data[deleteDataIndex][estadoIdx] = 'n';

          result.deleted++;
        }

        else {
          result.errors.push({
            index: i,
            action: action,
            message: 'Acción no soportada'
          });
        }

      } catch (taskError) {
        result.errors.push({
          index: i,
          action: action,
          message: taskError.toString()
        });
      }
    }

    if (rowsToAppend.length > 0) {
      var appendValueRange = Sheets.newValueRange();
      appendValueRange.values = rowsToAppend;

      Sheets.Spreadsheets.Values.append(
        appendValueRange,
        PRESUPUESTO_SHEET_ID,
        'BD!A1',
        { valueInputOption: 'RAW' }
      );
    }

    var message = 'Cambios de presupuesto guardados correctamente';

    if (result.errors.length > 0) {
      message = 'Cambios procesados con algunas observaciones';
    }

    return apiSuccess(result, message);

  } catch (e) {
    return apiError('Error en saveProjectionTasks: ' + e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * Valida que la hoja de presupuesto tenga las columnas mínimas.
 */
function validateProjectionHeaders(headers) {
  getRequiredColumnIndex(headers, 'idx');
  getRequiredColumnIndex(headers, 'fecha_ejecucion');
  getRequiredColumnIndex(headers, 'tipo');
  getRequiredColumnIndex(headers, 'grupo');
  getRequiredColumnIndex(headers, 'categoria');
  getRequiredColumnIndex(headers, 'tercero');
  getRequiredColumnIndex(headers, 'valor');
  getRequiredColumnIndex(headers, 'descripcion');
  getRequiredColumnIndex(headers, 'estado');
}

/**
 * Calcula el siguiente ID disponible tomando el máximo idx existente.
 */
function getNextProjectionId(data, idxIdx) {
  var maxId = 0;

  for (var i = 1; i < data.length; i++) {
    var id = parseInt(data[i][idxIdx], 10);
    if (!isNaN(id) && id > maxId) {
      maxId = id;
    }
  }

  return maxId + 1;
}

/**
 * Busca el índice interno dentro de data por idx.
 * Retorna índice de array data, no número de fila.
 */
function findProjectionDataIndexById(data, headers, id) {
  var idxIdx = getRequiredColumnIndex(headers, 'idx');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idxIdx]) === String(id)) {
      return i;
    }
  }

  return -1;
}

/**
 * Convierte un record de presupuesto a fila según headers.
 */
function buildProjectionRow(headers, record, id) {
  var idxIdx = getRequiredColumnIndex(headers, 'idx');

  return headers.map(function(header, index) {
    if (index === idxIdx) return id;

    return record.hasOwnProperty(header)
      ? record[header]
      : '';
  });
}

/**
 * Limpia y normaliza el record recibido desde el front.
 */
function sanitizeProjectionRecord(record) {
  record = record || {};

  return {
    fecha_ejecucion: normalizeProjectionDateValue(record.fecha_ejecucion),
    tipo: record.tipo || '',
    grupo: record.grupo || '',
    categoria: record.categoria || '',
    tercero: record.tercero || '',
    valor: parseProjectionNumber(record.valor),
    descripcion: record.descripcion || '',
    estado: record.estado || 'h'
  };
}

/**
 * Convierte valor a número conservando negativos.
 */
function parseProjectionNumber(value) {
  if (value === 0 || value === '0') return 0;

  if (value === null || value === undefined || value === '') {
    return '';
  }

  var clean = String(value)
    .replace(/\./g, '')
    .replace(/,/g, '')
    .trim();

  var num = Number(clean);

  if (isNaN(num)) {
    return value;
  }

  return num;
}

/**
 * Valida record de presupuesto.
 * Obligatorios:
 * fecha_ejecucion, tipo, grupo, valor, descripcion.
 *
 * Opcionales:
 * categoria, tercero.
 */
function validateProjectionRecord(record) {
  var errors = [];

  if (!record.fecha_ejecucion) {
    errors.push('fecha_ejecucion es requerida');
  }

  if (!record.tipo) {
    errors.push('tipo es requerido');
  }

  if (!record.grupo) {
    errors.push('grupo es requerido');
  }

  if (
    record.valor === '' ||
    record.valor === null ||
    record.valor === undefined ||
    isNaN(Number(record.valor))
  ) {
    errors.push('valor debe ser numérico');
  }

  if (!record.descripcion) {
    errors.push('descripcion es requerida');
  }

  return errors;
}

/**
 * payload esperado:
 * {
 *   id: 123
 * }
 */
function softDeleteTransaction(payload) {
  var lock = LockService.getScriptLock();

  try {
    var id = payload.id;

    if (!id) {
      return apiError('El id es requerido');
    }

    lock.waitLock(10000);

    var data = Sheets.Spreadsheets.Values.get(BD_SHEET_ID, 'BD!A:Z').values;
    if (!data || data.length <= 1) {
      return apiError('No hay datos en la BD');
    }

    var headers = data[0];
    var estadoIdx = getRequiredColumnIndex(headers, 'estado');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var rowNum = i + 1;
        var colLetter = columnToLetter(estadoIdx + 1);
        var updateRange = 'BD!' + colLetter + rowNum;

        var valueRange = Sheets.newValueRange();
        valueRange.values = [['n']];

        Sheets.Spreadsheets.Values.update(
          valueRange,
          BD_SHEET_ID,
          updateRange,
          { valueInputOption: 'RAW' }
        );

        return apiSuccess({ id: id }, 'Transacción eliminada lógicamente');
      }
    }

    return apiError('ID no encontrado', { id: id });
  } catch (e) {
    return apiError(e.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * payload esperado:
 * {
 *   username: "..."
 * }
 */
function logAccess(payload) {
  try {
    var username = payload.username || '';
    if (!username) {
      return apiError('username es requerido');
    }

    var valueRange = Sheets.newValueRange();
    valueRange.values = [[
      username,
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    ]];

    Sheets.Spreadsheets.Values.append(
      valueRange,
      PARAMS_SHEET_ID,
      'Logs!A1',
      { valueInputOption: 'USER_ENTERED' }
    );

    return apiSuccess({ username: username }, 'Acceso registrado');
  } catch (e) {
    return apiError(e.toString());
  }
}

/**
 * =========================
 * HELPERS VARIOS
 * =========================
 */
function columnToLetter(column) {
  var temp = '';
  while (column > 0) {
    var remainder = (column - 1) % 26;
    temp = String.fromCharCode(65 + remainder) + temp;
    column = Math.floor((column - 1) / 26);
  }
  return temp;
}

function normalizeProjectionDateValue(value) {
  if (value === null || value === undefined || value === '') return '';

  // Si por alguna razón llega como número serial de Google Sheets
  if (typeof value === 'number') {
    var dateFromSerial = new Date(Math.round((value - 25569) * 86400 * 1000));
    return Utilities.formatDate(
      dateFromSerial,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  var s = String(value).trim();

  // Ya viene ISO
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
    return s.substring(0, 10);
  }

  // yyyy/mm/dd
  if (/^\d{4}\/\d{1,2}\/\d{1,2}$/.test(s)) {
    var p1 = s.split('/');
    return p1[0] + '-' + pad2Projection(p1[1]) + '-' + pad2Projection(p1[2]);
  }

  // dd/mm/yyyy o dd-mm-yyyy
  if (/^\d{1,2}[\/-]\d{1,2}[\/-]\d{4}$/.test(s)) {
    var p2 = s.split(/[\/-]/);

    var a = parseInt(p2[0], 10);
    var b = parseInt(p2[1], 10);
    var y = p2[2];

    var day;
    var month;

    // En Colombia normalmente llega dd/mm/yyyy.
    // Si el segundo número es mayor a 12, asumimos mm/dd/yyyy.
    if (b > 12) {
      month = a;
      day = b;
    } else {
      day = a;
      month = b;
    }

    return y + '-' + pad2Projection(month) + '-' + pad2Projection(day);
  }

  return s;
}

function pad2Projection(value) {
  value = String(value);
  return value.length === 1 ? '0' + value : value;
}

function isProjectionAdjustmentRecord(record) {
  record = record || {};

  // Excluye únicamente descripción exactamente igual a "Ajuste".
  // No excluye: "Ajustes", "ajuste", "AJUSTE", "Ajuste caja".
  return String(record.descripcion || '').trim() === 'Ajuste';
}


/**
 * NEPTUNO FZ - ELIMINACION LOGICA SEGURA V2
 *
 * Uso exclusivo en copia de pruebas.
 * No conectada a doPost().
 *
 * Busca por idx unico, protege posibles
 * movimientos vinculados y verifica escritura.
 */

function softDeleteTransactionV2_(
  spreadsheetId,
  payload,
  allowProduction
) {


  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  
  const isTest = spreadsheetId === TEST_ID;

  const isProduction =
    spreadsheetId === BD_SHEET_ID &&
    allowProduction === true &&
    TRANSACTION_V3_ENABLED === true;

  if (!isTest && !isProduction) {
    return apiError(
      'BD no autorizada para eliminacion V2'
    );
  }


  const lock = LockService.getScriptLock();
  let locked = false;
  let writeAttempted = false;

  try {
    payload = payload || {};

    const id = String(
      payload.id == null ? '' : payload.id
    ).trim();

    if (!/^[1-9]\d*$/.test(id) ||
        !Number.isSafeInteger(Number(id))) {
      return apiError('ID invalido');
    }

    lock.waitLock(10000);
    locked = true;

    const response = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z',
      { valueRenderOption: 'UNFORMATTED_VALUE' }
    );

    const data = response.values || [];

    if (data.length < 2) {
      return apiError('BD sin registros');
    }

    const headers = data[0];

    const idCol = getRequiredColumnIndex(headers, 'idx');
    const estadoCol = getRequiredColumnIndex(
      headers, 'estado'
    );
    const medioCol = getRequiredColumnIndex(
      headers, 'medio'
    );
    const tipoCol = getRequiredColumnIndex(
      headers, 'tipo'
    );
    const requestCol = getRequiredColumnIndex(
      headers, 'request_id'
    );

    const matches = [];

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol] ?? '').trim() === id) {
        matches.push(i);
      }
    }

    if (matches.length === 0) {
      return apiError('ID no encontrado', { id: id });
    }

    if (matches.length !== 1) {
      return apiError(
        'ID duplicado. Eliminacion bloqueada',
        { id: id, coincidencias: matches.length }
      );
    }

    const index = matches[0];
    const row = data[index];

    if (String(row[estadoCol] || '').trim() !== 'h') {
      return apiError('El movimiento no esta activo');
    }

    const medio = String(row[medioCol] || '').trim();
    const tipo = String(row[tipoCol] || '').trim();
    const requestId = String(
      row[requestCol] || ''
    ).trim();

    
    // ==========================================
    // V3 - IDENTIFICAR ELIMINACION VINCULADA
    // ==========================================

    let targetIndices = [index];
    let linkedDelete = false;

    const relacionados = requestId
      ? data.slice(1).map(function(r, i) {
          return {
            index: i + 1,
            row: r
          };
        }).filter(function(item) {
          return String(
            item.row[requestCol] || ''
          ).trim() === requestId;
        })
      : [];

    const log = requestId
      ? readIdempotencyLogV3_(
          spreadsheetId
        ).filter(function(entry) {
          return entry.requestId === requestId;
        })
      : [];

    if (log.length > 1) {
      return apiError(
        'Registro tecnico duplicado. Eliminacion bloqueada'
      );
    }

    if (log.length === 1) {

      const entry = log[0];

      if (entry.status !== 'COMMITTED') {
        return apiError(
          'Solicitud no confirmada. Eliminacion bloqueada'
        );
      }

      const principalId = String(entry.principalId);
      const pasivoId = String(entry.pasivoId || '');

      const expectedIds = pasivoId
        ? [principalId, pasivoId]
        : [principalId];

      const idsRelacionados = relacionados.map(
        function(item) {
          return String(item.row[idCol] || '');
        }
      );

      const parejaValida =
        relacionados.length === expectedIds.length &&
        expectedIds.every(function(expectedId) {
          return idsRelacionados.filter(function(actualId) {
            return actualId === expectedId;
          }).length === 1;
        });

      if (!parejaValida) {
        return apiError(
          'Movimientos vinculados inconsistentes. ' +
          'Eliminacion bloqueada.'
        );
      }

      if (pasivoId) {

        if (expectedIds.indexOf(id) === -1) {
          return apiError(
            'El movimiento no pertenece a la pareja'
          );
        }

        targetIndices = relacionados.map(function(item) {
          return item.index;
        });

        linkedDelete = true;
      } else if (principalId !== id) {
        return apiError(
          'Movimiento no coincide con registro tecnico'
        );
      }
    }

    // Si hay varias filas con el mismo request_id,
    // exigir una pareja tecnica V3 verificada.
    if (relacionados.length > 1 && !linkedDelete) {
      return apiError(
        'Solicitud vinculada sin validacion tecnica. ' +
        'Eliminacion bloqueada.'
      );
    }

    // ==========================================
    // VALIDAR ESTADOS Y UNICIDAD
    // ==========================================

    const idsObjetivo = targetIndices.map(function(i) {
      return String(data[i][idCol] || '');
    });

    if (new Set(idsObjetivo).size !== targetIndices.length) {
      return apiError(
        'IDs duplicados en eliminacion vinculada'
      );
    }

    const estadosValidos = targetIndices.every(function(i) {
      return String(
        data[i][estadoCol] || ''
      ).trim() === 'h';
    });

    if (!estadosValidos) {
      return apiError(
        'Uno de los movimientos vinculados no esta activo'
      );
    }

    // ==========================================
    // VALIDAR PASIVO AUTOMATICO
    // ==========================================

    if (linkedDelete) {

      const principal = relacionados.find(function(item) {
        return String(item.row[idCol]) ===
          String(log[0].principalId);
      });

      const pasivo = relacionados.find(function(item) {
        return String(item.row[idCol]) ===
          String(log[0].pasivoId);
      });

      if (!principal || !pasivo) {
        return apiError('Pareja incompleta');
      }

      const principalRow = principal.row;
      const pasivoRow = pasivo.row;

      const pasivoCorrecto =
        String(pasivoRow[tipoCol]).toLowerCase() === 'pasivo' &&
        String(pasivoRow[medioCol]).toLowerCase() === 'bridge';

      const principalCorrecto =
        String(principalRow[tipoCol]).toLowerCase() !== 'pasivo';

      if (!pasivoCorrecto || !principalCorrecto) {
        return apiError(
          'Estructura financiera de pareja invalida'
        );
      }

    } else {

      // Mantener proteccion conservadora
      // para movimientos sin pareja V3 verificada.
      const catalogo = getFormData();

      if (!catalogo.success) {
        return apiError(
          'No se pudo validar el catalogo de medios'
        );
      }

      const fuentes = catalogo.data.formasPago || [];

      const fuente = fuentes.filter(function(f) {
        return String(f.medio || '').trim() === medio;
      });

      if (fuente.length !== 1) {
        return apiError(
          'Medio inexistente o ambiguo. Eliminacion bloqueada'
        );
      }

      const generaPasivo =
        String(fuente[0].genera_pasivo || 'n')
          .trim().toLowerCase() === 's';

      if (
        generaPasivo ||
        (
          tipo.toLowerCase() === 'pasivo' &&
          medio.toLowerCase() === 'bridge'
        )
      ) {
        return apiError(
          'Movimiento con posible pasivo automatico. ' +
          'Eliminacion vinculada pendiente.'
        );
      }
    }


    
    // ==========================================
    // ESCRITURA LOGICA V2 / V3
    // ==========================================

    const rowNumber = index + 1;

    const sheetMetadata = Sheets.Spreadsheets.get(
      spreadsheetId,
      {
        fields: 'sheets(properties(sheetId,title))'
      }
    );

    const bdSheet = (sheetMetadata.sheets || [])
      .find(function(s) {
        return s.properties.title === 'BD';
      });

    if (!bdSheet) {
      return apiError('Hoja BD no encontrada');
    }

    const sheetId = bdSheet.properties.sheetId;

    // Una sola solicitud batchUpdate.
    const requests = targetIndices.map(function(i) {
      return {
        updateCells: {
          start: {
            sheetId: sheetId,
            rowIndex: i,
            columnIndex: estadoCol
          },
          rows: [{
            values: [{
              userEnteredValue: {
                stringValue: 'n'
              }
            }]
          }],
          fields: 'userEnteredValue'
        }
      };
    });

    writeAttempted = true;

    Sheets.Spreadsheets.batchUpdate(
      { requests: requests },
      spreadsheetId
    );

    // ==========================================
    // VERIFICACION POSTERIOR POR ID
    // ==========================================

    const checkData = Sheets.Spreadsheets.Values.get(
      spreadsheetId,
      'BD!A:Z',
      { valueRenderOption: 'UNFORMATTED_VALUE' }
    ).values || [];

    const checkHeaders = checkData[0] || [];

    const checkIdCol = getRequiredColumnIndex(
      checkHeaders, 'idx'
    );

    const checkEstadoCol = getRequiredColumnIndex(
      checkHeaders, 'estado'
    );

    const verifiedIds = [];

    for (const targetIndex of targetIndices) {

      const originalRow = data[targetIndex];
      const targetId = String(
        originalRow[idCol] || ''
      );

      const matches = checkData.slice(1)
        .filter(function(r) {
          return String(r[checkIdCol] || '') === targetId;
        });

      if (matches.length !== 1) {
        return apiError(
          'Escritura enviada, pero ID no verificable.',
          {
            id: targetId,
            writeAttempted: true
          }
        );
      }

      const actualRow = matches[0];

      const mismosCampos = headers.every(
        function(header, col) {
          if (col === estadoCol) return true;

          return String(originalRow[col] ?? '') ===
            String(actualRow[col] ?? '');
        }
      );

      const estadoCorrecto =
        String(actualRow[checkEstadoCol] || '') === 'n';

      if (!mismosCampos || !estadoCorrecto) {
        return apiError(
          'Integridad no verificada despues de eliminar.',
          {
            id: targetId,
            writeAttempted: true,
            mismosCampos: mismosCampos,
            estadoCorrecto: estadoCorrecto
          }
        );
      }

      verifiedIds.push(targetId);
    }

    return apiSuccess({
      id: id,
      row: rowNumber,
      deleted: true,
      estado: 'n',
      verified: true,
      linkedDelete: linkedDelete,
      deletedIds: verifiedIds,
      deletedCount: verifiedIds.length
    }, linkedDelete
      ? 'Movimientos vinculados eliminados logicamente'
      : 'Transaccion eliminada logicamente'
    );


  } catch (error) {
    return apiError(
      'Error en softDeleteTransactionV2_: ' +
      error.toString(),
      {
        id: payload && payload.id,
        writeAttempted: writeAttempted
      }
    );
  } finally {
    if (locked) lock.releaseLock();
  }
}



