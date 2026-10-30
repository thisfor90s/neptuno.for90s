/**
 * ESTE ES EL CODIGO FUENTE DEL BACKEND DE LA APLICACIÓN.
 * EL LLAMADO NO SE HACE A ESTE ARCHIVO, ESTO ESO SOLO DE REFERENCIA 
 * Y PARA TENERLO CONSOLIDADO EN UN SOLO LUGAR. 
 */


/**
 * VARIABLES GLOBALES
 */
var PARAMS_SHEET_ID = '1UJPefwcIOhF65ujBDE5FGEwJaHr-1-hddgC5BLWdGfE';
var BD_SHEET_ID = '1IcQA1aWWz6aOZijOP8Nwpmr3Tp6bk0O6n4VBhCXWRQ0';
/**
 * BD independiente para presupuesto / proyección.
 * Dentro de este Spreadsheet debe existir una hoja llamada BD.
 *
 * Headers esperados:
 * idx;fecha_ejecucion;tipo;grupo;categoria;tercero;valor;descripcion;estado
 */
var PRESUPUESTO_SHEET_ID = '19qER6hdcfHqtH9MHxkE8yqX8I15ntuJ1SM0x8KMk4ag';

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
        return jsonOutput(saveTransaction(payload));

      case 'softDeleteTransaction':
        return jsonOutput(softDeleteTransaction(payload));

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