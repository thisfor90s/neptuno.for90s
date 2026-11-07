// PRUEBAS DE REGRESION. Ejecutar solo sobre copia, previa revision.
// Las funciones de escritura no deben ejecutarse dos veces sin revisar sus precondiciones.

/**
 * NEPTUNO FZ - INTEGRACION INTERNA
 *
 * Preparacion del backend.
 * Exclusivamente copia de pruebas.
 *
 * No conectado a doPost().
 * No habilita escrituras productivas.
 */
function dispatchTransactionTest_(action, payload) {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  if (TEST_ID === BD_SHEET_ID) {
    return apiError(
      'La BD de pruebas coincide con produccion'
    );
  }

  payload = payload || {};

  if (!payload || typeof payload !== 'object' ||
      Array.isArray(payload)) {
    return apiError('Payload invalido');
  }

  switch (action) {

    case 'saveTransaction':
      return saveTransactionV2_(
        TEST_ID,
        payload
      );

    case 'updateTransaction':
      return updateTransactionOnSheet_(
        payload,
        TEST_ID,
        false
      );

    case 'softDeleteTransaction':
      return softDeleteTransactionV2_(
        TEST_ID,
        payload
      );

    default:
      return apiError(
        'Operacion de integracion no autorizada'
      );
  }
}



/**
 * TEST - CONTRATO DE INTEGRACION
 *
 * Solo escenarios de rechazo.
 * No se esperan escrituras.
 */
function TEST_Backend_02_ContratoIntegracion() {

  const casos = [
    {
      nombre: 'Accion inexistente',
      action: 'accionDesconocida',
      payload: {}
    },
    {
      nombre: 'Guardar sin registro',
      action: 'saveTransaction',
      payload: {}
    },
    {
      nombre: 'Actualizar sin ID',
      action: 'updateTransaction',
      payload: {
        record: { descripcion: 'TEST' }
      }
    },
    {
      nombre: 'Eliminar sin ID',
      action: 'softDeleteTransaction',
      payload: {}
    }
  ];

  const resultados = [];

  casos.forEach(function(caso) {

    const respuesta = dispatchTransactionTest_(
      caso.action,
      caso.payload
    );

    const correcto =
      respuesta &&
      respuesta.success === false &&
      typeof respuesta.message === 'string';

    resultados.push({
      caso: caso.nombre,
      correcto: correcto,
      mensaje: respuesta && respuesta.message
    });

    if (!correcto) {
      throw new Error(
        'Contrato inesperado: ' + caso.nombre
      );
    }
  });

  Logger.log(
    'RESULTADOS: ' +
    JSON.stringify(resultados, null, 2)
  );

  Logger.log(
    'TEST BACKEND 02 SUPERADO'
  );
}


/**
 * TEST DELETE V2 - RECHAZOS
 *
 * No realiza escrituras previstas.
 * Solo copia de BD.
 */
function TEST_DeleteV2_Rechazos() {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  if (TEST_ID === BD_SHEET_ID) {
    throw new Error('Produccion no autorizada');
  }

  function snapshot() {
    const rows = Sheets.Spreadsheets.Values.get(
      TEST_ID,
      'BD!A:Z',
      { valueRenderOption: 'UNFORMATTED_VALUE' }
    ).values || [];

    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(rows),
      Utilities.Charset.UTF_8
    );

    return bytes.map(function(b) {
      return ('0' + ((b + 256) % 256).toString(16))
        .slice(-2);
    }).join('');
  }

  const antes = snapshot();

  const casos = [
    {
      nombre: 'ID invalido',
      id: 'ABC',
      mensaje: 'ID invalido'
    },
    {
      nombre: 'ID inexistente',
      id: '999999999',
      mensaje: 'ID no encontrado'
    },
    {
      nombre: 'Pasivo automatico protegido',
      id: '3672',
      mensaje: 'Movimiento con posible pasivo automatico.'
    },
    {
      nombre: 'Principal vinculado protegido',
      id: '3671',
      mensaje: 'Movimiento con posible pasivo automatico.'
    }
  ];

  const resultados = [];

  casos.forEach(function(caso) {

    const result = softDeleteTransactionV2_(
      TEST_ID,
      { id: caso.id }
    );

    const correcto =
      result.success === false &&
      String(result.message || '')
        .startsWith(caso.mensaje);

    resultados.push({
      caso: caso.nombre,
      correcto: correcto,
      mensaje: result.message
    });

    if (!correcto) {
      throw new Error(
        'Resultado inesperado en ' +
        caso.nombre + ': ' + result.message
      );
    }
  });

  const despues = snapshot();

  if (antes !== despues) {
    throw new Error(
      'La BD cambio durante las pruebas negativas'
    );
  }

  Logger.log(
    'RESULTADOS: ' +
    JSON.stringify(resultados, null, 2)
  );

  Logger.log(
    'TEST DELETE V2 RECHAZOS SUPERADO'
  );
}


/**
 * TEST UPDATE V2 - VERIFICACION POSTERIOR
 *
 * Exclusivamente copia de BD.
 * Actualiza y restaura una descripcion.
 * No crea movimientos.
 */
function TEST_UpdateV2_Verificacion() {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  const ID = '3670';
  const ORIGINAL = 'TEST V3 EDITADO';
  const TEMPORAL = 'TEST UPDATE V2 VERIFICADO';

  if (TEST_ID === BD_SHEET_ID) {
    throw new Error('Produccion no autorizada');
  }

  function leer() {
    const data = Sheets.Spreadsheets.Values.get(
      TEST_ID,
      'BD!A:Z',
      { valueRenderOption: 'FORMATTED_VALUE' }
    ).values || [];

    const headers = data[0] || [];
    const idCol = getRequiredColumnIndex(headers, 'idx');

    const matches = data.slice(1)
      .filter(function(row) {
        return String(row[idCol] || '') === ID;
      });

    if (matches.length !== 1) {
      throw new Error('ID inexistente o duplicado');
    }

    const obj = {};
    headers.forEach(function(h, i) {
      obj[h] = matches[0][i] == null
        ? '' : matches[0][i];
    });

    return obj;
  }

  function originales(obj) {
    const campos = [
      'fecha_registro', 'fecha_pago',
      'tipo', 'grupo', 'categoria',
      'medio', 'tercero', 'valor', 'descripcion'
    ];

    const expected = {};
    campos.forEach(function(campo) {
      expected[campo] = obj[campo];
    });

    return expected;
  }

  const antes = leer();

  if (
    antes.descripcion !== ORIGINAL ||
    antes.estado !== 'h' ||
    antes.medio !== 'Cash' ||
    antes.request_id !==
      'TEST_V3_INMUTABLE_20261107_001'
  ) {
    throw new Error(
      'El movimiento no coincide con el registro de prueba'
    );
  }

  // Primera actualización.
  const actualizado = updateTransactionOnSheet_(
    {
      id: ID,
      record: { descripcion: TEMPORAL },
      expected: originales(antes),
      dryRun: false
    },
    TEST_ID,
    false
  );

  Logger.log(
    'ACTUALIZAR: ' + JSON.stringify(actualizado)
  );

  if (
    !actualizado.success ||
    actualizado.data.updated !== true ||
    actualizado.data.verified !== true
  ) {
    throw new Error(
      'Actualización no verificada: ' +
      actualizado.message
    );
  }

  const intermedio = leer();

  if (intermedio.descripcion !== TEMPORAL) {
    throw new Error('Descripción incorrecta');
  }

  // Restaurar mediante la misma API.
  const restaurado = updateTransactionOnSheet_(
    {
      id: ID,
      record: { descripcion: ORIGINAL },
      expected: originales(intermedio),
      dryRun: false
    },
    TEST_ID,
    false
  );

  Logger.log(
    'RESTAURAR: ' + JSON.stringify(restaurado)
  );

  if (
    !restaurado.success ||
    restaurado.data.updated !== true ||
    restaurado.data.verified !== true
  ) {
    throw new Error(
      'Restauración no verificada: ' +
      restaurado.message
    );
  }

  const final = leer();

  if (JSON.stringify(final) !== JSON.stringify(antes)) {
    throw new Error(
      'El movimiento no quedó igual al original'
    );
  }

  Logger.log(
    'RESUMEN UPDATE V2: ' + JSON.stringify({
      id: ID,
      actualizado: true,
      verificado: true,
      restaurado: true,
      registroOriginalConservado: true
    }, null, 2)
  );

  Logger.log(
    'TEST UPDATE V2 VERIFICACION SUPERADO'
  );
}


/**
 * NEPTUNO FZ - TEST INTEGRACION BACKEND 01
 *
 * Verifica prioridad de idempotencia.
 * Solo copia de BD.
 * No realiza escrituras previstas.
 */
function TEST_Backend_01_RecuperacionPrioritaria() {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  const REQUEST_ID =
    'TEST_V3_INMUTABLE_20261107_001';

  if (TEST_ID === BD_SHEET_ID) {
    throw new Error('Produccion no autorizada');
  }

  function snapshot() {
    const ranges = [
      'BD!A:Z',
      "'Idempotency_Log'!A:G"
    ];

    const response = Sheets.Spreadsheets.Values.batchGet(
      TEST_ID,
      {
        ranges: ranges,
        valueRenderOption: 'UNFORMATTED_VALUE'
      }
    );

    const data = (response.valueRanges || [])
      .map(function(r) {
        return r.values || [];
      });

    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(data),
      Utilities.Charset.UTF_8
    );

    return bytes.map(function(b) {
      return ('0' + ((b + 256) % 256).toString(16))
        .slice(-2);
    }).join('');
  }

  const record = {
    fecha_registro: '2026-11-07',
    fecha_pago: '2026-11-07',
    tipo: 'Gasto',
    grupo: 'Pruebas',
    categoria: 'TEST',
    medio: 'Cash',
    tercero: 'Pruebas',
    valor: 5000,
    descripcion: 'TEST V3 ORIGINAL',
    genera_pasivo: 'n'
  };

  const logs = readIdempotencyLogV3_(TEST_ID)
    .filter(function(entry) {
      return entry.requestId === REQUEST_ID;
    });

  if (
    logs.length !== 1 ||
    logs[0].principalId !== '3670' ||
    logs[0].status !== 'COMMITTED'
  ) {
    throw new Error(
      'Registro tecnico de referencia no valido'
    );
  }

  const antes = snapshot();

  // 1. Recuperar solicitud original.
  const recuperado = saveTransactionV2_(
    TEST_ID,
    {
      requestId: REQUEST_ID,
      record: record
    }
  );

  Logger.log(
    'RECUPERACION: ' + JSON.stringify(recuperado)
  );

  if (
    !recuperado.success ||
    recuperado.data.replayed !== true ||
    recuperado.data.alreadyProcessed !== true ||
    String(recuperado.data.newId) !== '3670'
  ) {
    throw new Error(
      'No se recupero el resultado original'
    );
  }

  // 2. Mismo requestId, medio inexistente.
  // Debe detectar primero el conflicto de firma.
  const conflicto = saveTransactionV2_(
    TEST_ID,
    {
      requestId: REQUEST_ID,
      record: Object.assign({}, record, {
        medio: 'MEDIO_INEXISTENTE_TEST'
      })
    }
  );

  Logger.log(
    'CONFLICTO: ' + JSON.stringify(conflicto)
  );

  if (
    conflicto.success !== false ||
    conflicto.data.alreadyProcessed !== true ||
    conflicto.data.samePayload !== false ||
    conflicto.message !==
      'Conflicto: requestId reutilizado con datos diferentes.'
  ) {
    throw new Error(
      'La prioridad de idempotencia es incorrecta'
    );
  }

  const despues = snapshot();

  if (antes !== despues) {
    throw new Error(
      'La BD o el registro tecnico cambiaron'
    );
  }

  Logger.log(
    'RESUMEN BACKEND 01: ' + JSON.stringify({
      id: '3670',
      recuperacionCorrecta: true,
      conflictoPrioritario: true,
      escritura: false,
      integridadConservada: true
    }, null, 2)
  );

  Logger.log(
    'TEST BACKEND 01 SUPERADO'
  );
}

/**
 * NEPTUNO FZ - REGRESION FINAL BACKEND 03
 *
 * Solo lectura y rechazos controlados.
 * No realiza escrituras.
 * No modifica el interruptor de migracion.
 */
function TEST_Backend_03_RegresionFinal() {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  const resultados = [];
  const errores = [];

  function verificar(nombre, condicion, detalle) {
    const correcto = condicion === true;

    resultados.push({
      caso: nombre,
      correcto: correcto,
      detalle: detalle || ''
    });

    if (!correcto) {
      errores.push(nombre);
    }
  }

  function snapshot() {
    const response = Sheets.Spreadsheets.Values.batchGet(
      TEST_ID,
      {
        ranges: [
          'BD!A:Z',
          "'Idempotency_Log'!A:G"
        ],
        valueRenderOption: 'UNFORMATTED_VALUE'
      }
    );

    const data = (response.valueRanges || [])
      .map(function(r) {
        return r.values || [];
      });

    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(data),
      Utilities.Charset.UTF_8
    );

    return bytes.map(function(b) {
      return ('0' + ((b + 256) % 256).toString(16))
        .slice(-2);
    }).join('');
  }

  // ======================================
  // 1. CONTROL DE MIGRACION
  // ======================================

  if (TEST_ID === BD_SHEET_ID) {
    throw new Error(
      'La copia coincide con produccion'
    );
  }

  if (TRANSACTION_V3_ENABLED !== false) {
    throw new Error(
      'DETENIDO: interruptor V3 activado'
    );
  }

  const antes = snapshot();

  verificar(
    'Interruptor desactivado',
    TRANSACTION_V3_ENABLED === false
  );

  // ======================================
  // 2. GUARDADO PRODUCTIVO BLOQUEADO
  // ======================================

  const guardar = saveTransactionV2_(
    BD_SHEET_ID,
    {
      requestId: 'TEST_BACKEND03_BLOQUEO_001',
      record: {}
    },
    true
  );

  verificar(
    'Guardado V3 bloqueado',
    guardar.success === false &&
    guardar.message ===
      'Guardado V3 productivo no habilitado',
    guardar.message
  );

  // ======================================
  // 3. ELIMINACION PRODUCTIVA BLOQUEADA
  // ======================================

  const eliminar = softDeleteTransactionV2_(
    BD_SHEET_ID,
    { id: '3670' },
    true
  );

  verificar(
    'Eliminacion V2 bloqueada',
    eliminar.success === false &&
    eliminar.message ===
      'BD no autorizada para eliminacion V2',
    eliminar.message
  );

  // ======================================
  // 4. RUTA HTTP DE ACTUALIZACION
  // ======================================

  const http = doPost({
    postData: {
      contents: JSON.stringify({
        action: 'updateTransaction',
        payload: {
          id: '3670',
          record: {
            descripcion: 'NO ESCRIBIR'
          },
          dryRun: false
        }
      })
    }
  });

  const httpResult = JSON.parse(
    http.getContent()
  );

  verificar(
    'Ruta HTTP de edicion deshabilitada',
    httpResult.success === false &&
    httpResult.message ===
      'Actualizacion no habilitada',
    httpResult.message
  );

  // ======================================
  // 5. CATALOGOS
  // ======================================

  const catalogos = getFormData();

  verificar(
    'Lectura de catalogos',
    catalogos.success === true &&
    Array.isArray(catalogos.data.formasPago) &&
    Array.isArray(catalogos.data.tablaPG) &&
    Array.isArray(catalogos.data.flashData)
  );

  // ======================================
  // 6. MOVIMIENTOS
  // ======================================

  const movimientos = getListData({
    onlyCurrentYear: false
  });

  verificar(
    'Lectura de movimientos',
    movimientos.success === true &&
    Array.isArray(movimientos.data.items)
  );

  // ======================================
  // 7. PROYECCION
  // ======================================

  const proyeccion = getProjectionData({
    onlyCurrentYear: false
  });

  verificar(
    'Lectura de proyeccion',
    proyeccion.success === true &&
    Array.isArray(proyeccion.data.budget) &&
    Array.isArray(proyeccion.data.execution)
  );

  // ======================================
  // 8. INTEGRIDAD DE COPIA
  // ======================================

  const despues = snapshot();

  verificar(
    'Integridad de copia',
    antes === despues
  );

  // ======================================
  // RESULTADO
  // ======================================

  Logger.log(
    'RESULTADOS BACKEND 03: ' +
    JSON.stringify(resultados, null, 2)
  );

  Logger.log(
    'RESUMEN BACKEND 03: ' +
    JSON.stringify({
      casos: resultados.length,
      superados: resultados.length - errores.length,
      fallidos: errores.length,
      copiaIntacta: antes === despues,
      interruptorActivo: TRANSACTION_V3_ENABLED,
      errores: errores
    }, null, 2)
  );

  if (errores.length) {
    throw new Error(
      'TEST BACKEND 03 FALLIDO: ' +
      errores.join(' | ')
    );
  }

  Logger.log(
    'TEST BACKEND 03 SUPERADO - CIERRE TECNICO'
  );
}


/**
 * TEST FRONT V3 - INTEGRIDAD FAB
 *
 * Verifica los movimientos 3674 y 3675.
 * Solo lectura sobre copia.
 * No realiza escrituras.
 */
function TEST_FrontV3_IntegridadFAB() {

  const TEST_ID =
    '1LNT5Et3GzxVx0X_mb5YrD0ZQWz0O76GRvF-JtPw0Gz0';

  if (TEST_ID ===
      '1IcQA1aWWz6aOZijOP8Nwpmr3Tp6bk0O6n4VBhCXWRQ0') {
    throw new Error('Destino productivo no autorizado');
  }

  const response = Sheets.Spreadsheets.Values.get(
    TEST_ID,
    'BD!A:Z',
    { valueRenderOption: 'UNFORMATTED_VALUE' }
  );

  const data = response.values || [];
  const headers = data[0] || [];

  const idCol = getRequiredColumnIndex(headers, 'idx');
  const requestCol = getRequiredColumnIndex(
    headers, 'request_id'
  );
  const estadoCol = getRequiredColumnIndex(
    headers, 'estado'
  );

  const logs = readIdempotencyLogV3_(TEST_ID);

  const resultados = [];

  ['3674', '3675'].forEach(function(id) {

    const registros = data.slice(1).filter(function(row) {
      return String(row[idCol] ?? '') === id;
    });

    if (registros.length !== 1) {
      throw new Error('ID inexistente o duplicado: ' + id);
    }

    const row = registros[0];
    const requestId = String(row[requestCol] || '');

    const relacionados = data.slice(1).filter(
      function(r) {
        return String(r[requestCol] || '') === requestId;
      }
    );

    const entradas = logs.filter(function(entry) {
      return entry.requestId === requestId;
    });

    const correcto =
      requestId.length > 0 &&
      relacionados.length === 1 &&
      entradas.length === 1 &&
      entradas[0].status === 'COMMITTED' &&
      entradas[0].principalId === id &&
      entradas[0].pasivoId === '' &&
      String(row[estadoCol]) === 'h';

    resultados.push({
      id: id,
      requestId: requestId,
      registros: relacionados.length,
      entradasTecnicas: entradas.length,
      correcto: correcto
    });

    if (!correcto) {
      throw new Error(
        'Integridad incorrecta en movimiento ' + id
      );
    }
  });

  Logger.log(
    'RESULTADOS FAB V3: ' +
    JSON.stringify(resultados, null, 2)
  );

  Logger.log(
    'TEST FRONT V3 FAB SUPERADO'
  );
}

