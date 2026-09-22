/**
 * ROIS Consultoría — Orchestrator central
 * ============================================================
 * Este es un proyecto Apps Script STANDALONE — NO va pegado dentro
 * de ningún Google Sheets. Se crea UNA sola vez y sirve a TODOS los
 * clientes del Sistema de Finanzas Personales ROIS.
 *
 * Cómo instalarlo (una sola vez):
 *  1. https://script.google.com → Proyecto nuevo (NO "Extensiones →
 *     Apps Script" desde un Sheets — eso lo dejaría bound a esa hoja).
 *  2. Bórralo todo y pega este archivo completo. Nómbralo
 *     "ROIS Orchestrator".
 *  3. Selector de función → elige "procesarPendientes" y presiona
 *     Ejecutar una vez, solo para autorizar los permisos que pida
 *     (Drive, Sheets). Puede que no haya clientes pendientes todavía
 *     — no importa, el propósito es solo autorizar.
 *  4. Implementar → Nueva implementación → tipo "Aplicación web".
 *     Ejecutar como: Yo. Acceso: Cualquier usuario. Implementar.
 *  5. Copia la URL /exec. Esa es LA ÚNICA URL de Web App que existe
 *     en todo el sistema — se pega UNA vez en registro-financiero
 *     (constante WEBAPP_URL en index.html) y nunca se vuelve a tocar
 *     por cliente.
 *  6. Para dar de alta clientes nuevos: corre "procesarPendientes"
 *     manualmente (o instala un trigger de tiempo — Activadores →
 *     Añadir activador → procesarPendientes → basado en tiempo →
 *     cada hora, gratis dentro de las cuotas normales de Apps
 *     Script). Revisa el registro de ejecución (Ver → Registros) para
 *     ver la tabla de resultados: nombre, Sheet URL, sheetId, token.
 *  7. Por cada cliente en esa tabla, arma su link personalizado:
 *     https://solucionesrois-creator.github.io/registro-financiero/?sheetId=XXX&token=YYY
 *     Se lo compartes tal cual — la PWA guarda sheetId+token solos,
 *     el cliente no captura nada.
 *
 * Todo esto corre dentro de las cuotas gratuitas de Apps Script/Drive
 * (sin servicios de pago, sin infraestructura adicional).
 */

const CFG = {
  TEMPLATE_ID: '15AXSPZ1YHqklCBYVE0FGKkwQ_uPI3mDhMb-JQs0ArlE',
  CARPETA_CLIENTES_ID: '1f94cklORAVa5lVSOJWEPOETe3d5_4_ro',
  FORMS: [
    { origen: 'ROIS', sheetId: '1JptDVWURXpBEidUOgHue9xvnR1AwdmuXCI_fciGccVc' },
    { origen: 'Brenda', sheetId: '1Gq0VPhCMAJTf6JLmx2TvMOWlcMxjN6iuEPbWbDg4eWM' },
  ],
  SHEETS: { SETUP: '⚙️ MI SETUP', REGISTROS: 'REGISTROS' },
  CAP: { CUENTAS: 6, DEUDAS: 8, FIJOS: 12 },
  SETUP_ROWS: { cuentasStart: 6, deudasStart: 15, fijosStart: 26, ahorroRow: 41, satRow: 42 },
  TIPOS_DEUDA_VALIDOS: ['Personal', 'Tarjeta de crédito', 'Préstamo digital', 'Hipotecario', 'Otro'],
  MAX_ROWS_REGISTROS: 1000,
  INDICE_PROP_KEY: 'INDICE_SHEET_ID',
};

// ============================================================
// ÍNDICE DE CLIENTES — se autocrea la primera vez, guarda su ID
// en Script Properties (no requiere configuración manual).
// ============================================================
function getIndiceSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(CFG.INDICE_PROP_KEY);
  let ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; }
  }
  if (!id) {
    ss = SpreadsheetApp.create('ROIS — Índice de Clientes');
    props.setProperty(CFG.INDICE_PROP_KEY, ss.getId());
    const sh = ss.getSheets()[0];
    sh.setName('CLIENTES');
    sh.getRange(1, 1, 1, 7).setValues([[
      'Nombre', 'Email', 'Canal', 'Sheet ID', 'Sheet URL', 'Token', 'Fecha alta',
    ]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('Índice de clientes creado: ' + ss.getUrl());
  }
  return ss.getSheetByName('CLIENTES') || ss.getSheets()[0];
}

function registrarEnIndice_(cliente, clienteSs, token) {
  const sh = getIndiceSheet_();
  sh.appendRow([
    cliente.nombre, cliente.email, cliente.origen,
    clienteSs.getId(), clienteSs.getUrl(), token, new Date(),
  ]);
}

function validarToken_(sheetId, token) {
  const sh = getIndiceSheet_();
  const data = sh.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][3]) === String(sheetId)) return String(data[r][5]) === String(token);
  }
  return false;
}

// ============================================================
// ONBOARDING — lee los Sheets de respuestas del Form, clona el
// template por cada cliente pendiente, llena SETUP con sus datos
// reales y marca la fila como Procesado.
// ============================================================
function procesarPendientes() {
  const resumen = [];

  CFG.FORMS.forEach(function (form) {
    const formSs = SpreadsheetApp.openById(form.sheetId);
    const formSh = formSs.getSheets()[0];
    const data = formSh.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0];
    const colEstado = headers.indexOf('Estado');
    const idx = mapHeaders_(headers);

    for (let r = 1; r < data.length; r++) {
      const row = data[r];
      if (!row[idx.nombre]) continue; // fila vacía
      const estado = colEstado === -1 ? '' : String(row[colEstado] || '').trim().toLowerCase();
      if (estado !== '' && estado !== 'pendiente') continue;

      const cliente = parseClienteForm_(row, idx, form.origen);
      Logger.log('DEBUG cliente parseado: ' + JSON.stringify(cliente));
      const resultado = onboardCliente_(cliente);
      resumen.push(resultado);

      if (colEstado !== -1) {
        formSh.getRange(r + 1, colEstado + 1).setValue('Procesado');
      }
    }
  });

  if (resumen.length === 0) {
    Logger.log('No hay clientes nuevos pendientes de procesar.');
    return 'No hay clientes nuevos pendientes de procesar.';
  }
  Logger.log(JSON.stringify(resumen, null, 2));
  return resumen;
}

/** Mapea encabezados EXACTOS del Form a índices de columna. */
function mapHeaders_(headers) {
  function idx(label) { return headers.indexOf(label); }
  return {
    nombre: idx('Nombre completo'),
    email: idx('Correo de Google (Gmail)'),
    cuentas: idx('¿Con qué bancos o billeteras digitales manejas tu dinero?'),
    gastosFijos: idx('¿Cuáles de estos gastos fijos tienes? (selecciona todos los que apliquen)'),
    tieneDeudas: idx('¿Tienes deudas activas en este momento?'),
    deuda1Nombre: idx('¿Cómo se llama esta deuda?'),
    ahorroAuto: idx('¿Apartas un porcentaje de cada ingreso para ahorro automático?'),
    ahorroPct: idx('¿Qué porcentaje quieres apartar de cada ingreso?'),
    satActivo: idx('¿Estás dado de alta en el SAT como persona física?'),
  };
}

/**
 * Extrae los datos de una fila del Form. Las deudas usan búsqueda
 * posicional relativa a "¿Cómo se llama esta deuda?" (columnas
 * 20-34 del Form, 5 grupos de 3: nombre/tipo/situación) porque ese
 * encabezado se repite 5 veces y headers.indexOf() solo encuentra
 * la primera ocurrencia.
 */
function parseClienteForm_(row, idx, origen) {
  const cuentas = splitLista_(row[idx.cuentas]);
  const fijos = splitLista_(row[idx.gastosFijos]);
  const deudas = [];
  // Comparación por prefijo, no exacta: el Form guarda la frase completa
  // de la opción elegida (ej. "Sí, tengo deudas activas"), no solo "Sí".
  if (String(row[idx.tieneDeudas]).trim().toLowerCase().indexOf('s') === 0) {
    for (let g = 0; g < 5; g++) {
      const base = idx.deuda1Nombre + g * 3;
      const nombre = row[base];
      if (!nombre) continue;
      deudas.push({
        nombre: String(nombre).trim(),
        tipoForm: String(row[base + 1] || '').trim(),
      });
    }
  }
  return {
    origen: origen,
    nombre: String(row[idx.nombre]).trim(),
    email: String(row[idx.email] || '').trim(),
    cuentas: cuentas,
    fijos: fijos,
    deudas: deudas,
    ahorroPct: (String(row[idx.ahorroAuto]).trim().toLowerCase().indexOf('s') === 0 && row[idx.ahorroPct])
      ? Number(String(row[idx.ahorroPct]).replace('%', '').trim()) / 100
      : null,
  };
}

function splitLista_(value) {
  if (!value) return [];
  return String(value).split(',').map(function (s) { return s.trim(); }).filter(Boolean);
}

function onboardCliente_(cliente) {
  const carpeta = DriveApp.getFolderById(CFG.CARPETA_CLIENTES_ID);
  const nombreArchivo = 'Finanzas — ' + cliente.nombre + ' | ROIS';
  const copia = DriveApp.getFileById(CFG.TEMPLATE_ID).makeCopy(nombreArchivo, carpeta);
  const ss = SpreadsheetApp.openById(copia.getId());

  llenarSetup_(ss, cliente);

  const token = Utilities.getUuid();
  registrarEnIndice_(cliente, ss, token);

  return {
    nombre: cliente.nombre,
    origen: cliente.origen,
    sheetUrl: ss.getUrl(),
    sheetId: ss.getId(),
    token: token,
    linkPersonalizado: 'https://solucionesrois-creator.github.io/registro-financiero/?sheetId=' + ss.getId() + '&token=' + token,
  };
}

/**
 * Escribe SOLO estructura (nombres) en ⚙️ MI SETUP — nunca montos,
 * igual que la convención manual ya establecida. Las fórmulas de
 * VENCIMIENTOS/DASHBOARD/DEUDAS ya vienen listas en el clon (son
 * copia exacta del template) y jalan estos nombres solas.
 */
function llenarSetup_(ss, cliente) {
  const sh = ss.getSheetByName(CFG.SHEETS.SETUP);
  if (!sh) throw new Error('El clon no tiene la hoja ' + CFG.SHEETS.SETUP + ' — revisa el template.');

  // Cuentas (col A), tope CFG.CAP.CUENTAS
  const cuentas = cliente.cuentas.slice(0, CFG.CAP.CUENTAS);
  cuentas.forEach(function (nombre, i) {
    sh.getRange(CFG.SETUP_ROWS.cuentasStart + i, 1).setValue(nombre);
  });

  // Deudas (col A nombre, col B tipo), tope CFG.CAP.DEUDAS
  const deudas = cliente.deudas.slice(0, CFG.CAP.DEUDAS);
  deudas.forEach(function (deuda, i) {
    const row = CFG.SETUP_ROWS.deudasStart + i;
    sh.getRange(row, 1).setValue(deuda.nombre);
    const tipo = CFG.TIPOS_DEUDA_VALIDOS.indexOf(deuda.tipoForm) !== -1 ? deuda.tipoForm : 'Otro';
    sh.getRange(row, 2).setValue(tipo);
  });

  // Servicios fijos (col A), tope CFG.CAP.FIJOS
  const fijos = cliente.fijos.slice(0, CFG.CAP.FIJOS);
  fijos.forEach(function (nombre, i) {
    sh.getRange(CFG.SETUP_ROWS.fijosStart + i, 1).setValue(nombre);
  });

  // % de ahorro automático, solo si el cliente dio un valor explícito
  if (cliente.ahorroPct !== null && !isNaN(cliente.ahorroPct)) {
    sh.getRange(CFG.SETUP_ROWS.ahorroRow, 2).setValue(cliente.ahorroPct);
  }
}

// ============================================================
// WEB APP CENTRAL — recibe transacciones de TODOS los clientes.
// Un solo despliegue, un solo /exec, ruteo por sheetId + token.
// ============================================================

/**
 * Espera un POST con JSON:
 * { sheetId, token, fecha, tipo, concepto, categoria, subcategoria,
 *   monto, cuenta, metodo, empresaPersonal, notas }
 *
 * sheetId/token: identifican al cliente — vienen del Índice generado
 *   por procesarPendientes(). Sin un token válido para ese sheetId,
 *   se rechaza la escritura.
 * fecha: "YYYY-MM-DD" (se acepta con hora, se usa solo la fecha).
 * tipo: EXACTO uno de "Ingreso" | "Egreso" | "Cargo TC" | "Saldo inicial".
 * monto: número (o string numérico).
 *
 * Responde JSON: { status, rowNumber, timestamp } o { status:'error', message }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ status: 'error', message: 'Sin cuerpo en la petición.' });
    }
    const data = JSON.parse(e.postData.contents);

    if (!data.sheetId || !data.token) {
      return jsonResponse_({ status: 'error', message: 'Falta sheetId o token.' });
    }
    if (!validarToken_(data.sheetId, data.token)) {
      return jsonResponse_({ status: 'error', message: 'Token inválido para este sheetId.' });
    }

    const required = ['fecha', 'tipo', 'concepto', 'categoria', 'monto', 'cuenta', 'metodo'];
    for (const key of required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        return jsonResponse_({ status: 'error', message: 'Falta el campo: ' + key });
      }
    }

    const tiposValidos = ['Ingreso', 'Egreso', 'Cargo TC', 'Saldo inicial'];
    if (tiposValidos.indexOf(data.tipo) === -1) {
      return jsonResponse_({
        status: 'error',
        message: 'Tipo inválido: "' + data.tipo + '". Debe ser exactamente uno de: ' + tiposValidos.join(', '),
      });
    }

    const monto = Number(data.monto);
    if (isNaN(monto)) {
      return jsonResponse_({ status: 'error', message: 'Monto inválido: ' + data.monto });
    }

    const fecha = parseFechaLocal_(data.fecha);
    if (!fecha || isNaN(fecha.getTime())) {
      return jsonResponse_({ status: 'error', message: 'Fecha inválida: ' + data.fecha });
    }

    const ss = SpreadsheetApp.openById(data.sheetId);
    const sh = ss.getSheetByName(CFG.SHEETS.REGISTROS);
    if (!sh) {
      return jsonResponse_({ status: 'error', message: 'No se encontró la hoja REGISTROS en ese sheetId.' });
    }
    const targetRow = findNextEmptyRegistroRow_(sh);

    // B..I
    sh.getRange(targetRow, 2, 1, 8).setValues([[
      fecha, data.tipo, data.concepto, data.categoria, data.subcategoria || '',
      monto, data.cuenta, data.metodo,
    ]]);
    // M Empresa/Personal, O Notas
    sh.getRange(targetRow, 13).setValue(data.empresaPersonal || 'Personal');
    sh.getRange(targetRow, 15).setValue(data.notas || '');

    return jsonResponse_({
      status: 'ok',
      rowNumber: targetRow,
      timestamp: new Date().toISOString(),
    });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/** Ping simple para verificar que el Web App está desplegado. */
function doGet(e) {
  return jsonResponse_({ status: 'ok', message: 'ROIS Orchestrator activo.' });
}

/**
 * Convierte "YYYY-MM-DD" (o "YYYY-MM-DDTHH:mm:ss...") a una fecha
 * LOCAL a medianoche. Evita el bug clásico de new Date("YYYY-MM-DD"),
 * que interpreta la cadena como UTC y puede correr la fecha un día
 * hacia atrás en huso horario negativo (ej. México, UTC-6).
 */
function parseFechaLocal_(fechaStr) {
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(fechaStr));
  if (m) {
    const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    return new Date(y, mo - 1, d);
  }
  return new Date(fechaStr);
}

function findNextEmptyRegistroRow_(sh) {
  const maxRows = CFG.MAX_ROWS_REGISTROS;
  const fechas = sh.getRange(2, 2, maxRows - 1, 1).getValues();
  for (let i = 0; i < fechas.length; i++) {
    if (fechas[i][0] === '' || fechas[i][0] === null) return i + 2;
  }
  throw new Error('REGISTROS lleno (' + maxRows + ' filas) en ese Sheet.');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
