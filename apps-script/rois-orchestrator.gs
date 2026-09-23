/**
 * ROIS Consultoría — Orchestrator de onboarding
 * ============================================================
 * Proyecto Apps Script STANDALONE — NO va pegado dentro de ningún
 * Google Sheets. Se crea UNA sola vez.
 *
 * QUÉ HACE Y QUÉ YA NO HACE (arquitectura "Hacer una copia"):
 * Este Orchestrator SOLO prellena la estructura de cada cliente
 * (nombres de cuentas/deudas/servicios — nunca montos) en una copia
 * de staging que sigue siendo tuya. YA NO tiene doPost ni Web App:
 * cada cliente despliega su PROPIO Web App desde SU PROPIA copia
 * (ver rois-finanzas-template.gs), así que ROIS nunca vuelve a tocar
 * una transacción real de un cliente ya entregado.
 *
 * Cómo instalarlo (una sola vez):
 *  1. https://script.google.com → Proyecto nuevo (NO "Extensiones →
 *     Apps Script" desde un Sheets — eso lo dejaría bound a esa hoja).
 *  2. Bórralo todo y pega este archivo completo. Nómbralo
 *     "ROIS Orchestrator".
 *  3. Selector de función → elige "procesarPendientes" y presiona
 *     Ejecutar una vez, solo para autorizar los permisos que pida
 *     (Drive, Sheets).
 *  4. Para dar de alta clientes nuevos: corre "procesarPendientes"
 *     manualmente (o instala un trigger de tiempo — Activadores →
 *     Añadir activador → procesarPendientes → basado en tiempo →
 *     cada hora; gratis, dentro de las cuotas normales de Apps
 *     Script). Revisa el registro de ejecución (Ver → Registros) para
 *     ver la tabla de resultados: nombre, sheetUrl (staging, tuyo),
 *     linkCopia (esto es lo único que compartes con el cliente).
 *
 * DESPUÉS de dar de alta a un cliente:
 *  1. Comparte con él SOLO el "linkCopia" (nunca el sheetUrl de
 *     staging) — al abrirlo, Google le ofrece "Hacer una copia".
 *  2. Guíalo (llamada corta) para desplegar su propio Web App desde
 *     esa copia — ver instrucciones de despliegue en
 *     rois-finanzas-template.gs.
 *  3. Una vez que confirmes que ya tiene su copia y su Web App
 *     funcionando, borra el archivo de staging de tu Drive
 *     ("PERSONALIZADOS - CLIENTES") — ya cumplió su propósito y no
 *     hay razón para conservar ni siquiera los nombres de sus cuentas.
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
  SHEETS: { SETUP: '⚙️ MI SETUP' },
  CAP: { CUENTAS: 6, DEUDAS: 8, FIJOS: 12 },
  SETUP_ROWS: { cuentasStart: 6, deudasStart: 15, fijosStart: 26, ahorroRow: 41 },
  TIPOS_DEUDA_VALIDOS: ['Personal', 'Tarjeta de crédito', 'Préstamo digital', 'Hipotecario', 'Otro'],
  INDICE_PROP_KEY: 'INDICE_SHEET_ID',
};

// ============================================================
// ÍNDICE DE CLIENTES — solo datos de contacto (nombre/correo/canal),
// NUNCA cifras financieras. Se autocrea la primera vez y guarda su
// ID en Script Properties (no requiere configuración manual).
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
    sh.getRange(1, 1, 1, 6).setValues([[
      'Nombre', 'Email', 'Canal', 'Staging Sheet ID', 'Staging Sheet URL', 'Fecha alta',
    ]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    Logger.log('Índice de clientes creado: ' + ss.getUrl());
  }
  return ss.getSheetByName('CLIENTES') || ss.getSheets()[0];
}

function registrarEnIndice_(cliente, clienteSs) {
  const sh = getIndiceSheet_();
  sh.appendRow([
    cliente.nombre, cliente.email, cliente.origen,
    clienteSs.getId(), clienteSs.getUrl(), new Date(),
  ]);
}

// ============================================================
// ONBOARDING — lee los Sheets de respuestas del Form, clona el
// template por cada cliente pendiente, llena SETUP con su estructura
// real (nunca montos) y marca la fila como Procesado.
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
  const nombreArchivo = 'Finanzas — ' + cliente.nombre + ' | ROIS (staging)';
  const copia = DriveApp.getFileById(CFG.TEMPLATE_ID).makeCopy(nombreArchivo, carpeta);
  const ss = SpreadsheetApp.openById(copia.getId());

  llenarSetup_(ss, cliente);
  registrarEnIndice_(cliente, ss);

  return {
    nombre: cliente.nombre,
    origen: cliente.origen,
    stagingSheetUrl: ss.getUrl(),
    linkCopia: 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/copy',
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
