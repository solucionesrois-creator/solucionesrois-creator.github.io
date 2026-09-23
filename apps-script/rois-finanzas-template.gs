/**
 * ROIS Consultoría — Template base de Finanzas Personales
 * ============================================================
 * ARQUITECTURA "HACER UNA COPIA" (privacidad real, no de procedimiento):
 * Este archivo SÍ incluye doPost/doGet — a propósito. Cada cliente
 * despliega su PROPIO Web App desde SU PROPIA copia del Sheet, bajo
 * SU PROPIA cuenta de Google. ROIS nunca ve sus cifras reales porque
 * el script nunca corre bajo la cuenta de ROIS para un cliente ya
 * entregado. Por eso doPost usa SpreadsheetApp.getActiveSpreadsheet()
 * — nunca un ID fijo ni un ID recibido por parámetro.
 *
 * Flujo completo:
 *  1. El Orchestrator (apps-script/rois-orchestrator.gs) clona este
 *     template y llena ⚙️ MI SETUP con la estructura del cliente
 *     (nombres de cuentas/deudas/servicios — nunca montos).
 *  2. Daniel comparte con el cliente el link "…/copy" de ese archivo
 *     prellenado (no el link "/edit").
 *  3. El cliente abre ese link → Google le pide "Hacer una copia" →
 *     la copia resultante es 100% suya, Daniel pierde acceso.
 *  4. El cliente (guiado por Daniel en una llamada corta) despliega
 *     SU PROPIO Web App desde esa copia: Extensiones → Apps Script →
 *     Implementar → Nueva implementación → Aplicación web →
 *     Ejecutar como: Yo (el cliente) → Acceso: Cualquier usuario.
 *  5. Esa URL /exec es solo del cliente. Nunca vuelve a pasar por
 *     ROIS.
 *
 * Cómo usarlo AQUÍ, en el template maestro (antes de clonar):
 *  1. Crea un Google Sheets en blanco (o duplica uno vacío).
 *  2. Extensiones → Apps Script. Borra el contenido de Code.gs y pega
 *     este archivo completo.
 *  3. En el selector de función (arriba, junto a "Depurar") elige
 *     "construirTemplateROIS" y presiona Ejecutar. Autoriza los
 *     permisos que pida (es tu propia hoja).
 *  4. Revisa el resultado: 7 hojas creadas, formato aplicado,
 *     fórmulas conectadas.
 *  5. NO despliegues Web App en el archivo maestro ni en ningún
 *     archivo de staging que aún esté en tu Drive — solo se despliega
 *     DESPUÉS de que el cliente hizo su propia copia, y lo hace él.
 *
 * Decisiones de diseño (para que Daniel las conozca):
 *  - Capacidad fija: 6 cuentas, 8 deudas, 12 servicios fijos por
 *    cliente. Si un cliente necesita más filas, se insertan filas
 *    dentro del bloque correspondiente en ⚙️ MI SETUP y se copian las
 *    fórmulas de la fila de arriba hacia abajo en REGISTROS/DASHBOARD/
 *    VENCIMIENTOS/DEUDAS/TABLAS.
 *  - "Monto original" de una deuda = el "Monto actual" que el cliente
 *    capturó la primera vez en ⚙️ MI SETUP. A partir de ahí el sistema
 *    resta abonos a capital registrados en REGISTROS.
 *  - % de ahorro automático y % de reserva SAT se capturan como
 *    decimales (0.10 = 10%), NO en formato de porcentaje de Sheets,
 *    para evitar el error clásico de "escribí 10 y se volvió 1000%".
 *  - Se agregó una sección de "Presupuesto por categoría" en
 *    ⚙️ MI SETUP porque el DASHBOARD la necesita y no existía en la
 *    lista original de campos — ver nota en el mensaje de entrega.
 *  - CONVENCIÓN OBLIGATORIA para que "Monto actual" en DEUDAS baje
 *    solo y para que VENCIMIENTOS marque "Pagado": la columna
 *    Subcategoría (F) de REGISTROS debe tener EXACTAMENTE el mismo
 *    texto que el nombre del acreedor o del servicio fijo en
 *    ⚙️ MI SETUP columna A (ej. acreedor "Plata Card" → Subcategoría
 *    "Plata Card"). Si no coincide letra por letra, el abono/pago no
 *    se refleja. El Orchestrator (personalizarSubcategoria_, en
 *    rois-orchestrator.gs) deja un dropdown de SUGERENCIA con todos
 *    los nombres de deudas + servicios de ese cliente en Subcategoría
 *    de REGISTROS y de 📝 REGISTRAR — permite texto libre también
 *    (setAllowInvalid true), para no bloquear conceptos que no son ni
 *    deuda ni servicio fijo.
 */

const CONFIG = {
  COLORS: {
    HEADER_BG: '#1A2332',
    HEADER_FG: '#FFFFFF',
    INGRESO_BG: '#D5F5E3',
    EGRESO_BG: '#FADBD8',
    CARGO_TC_BG: '#D6EAF8',
    SETUP_YELLOW: '#FFF9C4',
    SECTION_BG: '#D5D8DC',
    SUBSECTION_BG: '#ECF0F1',
  },
  CAP: {
    CUENTAS: 6,
    DEUDAS: 8,
    FIJOS: 12,
  },
  SHEETS: {
    SETUP: '⚙️ MI SETUP',
    REGISTRAR: '📝 REGISTRAR',
    REGISTROS: 'REGISTROS',
    DASHBOARD: 'DASHBOARD',
    VENCIMIENTOS: 'VENCIMIENTOS',
    DEUDAS: 'DEUDAS',
    TABLAS: 'TABLAS',
  },
  MAX_ROWS_REGISTROS: 1000,
  // Estas 11 siempre están activas y son las únicas que se rastrean en
  // PRESUPUESTO/DASHBOARD (presupuesto vs real). Cambiar esta lista
  // cambia también esas secciones — no lo hagas sin avisar a Daniel.
  CATEGORIAS_PRESUPUESTO: [
    'Ahorro / Donativo', 'SAT Reserva', 'Pago deuda', 'Empresa', 'Alimentación',
    'Despensa/Reserva', 'Higiene y limpieza', 'Hogar', 'Transporte',
    'Personal', 'Imprevisto',
  ],
  // Disponibles para categorizar en REGISTROS/📝 REGISTRAR, pero SIN
  // fila propia en PRESUPUESTO — evita tener que personalizar el
  // template por cliente solo por esto.
  CATEGORIAS_OPCIONALES: [
    'Restaurantes', 'Salud y medicamentos', 'Ropa y calzado', 'Educación y cursos',
    'Entretenimiento', 'Mascotas', 'Regalos y fechas especiales', 'Viajes',
    'Cuidado personal', 'Mantenimiento del hogar', 'Ahorro o inversión adicional',
  ],
};
CONFIG.CATEGORIAS_TODAS = CONFIG.CATEGORIAS_PRESUPUESTO.concat(CONFIG.CATEGORIAS_OPCIONALES);

// Filas fijas dentro de 📝 REGISTRAR (columna B tiene los valores).
const PANEL_ROWS = {
  fecha: 4, tipo: 5, concepto: 6, categoria: 7, subcategoria: 8,
  monto: 9, cuenta: 10, metodo: 11, empresaPersonal: 12, notas: 13,
};

// Filas fijas dentro de ⚙️ MI SETUP — referenciadas por fórmulas en
// todas las demás hojas. Si cambias la capacidad en CONFIG.CAP,
// actualiza también estos números.
const SETUP_ROWS = {
  cuentasHeader: 5,
  cuentasStart: 6,
  cuentasEnd: 6 + CONFIG.CAP.CUENTAS - 1, // 11
  deudasHeader: 14,
  deudasStart: 15,
  deudasEnd: 15 + CONFIG.CAP.DEUDAS - 1, // 22
  fijosHeader: 25,
  fijosStart: 26,
  fijosEnd: 26 + CONFIG.CAP.FIJOS - 1, // 37
  metaRow: 40,
  ahorroRow: 41,
  satRow: 42,
  presupuestoHeader: 45,
  presupuestoStart: 46,
  presupuestoEnd: 46 + CONFIG.CATEGORIAS_PRESUPUESTO.length - 1, // 56
};

// ============================================================
// ORQUESTADOR PRINCIPAL
// ============================================================
function construirTemplateROIS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try { ss.setSpreadsheetLocale('es_MX'); } catch (e) { /* no crítico */ }

  buildSetup_(ss);
  buildRegistrar_(ss);
  buildRegistros_(ss);
  buildTablas_(ss);
  buildDeudas_(ss);
  buildVencimientos_(ss);
  buildDashboard_(ss);
  buildGraficas_(ss);
  reorderSheets_(ss);

  // Elimina la hoja "Hoja 1" / "Sheet1" en blanco que Sheets crea por defecto
  const defaultSheet = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'Template ROIS construido correctamente.\n\n' +
    'Ve a ⚙️ MI SETUP y llena únicamente las celdas amarillas.'
  );
}

function reorderSheets_(ss) {
  const order = [
    CONFIG.SHEETS.SETUP, CONFIG.SHEETS.REGISTRAR, CONFIG.SHEETS.REGISTROS,
    CONFIG.SHEETS.DASHBOARD, CONFIG.SHEETS.VENCIMIENTOS, CONFIG.SHEETS.DEUDAS,
    CONFIG.SHEETS.TABLAS,
  ];
  order.forEach((name, i) => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.setActiveSheet(sh) && ss.moveActiveSheet(i + 1);
  });
  ss.setActiveSheet(ss.getSheetByName(CONFIG.SHEETS.SETUP));
}

function getOrCreateSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (sh) {
    sh.clear();
    sh.clearFormats();
    sh.clearConditionalFormatRules();
    sh.getDataRange().clearDataValidations();
  } else {
    sh = ss.insertSheet(name);
  }
  return sh;
}

function setSectionHeader_(range, label) {
  range.merge()
    .setValue(label)
    .setBackground(CONFIG.COLORS.SUBSECTION_BG)
    .setFontWeight('bold')
    .setFontSize(11);
}

function setColumnHeaders_(range, labels) {
  range.setValues([labels])
    .setBackground(CONFIG.COLORS.SECTION_BG)
    .setFontWeight('bold');
}

function markYellow_(range) {
  return range.setBackground(CONFIG.COLORS.SETUP_YELLOW);
}

// ============================================================
// HOJA 1: ⚙️ MI SETUP
// ============================================================
function buildSetup_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.SETUP);
  sh.setTabColor('#F1C40F');
  sh.setColumnWidths(1, 5, 160);

  sh.getRange('A1:E1').merge()
    .setValue('⚙️ MI SETUP — Configuración del cliente')
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontSize(13).setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange('A2:E2').merge()
    .setValue('Llena SOLO las celdas amarillas. El resto del sistema se actualiza solo.')
    .setFontStyle('italic').setFontColor('#7F8C8D')
    .setHorizontalAlignment('center');

  // --- CUENTAS BANCARIAS ---
  setSectionHeader_(sh.getRange(4, 1, 1, 3), '🏦 CUENTAS BANCARIAS');
  setColumnHeaders_(sh.getRange(SETUP_ROWS.cuentasHeader, 1, 1, 3),
    ['Nombre de cuenta', 'Tipo', 'Saldo inicial']);
  const nCuentas = CONFIG.CAP.CUENTAS;
  markYellow_(sh.getRange(SETUP_ROWS.cuentasStart, 1, nCuentas, 3));
  sh.getRange(SETUP_ROWS.cuentasStart, 3, nCuentas, 1).setNumberFormat('$#,##0.00');
  sh.getRange(SETUP_ROWS.cuentasStart, 2, nCuentas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Débito', 'Tarjeta de crédito', 'Efectivo', 'Otro'], true)
      .setAllowInvalid(true)
      .build()
  );

  // --- DEUDAS ---
  setSectionHeader_(sh.getRange(SETUP_ROWS.deudasHeader - 1, 1, 1, 5), '💳 DEUDAS');
  setColumnHeaders_(sh.getRange(SETUP_ROWS.deudasHeader, 1, 1, 5),
    ['Nombre', 'Tipo', 'Monto actual', 'Tasa mensual %', 'Pago mensual']);
  const nDeudas = CONFIG.CAP.DEUDAS;
  markYellow_(sh.getRange(SETUP_ROWS.deudasStart, 1, nDeudas, 5));
  sh.getRange(SETUP_ROWS.deudasStart, 3, nDeudas, 1).setNumberFormat('$#,##0.00');
  sh.getRange(SETUP_ROWS.deudasStart, 4, nDeudas, 1).setNumberFormat('0.000')
    .setNote('Ingresa como número, ej. 8.325 para 8.325% mensual. NO uses el signo %.');
  sh.getRange(SETUP_ROWS.deudasStart, 5, nDeudas, 1).setNumberFormat('$#,##0.00');
  sh.getRange(SETUP_ROWS.deudasStart, 2, nDeudas, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Personal', 'Tarjeta de crédito', 'Préstamo digital', 'Hipotecario', 'Otro'], true)
      .setAllowInvalid(true)
      .build()
  );

  // --- SERVICIOS FIJOS ---
  setSectionHeader_(sh.getRange(SETUP_ROWS.fijosHeader - 1, 1, 1, 4), '🧾 SERVICIOS FIJOS');
  setColumnHeaders_(sh.getRange(SETUP_ROWS.fijosHeader, 1, 1, 4),
    ['Nombre', 'Día de cargo', 'Monto', 'Cuenta/TC']);
  const nFijos = CONFIG.CAP.FIJOS;
  markYellow_(sh.getRange(SETUP_ROWS.fijosStart, 1, nFijos, 4));
  sh.getRange(SETUP_ROWS.fijosStart, 2, nFijos, 1).setNumberFormat('0')
    .setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(1, 31).setAllowInvalid(true).build());
  sh.getRange(SETUP_ROWS.fijosStart, 3, nFijos, 1).setNumberFormat('$#,##0.00');
  sh.getRange(SETUP_ROWS.fijosStart, 4, nFijos, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(sh.getRange(SETUP_ROWS.cuentasStart, 1, nCuentas, 1), true)
      .setAllowInvalid(true)
      .build()
  );

  // --- CONFIGURACIÓN GENERAL ---
  setSectionHeader_(sh.getRange(SETUP_ROWS.metaRow - 1, 1, 1, 2), '🎯 CONFIGURACIÓN GENERAL');
  sh.getRange(SETUP_ROWS.metaRow, 1).setValue('Meta de ingreso mensual (opcional)');
  markYellow_(sh.getRange(SETUP_ROWS.metaRow, 2)).setNumberFormat('$#,##0.00');

  sh.getRange(SETUP_ROWS.ahorroRow, 1).setValue('% de ahorro automático (decimal)');
  markYellow_(sh.getRange(SETUP_ROWS.ahorroRow, 2))
    .setNumberFormat('0.00').setValue(0.10)
    .setNote('Escribe como decimal: 0.10 = 10%. No uses formato de porcentaje.');

  sh.getRange(SETUP_ROWS.satRow, 1).setValue('% de reserva SAT (decimal, opcional)');
  markYellow_(sh.getRange(SETUP_ROWS.satRow, 2))
    .setNumberFormat('0.00').setValue(0)
    .setNote('Escribe como decimal: 0.10 = 10%. Déjalo en 0 si no aplica.');

  // --- PRESUPUESTO POR CATEGORÍA ---
  // Agregado porque el DASHBOARD pide "Presupuesto vs Real" y este
  // campo no existía en ninguna otra sección de captura del cliente.
  setSectionHeader_(sh.getRange(SETUP_ROWS.presupuestoHeader - 1, 1, 1, 2),
    '📊 PRESUPUESTO MENSUAL POR CATEGORÍA');
  setColumnHeaders_(sh.getRange(SETUP_ROWS.presupuestoHeader, 1, 1, 2), ['Categoría', 'Presupuesto']);
  const cats = CONFIG.CATEGORIAS_PRESUPUESTO;
  const catRows = cats.map(c => [c]);
  sh.getRange(SETUP_ROWS.presupuestoStart, 1, cats.length, 1).setValues(catRows).setFontWeight('bold');
  markYellow_(sh.getRange(SETUP_ROWS.presupuestoStart, 2, cats.length, 1)).setNumberFormat('$#,##0.00');

  sh.setFrozenRows(2);
}

// ============================================================
// HOJA 2: 📝 REGISTRAR — captura rápida desde la computadora, sin
// depender del celular. Se usa junto con el menú ROIS ▸ Registrar
// movimiento (ver onOpen_/registrarDesdePanel_ más abajo).
// ============================================================
function buildRegistrar_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.REGISTRAR);
  sh.setTabColor('#2ECC71');
  sh.setColumnWidths(1, 1, 180);
  sh.setColumnWidths(2, 1, 260);

  sh.getRange('A1:B1').merge()
    .setValue('📝 Registrar movimiento')
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(12);
  sh.getRange('A2:B2').merge()
    .setValue('Llena los campos y usa el menú ROIS ▸ Registrar movimiento (arriba). No necesitas el celular.')
    .setFontStyle('italic').setFontColor('#7F8C8D');

  const labels = [
    [PANEL_ROWS.fecha, 'Fecha'],
    [PANEL_ROWS.tipo, 'Tipo'],
    [PANEL_ROWS.concepto, 'Concepto'],
    [PANEL_ROWS.categoria, 'Categoría'],
    [PANEL_ROWS.subcategoria, 'Subcategoría (opcional)'],
    [PANEL_ROWS.monto, 'Monto'],
    [PANEL_ROWS.cuenta, 'Cuenta'],
    [PANEL_ROWS.metodo, 'Método'],
    [PANEL_ROWS.empresaPersonal, 'Empresa/Personal'],
    [PANEL_ROWS.notas, 'Notas (opcional)'],
  ];
  labels.forEach(([row, label]) => sh.getRange(row, 1).setValue(label).setFontWeight('bold'));

  markYellow_(sh.getRange(PANEL_ROWS.fecha, 2)).setValue(new Date()).setNumberFormat('dd/mm/yyyy');
  markYellow_(sh.getRange(PANEL_ROWS.tipo, 2)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ingreso', 'Egreso', 'Cargo TC', 'Saldo inicial'], true)
      .setAllowInvalid(false).build()
  );
  markYellow_(sh.getRange(PANEL_ROWS.concepto, 2));
  markYellow_(sh.getRange(PANEL_ROWS.categoria, 2)).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.CATEGORIAS_TODAS, true).setAllowInvalid(true).build()
  );
  markYellow_(sh.getRange(PANEL_ROWS.subcategoria, 2));
  markYellow_(sh.getRange(PANEL_ROWS.monto, 2)).setNumberFormat('$#,##0.00');

  const setupSh = ss.getSheetByName(CONFIG.SHEETS.SETUP);
  markYellow_(sh.getRange(PANEL_ROWS.cuenta, 2)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(setupSh.getRange(SETUP_ROWS.cuentasStart, 1, CONFIG.CAP.CUENTAS, 1), true)
      .setAllowInvalid(true).build()
  );
  markYellow_(sh.getRange(PANEL_ROWS.metodo, 2)).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Transferencia', 'Efectivo', 'Domiciliado', 'Débito', 'SPEI'], true)
      .setAllowInvalid(true).build()
  );
  markYellow_(sh.getRange(PANEL_ROWS.empresaPersonal, 2)).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Empresa', 'Personal'], true).setAllowInvalid(true).build()
  );
  markYellow_(sh.getRange(PANEL_ROWS.notas, 2));

  sh.setFrozenRows(2);
}

// ============================================================
// HOJA 3: REGISTROS
// ============================================================
function buildRegistros_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.REGISTROS);
  const maxRows = CONFIG.MAX_ROWS_REGISTROS;

  const headers = ['#', 'Fecha', 'Tipo', 'Concepto', 'Categoría', 'Subcategoría',
    'Monto', 'Cuenta', 'Método', 'Ref. Ahorro (auto)', 'Ref. SAT (auto)',
    'Ref. Disponible (auto)', 'Empresa/Personal', 'Mueve dinero (auto)', 'Notas', 'Mes-Año (auto)'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(11);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 1, 40);
  sh.setColumnWidths(4, 1, 200);

  // Fórmulas automáticas para todas las filas de datos
  const aFormulas = [], jFormulas = [], kFormulas = [], lFormulas = [], nFormulas = [], pFormulas = [];
  for (let r = 2; r <= maxRows; r++) {
    aFormulas.push([`=IF(B${r}="","",ROW()-1)`]);
    jFormulas.push([`=IF(C${r}="Ingreso",G${r}*'${CONFIG.SHEETS.SETUP}'!$B$${SETUP_ROWS.ahorroRow},"")`]);
    kFormulas.push([`=IF(C${r}="Ingreso",G${r}*'${CONFIG.SHEETS.SETUP}'!$B$${SETUP_ROWS.satRow},"")`]);
    lFormulas.push([`=IF(C${r}="Ingreso",G${r}-N(J${r})-N(K${r}),"")`]);
    nFormulas.push([`=IF(B${r}="","",IF(C${r}="Ingreso","✅ Entrada de efectivo",IF(C${r}="Cargo TC","❌ No — cargo a tarjeta",IF(C${r}="Egreso","💸 Salida de efectivo",""))))`]);
    pFormulas.push([`=IF(B${r}="","",PROPER(TEXT(B${r},"mmmm yyyy")))`]);
  }
  sh.getRange(2, 1, maxRows - 1, 1).setFormulas(aFormulas);
  sh.getRange(2, 10, maxRows - 1, 1).setFormulas(jFormulas).setNumberFormat('$#,##0.00');
  sh.getRange(2, 11, maxRows - 1, 1).setFormulas(kFormulas).setNumberFormat('$#,##0.00');
  sh.getRange(2, 12, maxRows - 1, 1).setFormulas(lFormulas).setNumberFormat('$#,##0.00');
  sh.getRange(2, 14, maxRows - 1, 1).setFormulas(nFormulas);
  sh.getRange(2, 16, maxRows - 1, 1).setFormulas(pFormulas);
  sh.getRange(2, 7, maxRows - 1, 1).setNumberFormat('$#,##0.00');
  sh.getRange(2, 2, maxRows - 1, 1).setNumberFormat('dd/mm/yyyy');

  // Validaciones
  sh.getRange(2, 3, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ingreso', 'Egreso', 'Cargo TC', 'Saldo inicial'], true)
      .setAllowInvalid(false).build()
  );
  sh.getRange(2, 5, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.CATEGORIAS_TODAS, true)
      .setAllowInvalid(true).build()
  );
  const setupSh = ss.getSheetByName(CONFIG.SHEETS.SETUP);
  sh.getRange(2, 8, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(setupSh.getRange(SETUP_ROWS.cuentasStart, 1, CONFIG.CAP.CUENTAS, 1), true)
      .setAllowInvalid(true).build()
  );
  sh.getRange(2, 9, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Transferencia', 'Efectivo', 'Domiciliado', 'Débito', 'SPEI'], true)
      .setAllowInvalid(true).build()
  );
  sh.getRange(2, 13, maxRows - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Empresa', 'Personal'], true)
      .setAllowInvalid(true).build()
  );

  // Formato condicional por tipo de movimiento
  const fullRange = sh.getRange(2, 1, maxRows - 1, headers.length);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="Ingreso"').setBackground(CONFIG.COLORS.INGRESO_BG)
      .setRanges([fullRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="Egreso"').setBackground(CONFIG.COLORS.EGRESO_BG)
      .setRanges([fullRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$C2="Cargo TC"').setBackground(CONFIG.COLORS.CARGO_TC_BG)
      .setRanges([fullRange]).build(),
  ];
  sh.setConditionalFormatRules(rules);
}

// ============================================================
// HOJA 6: TABLAS (se construye antes que DASHBOARD porque el
// selector de mes del DASHBOARD toma su lista de aquí)
// ============================================================
function buildTablas_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.TABLAS);
  const R = CONFIG.SHEETS.REGISTROS, S = CONFIG.SHEETS.SETUP, D = CONFIG.SHEETS.DEUDAS;

  // Tabla 1: Ingresos vs Egresos por mes (mes actual + próximos 11) — el
  // selector del DASHBOARD (B3) toma su lista de esta columna A, por
  // eso arranca en el mes actual hacia adelante y no hacia atrás: un
  // cliente nuevo no tiene meses pasados que revisar, pero sí necesita
  // planear los que vienen.
  setSectionHeader_(sh.getRange(1, 1, 1, 4), 'Ingresos vs Egresos por mes (mes actual + próximos 11)');
  setColumnHeaders_(sh.getRange(2, 1, 1, 4), ['Mes-Año', 'Ingresos', 'Egresos', 'Balance neto']);
  for (let i = 0; i < 12; i++) {
    const row = 3 + i;
    sh.getRange(row, 1).setFormula(`=PROPER(TEXT(EDATE(TODAY(),${i}),"mmmm yyyy"))`);
    sh.getRange(row, 2).setFormula(`=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Ingreso",${R}!$P:$P,A${row})`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 3).setFormula(`=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Egreso",${R}!$P:$P,A${row})`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 4).setFormula(`=B${row}-C${row}`).setNumberFormat('$#,##0.00');
  }

  // Tabla 2: Gastos por categoría del mes actual
  setSectionHeader_(sh.getRange(16, 1, 1, 2), 'Gastos por categoría (mes actual)');
  setColumnHeaders_(sh.getRange(17, 1, 1, 2), ['Categoría', 'Monto']);
  const cats = CONFIG.CATEGORIAS_PRESUPUESTO;
  cats.forEach((cat, i) => {
    const row = 18 + i;
    sh.getRange(row, 1).setValue(cat);
    sh.getRange(row, 2).setFormula(
      `=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Egreso",${R}!$E:$E,A${row},${R}!$P:$P,PROPER(TEXT(TODAY(),"mmmm yyyy")))`
    ).setNumberFormat('$#,##0.00');
  });

  // Tabla 3: Progreso de deudas
  setSectionHeader_(sh.getRange(30, 1, 1, 2), 'Progreso de deudas');
  setColumnHeaders_(sh.getRange(31, 1, 1, 2), ['Acreedor', '% Completado']);
  for (let i = 0; i < CONFIG.CAP.DEUDAS; i++) {
    const row = 32 + i;
    const deudaRow = 3 + i;
    sh.getRange(row, 1).setFormula(`=IF(${D}!A${deudaRow}="","",${D}!A${deudaRow})`);
    sh.getRange(row, 2).setFormula(`=IF(${D}!A${deudaRow}="","",${D}!K${deudaRow})`).setNumberFormat('0.0%');
  }

  // Tabla 4: Meta vs Ingresos reales (mes actual + próximos 11)
  setSectionHeader_(sh.getRange(41, 1, 1, 4), 'Meta vs Ingresos reales (mes actual + próximos 11)');
  setColumnHeaders_(sh.getRange(42, 1, 1, 4), ['Mes-Año', 'Meta', 'Ingreso real', 'Cumplimiento %']);
  for (let i = 0; i < 12; i++) {
    const row = 43 + i;
    sh.getRange(row, 1).setFormula(`=PROPER(TEXT(EDATE(TODAY(),${i}),"mmmm yyyy"))`);
    sh.getRange(row, 2).setFormula(`='${S}'!$B$${SETUP_ROWS.metaRow}`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 3).setFormula(`=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Ingreso",${R}!$P:$P,A${row})`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 4).setFormula(`=IF(B${row}=0,"",C${row}/B${row})`).setNumberFormat('0.0%');
  }
}

// ============================================================
// HOJA 5: DEUDAS
// ============================================================
function buildDeudas_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.DEUDAS);
  const S = CONFIG.SHEETS.SETUP;
  const R = CONFIG.SHEETS.REGISTROS;
  const n = CONFIG.CAP.DEUDAS;
  const dataStart = 3, dataEnd = 2 + n;

  sh.getRange(1, 1, 1, 14).merge()
    .setValue('💳 Seguimiento de deudas')
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(12);

  const headers = ['Acreedor', 'Tipo', 'Monto original', 'Monto actual', 'Tasa %',
    'Pago mensual', 'Interés mensual', 'Abono a capital', 'Meses restantes',
    'Fecha estimada liquidación', '% Completado', 'Barra de progreso', 'Prioridad', 'Estrategia'];
  setColumnHeaders_(sh.getRange(2, 1, 1, headers.length), headers);

  for (let i = 0; i < n; i++) {
    const row = dataStart + i;
    const setupRow = SETUP_ROWS.deudasStart + i;
    sh.getRange(row, 1).setFormula(`='${S}'!A${setupRow}`); // Acreedor
    sh.getRange(row, 2).setFormula(`='${S}'!B${setupRow}`); // Tipo
    sh.getRange(row, 3).setFormula(`='${S}'!C${setupRow}`).setNumberFormat('$#,##0.00'); // Monto original (baseline)
    sh.getRange(row, 4).setFormula(
      `=IF(A${row}="","",MAX(0,C${row}-SUMIFS(${R}!$G:$G,${R}!$C:$C,"Egreso",${R}!$F:$F,A${row})))`
    ).setNumberFormat('$#,##0.00'); // Monto actual = original - abonos registrados (empareja por Subcategoría, no Concepto)
    sh.getRange(row, 5).setFormula(`='${S}'!D${setupRow}`).setNumberFormat('0.000'); // Tasa %
    sh.getRange(row, 6).setFormula(`='${S}'!E${setupRow}`).setNumberFormat('$#,##0.00'); // Pago mensual
    sh.getRange(row, 7).setFormula(`=IF(A${row}="","",D${row}*E${row}/100)`).setNumberFormat('$#,##0.00'); // Interés mensual
    sh.getRange(row, 8).setFormula(`=IF(A${row}="","",MAX(0,F${row}-G${row}))`).setNumberFormat('$#,##0.00'); // Abono a capital
    sh.getRange(row, 9).setFormula(`=IF(A${row}="","",IF(H${row}<=0,"Sin acuerdo",IF(D${row}<=0,0,ROUNDUP(D${row}/H${row},0))))`); // Meses restantes
    sh.getRange(row, 10).setFormula(`=IF(A${row}="","",IF(ISNUMBER(I${row}),TEXT(EDATE(TODAY(),I${row}),"dd/mm/yyyy"),"—"))`); // Fecha estimada
    sh.getRange(row, 11).setFormula(`=IF(A${row}="","",IF(C${row}=0,0,(C${row}-D${row})/C${row}))`).setNumberFormat('0.0%'); // % Completado
    sh.getRange(row, 12).setFormula(`=IF(A${row}="","",REPT("█",ROUND(K${row}*10,0))&REPT("░",10-ROUND(K${row}*10,0)))`); // Barra
  }
  sh.getRange(dataStart, 13, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Urgente', 'Media', 'Baja'], true).setAllowInvalid(true).build()
  );
  sh.getRange(dataStart, 14, n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pago activo', 'Sin acuerdo — moratorio', 'Saldada'], true)
      .setAllowInvalid(true).build()
  );
  sh.getRange(dataStart, 11, n, 1).setBackground(null);
  // Color scale sobre % Completado para dar la sensación de barra visual
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .setGradientMaxpointWithValue('#2ECC71', SpreadsheetApp.InterpolationType.NUMBER, '1')
      .setGradientMinpointWithValue('#FADBD8', SpreadsheetApp.InterpolationType.NUMBER, '0')
      .setRanges([sh.getRange(dataStart, 11, n, 1)])
      .build(),
  ]);

  // Totales
  const totalsRow = dataEnd + 2;
  sh.getRange(totalsRow, 1).setValue('Total deuda actual:').setFontWeight('bold');
  sh.getRange(totalsRow, 2).setFormula(`=SUM(D${dataStart}:D${dataEnd})`).setNumberFormat('$#,##0.00');
  sh.getRange(totalsRow + 1, 1).setValue('Costo mensual en intereses:').setFontWeight('bold');
  sh.getRange(totalsRow + 1, 2).setFormula(`=SUM(G${dataStart}:G${dataEnd})`).setNumberFormat('$#,##0.00');
  sh.getRange(totalsRow + 2, 1).setValue('Abono total mensual:').setFontWeight('bold');
  sh.getRange(totalsRow + 2, 2).setFormula(`=SUM(F${dataStart}:F${dataEnd})`).setNumberFormat('$#,##0.00');

  // Escenarios de liquidación para la deuda más grande
  const escRow = totalsRow + 4;
  sh.getRange(escRow, 1, 1, 7).merge()
    .setValue('🎯 ESCENARIOS — Liquidar la deuda más grande')
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.HEADER_FG).setFontWeight('bold');
  setColumnHeaders_(sh.getRange(escRow + 1, 1, 1, 7),
    ['Escenario', 'Abono/mes', 'Interés/mes', 'Capital real/mes', 'Meses', 'Fecha libre', 'Total pagado (aprox)']);

  const maxMontoRef = `MAX($D$${dataStart}:$D$${dataEnd})`;
  const maxTasaRef = `INDEX($E$${dataStart}:$E$${dataEnd},MATCH(${maxMontoRef},$D$${dataStart}:$D$${dataEnd},0))`;
  const maxPagoRef = `INDEX($F$${dataStart}:$F$${dataEnd},MATCH(${maxMontoRef},$D$${dataStart}:$D$${dataEnd},0))`;
  const escenarios = [
    ['Mínimo', `=${maxPagoRef}`],
    ['Acelerado', `=${maxPagoRef}*1.5`],
    ['Agresivo', `=${maxPagoRef}*2.5`],
  ];
  escenarios.forEach((esc, i) => {
    const row = escRow + 2 + i;
    sh.getRange(row, 1).setValue(esc[0]);
    sh.getRange(row, 2).setFormula(esc[1]).setNumberFormat('$#,##0.00');
    sh.getRange(row, 3).setFormula(`=${maxMontoRef}*${maxTasaRef}/100`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 4).setFormula(`=MAX(0,B${row}-C${row})`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 5).setFormula(`=IF(D${row}<=0,"Sin acuerdo",ROUNDUP(${maxMontoRef}/D${row},0))`);
    sh.getRange(row, 6).setFormula(`=IF(ISNUMBER(E${row}),TEXT(EDATE(TODAY(),E${row}),"mmm yyyy"),"—")`);
    sh.getRange(row, 7).setFormula(`=IF(ISNUMBER(E${row}),B${row}*E${row},"—")`).setNumberFormat('$#,##0.00');
  });
  sh.getRange(escRow + 4, 1, 1, 7).setBackground('#D5F5E3').setFontWeight('bold');

  sh.setFrozenRows(2);
}

// ============================================================
// HOJA 4: VENCIMIENTOS
// ============================================================
function buildVencimientos_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.VENCIMIENTOS);
  const S = CONFIG.SHEETS.SETUP;
  const R = CONFIG.SHEETS.REGISTROS;
  const n = CONFIG.CAP.FIJOS;

  sh.getRange(1, 1, 1, 7).merge()
    .setValue('🧾 Control de vencimientos')
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontWeight('bold').setFontSize(12);
  setColumnHeaders_(sh.getRange(2, 1, 1, 7),
    ['Servicio', 'Día de cargo', 'Monto', 'Cuenta/TC', 'Estado', 'Próx. fecha', 'Alerta']);

  for (let i = 0; i < n; i++) {
    const row = 3 + i;
    const setupRow = SETUP_ROWS.fijosStart + i;
    sh.getRange(row, 1).setFormula(`='${S}'!A${setupRow}`);
    sh.getRange(row, 2).setFormula(`='${S}'!B${setupRow}`);
    sh.getRange(row, 3).setFormula(`='${S}'!C${setupRow}`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 4).setFormula(`='${S}'!D${setupRow}`);
    // Próx. fecha: si el día de cargo ya pasó este mes, usa el próximo mes
    sh.getRange(row, 6).setFormula(
      `=IF(A${row}="","",IF(B${row}>=DAY(TODAY()),DATE(YEAR(TODAY()),MONTH(TODAY()),B${row}),EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),B${row}),1)))`
    ).setNumberFormat('dd/mm/yyyy');
    // Estado: revisa si ya hay un egreso con Subcategoría = este servicio
    // (misma convención que DEUDAS — no el Concepto, que es texto libre)
    // en el mes-año actual.
    sh.getRange(row, 5).setFormula(
      `=IF(A${row}="","",IF(COUNTIFS(${R}!$F:$F,A${row},${R}!$P:$P,PROPER(TEXT(TODAY(),"mmmm yyyy")))>0,"Pagado","Pendiente"))`
    );
    sh.getRange(row, 7).setFormula(
      `=IF(A${row}="","",IF(AND(F${row}-TODAY()<=3,F${row}-TODAY()>=0,E${row}="Pendiente"),"⚠️ PRÓXIMO",""))`
    );
  }

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$G3<>""`).setBackground('#FADBD8').setFontColor('#C0392B').setBold(true)
      .setRanges([sh.getRange(3, 1, n, 7)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$E3="Pagado"`).setBackground(CONFIG.COLORS.INGRESO_BG)
      .setRanges([sh.getRange(3, 1, n, 7)]).build(),
  ]);

  sh.setFrozenRows(2);
}

// ============================================================
// HOJA 3: DASHBOARD
// ============================================================
function buildDashboard_(ss) {
  const sh = getOrCreateSheet_(ss, CONFIG.SHEETS.DASHBOARD);
  const S = CONFIG.SHEETS.SETUP;
  const R = CONFIG.SHEETS.REGISTROS;
  const D = CONFIG.SHEETS.DEUDAS;
  const T = CONFIG.SHEETS.TABLAS;

  sh.getRange('A1:F1').merge()
    .setValue('📊 ROIS — Dashboard Financiero')
    .setBackground(CONFIG.COLORS.HEADER_BG).setFontColor(CONFIG.COLORS.HEADER_FG)
    .setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');

  sh.getRange('A3').setValue('Ver mes:').setFontWeight('bold');
  sh.getRange('B3').setFormula('=PROPER(TEXT(TODAY(),"mmmm yyyy"))')
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireValueInRange(ss.getSheetByName(T).getRange(3, 1, 12, 1), true)
        .setAllowInvalid(true).build()
    ).setBackground('#FCF3CF').setFontWeight('bold');

  // KPIs del mes
  sh.getRange('A5').setValue('INGRESOS DEL MES').setFontWeight('bold');
  sh.getRange('B5').setFormula(`=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Ingreso",${R}!$P:$P,$B$3)`)
    .setNumberFormat('$#,##0.00').setBackground(CONFIG.COLORS.INGRESO_BG).setFontWeight('bold');
  sh.getRange('A6').setValue('EGRESOS DEL MES').setFontWeight('bold');
  sh.getRange('B6').setFormula(`=SUMIFS(${R}!$G:$G,${R}!$C:$C,"Egreso",${R}!$P:$P,$B$3)+SUMIFS(${R}!$G:$G,${R}!$C:$C,"Cargo TC",${R}!$P:$P,$B$3)`)
    .setNumberFormat('$#,##0.00').setBackground(CONFIG.COLORS.EGRESO_BG).setFontWeight('bold');
  sh.getRange('A7').setValue('BALANCE NETO').setFontWeight('bold');
  sh.getRange('B7').setFormula('=B5-B6').setNumberFormat('$#,##0.00').setFontWeight('bold');

  // Ahorro / Donativo / meta
  setSectionHeader_(sh.getRange(9, 1, 1, 2), '🙏 Ahorro / Donativo del mes');
  sh.getRange('A10').setValue('Ahorro generado:');
  sh.getRange('B10').setFormula(`=SUMIFS(${R}!$J:$J,${R}!$P:$P,$B$3)`).setNumberFormat('$#,##0.00');
  sh.getRange('A11').setValue('Reserva SAT:');
  sh.getRange('B11').setFormula(`=SUMIFS(${R}!$K:$K,${R}!$P:$P,$B$3)`).setNumberFormat('$#,##0.00');
  sh.getRange('A12').setValue('Meta de ingreso:');
  sh.getRange('B12').setFormula(`='${S}'!$B$${SETUP_ROWS.metaRow}`).setNumberFormat('$#,##0.00');
  sh.getRange('A13').setValue('Cumplimiento de meta:');
  sh.getRange('B13').setFormula('=IF(B12=0,"—",B5/B12)').setNumberFormat('0.0%');

  // Saldo por cuenta
  setSectionHeader_(sh.getRange(15, 1, 1, 2), '🏦 Saldo por cuenta (histórico)');
  const nCuentas = CONFIG.CAP.CUENTAS;
  for (let i = 0; i < nCuentas; i++) {
    const row = 16 + i;
    const setupRow = SETUP_ROWS.cuentasStart + i;
    sh.getRange(row, 1).setFormula(`='${S}'!A${setupRow}`);
    sh.getRange(row, 2).setFormula(
      `=IF(A${row}="","",'${S}'!C${setupRow}` +
      `+SUMIFS(${R}!$G:$G,${R}!$H:$H,A${row},${R}!$C:$C,"Ingreso")` +
      `-SUMIFS(${R}!$G:$G,${R}!$H:$H,A${row},${R}!$C:$C,"Egreso")` +
      `-SUMIFS(${R}!$G:$G,${R}!$H:$H,A${row},${R}!$C:$C,"Cargo TC"))`
    ).setNumberFormat('$#,##0.00');
  }
  const totalRow = 16 + nCuentas;
  sh.getRange(totalRow, 1).setValue('TOTAL').setFontWeight('bold');
  sh.getRange(totalRow, 2).setFormula(`=SUM(B16:B${totalRow - 1})`).setNumberFormat('$#,##0.00')
    .setFontWeight('bold').setBackground(CONFIG.COLORS.SECTION_BG);

  // Presupuesto vs Real
  const presRow = totalRow + 3;
  setSectionHeader_(sh.getRange(presRow, 1, 1, 5), '📋 Presupuesto vs Real (mes seleccionado)');
  setColumnHeaders_(sh.getRange(presRow + 1, 1, 1, 5),
    ['Categoría', 'Presupuesto', 'Real', 'Diferencia', '% Ejecutado']);
  const cats = CONFIG.CATEGORIAS_PRESUPUESTO;
  cats.forEach((cat, i) => {
    const row = presRow + 2 + i;
    const setupRow = SETUP_ROWS.presupuestoStart + i;
    sh.getRange(row, 1).setValue(cat);
    sh.getRange(row, 2).setFormula(`='${S}'!B${setupRow}`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 3).setFormula(
      `=SUMIFS(${R}!$G:$G,${R}!$E:$E,A${row},${R}!$P:$P,$B$3)`
    ).setNumberFormat('$#,##0.00');
    sh.getRange(row, 4).setFormula(`=C${row}-B${row}`).setNumberFormat('$#,##0.00');
    sh.getRange(row, 5).setFormula(`=IF(B${row}=0,"—",C${row}/B${row})`).setNumberFormat('0%');
  });

  // Estado de deudas
  const deudaRow = presRow + 3 + cats.length;
  setSectionHeader_(sh.getRange(deudaRow, 1, 1, 2), '💳 Estado de deudas activas');
  sh.getRange(deudaRow + 1, 1).setValue('Total deuda actual:');
  sh.getRange(deudaRow + 1, 2).setFormula(`=${D}!B` + (2 + CONFIG.CAP.DEUDAS + 2)).setNumberFormat('$#,##0.00');
  sh.getRange(deudaRow + 2, 1).setValue('Costo mensual en intereses:');
  sh.getRange(deudaRow + 2, 2).setFormula(`=${D}!B` + (2 + CONFIG.CAP.DEUDAS + 3)).setNumberFormat('$#,##0.00');
  sh.getRange(deudaRow + 3, 1).setValue('Abono total mensual:');
  sh.getRange(deudaRow + 3, 2).setFormula(`=${D}!B` + (2 + CONFIG.CAP.DEUDAS + 4)).setNumberFormat('$#,##0.00');

  sh.setColumnWidths(1, 1, 220);
  sh.setColumnWidths(2, 4, 120);
  sh.setFrozenRows(3);
}

// ============================================================
// GRÁFICAS — se insertan en DASHBOARD, a la derecha del resumen
// numérico (columna H en adelante), leyendo de TABLAS. Se borran y
// se vuelven a crear cada vez que corre construirTemplateROIS, para
// que re-ejecutarlo no vaya acumulando gráficas duplicadas.
// ============================================================
function buildGraficas_(ss) {
  const dash = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  const tablas = ss.getSheetByName(CONFIG.SHEETS.TABLAS);
  const catCount = CONFIG.CATEGORIAS_PRESUPUESTO.length;
  const deudaCount = CONFIG.CAP.DEUDAS;

  dash.getCharts().forEach((chart) => dash.removeChart(chart));

  // 1. Ingresos vs Egresos por mes (columnas) — TABLAS!A2:C14
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.COLUMN)
      .addRange(tablas.getRange(2, 1, 13, 3))
      .setPosition(2, 8, 0, 0)
      .setOption('title', 'Ingresos vs Egresos por mes')
      .setOption('width', 480).setOption('height', 300)
      .build()
  );

  // 2. Evolución del balance neto (línea) — TABLAS!A2, D2:D14
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.LINE)
      .addRange(tablas.getRange(2, 1, 13, 1))
      .addRange(tablas.getRange(2, 4, 13, 1))
      .setPosition(20, 8, 0, 0)
      .setOption('title', 'Evolución del balance neto')
      .setOption('width', 480).setOption('height', 300)
      .build()
  );

  // 3. Gastos por categoría del mes (pastel) — TABLAS!A17:B28
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.PIE)
      .addRange(tablas.getRange(17, 1, catCount + 1, 2))
      .setPosition(38, 8, 0, 0)
      .setOption('title', 'Gastos por categoría (mes actual)')
      .setOption('width', 480).setOption('height', 320)
      .build()
  );

  // 4. Progreso de liquidación de deudas (barras) — TABLAS!A31:B39
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(tablas.getRange(31, 1, deudaCount + 1, 2))
      .setPosition(58, 8, 0, 0)
      .setOption('title', 'Progreso de liquidación de deudas')
      .setOption('width', 480).setOption('height', 300)
      .build()
  );

  // 5. Ingresos reales vs Meta (combo) — TABLAS!A42:C54
  dash.insertChart(
    dash.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(tablas.getRange(42, 1, 13, 3))
      .setPosition(76, 8, 0, 0)
      .setOption('title', 'Ingresos reales vs Meta')
      .setOption('width', 480).setOption('height', 300)
      .setOption('series', { 0: { type: 'line' }, 1: { type: 'line' } })
      .build()
  );
}

// ============================================================
// MENÚ — botón "Registrar" dentro del propio Sheet (sin celular)
// ============================================================

/**
 * Trigger simple: se ejecuta solo cuando el USUARIO (dueño de esta
 * copia) abre el Sheet en su navegador — nunca durante el clonado
 * automático del Orchestrator, así que no interfiere con el
 * onboarding.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('ROIS')
    .addItem('▶ Registrar movimiento', 'registrarDesdePanel_')
    .addToUi();
}

/**
 * Lee 📝 REGISTRAR, valida y escribe en REGISTROS — misma lógica que
 * doPost, para que el celular y la computadora nunca diverjan.
 */
function registrarDesdePanel_() {
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.REGISTRAR);
  const datos = {
    fecha: sh.getRange(PANEL_ROWS.fecha, 2).getValue(),
    tipo: sh.getRange(PANEL_ROWS.tipo, 2).getValue(),
    concepto: sh.getRange(PANEL_ROWS.concepto, 2).getValue(),
    categoria: sh.getRange(PANEL_ROWS.categoria, 2).getValue(),
    subcategoria: sh.getRange(PANEL_ROWS.subcategoria, 2).getValue(),
    monto: sh.getRange(PANEL_ROWS.monto, 2).getValue(),
    cuenta: sh.getRange(PANEL_ROWS.cuenta, 2).getValue(),
    metodo: sh.getRange(PANEL_ROWS.metodo, 2).getValue(),
    empresaPersonal: sh.getRange(PANEL_ROWS.empresaPersonal, 2).getValue(),
    notas: sh.getRange(PANEL_ROWS.notas, 2).getValue(),
  };
  try {
    const targetRow = escribirTransaccion_(datos);
    sh.getRange(PANEL_ROWS.fecha, 2).setValue(new Date());
    [PANEL_ROWS.concepto, PANEL_ROWS.categoria, PANEL_ROWS.subcategoria, PANEL_ROWS.monto, PANEL_ROWS.notas]
      .forEach((row) => sh.getRange(row, 2).clearContent());
    ui.alert('✅ Registrado en REGISTROS, fila ' + targetRow + '.');
  } catch (err) {
    ui.alert('❌ ' + (err && err.message ? err.message : err));
  }
}

// ============================================================
// WEB APP — recibe transacciones desde el celular (PWA/Shortcut).
// Corre SIEMPRE sobre este mismo Sheet (getActiveSpreadsheet), nunca
// sobre un ID ajeno — esa es la garantía de privacidad de la
// arquitectura "Hacer una copia".
// ============================================================

/**
 * Valida y escribe una transacción en REGISTROS. Usada tanto por
 * doPost (celular) como por registrarDesdePanel_ (computadora) para
 * que ambos caminos compartan exactamente la misma regla.
 * Lanza un Error con mensaje legible si algo no es válido.
 */
function escribirTransaccion_(datos) {
  const required = ['fecha', 'tipo', 'concepto', 'categoria', 'monto', 'cuenta', 'metodo'];
  for (const key of required) {
    if (datos[key] === undefined || datos[key] === null || datos[key] === '') {
      throw new Error('Falta el campo: ' + key);
    }
  }

  const tiposValidos = ['Ingreso', 'Egreso', 'Cargo TC', 'Saldo inicial'];
  if (tiposValidos.indexOf(datos.tipo) === -1) {
    throw new Error('Tipo inválido: "' + datos.tipo + '". Debe ser exactamente uno de: ' + tiposValidos.join(', '));
  }

  const monto = Number(datos.monto);
  if (isNaN(monto)) throw new Error('Monto inválido: ' + datos.monto);

  const fecha = parseFechaLocal_(datos.fecha);
  if (!fecha || isNaN(fecha.getTime())) throw new Error('Fecha inválida: ' + datos.fecha);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.SHEETS.REGISTROS);
  const targetRow = findNextEmptyRegistroRow_(sh);

  // B..I
  sh.getRange(targetRow, 2, 1, 8).setValues([[
    fecha, datos.tipo, datos.concepto, datos.categoria, datos.subcategoria || '',
    monto, datos.cuenta, datos.metodo,
  ]]);
  // M Empresa/Personal, O Notas
  sh.getRange(targetRow, 13).setValue(datos.empresaPersonal || 'Personal');
  sh.getRange(targetRow, 15).setValue(datos.notas || '');

  return targetRow;
}

/**
 * Espera un POST con JSON:
 * { fecha, tipo, concepto, categoria, subcategoria, monto, cuenta,
 *   metodo, empresaPersonal, notas }
 *
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
    const targetRow = escribirTransaccion_(data);
    return jsonResponse_({ status: 'ok', rowNumber: targetRow, timestamp: new Date().toISOString() });
  } catch (err) {
    return jsonResponse_({ status: 'error', message: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/** Ping simple para verificar que el Web App está desplegado. */
function doGet(e) {
  return jsonResponse_({ status: 'ok', message: 'ROIS Finanzas Web App activo.' });
}

/**
 * Convierte "YYYY-MM-DD" (o "YYYY-MM-DDTHH:mm:ss...") a una fecha
 * LOCAL a medianoche. También acepta un objeto Date directo (lo que
 * entrega una celda con formato de fecha, ej. desde 📝 REGISTRAR).
 * Evita el bug clásico de new Date("YYYY-MM-DD"), que interpreta la
 * cadena como UTC y puede correr la fecha un día hacia atrás en huso
 * horario negativo (ej. México, UTC-6).
 */
function parseFechaLocal_(fechaStr) {
  if (fechaStr instanceof Date) {
    return new Date(fechaStr.getFullYear(), fechaStr.getMonth(), fechaStr.getDate());
  }
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(fechaStr));
  if (m) {
    const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    return new Date(y, mo - 1, d);
  }
  return new Date(fechaStr);
}

function findNextEmptyRegistroRow_(sh) {
  const maxRows = CONFIG.MAX_ROWS_REGISTROS;
  const fechas = sh.getRange(2, 2, maxRows - 1, 1).getValues();
  for (let i = 0; i < fechas.length; i++) {
    if (fechas[i][0] === '' || fechas[i][0] === null) return i + 2;
  }
  throw new Error('REGISTROS lleno (' + maxRows + ' filas). Amplía MAX_ROWS_REGISTROS y vuelve a ejecutar construirTemplateROIS, o inserta filas manualmente copiando las fórmulas de la última fila.');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================
 * DESPLIEGUE DEL WEB APP — lo hace EL CLIENTE, en SU copia, UNA vez:
 * ============================================================
 *
 * 1. En el editor de Apps Script (Extensiones → Apps Script desde SU
 *    copia del Sheet): Implementar → Nueva implementación.
 * 2. Tipo de implementación: "Aplicación web" (ícono de engrane si no
 *    aparece directo).
 * 3. Descripción: "ROIS Finanzas - Web App v1".
 * 4. "Ejecutar como": Yo (la cuenta del cliente — es SU hoja).
 * 5. "Quién tiene acceso": "Cualquier usuario" (necesario para que
 *    el celular pueda hacer POST sin login interactivo).
 * 6. Clic en Implementar. Va a aparecer una pantalla de Google que
 *    dice "Esta app no está verificada" — es NORMAL y esperado (la
 *    app es literalmente suya); debe darle clic en "Avanzado" y
 *    luego en "Ir a [nombre del proyecto] (no seguro)". Adviértele
 *    esto ANTES de que lo vea, para que no piense que algo salió mal.
 * 7. Copia la "URL de la aplicación web" — termina en /exec. Esa URL
 *    solo la tiene el cliente; se pega en su PWA (registro-financiero)
 *    o en su Shortcut de iPhone.
 * 8. Prueba abriendo la URL /exec en el navegador — debe responder
 *    {"status":"ok","message":"ROIS Finanzas Web App activo."}
 * 9. Si se vuelve a editar el script, los cambios NO se reflejan en
 *    la URL ya desplegada hasta hacer Implementar → Administrar
 *    implementaciones → ✏️ → Nueva versión → Implementar.
 */
