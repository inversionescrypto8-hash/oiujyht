/**
 * ============================================================================
 *  SISTEMA DE GESTIÓN DE INVENTARIO, COMPRAS, VENTAS, COSTOS Y FACTURACIÓN
 *  Archivo principal: menú, constantes globales y disparadores.
 * ============================================================================
 *
 *  Cómo usar:
 *   1. Abre/crea una hoja de cálculo de Google.
 *   2. Menú: Extensiones > Apps Script.
 *   3. Pega los 4 archivos .gs y el appsscript.json de este proyecto.
 *   4. Guarda y recarga la hoja de cálculo.
 *   5. Aparecerá el menú "📊 Gestión".  Ejecuta "🚀 Inicializar / Reconstruir
 *      sistema" la primera vez para que se construya todo automáticamente.
 * ============================================================================
 */

/** Nombres de las hojas (pestañas) del libro. */
var SHEETS = {
  DASHBOARD: 'Dashboard',
  CATALOGO: 'Catálogo',
  COMPRAS: 'Compras',
  INVENTARIO: 'Inventario',
  TRASLADOS: 'Traslados',
  VENTAS: 'Ventas',
  AJUSTES: 'Ajustes',
  CLIENTES: 'Clientes',
  FACTURA: 'Factura',
  CONFIG: 'Config'
};

/** Ubicaciones del negocio. */
var UBIC = {
  CASA: 'Casa',
  ML: 'Mercado Libre'
};

/** Motivos válidos para los ajustes de inventario. */
var MOTIVOS_AJUSTE = ['Producto dañado', 'Pérdida', 'Corrección de conteo', 'Ajuste administrativo', 'Devolución'];

/** Estados del catálogo. */
var ESTADOS = ['Activo', 'Inactivo'];

/** Paleta de colores corporativos usada en todo el libro. */
var COLOR = {
  PRIMARIO: '#1f3a5f',   // azul profundo
  SECUNDARIO: '#2e6da4', // azul medio
  ACENTO: '#16a085',     // verde
  CLARO: '#eef3f8',      // gris azulado claro
  ALERTA: '#c0392b',     // rojo
  ADVERT: '#e67e22',     // naranja
  TEXTO: '#2c3e50',
  BLANCO: '#ffffff',
  BORDE: '#cfd8e3'
};

/** Número máximo de filas con fórmulas pre-cargadas en hojas de movimientos. */
var MAX_ROWS = 2000;

/**
 * Se ejecuta automáticamente al abrir el libro. Crea el menú personalizado.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 Gestión')
    .addItem('🚀 Inicializar / Reconstruir sistema', 'inicializarSistema')
    .addSeparator()
    .addItem('🔄 Actualizar Dashboard y gráficos', 'actualizarDashboard')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🧾 Facturación')
      .addItem('➕ Nueva factura', 'nuevaFactura')
      .addItem('💾 Guardar venta y registrar inventario', 'guardarFactura')
      .addItem('📄 Generar / Descargar PDF', 'generarPDFFactura')
      .addItem('✉️ Enviar por correo', 'enviarFacturaCorreo')
      .addItem('🟢 Compartir por WhatsApp', 'compartirFacturaWhatsApp'))
    .addSeparator()
    .addItem('❓ Ayuda', 'mostrarAyuda')
    .addToUi();
}

/**
 * Disparador instalable simple: mantiene el número de factura y la fecha
 * coherentes cuando el usuario cambia el cliente en la hoja Factura.
 * (La mayor parte de la lógica es por fórmulas; este hook es de apoyo.)
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== SHEETS.FACTURA) return;
    // Nada obligatorio aquí: el autocompletado del cliente y los totales
    // se resuelven con fórmulas. Se deja el hook por extensibilidad.
  } catch (err) {
    // Silencioso: onEdit no debe interrumpir la edición del usuario.
  }
}

/** Devuelve la hoja por nombre o lanza un error claro si no existe. */
function getSheet_(nombre) {
  var sh = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!sh) {
    throw new Error('No se encontró la hoja "' + nombre + '". Ejecuta "Inicializar / Reconstruir sistema".');
  }
  return sh;
}

/** Atajo para mostrar alertas. */
function alerta_(titulo, mensaje) {
  SpreadsheetApp.getUi().alert(titulo, mensaje, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Lee un parámetro de la hoja Config por su etiqueta. */
function getParametro_(etiqueta) {
  var sh = getSheet_(SHEETS.CONFIG);
  var datos = sh.getRange('A1:B40').getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === etiqueta) return datos[i][1];
  }
  return '';
}

/** Escribe un parámetro de la hoja Config por su etiqueta. */
function setParametro_(etiqueta, valor) {
  var sh = getSheet_(SHEETS.CONFIG);
  var datos = sh.getRange('A1:B40').getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]).trim() === etiqueta) {
      sh.getRange(i + 1, 2).setValue(valor);
      return;
    }
  }
}

/** Muestra ayuda rápida en un diálogo. */
function mostrarAyuda() {
  var msg =
    'GUÍA RÁPIDA\n\n' +
    '1) Catálogo: registra cada producto UNA sola vez (Código, Nombre, Categoría, Estado, Stock mínimo).\n' +
    '2) Compras: cada compra suma stock a Casa y recalcula el costo promedio.\n' +
    '3) Traslados: mueve unidades entre Casa y Mercado Libre.\n' +
    '4) Ventas: descuenta inventario y calcula la utilidad con el costo promedio vigente.\n' +
    '5) Ajustes: registra daños, pérdidas y correcciones sin tocar compras/ventas.\n' +
    '6) Clientes: base de datos para la facturación.\n' +
    '7) Factura: usa el menú 🧾 Facturación para crear, guardar, exportar PDF y compartir.\n\n' +
    'El Dashboard y los stocks se actualizan automáticamente con fórmulas.';
  alerta_('Ayuda del sistema', msg);
}
/**
 * ============================================================================
 *  CONSTRUCCIÓN DEL LIBRO  (Setup.gs)
 *  Crea/reconstruye todas las hojas, fórmulas, validaciones y formato.
 *
 *  Estrategia de automatización:
 *   - Cada movimiento (compra, venta, traslado, ajuste) tiene columnas "clave"
 *     ocultas con concatenaciones del tipo  Código|Ubicación.
 *   - El Catálogo usa ARRAYFORMULA + SUMIF sobre esas claves para calcular,
 *     de forma 100% automática y por producto:
 *       · Costo promedio ponderado (según TODAS las compras).
 *       · Stock en Casa, Stock en Mercado Libre y Stock Total.
 *       · Valor del inventario y semáforo de stock (Disponible/Bajo/Agotado).
 * ============================================================================
 */

/* Rango de productos del catálogo (filas 2..1000 => 999 productos).
 * Importante: se mantiene en 1000 para que los rangos de validación que
 * apuntan a otras hojas (p. ej. Clientes!B2:B1000) quepan en una hoja recién
 * creada (Google Sheets crea las hojas con 1000 filas por defecto). */
var CAT_LAST = 1000;

/* Formatos numéricos reutilizables. */
var FMT_MONEY = '"$"#,##0';
var FMT_INT = '#,##0';
var FMT_PCT = '0.0%';
var FMT_DATE = 'dd/mm/yyyy';

/**
 * Punto de entrada del menú: construye o reconstruye todo el sistema.
 */
function inicializarSistema() {
  var ss = SpreadsheetApp.getActive();
  var orden = [
    SHEETS.DASHBOARD, SHEETS.CATALOGO, SHEETS.COMPRAS, SHEETS.VENTAS,
    SHEETS.TRASLADOS, SHEETS.AJUSTES, SHEETS.INVENTARIO, SHEETS.CLIENTES,
    SHEETS.FACTURA, SHEETS.CONFIG
  ];

  // 1. Crear las hojas que falten.
  orden.forEach(function (n) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
  });

  // 2. Construir cada módulo. Config primero (de él dependen validaciones).
  buildConfig_(ss);
  buildCatalogo_(ss);
  buildCompras_(ss);
  buildVentas_(ss);
  buildTraslados_(ss);
  buildAjustes_(ss);
  buildClientes_(ss);
  buildInventario_(ss);
  buildDashboard_(ss);
  buildFactura_(ss); // definida en Facturacion.gs

  // 3. Eliminar hojas sobrantes (p. ej. "Hoja 1" / "Sheet1").
  ss.getSheets().forEach(function (sh) {
    if (orden.indexOf(sh.getName()) < 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(sh);
    }
  });

  // 4. Reordenar las pestañas y colorearlas.
  orden.forEach(function (n, i) {
    var sh = ss.getSheetByName(n);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
  });
  ss.getSheetByName(SHEETS.DASHBOARD).setTabColor(COLOR.PRIMARIO);
  ss.getSheetByName(SHEETS.FACTURA).setTabColor(COLOR.ACENTO);
  ss.getSheetByName(SHEETS.CONFIG).setTabColor('#7f8c8d');

  // 5. Crear gráficos del dashboard. (definida en Dashboard.gs)
  actualizarDashboard();

  ss.setActiveSheet(ss.getSheetByName(SHEETS.DASHBOARD));
  SpreadsheetApp.getUi().alert(
    '✅ Sistema construido',
    'Todas las hojas, fórmulas y automatizaciones quedaron listas.\n\n' +
    'Siguiente paso: abre la pestaña "Config" y personaliza los datos de tu ' +
    'empresa (nombre, NIT, teléfono, logo, IVA, etc.).',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ------------------------------------------------------------------ */
/*  Utilidades de construcción                                         */
/* ------------------------------------------------------------------ */

/** Limpia por completo una hoja antes de reconstruirla. */
function clearSheet_(sh) {
  sh.clear();
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart(); } catch (e) {}
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  try { sh.setConditionalFormatRules([]); } catch (e) {}
  try { sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearDataValidations(); } catch (e) {}
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.clearNotes();
  sh.setFrozenRows(0);
  sh.setFrozenColumns(0);
  // Mostrar todas las columnas/filas que pudieran estar ocultas.
  if (sh.getMaxColumns() > 0) sh.showColumns(1, sh.getMaxColumns());
  if (sh.getMaxRows() > 0) sh.showRows(1, sh.getMaxRows());
}

/** Garantiza un mínimo de filas. */
function ensureRows_(sh, n) {
  var m = sh.getMaxRows();
  if (m < n) sh.insertRowsAfter(m, n - m);
}

/** Garantiza un mínimo de columnas. */
function ensureCols_(sh, n) {
  var m = sh.getMaxColumns();
  if (m < n) sh.insertColumnsAfter(m, n - m);
}

/** Da formato a una fila de encabezados. */
function styleHeader_(sh, row, numCols) {
  var rng = sh.getRange(row, 1, 1, numCols);
  rng.setBackground(COLOR.PRIMARIO)
    .setFontColor(COLOR.BLANCO)
    .setFontWeight('bold')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sh.setRowHeight(row, 30);
}

/** Aplica bordes suaves a un rango de datos. */
function softBorders_(rng) {
  rng.setBorder(true, true, true, true, true, true, COLOR.BORDE, SpreadsheetApp.BorderStyle.SOLID);
}

/* ------------------------------------------------------------------ */
/*  CONFIG                                                             */
/* ------------------------------------------------------------------ */
function buildConfig_(ss) {
  var sh = ss.getSheetByName(SHEETS.CONFIG);
  clearSheet_(sh);
  ensureCols_(sh, 12);
  ensureRows_(sh, 60);

  var params = [
    ['Parámetro', 'Valor'],
    ['Nombre de la empresa', 'Mi Negocio S.A.S.'],
    ['NIT / Identificación', '000.000.000-0'],
    ['Dirección', 'Calle 00 # 00-00, Ciudad'],
    ['Teléfono', '+57 300 000 0000'],
    ['Correo electrónico', 'correo@minegocio.com'],
    ['Sitio web', 'www.minegocio.com'],
    ['Logo (URL de imagen)', ''],
    ['Prefijo de factura', 'FAC-'],
    ['Próximo número de factura', 1],
    ['Moneda (símbolo)', '$'],
    ['ID carpeta Drive para PDF (opcional)', ''],
    ['Mensaje pie de factura', '¡Gracias por su compra!']
  ];
  sh.getRange(1, 1, params.length, 2).setValues(params);
  styleHeader_(sh, 1, 2);
  sh.getRange(2, 1, params.length - 1, 1).setFontWeight('bold').setBackground(COLOR.CLARO);
  sh.getRange(2, 2, params.length - 1, 1).setBackground(COLOR.BLANCO);
  softBorders_(sh.getRange(1, 1, params.length, 2));
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 280);

  // Tabla de IVA por categoría: D = Categoría, E = IVA %.
  // El IVA se deja EN BLANCO: escribe 19, 5, 0, etc. según cada categoría.
  var cats = ['General', 'Electrónica', 'Hogar', 'Ropa y calzado', 'Accesorios', 'Papelería', 'Otros'];
  writeList_(sh, 1, 4, 'Categorías', cats);
  sh.getRange(1, 5).setValue('IVA %')
    .setBackground(COLOR.SECUNDARIO).setFontColor(COLOR.BLANCO).setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange(2, 5, cats.length, 1).setNumberFormat('0').setBackground('#fffbe6');
  softBorders_(sh.getRange(1, 5, cats.length + 1, 1));
  sh.getRange('E1').setNote('IVA por categoría. Déjalo en blanco o pon 0 si no aplica; escribe 19 o 5 según el producto.');

  writeList_(sh, 1, 6, 'Ubicaciones', [UBIC.CASA, UBIC.ML]);
  writeList_(sh, 1, 8, 'Estados', ESTADOS);
  writeList_(sh, 1, 10, 'Motivos de ajuste', MOTIVOS_AJUSTE);

  sh.setColumnWidth(4, 150);
  sh.setColumnWidth(5, 80);
  sh.setColumnWidth(6, 150);
  sh.setColumnWidth(8, 110);
  sh.setColumnWidth(10, 180);

  sh.getRange(10, 2).setNumberFormat('0');           // próximo número de factura
  sh.setFrozenRows(1);
  sh.getRange('B8').setNote('Pega aquí la URL pública de tu logo (PNG/JPG) para que aparezca en la factura.');
}

/** Escribe una lista vertical con encabezado en (row,col). */
function writeList_(sh, row, col, titulo, valores) {
  sh.getRange(row, col).setValue(titulo)
    .setBackground(COLOR.SECUNDARIO).setFontColor(COLOR.BLANCO).setFontWeight('bold')
    .setHorizontalAlignment('center');
  var vals = valores.map(function (v) { return [v]; });
  sh.getRange(row + 1, col, vals.length, 1).setValues(vals);
  softBorders_(sh.getRange(row, col, vals.length + 1, 1));
}

/* ------------------------------------------------------------------ */
/*  CATÁLOGO                                                           */
/* ------------------------------------------------------------------ */
function buildCatalogo_(ss) {
  var sh = ss.getSheetByName(SHEETS.CATALOGO);
  clearSheet_(sh);
  ensureCols_(sh, 12);
  ensureRows_(sh, CAT_LAST + 9);

  var headers = ['Código', 'Nombre', 'Categoría', 'Estado', 'Stock Mínimo',
    'Costo Promedio', 'Stock Casa', 'Stock Mercado Libre', 'Stock Total',
    'Valor Inventario', 'Estado Stock', 'IVA %'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  // Costo promedio ponderado = total comprado ($) / total comprado (unid).
  sh.getRange('F2').setFormula(
    `=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",` +
    `IFERROR(SUMIF('Compras'!$B$2:$B,$A$2:$A${CAT_LAST},'Compras'!$F$2:$F)/` +
    `SUMIF('Compras'!$B$2:$B,$A$2:$A${CAT_LAST},'Compras'!$D$2:$D),0)))`
  );

  // Stock Casa = compras + traslados entrantes - traslados salientes - ventas + ajustes.
  sh.getRange('G2').setFormula(
    `=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",` +
    `SUMIF('Compras'!$B$2:$B,$A$2:$A${CAT_LAST},'Compras'!$D$2:$D)` +
    `+SUMIF('Traslados'!$H$2:$H,$A$2:$A${CAT_LAST}&"|${UBIC.CASA}",'Traslados'!$D$2:$D)` +
    `-SUMIF('Traslados'!$G$2:$G,$A$2:$A${CAT_LAST}&"|${UBIC.CASA}",'Traslados'!$D$2:$D)` +
    `-SUMIF('Ventas'!$M$2:$M,$A$2:$A${CAT_LAST}&"|${UBIC.CASA}",'Ventas'!$F$2:$F)` +
    `+SUMIF('Ajustes'!$H$2:$H,$A$2:$A${CAT_LAST}&"|${UBIC.CASA}",'Ajustes'!$D$2:$D)))`
  );

  // Stock Mercado Libre = traslados entrantes - traslados salientes - ventas + ajustes.
  sh.getRange('H2').setFormula(
    `=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",` +
    `SUMIF('Traslados'!$H$2:$H,$A$2:$A${CAT_LAST}&"|${UBIC.ML}",'Traslados'!$D$2:$D)` +
    `-SUMIF('Traslados'!$G$2:$G,$A$2:$A${CAT_LAST}&"|${UBIC.ML}",'Traslados'!$D$2:$D)` +
    `-SUMIF('Ventas'!$M$2:$M,$A$2:$A${CAT_LAST}&"|${UBIC.ML}",'Ventas'!$F$2:$F)` +
    `+SUMIF('Ajustes'!$H$2:$H,$A$2:$A${CAT_LAST}&"|${UBIC.ML}",'Ajustes'!$D$2:$D)))`
  );

  sh.getRange('I2').setFormula(`=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",$G$2:$G${CAT_LAST}+$H$2:$H${CAT_LAST}))`);
  sh.getRange('J2').setFormula(`=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",$I$2:$I${CAT_LAST}*$F$2:$F${CAT_LAST}))`);
  sh.getRange('K2').setFormula(
    `=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",` +
    `IF($I$2:$I${CAT_LAST}<=0,"Agotado",` +
    `IF($I$2:$I${CAT_LAST}<=$E$2:$E${CAT_LAST},"Stock bajo","Disponible"))))`
  );

  // IVA % por producto, según su categoría (tabla Config!D:E).
  sh.getRange('L2').setFormula(
    `=ARRAYFORMULA(IF($A$2:$A${CAT_LAST}="","",` +
    `IFERROR(VLOOKUP($C$2:$C${CAT_LAST},'Config'!$D$2:$E$50,2,FALSE),0)))`
  );

  // Formatos.
  sh.getRange(2, 5, CAT_LAST - 1, 1).setNumberFormat(FMT_INT);          // stock mínimo
  sh.getRange(2, 6, CAT_LAST - 1, 1).setNumberFormat(FMT_MONEY);        // costo promedio
  sh.getRange(2, 7, CAT_LAST - 1, 3).setNumberFormat(FMT_INT);          // stocks
  sh.getRange(2, 10, CAT_LAST - 1, 1).setNumberFormat(FMT_MONEY);       // valor inventario
  sh.getRange(2, 12, CAT_LAST - 1, 1).setNumberFormat(FMT_INT);         // IVA %

  // Validaciones.
  setDropdownFromRange_(sh.getRange(2, 3, CAT_LAST - 1, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('D2:D50'));   // categoría
  setDropdownFromRange_(sh.getRange(2, 4, CAT_LAST - 1, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('H2:H3'));    // estado

  // Semáforo de stock.
  var rngK = sh.getRange(2, 11, CAT_LAST - 1, 1);
  var rules = [];
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Agotado')
    .setBackground('#f8d7da').setFontColor(COLOR.ALERTA).setBold(true).setRanges([rngK]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Stock bajo')
    .setBackground('#fff3cd').setFontColor('#8a6d3b').setBold(true).setRanges([rngK]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Disponible')
    .setBackground('#d4edda').setFontColor('#1e6b3a').setRanges([rngK]).build());
  sh.setConditionalFormatRules(rules);

  // Anchos y vista.
  var widths = [110, 240, 140, 90, 110, 120, 95, 150, 100, 130, 120, 70];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);
  softBorders_(sh.getRange(1, 1, CAT_LAST, headers.length));
  sh.getRange('A1').setNote('Registra cada producto UNA sola vez. Las columnas en azul claro (F a L) se calculan solas. El IVA % viene de la categoría (hoja Config).');
  sh.getRange(2, 6, CAT_LAST - 1, 7).setBackground('#f4f8fc');
}

/* ------------------------------------------------------------------ */
/*  COMPRAS                                                            */
/* ------------------------------------------------------------------ */
function buildCompras_(ss) {
  var sh = ss.getSheetByName(SHEETS.COMPRAS);
  clearSheet_(sh);
  ensureCols_(sh, 6);
  ensureRows_(sh, MAX_ROWS + 1);

  var headers = ['Fecha', 'Código', 'Producto', 'Cantidad', 'Costo Unitario', 'Costo Total'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('C2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",IFERROR(VLOOKUP($B$2:$B,'Catálogo'!$A:$B,2,FALSE),"⚠ Código no existe")))`);
  sh.getRange('F2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",$D$2:$D*$E$2:$E))`);

  sh.getRange(2, 1, MAX_ROWS, 1).setNumberFormat(FMT_DATE);
  sh.getRange(2, 4, MAX_ROWS, 1).setNumberFormat(FMT_INT);
  sh.getRange(2, 5, MAX_ROWS, 2).setNumberFormat(FMT_MONEY);

  setDropdownFromRange_(sh.getRange(2, 2, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CATALOGO).getRange('A2:A' + CAT_LAST));

  var widths = [110, 120, 240, 100, 130, 130];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  shadeAutoCols_(sh, [3, 6]);
  softBorders_(sh.getRange(1, 1, MAX_ROWS, headers.length));
  sh.getRange('A1').setNote('Cada compra suma stock a "Casa" y recalcula el costo promedio del producto automáticamente.');
}

/* ------------------------------------------------------------------ */
/*  VENTAS                                                             */
/* ------------------------------------------------------------------ */
function buildVentas_(ss) {
  var sh = ss.getSheetByName(SHEETS.VENTAS);
  clearSheet_(sh);
  ensureCols_(sh, 13);
  ensureRows_(sh, MAX_ROWS + 1);

  var headers = ['Fecha', 'N° Factura', 'Cliente', 'Código', 'Producto', 'Cantidad',
    'Ubicación', 'Valor Recibido', 'Costo Prom. Unit.', 'Costo Total', 'Utilidad', 'Margen %', 'clave'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('E2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",IFERROR(VLOOKUP($D$2:$D,'Catálogo'!$A:$B,2,FALSE),"⚠ Código no existe")))`);
  sh.getRange('I2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",IFERROR(VLOOKUP($D$2:$D,'Catálogo'!$A:$F,6,FALSE),0)))`);
  sh.getRange('J2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",$F$2:$F*$I$2:$I))`);
  sh.getRange('K2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",$H$2:$H-$J$2:$J))`);
  sh.getRange('L2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",IF($H$2:$H=0,0,($H$2:$H-$J$2:$J)/$H$2:$H)))`);
  sh.getRange('M2').setFormula(`=ARRAYFORMULA(IF($D$2:$D="","",$D$2:$D&"|"&$G$2:$G))`);

  sh.getRange(2, 1, MAX_ROWS, 1).setNumberFormat(FMT_DATE);
  sh.getRange(2, 6, MAX_ROWS, 1).setNumberFormat(FMT_INT);
  sh.getRange(2, 8, MAX_ROWS, 4).setNumberFormat(FMT_MONEY);
  sh.getRange(2, 12, MAX_ROWS, 1).setNumberFormat(FMT_PCT);

  setDropdownFromRange_(sh.getRange(2, 3, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CLIENTES).getRange('B2:B' + CAT_LAST));
  setDropdownFromRange_(sh.getRange(2, 4, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CATALOGO).getRange('A2:A' + CAT_LAST));
  setDropdownFromRange_(sh.getRange(2, 7, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('F2:F3'));

  var widths = [105, 110, 170, 110, 220, 90, 130, 130, 130, 120, 120, 90, 10];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(13); // clave
  sh.setFrozenRows(1);
  shadeAutoCols_(sh, [5, 9, 10, 11, 12]);
  softBorders_(sh.getRange(1, 1, MAX_ROWS, 12));
  sh.getRange('A1').setNote('Cada venta descuenta el inventario de la ubicación elegida y calcula la utilidad con el costo promedio vigente.');
}

/* ------------------------------------------------------------------ */
/*  TRASLADOS                                                          */
/* ------------------------------------------------------------------ */
function buildTraslados_(ss) {
  var sh = ss.getSheetByName(SHEETS.TRASLADOS);
  clearSheet_(sh);
  ensureCols_(sh, 8);
  ensureRows_(sh, MAX_ROWS + 1);

  var headers = ['Fecha', 'Código', 'Producto', 'Cantidad', 'Origen', 'Destino', 'claveOrigen', 'claveDestino'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('C2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",IFERROR(VLOOKUP($B$2:$B,'Catálogo'!$A:$B,2,FALSE),"⚠ Código no existe")))`);
  sh.getRange('G2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",$B$2:$B&"|"&$E$2:$E))`);
  sh.getRange('H2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",$B$2:$B&"|"&$F$2:$F))`);

  sh.getRange(2, 1, MAX_ROWS, 1).setNumberFormat(FMT_DATE);
  sh.getRange(2, 4, MAX_ROWS, 1).setNumberFormat(FMT_INT);

  setDropdownFromRange_(sh.getRange(2, 2, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CATALOGO).getRange('A2:A' + CAT_LAST));
  setDropdownFromRange_(sh.getRange(2, 5, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('F2:F3'));
  setDropdownFromRange_(sh.getRange(2, 6, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('F2:F3'));

  var widths = [110, 120, 240, 100, 140, 140, 10, 10];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(7, 2);
  sh.setFrozenRows(1);
  shadeAutoCols_(sh, [3]);
  softBorders_(sh.getRange(1, 1, MAX_ROWS, 6));
  sh.getRange('A1').setNote('Mueve unidades entre Casa y Mercado Libre. Actualiza ambas existencias automáticamente.');
}

/* ------------------------------------------------------------------ */
/*  AJUSTES                                                            */
/* ------------------------------------------------------------------ */
function buildAjustes_(ss) {
  var sh = ss.getSheetByName(SHEETS.AJUSTES);
  clearSheet_(sh);
  ensureCols_(sh, 8);
  ensureRows_(sh, MAX_ROWS + 1);

  var headers = ['Fecha', 'Código', 'Producto', 'Cantidad (+/-)', 'Ubicación', 'Motivo', 'Nota', 'clave'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('C2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",IFERROR(VLOOKUP($B$2:$B,'Catálogo'!$A:$B,2,FALSE),"⚠ Código no existe")))`);
  sh.getRange('H2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","",$B$2:$B&"|"&$E$2:$E))`);

  sh.getRange(2, 1, MAX_ROWS, 1).setNumberFormat(FMT_DATE);
  sh.getRange(2, 4, MAX_ROWS, 1).setNumberFormat('+#,##0;-#,##0;0');

  setDropdownFromRange_(sh.getRange(2, 2, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CATALOGO).getRange('A2:A' + CAT_LAST));
  setDropdownFromRange_(sh.getRange(2, 5, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('F2:F3'));
  setDropdownFromRange_(sh.getRange(2, 6, MAX_ROWS, 1), ss.getSheetByName(SHEETS.CONFIG).getRange('J2:J20'));

  var widths = [110, 120, 240, 130, 140, 180, 260, 10];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.hideColumns(8);
  sh.setFrozenRows(1);
  shadeAutoCols_(sh, [3]);
  softBorders_(sh.getRange(1, 1, MAX_ROWS, 7));
  sh.getRange('D1').setNote('Usa valores POSITIVOS para sumar y NEGATIVOS para restar (p. ej. -2 por producto dañado).');
}

/* ------------------------------------------------------------------ */
/*  CLIENTES                                                           */
/* ------------------------------------------------------------------ */
function buildClientes_(ss) {
  var sh = ss.getSheetByName(SHEETS.CLIENTES);
  clearSheet_(sh);
  ensureCols_(sh, 8);
  ensureRows_(sh, MAX_ROWS + 1);

  var headers = ['ID', 'Nombre', 'Documento / NIT', 'Teléfono', 'Correo', 'Dirección', 'Ciudad', 'Notas'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('A2').setFormula(`=ARRAYFORMULA(IF($B$2:$B="","","CLI-"&TEXT(ROW($B$2:$B)-1,"000")))`);

  var widths = [90, 220, 150, 140, 220, 240, 130, 240];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  shadeAutoCols_(sh, [1]);
  softBorders_(sh.getRange(1, 1, MAX_ROWS, headers.length));
  sh.getRange('D1').setNote('Para WhatsApp usa el formato internacional sin signos, p. ej. 573001112233.');
}

/* ------------------------------------------------------------------ */
/*  INVENTARIO (reporte consolidado, solo lectura)                     */
/* ------------------------------------------------------------------ */
function buildInventario_(ss) {
  var sh = ss.getSheetByName(SHEETS.INVENTARIO);
  clearSheet_(sh);
  ensureCols_(sh, 9);
  ensureRows_(sh, CAT_LAST + 2);

  var headers = ['Código', 'Producto', 'Categoría', 'Costo Promedio', 'Stock Casa',
    'Stock Mercado Libre', 'Stock Total', 'Valor Inventario', 'Estado'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  styleHeader_(sh, 1, headers.length);

  sh.getRange('A2').setFormula(
    `=IFERROR(QUERY('Catálogo'!$A$2:$K$${CAT_LAST},` +
    `"select Col1,Col2,Col3,Col6,Col7,Col8,Col9,Col10,Col11 where Col1 is not null order by Col2",0),"")`
  );

  sh.getRange(2, 4, CAT_LAST, 1).setNumberFormat(FMT_MONEY);
  sh.getRange(2, 5, CAT_LAST, 3).setNumberFormat(FMT_INT);
  sh.getRange(2, 8, CAT_LAST, 1).setNumberFormat(FMT_MONEY);

  var widths = [110, 240, 140, 130, 95, 150, 100, 140, 120];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setFrozenRows(1);
  softBorders_(sh.getRange(1, 1, CAT_LAST, headers.length));
  sh.getRange('A1').setNote('Reporte automático del Catálogo. Es de solo lectura: no escribas en estas celdas.');
}

/* ------------------------------------------------------------------ */
/*  DASHBOARD (KPIs + tablas; los gráficos los crea actualizarDashboard)*/
/* ------------------------------------------------------------------ */
function buildDashboard_(ss) {
  var sh = ss.getSheetByName(SHEETS.DASHBOARD);
  clearSheet_(sh);
  ensureCols_(sh, 20);
  ensureRows_(sh, 100);

  // Anchos base.
  for (var c = 1; c <= 12; c++) sh.setColumnWidth(c, 110);

  // Título.
  sh.getRange('A1:L1').merge().setValue('DASHBOARD EJECUTIVO')
    .setBackground(COLOR.PRIMARIO).setFontColor(COLOR.BLANCO).setFontSize(20)
    .setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 46);
  sh.getRange('A2:L2').merge()
    .setFormula('="Mes en curso: "&TEXT(TODAY(),"mmmm \\d\\e yyyy")&"   ·   Actualizado: "&TEXT(NOW(),"dd/mm/yyyy hh:mm")')
    .setBackground(COLOR.CLARO).setFontColor(COLOR.TEXTO).setHorizontalAlignment('center').setFontWeight('bold');
  sh.setRowHeight(2, 24);

  var df = `SUMIFS('Ventas'!$H:$H,'Ventas'!$A:$A,">="&EOMONTH(TODAY(),-1)+1,'Ventas'!$A:$A,"<="&EOMONTH(TODAY(),0))`;
  var gf = `SUMIFS('Ventas'!$K:$K,'Ventas'!$A:$A,">="&EOMONTH(TODAY(),-1)+1,'Ventas'!$A:$A,"<="&EOMONTH(TODAY(),0))`;
  var cf = `SUMIFS('Compras'!$F:$F,'Compras'!$A:$A,">="&EOMONTH(TODAY(),-1)+1,'Compras'!$A:$A,"<="&EOMONTH(TODAY(),0))`;

  // Fila 1 de tarjetas (4 KPIs).
  card_(sh, 4, 1, 'VENTAS DEL MES', '=' + df, FMT_MONEY, COLOR.SECUNDARIO);
  card_(sh, 4, 4, 'GANANCIA DEL MES', '=' + gf, FMT_MONEY, COLOR.ACENTO);
  card_(sh, 4, 7, 'COMPRAS DEL MES', '=' + cf, FMT_MONEY, COLOR.ADVERT);
  card_(sh, 4, 10, 'VALOR DEL INVENTARIO', `=SUM('Catálogo'!$J$2:$J$${CAT_LAST})`, FMT_MONEY, COLOR.PRIMARIO);

  // Fila 2 de tarjetas (4 KPIs).
  card_(sh, 7, 1, 'INVENTARIO EN CASA (UNID)', `=SUM('Catálogo'!$G$2:$G$${CAT_LAST})`, FMT_INT, '#34495e');
  card_(sh, 7, 4, 'INVENTARIO EN ML (UNID)', `=SUM('Catálogo'!$H$2:$H$${CAT_LAST})`, FMT_INT, '#34495e');
  card_(sh, 7, 7, 'PRODUCTOS CON STOCK BAJO', `=COUNTIF('Catálogo'!$K$2:$K$${CAT_LAST},"Stock bajo")`, FMT_INT, COLOR.ADVERT);
  card_(sh, 7, 10, 'PRODUCTOS AGOTADOS', `=COUNTIF('Catálogo'!$K$2:$K$${CAT_LAST},"Agotado")`, FMT_INT, COLOR.ALERTA);

  // Tabla: productos más vendidos.
  sh.getRange('A11:E11').merge().setValue('🏆 PRODUCTOS MÁS VENDIDOS')
    .setBackground(COLOR.PRIMARIO).setFontColor(COLOR.BLANCO).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('A12').setFormula(
    `=IFERROR(QUERY('Ventas'!$E$2:$F,` +
    `"select Col1, sum(Col2) where Col1 is not null group by Col1 order by sum(Col2) desc limit 8 ` +
    `label Col1 'Producto', sum(Col2) 'Unidades vendidas'",0),"Aún no hay ventas registradas")`
  );

  // Tabla: alertas de inventario.
  sh.getRange('G11:L11').merge().setValue('⚠️ ALERTAS DE INVENTARIO (bajo / agotado)')
    .setBackground(COLOR.ALERTA).setFontColor(COLOR.BLANCO).setFontWeight('bold').setHorizontalAlignment('left');
  sh.getRange('G12').setFormula(
    `=IFERROR(QUERY('Catálogo'!$A$2:$K$${CAT_LAST},` +
    `"select Col2, Col7, Col8, Col9, Col11 where Col1 is not null and (Col11='Agotado' or Col11='Stock bajo') ` +
    `order by Col9 asc limit 30 label Col2 'Producto', Col7 'Casa', Col8 'ML', Col9 'Total', Col11 'Estado'",0),"Sin alertas 🎉")`
  );

  // Tabla auxiliar (oculta) para gráficos: últimos 6 meses.
  var aux = [['MesIni', 'MesFin', 'Mes', 'Ventas', 'Utilidad', 'Compras']];
  for (var i = 0; i < 6; i++) {
    var off = -5 + i;
    var ini = `DATE(YEAR(EOMONTH(TODAY(),${off})),MONTH(EOMONTH(TODAY(),${off})),1)`;
    var fin = `EOMONTH(TODAY(),${off})`;
    aux.push([
      `=${ini}`,
      `=${fin}`,
      `=TEXT(N${i + 2},"mmm yy")`,
      `=SUMIFS('Ventas'!$H:$H,'Ventas'!$A:$A,">="&N${i + 2},'Ventas'!$A:$A,"<="&O${i + 2})`,
      `=SUMIFS('Ventas'!$K:$K,'Ventas'!$A:$A,">="&N${i + 2},'Ventas'!$A:$A,"<="&O${i + 2})`,
      `=SUMIFS('Compras'!$F:$F,'Compras'!$A:$A,">="&N${i + 2},'Compras'!$A:$A,"<="&O${i + 2})`
    ]);
  }
  sh.getRange(1, 14, aux.length, 6).setValues(aux);   // N1:S7
  sh.getRange(2, 17, 6, 3).setNumberFormat(FMT_MONEY);
  sh.hideColumns(14, 6); // ocultar columnas auxiliares N..S

  sh.setFrozenRows(2);
  sh.setHiddenGridlines(true);
}

/* ------------------------------------------------------------------ */
/*  Helpers de Dashboard / validaciones                                */
/* ------------------------------------------------------------------ */

/** Dibuja una tarjeta KPI de 3 columnas de ancho y 2 filas de alto. */
function card_(sh, row, col, etiqueta, formula, formato, color) {
  var lab = sh.getRange(row, col, 1, 3).merge();
  lab.setValue(etiqueta).setBackground(color).setFontColor(COLOR.BLANCO)
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  var val = sh.getRange(row + 1, col, 1, 3).merge();
  val.setFormula(formula).setNumberFormat(formato)
    .setBackground(COLOR.BLANCO).setFontColor(color).setFontSize(20).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(row, 24);
  sh.setRowHeight(row + 1, 40);
  softBorders_(sh.getRange(row, col, 2, 3));
}

/** Aplica una validación de tipo lista que referencia un rango. */
function setDropdownFromRange_(targetRange, sourceRange) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true)
    .setAllowInvalid(true)
    .build();
  targetRange.setDataValidation(rule);
}

/** Sombrea columnas calculadas para indicar que son automáticas. */
function shadeAutoCols_(sh, cols) {
  cols.forEach(function (c) {
    sh.getRange(2, c, MAX_ROWS, 1).setBackground('#f4f8fc');
  });
}
/**
 * ============================================================================
 *  DASHBOARD - Gráficos  (Dashboard.gs)
 *  Crea/recrea los gráficos de ventas, utilidades y productos más vendidos.
 * ============================================================================
 */

/**
 * Reconstruye los gráficos del Dashboard usando las tablas auxiliares.
 * Se invoca desde el menú y al inicializar el sistema.
 */
function actualizarDashboard() {
  var sh = getSheet_(SHEETS.DASHBOARD);

  // Eliminar gráficos previos para evitar duplicados.
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });

  // Banner de sección de gráficos (debajo de las tablas, fila 44).
  sh.getRange('A44:L44').merge().setValue('📈 TENDENCIAS (últimos 6 meses)')
    .setBackground(COLOR.SECUNDARIO).setFontColor(COLOR.BLANCO)
    .setFontWeight('bold').setHorizontalAlignment('left');

  // 1) Columnas: Ventas, Utilidad y Compras por mes (rango N..S oculto: Mes/Ventas/Util/Compras).
  var rangoMes = sh.getRange('P1:S7'); // Mes, Ventas, Utilidad, Compras
  var chartBarras = sh.newChart()
    .asColumnChart()
    .addRange(rangoMes)
    .setNumHeaders(1)
    .setOption('title', 'Ventas, Utilidad y Compras por mes')
    .setOption('legend', { position: 'top' })
    .setOption('colors', [COLOR.SECUNDARIO, COLOR.ACENTO, COLOR.ADVERT])
    .setOption('height', 300)
    .setOption('width', 620)
    .setPosition(45, 1, 0, 0)
    .build();
  sh.insertChart(chartBarras);

  // 2) Línea: evolución de Ventas vs Utilidad.
  var chartLinea = sh.newChart()
    .asLineChart()
    .addRange(sh.getRange('P1:R7')) // Mes, Ventas, Utilidad
    .setNumHeaders(1)
    .setOption('title', 'Evolución de Ventas y Utilidad')
    .setOption('legend', { position: 'top' })
    .setOption('colors', [COLOR.SECUNDARIO, COLOR.ACENTO])
    .setOption('curveType', 'function')
    .setOption('height', 300)
    .setOption('width', 540)
    .setPosition(45, 8, 0, 0)
    .build();
  sh.insertChart(chartLinea);

  // 3) Barras horizontales: productos más vendidos (tabla A12:B...).
  var ultima = ultimaFilaTabla_(sh, 'A', 12);
  if (ultima >= 13) {
    var chartTop = sh.newChart()
      .asBarChart()
      .addRange(sh.getRange('A12:B' + ultima))
      .setNumHeaders(1)
      .setOption('title', 'Productos más vendidos (unidades)')
      .setOption('legend', { position: 'none' })
      .setOption('colors', [COLOR.PRIMARIO])
      .setOption('height', 300)
      .setOption('width', 620)
      .setPosition(62, 1, 0, 0)
      .build();
    sh.insertChart(chartTop);
  }

  SpreadsheetApp.getActive().toast('Dashboard y gráficos actualizados.', '📊 Gestión', 4);
}

/** Devuelve el número de la última fila con datos en una columna a partir de fromRow. */
function ultimaFilaTabla_(sh, colLetter, fromRow) {
  var col = colLetter.charCodeAt(0) - 64;
  var last = sh.getLastRow();
  var fila = fromRow - 1;
  for (var r = fromRow; r <= last; r++) {
    var v = sh.getRange(r, col).getValue();
    if (v === '' || v === null) break;
    fila = r;
  }
  return fila;
}
/**
 * ============================================================================
 *  FACTURACIÓN  (Facturacion.gs)
 *  - Construye la hoja "Factura" con apariencia de documento comercial.
 *  - Numeración automática, autocompletado de cliente y totales por fórmula.
 *  - Registra la venta en el módulo Ventas (descuenta inventario).
 *  - Exporta a PDF y permite compartir por correo o WhatsApp.
 * ============================================================================
 */

/* Referencias de celdas de la hoja Factura. */
var FAC = {
  NUMERO: 'E2',
  FECHA: 'E3',
  PAGO: 'E4',
  ORIGEN: 'E5',
  CLIENTE: 'B8',
  ITEMS_FIRST: 13,
  ITEMS_LAST: 24,
  SUBTOTAL: 'F26',
  IVA: 'F27',
  TOTAL: 'F28'
};

/* ------------------------------------------------------------------ */
/*  Construcción de la plantilla de factura                            */
/* ------------------------------------------------------------------ */
function buildFactura_(ss) {
  var sh = ss.getSheetByName(SHEETS.FACTURA);
  clearSheet_(sh);
  ensureCols_(sh, 6);
  ensureRows_(sh, 40);
  sh.setHiddenGridlines(true);

  // Anchos de columna (documento ~ carta).
  var widths = [120, 230, 90, 95, 120, 130];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });

  // Formato base del documento.
  sh.getRange(1, 1, 40, 6)
    .setFontFamily('Arial').setFontColor(COLOR.TEXTO).setVerticalAlignment('middle');

  // --- Encabezado: logo + datos de la empresa ---
  sh.getRange('A1:A5').merge()
    .setFormula(`=IF('Config'!$B$8="","",IMAGE('Config'!$B$8,1))`)
    .setHorizontalAlignment('center');
  sh.getRange('B1:C1').merge().setFormula(`='Config'!$B$2`)
    .setFontSize(16).setFontWeight('bold').setFontColor(COLOR.PRIMARIO);
  sh.getRange('B2:C2').merge().setFormula(`="NIT: "&'Config'!$B$3`);
  sh.getRange('B3:C3').merge().setFormula(`='Config'!$B$4`);
  sh.getRange('B4:C4').merge().setFormula(`="Tel: "&'Config'!$B$5&"   |   "&'Config'!$B$6`);
  sh.getRange('B5:C5').merge().setFormula(`='Config'!$B$7`).setFontColor(COLOR.SECUNDARIO);

  // --- Caja de metadatos de la factura ---
  sh.getRange('D1:F1').merge().setValue('FACTURA DE VENTA')
    .setBackground(COLOR.PRIMARIO).setFontColor(COLOR.BLANCO)
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  sh.getRange('D2').setValue('N°:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('E2:F2').merge().setFontWeight('bold').setFontColor(COLOR.ALERTA).setHorizontalAlignment('left');
  sh.getRange('D3').setValue('Fecha:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('E3:F3').merge().setNumberFormat(FMT_DATE).setHorizontalAlignment('left');
  sh.getRange('D4').setValue('Pago:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('E4:F4').merge().setHorizontalAlignment('left');
  sh.getRange('D5').setValue('Origen inv.:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('E5:F5').merge().setHorizontalAlignment('left').setValue(UBIC.CASA);

  var pagoRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Efectivo', 'Transferencia', 'Tarjeta', 'Nequi / Daviplata', 'Crédito'], true)
    .setAllowInvalid(true).build();
  sh.getRange('E4').setDataValidation(pagoRule);
  setDropdownFromRange_(sh.getRange('E5'), ss.getSheetByName(SHEETS.CONFIG).getRange('F2:F3'));

  // --- Datos del cliente ---
  sh.getRange('A7:F7').merge().setValue('DATOS DEL CLIENTE')
    .setBackground(COLOR.SECUNDARIO).setFontColor(COLOR.BLANCO).setFontWeight('bold');
  sh.getRange('A8').setValue('Cliente:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B8:C8').merge();
  sh.getRange('D8').setValue('Documento:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('E8:F8').merge().setFormula(`=IFERROR(VLOOKUP($B$8,'Clientes'!$B:$F,2,FALSE),"")`);
  sh.getRange('A9').setValue('Teléfono:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B9:C9').merge().setFormula(`=IFERROR(VLOOKUP($B$8,'Clientes'!$B:$F,3,FALSE),"")`);
  sh.getRange('D9').setValue('Correo:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('E9:F9').merge().setFormula(`=IFERROR(VLOOKUP($B$8,'Clientes'!$B:$F,4,FALSE),"")`);
  sh.getRange('A10').setValue('Dirección:').setFontWeight('bold').setHorizontalAlignment('right');
  sh.getRange('B10:F10').merge().setFormula(`=IFERROR(VLOOKUP($B$8,'Clientes'!$B:$F,5,FALSE),"")`);
  setDropdownFromRange_(sh.getRange('B8'), ss.getSheetByName(SHEETS.CLIENTES).getRange('B2:B' + CAT_LAST));

  // --- Tabla de productos ---
  sh.getRange('A12').setValue('Código');
  sh.getRange('B12:C12').merge().setValue('Descripción');
  sh.getRange('D12').setValue('Cant.');
  sh.getRange('E12').setValue('Vr. Unitario');
  sh.getRange('F12').setValue('Total');
  sh.getRange('A12:F12').setBackground(COLOR.PRIMARIO).setFontColor(COLOR.BLANCO)
    .setFontWeight('bold').setHorizontalAlignment('center');

  for (var r = FAC.ITEMS_FIRST; r <= FAC.ITEMS_LAST; r++) {
    sh.getRange('B' + r + ':C' + r).merge()
      .setFormula(`=IF($A${r}="","",IFERROR(VLOOKUP($A${r},'Catálogo'!$A:$B,2,FALSE),"⚠ Código no existe"))`);
    sh.getRange('F' + r)
      .setFormula(`=IF(OR($A${r}="",$D${r}="",$E${r}=""),"",$D${r}*$E${r})`)
      .setNumberFormat(FMT_MONEY);
    sh.getRange('D' + r).setNumberFormat(FMT_INT).setHorizontalAlignment('center');
    sh.getRange('E' + r).setNumberFormat(FMT_MONEY);
    // Columnas auxiliares (ocultas): H = IVA% del producto (según categoría), I = IVA $ del renglón.
    sh.getRange('H' + r).setFormula(`=IF($A${r}="","",IFERROR(VLOOKUP($A${r},'Catálogo'!$A:$L,12,FALSE),0))`);
    sh.getRange('I' + r).setFormula(`=IF($F${r}="","",ROUND($F${r}*$H${r}/100,0))`).setNumberFormat(FMT_MONEY);
  }
  setDropdownFromRange_(sh.getRange('A13:A24'), ss.getSheetByName(SHEETS.CATALOGO).getRange('A2:A' + CAT_LAST));
  softBorders_(sh.getRange('A12:F24'));
  sh.hideColumns(8, 2); // ocultar columnas auxiliares H e I

  // --- Observaciones ---
  sh.getRange('A25').setValue('Observaciones:').setFontWeight('bold');
  sh.getRange('A26:C29').merge().setVerticalAlignment('top').setWrap(true);
  softBorders_(sh.getRange('A26:C29'));

  // --- Totales ---
  sh.getRange('D26:E26').merge().setValue('Subtotal:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('F26').setFormula('=SUM(F13:F24)').setNumberFormat(FMT_MONEY);
  sh.getRange('D27:E27').merge().setValue('IVA:').setHorizontalAlignment('right').setFontWeight('bold');
  sh.getRange('F27').setFormula(`=SUM(I13:I24)`).setNumberFormat(FMT_MONEY);
  sh.getRange('D28:E28').merge().setValue('TOTAL A PAGAR:')
    .setHorizontalAlignment('right').setFontWeight('bold').setFontSize(13)
    .setBackground(COLOR.ACENTO).setFontColor(COLOR.BLANCO);
  sh.getRange('F28').setFormula('=$F$26+$F$27').setNumberFormat(FMT_MONEY)
    .setFontSize(13).setFontWeight('bold').setBackground(COLOR.ACENTO).setFontColor(COLOR.BLANCO);
  softBorders_(sh.getRange('D26:F28'));

  // --- Pie de página ---
  sh.getRange('A31:F31').merge().setFormula(`='Config'!$B$13`)
    .setHorizontalAlignment('center').setFontWeight('bold').setFontColor(COLOR.PRIMARIO).setFontSize(12);
  sh.getRange('A32:F32').merge()
    .setValue('Documento generado electrónicamente · Sistema de Gestión en Google Sheets')
    .setHorizontalAlignment('center').setFontColor('#95a5a6').setFontSize(8);

  // Altura de filas.
  sh.setRowHeight(6, 8);
  sh.setRowHeight(7, 22);
  sh.setRowHeight(11, 8);
  sh.setRowHeight(12, 24);
  sh.setRowHeight(25, 18);
  sh.setRowHeight(28, 28);
  sh.setRowHeight(30, 10);

  // Resaltar celdas de entrada (amarillo suave) para guiar al usuario.
  var inputColor = '#fffbe6';
  sh.getRange('B8').setBackground(inputColor);
  sh.getRange('E4:F4').setBackground(inputColor);
  sh.getRange('E5:F5').setBackground(inputColor);
  sh.getRange('A13:A24').setBackground(inputColor);
  sh.getRange('D13:E24').setBackground(inputColor);
  sh.getRange('A26:C29').setBackground(inputColor);

  sh.getRange('A12').setNote('Selecciona el código del producto: la descripción y el Total se completan solos. Escribe Cant. y Vr. Unitario.');
  sh.getRange(FAC.CLIENTE).setNote('Elige un cliente de la lista (se carga desde la hoja Clientes).');
}

/* ------------------------------------------------------------------ */
/*  Operaciones de factura                                             */
/* ------------------------------------------------------------------ */

/** Prepara una factura nueva: asigna número, fecha y limpia los campos. */
function nuevaFactura() {
  var ss = SpreadsheetApp.getActive();
  var sh = getSheet_(SHEETS.FACTURA);

  var prefijo = getParametro_('Prefijo de factura') || 'FAC-';
  var num = Number(getParametro_('Próximo número de factura')) || 1;
  sh.getRange(FAC.NUMERO).setValue(prefijo + padNum_(num));
  sh.getRange(FAC.FECHA).setValue(new Date());

  // Limpiar entradas de la factura anterior.
  sh.getRange('A13:A24').clearContent();
  sh.getRange('D13:E24').clearContent();
  sh.getRange(FAC.CLIENTE).clearContent();
  sh.getRange('A26').clearContent();
  sh.getRange(FAC.PAGO).setValue('Efectivo');
  sh.getRange(FAC.ORIGEN).setValue(UBIC.CASA);

  ss.setActiveSheet(sh);
  ss.toast('Factura ' + prefijo + padNum_(num) + ' lista para diligenciar.', '🧾 Facturación', 5);
}

/** Rellena con ceros a la izquierda (mínimo 4 dígitos). */
function padNum_(n) {
  var s = String(n);
  while (s.length < 4) s = '0' + s;
  return s;
}

/**
 * Registra la factura actual en el módulo Ventas (descuenta inventario)
 * y avanza el consecutivo de facturación.
 */
function guardarFactura() {
  var ss = SpreadsheetApp.getActive();
  var sh = getSheet_(SHEETS.FACTURA);
  var ui = SpreadsheetApp.getUi();
  SpreadsheetApp.flush();

  var numero = sh.getRange(FAC.NUMERO).getValue();
  if (!numero) {
    alerta_('Falta el número', 'Primero usa "Nueva factura" para generar el número.');
    return;
  }
  var fecha = sh.getRange(FAC.FECHA).getValue();
  if (!(fecha instanceof Date)) fecha = new Date();
  var cliente = sh.getRange(FAC.CLIENTE).getValue();
  var origen = sh.getRange(FAC.ORIGEN).getValue() || UBIC.CASA;

  // Leer líneas de la factura.
  var datos = sh.getRange('A13:F24').getValues();
  var lineas = [];
  for (var i = 0; i < datos.length; i++) {
    var cod = datos[i][0];
    var cant = datos[i][3];
    var vr = datos[i][4];
    var tot = datos[i][5];
    if (cod !== '' && cant !== '' && Number(cant) > 0) {
      var total = (tot === '' || tot === null) ? Number(cant) * Number(vr || 0) : Number(tot);
      lineas.push({ cod: cod, cant: Number(cant), total: total });
    }
  }
  if (lineas.length === 0) {
    alerta_('Sin productos', 'Agrega al menos un producto con su cantidad antes de guardar.');
    return;
  }

  var totalFac = sh.getRange(FAC.TOTAL).getValue();
  var resp = ui.alert('Confirmar venta',
    'Se registrarán ' + lineas.length + ' producto(s) de la factura ' + numero + '.\n' +
    'El inventario se descontará de: ' + origen + '.\n\n¿Deseas continuar?',
    ui.ButtonSet.YES_NO);
  if (resp !== ui.Button.YES) return;

  // Insertar en Ventas buscando filas vacías en la columna D (Código).
  var vsh = getSheet_(SHEETS.VENTAS);
  var maxR = vsh.getMaxRows();
  var colD = vsh.getRange(2, 4, maxR - 1, 1).getValues();
  var ptr = 0;
  lineas.forEach(function (l) {
    while (ptr < colD.length && colD[ptr][0] !== '') ptr++;
    var row = ptr + 2;
    ptr++;
    if (row > vsh.getMaxRows()) vsh.insertRowsAfter(vsh.getMaxRows(), 1);
    vsh.getRange(row, 1, 1, 4).setValues([[fecha, numero, cliente, l.cod]]); // A:D
    vsh.getRange(row, 6).setValue(l.cant);    // F Cantidad
    vsh.getRange(row, 7).setValue(origen);    // G Ubicación
    vsh.getRange(row, 8).setValue(l.total);   // H Valor Recibido
  });

  // Avanzar consecutivo.
  var cn = Number(getParametro_('Próximo número de factura')) || 1;
  setParametro_('Próximo número de factura', cn + 1);
  SpreadsheetApp.flush();

  alerta_('✅ Venta registrada',
    'La factura ' + numero + ' se registró en el módulo Ventas y el inventario fue actualizado.\n' +
    'Total facturado: ' + totalFac + '\n\n' +
    'Consejo: usa "Generar / Descargar PDF" para guardar o compartir la factura.');
}

/* ------------------------------------------------------------------ */
/*  PDF y compartir                                                    */
/* ------------------------------------------------------------------ */

/** Devuelve (o crea) la carpeta de Drive donde se guardan las facturas. */
function carpetaFacturas_() {
  var id = String(getParametro_('ID carpeta Drive para PDF (opcional)') || '').trim();
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* id inválido: usar carpeta por defecto */ }
  }
  var it = DriveApp.getFoldersByName('Facturas');
  return it.hasNext() ? it.next() : DriveApp.createFolder('Facturas');
}

/** Genera el PDF de la hoja Factura y devuelve el archivo de Drive. */
function crearPDFFactura_() {
  var ss = SpreadsheetApp.getActive();
  var sh = getSheet_(SHEETS.FACTURA);
  var numero = sh.getRange(FAC.NUMERO).getValue();
  if (!numero) throw new Error('No hay factura activa. Usa "Nueva factura" primero.');
  SpreadsheetApp.flush();

  var params = [
    'format=pdf',
    'size=A4',
    'portrait=true',
    'fitw=true',
    'gridlines=false',
    'printtitle=false',
    'sheetnames=false',
    'pagenumbers=false',
    'fzr=false',
    'gid=' + sh.getSheetId(),
    'range=A1:F33',
    'top_margin=0.50',
    'bottom_margin=0.50',
    'left_margin=0.50',
    'right_margin=0.50'
  ].join('&');

  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + params;
  var resp = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    throw new Error('No se pudo generar el PDF (código ' + resp.getResponseCode() + ').');
  }
  var nombre = ('Factura_' + numero).replace(/[^\w\-]+/g, '_') + '.pdf';
  var blob = resp.getBlob().setName(nombre);
  var file = carpetaFacturas_().createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) { /* dominio puede restringir */ }
  return file;
}

/** Genera el PDF y muestra enlaces para verlo y descargarlo. */
function generarPDFFactura() {
  try {
    var file = crearPDFFactura_();
    var descarga = 'https://drive.google.com/uc?export=download&id=' + file.getId();
    var html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:14px;line-height:1.6">' +
      '<p style="font-size:14px">✅ <b>La factura se generó en PDF.</b></p>' +
      '<p><a href="' + file.getUrl() + '" target="_blank" style="font-size:14px">👁️ Ver / Imprimir factura</a></p>' +
      '<p><a href="' + descarga + '" target="_blank" style="font-size:14px">⬇️ Descargar PDF</a></p>' +
      '<hr><p style="color:#7f8c8d;font-size:12px">El archivo quedó guardado en la carpeta "Facturas" de tu Google Drive.</p>' +
      '</div>'
    ).setWidth(360).setHeight(220);
    SpreadsheetApp.getUi().showModalDialog(html, 'Factura en PDF');
  } catch (e) {
    alerta_('Error al generar PDF', e.message);
  }
}

/** Envía la factura en PDF por correo al cliente. */
function enviarFacturaCorreo() {
  try {
    var sh = getSheet_(SHEETS.FACTURA);
    var correo = String(sh.getRange('E9').getValue() || '').trim();
    var numero = sh.getRange(FAC.NUMERO).getValue();
    if (!correo) {
      var ui = SpreadsheetApp.getUi();
      var r = ui.prompt('Correo del cliente', 'No hay correo en la factura. Escribe el correo destino:', ui.ButtonSet.OK_CANCEL);
      if (r.getSelectedButton() !== ui.Button.OK) return;
      correo = String(r.getResponseText() || '').trim();
      if (!correo) { alerta_('Sin correo', 'No se indicó un correo destino.'); return; }
    }
    var empresa = getParametro_('Nombre de la empresa');
    var file = crearPDFFactura_();
    MailApp.sendEmail({
      to: correo,
      subject: 'Factura ' + numero + ' - ' + empresa,
      htmlBody: '<p>Hola,</p><p>Adjuntamos la factura <b>' + numero + '</b>.</p>' +
        '<p>También puedes verla aquí: <a href="' + file.getUrl() + '">' + file.getUrl() + '</a></p>' +
        '<p>Gracias por tu compra.<br>' + empresa + '</p>',
      attachments: [file.getBlob()]
    });
    alerta_('✉️ Correo enviado', 'La factura ' + numero + ' fue enviada a ' + correo + '.');
  } catch (e) {
    alerta_('Error al enviar correo', e.message);
  }
}

/** Crea un enlace de WhatsApp con el PDF de la factura. */
function compartirFacturaWhatsApp() {
  try {
    var sh = getSheet_(SHEETS.FACTURA);
    var numero = sh.getRange(FAC.NUMERO).getValue();
    var tel = String(sh.getRange('B9').getValue() || '').replace(/\D/g, '');
    var empresa = getParametro_('Nombre de la empresa');
    var file = crearPDFFactura_();
    var mensaje = 'Hola, te compartimos tu factura ' + numero + ' de ' + empresa +
      '. Puedes verla o descargarla aquí: ' + file.getUrl();
    var wa = 'https://wa.me/' + tel + '?text=' + encodeURIComponent(mensaje);

    var aviso = tel ? '' : '<p style="color:#c0392b">No hay teléfono en la factura: al abrir WhatsApp deberás elegir el contacto.</p>';
    var html = HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;padding:14px;line-height:1.6">' +
      '<p style="font-size:14px">🟢 <b>Compartir factura ' + numero + ' por WhatsApp</b></p>' +
      aviso +
      '<p><a href="' + wa + '" target="_blank" style="font-size:15px;font-weight:bold;color:#128c7e">➡️ Abrir WhatsApp</a></p>' +
      '<hr><p style="color:#7f8c8d;font-size:12px">El PDF está disponible en tu Drive y en el enlace del mensaje.</p>' +
      '</div>'
    ).setWidth(380).setHeight(230);
    SpreadsheetApp.getUi().showModalDialog(html, 'Compartir por WhatsApp');
  } catch (e) {
    alerta_('Error al compartir', e.message);
  }
}
