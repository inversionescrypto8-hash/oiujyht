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
