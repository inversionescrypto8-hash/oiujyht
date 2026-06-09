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
