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
