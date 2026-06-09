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
