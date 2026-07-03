# -*- coding: utf-8 -*-
"""
Genera 'Sistema_Gestion_Inventario.xlsx' usando SOLO la librería estándar.
Un .xlsx es un ZIP de XML (OOXML/SpreadsheetML); aquí se construye a mano.

Compatibilidad: se usan fórmulas estándar (VLOOKUP, SUMIF, SUMIFS, IF,
EOMONTH, COUNTIF, INDEX, MATCH, LARGE) que funcionan igual en Excel y en
Google Sheets. NO se usan ARRAYFORMULA ni QUERY (exclusivas de Google y que
no sobreviven la importación de un .xlsx).
"""
import zipfile
import datetime
from xml.sax.saxutils import escape

# ----------------------------------------------------------------------
# Parámetros de tamaño (filas con fórmulas precargadas)
# ----------------------------------------------------------------------
CAT_ROWS = 200      # productos: filas 2..201
MOV_ROWS = 300      # movimientos: filas 2..301
INCLUDE_SAMPLES = True   # True = con datos de ejemplo; False = plantilla vacía
CLI_ROWS = 100      # clientes:   filas 2..101
CAT_LAST = CAT_ROWS + 1   # 201
MOV_LAST = MOV_ROWS + 1   # 301
CLI_LAST = CLI_ROWS + 1   # 101

CASA = "Casa"
ML = "Mercado Libre"

# Colores (ARGB)
C_PRIMARIO = "FF1F3A5F"
C_SECUND = "FF2E6DA4"
C_ACENTO = "FF16A085"
C_CLARO = "FFEEF3F8"
C_ALERTA = "FFC0392B"
C_ADVERT = "FFE67E22"
C_TEXTO = "FF2C3E50"
C_BLANCO = "FFFFFFFF"
C_BORDE = "FFCFD8E3"
C_DARK = "FF34495E"
C_INPUT = "FFFFFBE6"
C_AUTOCALC = "FFF4F8FC"
C_GRAY = "FF95A5A6"

# ----------------------------------------------------------------------
# Registros de estilos (deduplican y devuelven índices)
# ----------------------------------------------------------------------
class Reg:
    def __init__(self, start=0):
        self.items = []
        self.index = {}
        self.start = start

    def add(self, key, xml):
        if key in self.index:
            return self.index[key]
        idx = len(self.items) + self.start
        self.index[key] = idx
        self.items.append(xml)
        return idx


numfmts = Reg(start=164)   # ids personalizados >= 164
fonts = Reg()
fills = Reg()
borders = Reg()
xfs = Reg()

# numFmts
FMT_MONEY = numfmts.add('money', '<numFmt numFmtId="164" formatCode="&quot;$ &quot;#,##0"/>')
FMT_INT = numfmts.add('int', '<numFmt numFmtId="165" formatCode="#,##0"/>')
FMT_PCT = numfmts.add('pct', '<numFmt numFmtId="166" formatCode="0.0%"/>')
FMT_DATE = numfmts.add('date', '<numFmt numFmtId="167" formatCode="dd/mm/yyyy"/>')
FMT_SIGNED = numfmts.add('signed', '<numFmt numFmtId="168" formatCode="+#,##0;\\-#,##0;0"/>')

# fills: reservar 0 (none) y 1 (gray125)
fills.add('none', '<fill><patternFill patternType="none"/></fill>')
fills.add('gray125', '<fill><patternFill patternType="gray125"/></fill>')

def fill_solid(color):
    return fills.add('solid' + color,
                     '<fill><patternFill patternType="solid"><fgColor rgb="%s"/><bgColor indexed="64"/></patternFill></fill>' % color)

FILL_PRIM = fill_solid(C_PRIMARIO)
FILL_SEC = fill_solid(C_SECUND)
FILL_ACC = fill_solid(C_ACENTO)
FILL_LIGHT = fill_solid(C_CLARO)
FILL_INPUT = fill_solid(C_INPUT)
FILL_AUTO = fill_solid(C_AUTOCALC)
FILL_ALERT = fill_solid(C_ALERTA)
FILL_ADVERT = fill_solid(C_ADVERT)
FILL_DARK = fill_solid(C_DARK)

# fonts
def font(key, sz=11, bold=False, color=C_TEXTO, name="Arial"):
    b = '<b/>' if bold else ''
    return fonts.add(key, '<font>%s<sz val="%d"/><color rgb="%s"/><name val="%s"/></font>' % (b, sz, color, name))

F_DEFAULT = font('default')
F_HDR = font('hdr', sz=10, bold=True, color=C_BLANCO)
F_BOLD = font('bold', bold=True)
F_TITLE = font('title', sz=20, bold=True, color=C_BLANCO)
F_SUBHDR = font('subhdr', sz=11, bold=True, color=C_BLANCO)
F_ALERT = font('alert', bold=True, color=C_ALERTA)
F_SECOND = font('second', color=C_SECUND)
F_GRAYSMALL = font('graysmall', sz=8, color=C_GRAY)
F_TOTAL = font('total', sz=13, bold=True, color=C_BLANCO)
F_KPI = font('kpi', sz=18, bold=True, color=C_PRIMARIO)
F_KPIW = font('kpiw', sz=18, bold=True, color=C_BLANCO)
F_COMPANY = font('company', sz=16, bold=True, color=C_PRIMARIO)
F_FACTITLE = font('factitle', sz=13, bold=True, color=C_BLANCO)
F_FOOTER = font('footer', sz=12, bold=True, color=C_PRIMARIO)

# borders
BORDER_NONE = borders.add('none', '<border><left/><right/><top/><bottom/><diagonal/></border>')
def thin():
    e = '<left style="thin"><color rgb="%s"/></left><right style="thin"><color rgb="%s"/></right><top style="thin"><color rgb="%s"/></top><bottom style="thin"><color rgb="%s"/></bottom><diagonal/>' % (C_BORDE, C_BORDE, C_BORDE, C_BORDE)
    return borders.add('thin', '<border>' + e + '</border>')
BORDER_THIN = thin()

# cellXfs
def xf(numFmtId=0, fontId=0, fillId=0, borderId=0, halign=None, valign=None, wrap=False):
    al = ''
    if halign or valign or wrap:
        parts = []
        if halign: parts.append('horizontal="%s"' % halign)
        if valign: parts.append('vertical="%s"' % valign)
        if wrap: parts.append('wrapText="1"')
        al = '<alignment %s/>' % ' '.join(parts)
    applyFont = ' applyFont="1"' if fontId else ''
    applyFill = ' applyFill="1"' if fillId else ''
    applyBorder = ' applyBorder="1"' if borderId else ''
    applyNum = ' applyNumberFormat="1"' if numFmtId else ''
    applyAl = ' applyAlignment="1"' if al else ''
    key = (numFmtId, fontId, fillId, borderId, halign, valign, wrap)
    xml = '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0"%s%s%s%s%s>%s</xf>' % (
        numFmtId, fontId, fillId, borderId, applyNum, applyFont, applyFill, applyBorder, applyAl, al)
    return xfs.add(key, xml)

# Estilos base (índice 0 = default obligatorio)
S_DEFAULT = xf()
S_HDR = xf(fontId=F_HDR, fillId=FILL_PRIM, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_TITLE = xf(fontId=F_TITLE, fillId=FILL_PRIM, halign="center", valign="center")
S_BANNER_SEC = xf(fontId=F_SUBHDR, fillId=FILL_SEC, halign="left", valign="center")
S_BANNER_ALERT = xf(fontId=F_SUBHDR, fillId=FILL_ALERT, halign="left", valign="center")
S_BANNER_PRIM = xf(fontId=F_SUBHDR, fillId=FILL_PRIM, halign="left", valign="center")
S_SUBTITLE = xf(fontId=F_BOLD, fillId=FILL_LIGHT, halign="center", valign="center")
S_LABEL = xf(fontId=F_BOLD, fillId=FILL_LIGHT, borderId=BORDER_THIN, halign="left", valign="center")
S_LABEL_RIGHT = xf(fontId=F_BOLD, halign="right", valign="center")
S_TEXT = xf(borderId=BORDER_THIN, halign="left", valign="center")
S_TEXT_WRAP = xf(borderId=BORDER_THIN, halign="left", valign="top", wrap=True)
S_MONEY = xf(numFmtId=FMT_MONEY, borderId=BORDER_THIN, halign="right", valign="center")
S_INT = xf(numFmtId=FMT_INT, borderId=BORDER_THIN, halign="center", valign="center")
S_PCT = xf(numFmtId=FMT_PCT, borderId=BORDER_THIN, halign="center", valign="center")
S_DATE = xf(numFmtId=FMT_DATE, borderId=BORDER_THIN, halign="center", valign="center")
S_SIGNED = xf(numFmtId=FMT_SIGNED, borderId=BORDER_THIN, halign="center", valign="center")
# automáticas (fondo azul claro)
S_MONEY_A = xf(numFmtId=FMT_MONEY, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="right", valign="center")
S_INT_A = xf(numFmtId=FMT_INT, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="center", valign="center")
S_PCT_A = xf(numFmtId=FMT_PCT, fillId=FILL_AUTO, borderId=BORDER_THIN, halign="center", valign="center")
S_TEXT_A = xf(fillId=FILL_AUTO, borderId=BORDER_THIN, halign="left", valign="center")
# entrada (fondo amarillo)
S_INPUT = xf(fillId=FILL_INPUT, borderId=BORDER_THIN, halign="left", valign="center")
S_INPUT_INT = xf(numFmtId=FMT_INT, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="center", valign="center")
S_INPUT_MONEY = xf(numFmtId=FMT_MONEY, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="right", valign="center")
S_INPUT_DATE = xf(numFmtId=FMT_DATE, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="center", valign="center")
# KPIs
S_KPI_LAB_SEC = xf(fontId=F_SUBHDR, fillId=FILL_SEC, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ACC = xf(fontId=F_SUBHDR, fillId=FILL_ACC, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ADV = xf(fontId=F_SUBHDR, fillId=FILL_ADVERT, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_PRI = xf(fontId=F_SUBHDR, fillId=FILL_PRIM, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_DARK = xf(fontId=F_SUBHDR, fillId=FILL_DARK, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_LAB_ALE = xf(fontId=F_SUBHDR, fillId=FILL_ALERT, borderId=BORDER_THIN, halign="center", valign="center", wrap=True)
S_KPI_MONEY = xf(numFmtId=FMT_MONEY, fontId=F_KPI, borderId=BORDER_THIN, halign="center", valign="center")
S_KPI_INT = xf(numFmtId=FMT_INT, fontId=F_KPI, borderId=BORDER_THIN, halign="center", valign="center")
# factura
S_FACTITLE = xf(fontId=F_FACTITLE, fillId=FILL_PRIM, halign="center", valign="center")
S_COMPANY = xf(fontId=F_COMPANY, halign="left", valign="center")
S_TOTLAB = xf(fontId=F_TOTAL, fillId=FILL_ACC, borderId=BORDER_THIN, halign="right", valign="center")
S_TOTVAL = xf(numFmtId=FMT_MONEY, fontId=F_TOTAL, fillId=FILL_ACC, borderId=BORDER_THIN, halign="right", valign="center")
S_FOOTER = xf(fontId=F_FOOTER, halign="center", valign="center")
S_FOOTER_SM = xf(fontId=F_GRAYSMALL, halign="center", valign="center")
S_SECOND = xf(fontId=F_SECOND, halign="left", valign="center")

# catálogo de venta (PDF de productos con imagen)
F_CVNAME = font('cvname', sz=12, bold=True, color=C_PRIMARIO)
F_CVPRICE = font('cvprice', sz=16, bold=True, color=C_ACENTO)
F_CVTITLE = font('cvtitle', sz=22, bold=True, color=C_BLANCO)
S_CV_IMG = xf(borderId=BORDER_THIN, halign="center", valign="center")
S_CV_NAME = xf(fontId=F_CVNAME, borderId=BORDER_THIN, halign="left", valign="center", wrap=True)
S_CV_CODE = xf(fontId=F_GRAYSMALL, borderId=BORDER_THIN, halign="left", valign="center")
S_CV_PRICE = xf(numFmtId=FMT_MONEY, fontId=F_CVPRICE, borderId=BORDER_THIN, halign="right", valign="center")
S_CV_TITLE = xf(fontId=F_CVTITLE, fillId=FILL_PRIM, halign="center", valign="center")
S_CV_SUB = xf(fontId=F_BOLD, fillId=FILL_LIGHT, halign="center", valign="center")

# ----------------------------------------------------------------------
# Utilidades de celda
# ----------------------------------------------------------------------
def col_letter(c):
    s = ""
    while c > 0:
        c, r = divmod(c - 1, 26)
        s = chr(65 + r) + s
    return s

def ref(r, c):
    return col_letter(c) + str(r)

EPOCH = datetime.date(1899, 12, 30)
def serial(d):
    return (d - EPOCH).days


class Sheet:
    def __init__(self, name, hide_gridlines=False, freeze_row=0, freeze_col=0):
        self.name = name
        self.cells = {}          # (r,c) -> dict
        self.merges = []         # "A1:B2"
        self.cols = {}           # c -> width
        self.hidden_cols = set() # columnas ocultas (auxiliares)
        self.rows = {}           # r -> height
        self.validations = []    # (sqref, formula1)
        self.hide_gridlines = hide_gridlines
        self.freeze_row = freeze_row
        self.freeze_col = freeze_col
        self.max_r = 1
        self.max_c = 1

    def _touch(self, r, c):
        self.max_r = max(self.max_r, r)
        self.max_c = max(self.max_c, c)

    def text(self, r, c, value, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 's', 'v': value, 's': style}
        self._touch(r, c)

    def num(self, r, c, value, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'n', 'v': value, 's': style}
        self._touch(r, c)

    def date(self, r, c, d, style=S_DATE):
        self.cells[(r, c)] = {'t': 'n', 'v': serial(d), 's': style}
        self._touch(r, c)

    def formula(self, r, c, f, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'f', 'v': f, 's': style}
        self._touch(r, c)

    def blank(self, r, c, style=S_DEFAULT):
        self.cells[(r, c)] = {'t': 'b', 'v': None, 's': style}
        self._touch(r, c)

    def merge(self, r1, c1, r2, c2):
        self.merges.append(ref(r1, c1) + ":" + ref(r2, c2))

    def colw(self, c, w):
        self.cols[c] = w

    def hide_col(self, c, w=8):
        self.cols[c] = w
        self.hidden_cols.add(c)

    def rowh(self, r, h):
        self.rows[r] = h

    def validate_list(self, r1, c1, r2, c2, formula1):
        self.validations.append((ref(r1, c1) + ":" + ref(r2, c2), formula1))

    def xml(self):
        out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>']
        out.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        out.append('<dimension ref="A1:%s"/>' % ref(self.max_r, self.max_c))
        # sheetViews
        sv = '<sheetView workbookViewId="0"'
        if self.hide_gridlines:
            sv += ' showGridLines="0"'
        sv += '>'
        if self.freeze_row or self.freeze_col:
            tl = ref(self.freeze_row + 1, self.freeze_col + 1)
            sv += ('<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>'
                   % (self.freeze_col, self.freeze_row, tl))
            sv += '<selection pane="bottomRight" activeCell="%s" sqref="%s"/>' % (tl, tl)
        sv += '</sheetView>'
        out.append('<sheetViews>%s</sheetViews>' % sv)
        out.append('<sheetFormatPr defaultRowHeight="15"/>')
        # cols
        if self.cols:
            out.append('<cols>')
            for c in sorted(self.cols):
                hid = ' hidden="1"' if c in self.hidden_cols else ''
                out.append('<col min="%d" max="%d" width="%.2f" customWidth="1"%s/>' % (c, c, self.cols[c], hid))
            out.append('</cols>')
        # sheetData
        out.append('<sheetData>')
        rows = {}
        for (r, c), cell in self.cells.items():
            rows.setdefault(r, {})[c] = cell
        for r in sorted(rows):
            ht = ''
            if r in self.rows:
                ht = ' ht="%.2f" customHeight="1"' % self.rows[r]
            out.append('<row r="%d"%s>' % (r, ht))
            for c in sorted(rows[r]):
                cell = rows[r][c]
                rr = ref(r, c)
                s = cell['s']
                t = cell['t']
                if t == 's':
                    out.append('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                               % (rr, s, escape(str(cell['v']))))
                elif t == 'n':
                    v = cell['v']
                    vs = repr(v) if isinstance(v, float) else str(v)
                    out.append('<c r="%s" s="%d"><v>%s</v></c>' % (rr, s, vs))
                elif t == 'f':
                    out.append('<c r="%s" s="%d"><f>%s</f></c>' % (rr, s, escape(cell['v'])))
                else:
                    out.append('<c r="%s" s="%d"/>' % (rr, s))
            out.append('</row>')
        out.append('</sheetData>')
        # mergeCells
        if self.merges:
            out.append('<mergeCells count="%d">' % len(self.merges))
            for m in self.merges:
                out.append('<mergeCell ref="%s"/>' % m)
            out.append('</mergeCells>')
        # dataValidations
        if self.validations:
            out.append('<dataValidations count="%d">' % len(self.validations))
            for sqref, f1 in self.validations:
                out.append('<dataValidation type="list" allowBlank="1" showInputMessage="1" '
                           'showErrorMessage="0" sqref="%s"><formula1>%s</formula1></dataValidation>'
                           % (sqref, escape(f1)))
            out.append('</dataValidations>')
        out.append('<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>')
        out.append('</worksheet>')
        return ''.join(out)


# ----------------------------------------------------------------------
# Construcción de las hojas
# ----------------------------------------------------------------------
def q(name):
    """Nombre de hoja entre comillas para fórmulas."""
    return "'" + name + "'"

CAT = "Catálogo"
COMP = "Compras"
VEN = "Ventas"
TRA = "Traslados"
AJU = "Ajustes"
INV = "Inventario"
CLI = "Clientes"
FAC = "Factura"
CFG = "Config"
DASH = "Dashboard"
CVENTA = "Catálogo Venta"
CIERRE = "Cierre Diario"
PREST = "Préstamos"
ABONOS = "Abonos Préstamos"
BILL = "Billeteras"
TARJ = "Tarjetas"
MOV = "Movimientos"
PAGOST = "Pagos Tarjeta"

today = datetime.date(2026, 6, 4)

# ---------- CONFIG ----------
def build_config():
    s = Sheet(CFG, freeze_row=1)
    s.colw(1, 34); s.colw(2, 40); s.colw(4, 22); s.colw(5, 10); s.colw(6, 22); s.colw(8, 16); s.colw(10, 26)
    s.text(1, 1, "Parámetro", S_HDR); s.text(1, 2, "Valor", S_HDR)
    params = [
        ("Nombre de la empresa", "Mi Negocio S.A.S."),
        ("NIT / Identificación", "000.000.000-0"),
        ("Dirección", "Calle 00 # 00-00, Ciudad"),
        ("Teléfono", "+57 300 000 0000"),
        ("Correo electrónico", "correo@minegocio.com"),
        ("Sitio web", "www.minegocio.com"),
        ("Prefijo de factura", "FAC-"),
        ("Próximo número de factura", 1),
        ("Mensaje pie de factura", "¡Gracias por su compra!"),
    ]
    r = 2
    for k, v in params:
        s.text(r, 1, k, S_LABEL)
        if isinstance(v, (int, float)):
            s.num(r, 2, v, S_TEXT)
        else:
            s.text(r, 2, v, S_TEXT)
        r += 1
    # Costos fijos POR VENTA — valor de REFERENCIA (se precarga en cada fila nueva de Ventas,
    # pero cada venta puede cambiarlo individualmente allá si ese día costó distinto).
    # Filas 11-12 (no mueven B2..B10).
    s.text(11, 1, "Costo fijo por venta: envío (Mary) — referencia", S_LABEL); s.num(11, 2, 1000, S_INPUT_MONEY)
    s.text(12, 1, "Costo fijo por venta: bolsa y etiqueta — referencia", S_LABEL); s.num(12, 2, 1000, S_INPUT_MONEY)
    # Días de la tarjeta (genéricos): cierre y pago. Se usan para avisar vencimientos.
    s.text(13, 1, "Día de cierre de tarjeta", S_LABEL); s.num(13, 2, 30, S_INPUT_INT)
    s.text(14, 1, "Día de pago de tarjeta", S_LABEL); s.num(14, 2, 16, S_INPUT_INT)
    s.text(15, 1, "Avisar si faltan (días) o menos", S_LABEL); s.num(15, 2, 5, S_INPUT_INT)
    # Tabla IVA por categoría (columna D = categoría, columna E = IVA %).
    # Se deja el IVA EN BLANCO: escribe 19, 5, 0, etc. según cada categoría.
    # Puedes AGREGAR más categorías escribiéndolas debajo (hasta la fila 30).
    categorias = ["General", "Electrónica", "Hogar", "Ropa y calzado",
                  "Accesorios", "Papelería", "Arte", "Otros"]
    s.text(1, 4, "Categorías", S_SUBTITLE)
    s.text(1, 5, "IVA %", S_SUBTITLE)
    for i, v in enumerate(categorias):
        s.text(2 + i, 4, v, S_TEXT)
        s.blank(2 + i, 5, S_INPUT_INT)   # IVA en blanco, editable por categoría
    # Filas extra en blanco para nuevas categorías (hasta fila 30).
    for r in range(2 + len(categorias), 31):
        s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT_INT)
    s.colw(5, 80)
    # listas auxiliares
    def lista(col, titulo, valores):
        s.text(1, col, titulo, S_SUBTITLE)
        for i, v in enumerate(valores):
            s.text(2 + i, col, v, S_TEXT)
    lista(6, "Ubicaciones", [CASA, ML])
    lista(8, "Estados", ["Activo", "Inactivo"])
    lista(10, "Motivos de ajuste", ["Producto dañado", "Pérdida", "Corrección de conteo", "Ajuste administrativo", "Devolución"])
    # Columna L: lista para el "Catálogo Venta" = TODAS + categorías (se actualiza sola).
    s.text(1, 12, "Selección catálogo", S_SUBTITLE)
    s.text(2, 12, "TODAS", S_TEXT)
    for k in range(3, 31):           # L3..L30 reflejan D2..D29
        s.formula(k, 12, 'IF($D%d="","",$D%d)' % (k - 1, k - 1), S_TEXT)
    s.colw(12, 150)
    # ===== TESORERÍA: billeteras y tarjetas =====
    # Tarjetas de crédito (columna N) — tus tarjetas reales.
    tarjetas = ["TC Bancolombia 1 (1517)", "TC Bancolombia 2 (0554)", "TC Bancolombia 3 (4582)",
                "TC Bancolombia 4 (8366)", "TC Bancolombia 5 (0623)", "TC Bancolombia MM 6 (6423)",
                "TC Bancolombia MM 7 (2459)", "TC NU 8 (3750)"]
    s.text(1, 14, "Tarjetas de crédito", S_SUBTITLE)
    for i, v in enumerate(tarjetas):
        s.text(2 + i, 14, v, S_TEXT)
    for r2 in range(2 + len(tarjetas), 21):
        s.blank(r2, 14, S_INPUT)
    s.colw(14, 220)
    # Billeteras / formas de cobro (columna P) — donde tienes/recibes la plata.
    billeteras = ["Efectivo", "Nequi", "Bancolombia", "Bancolombia MM", "Daviplata",
                  "Davivienda", "Mercado Libre", "Skydrops"]
    s.text(1, 16, "Billeteras (formas de cobro)", S_SUBTITLE)
    for i, v in enumerate(billeteras):
        s.text(2 + i, 16, v, S_TEXT)
    for r2 in range(2 + len(billeteras), 21):
        s.blank(r2, 16, S_INPUT)
    s.colw(16, 200)
    # Medio de compra (columna R) = billeteras + tarjetas (para elegir cómo pagaste una compra).
    # "Inventario inicial" = productos que YA tenías (no descuenta de ninguna billetera ni tarjeta).
    medios_compra = ["Inventario inicial (ya pagado)"] + billeteras + tarjetas
    s.text(1, 18, "Medio de compra (billetera o tarjeta)", S_SUBTITLE)
    for i, v in enumerate(medios_compra):
        s.text(2 + i, 18, v, S_TEXT)
    for r2 in range(2 + len(medios_compra), 30):
        s.blank(r2, 18, S_INPUT)
    s.colw(18, 240)
    return s

# ---------- CATÁLOGO ----------
def build_catalogo():
    s = Sheet(CAT, freeze_row=1, freeze_col=2)
    headers = ["Código", "Nombre", "Categoría", "Estado", "Stock Mínimo",
               "Costo Promedio", "Stock Casa", "Stock Mercado Libre", "Stock Total",
               "Valor Inventario", "Estado Stock", "IVA %",
               "URL Imagen", "Precio Detal", "Precio Mayor", "_rank", "_key",
               "Medidas (ej. 30x40)", "_orden", "_rankAll", "_keyAll",
               "Fecha 1ª Compra", "Días en Bodega", "Vendidas (total)", "Reabastecer"]
    widths = [12, 32, 18, 11, 12, 15, 12, 18, 12, 16, 14, 8, 36, 14, 14, 8, 8, 16, 8, 8, 8,
              14, 13, 14, 30]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    s.hide_col(16, 8)  # columna _rank (auxiliar para el Catálogo Venta)
    s.hide_col(17, 8)  # columna _key  (auxiliar: categoría|posición)
    s.hide_col(19, 8)  # columna _orden (auxiliar: área para ordenar por medidas)
    s.hide_col(20, 8)  # columna _rankAll (auxiliar: orden global para "TODAS")
    s.hide_col(21, 8)  # columna _keyAll (auxiliar: "TODAS|posición")
    # cod, nombre, categoria, estado, stock min, url imagen, precio detal, precio mayor, medidas
    sample = [] if not INCLUDE_SAMPLES else [
        ("P001", "Audífonos Bluetooth", "Electrónica", "Activo", 5, "", 75000, 60000, ""),
        ("P002", "Cargador USB-C 20W", "Electrónica", "Activo", 8, "", 28000, 22000, ""),
        ("P003", "Camiseta básica", "Ropa y calzado", "Activo", 10, "", 25000, 18000, ""),
        ("P004", "Termo 1L acero", "Hogar", "Activo", 4, "", 45000, 36000, ""),
        ("P005", "Cuaderno argollado", "Papelería", "Activo", 15, "", 12000, 8000, ""),
        ("P006", "Mouse inalámbrico", "Electrónica", "Activo", 6, "", 38000, 30000, ""),
        ("A001", "Cuadro abstracto azul", "Arte", "Activo", 1, "", 120000, 90000, "20x20"),
        ("A002", "Paisaje montañas", "Arte", "Activo", 1, "", 180000, 140000, "30x30"),
        ("A003", "Retrato moderno", "Arte", "Activo", 1, "", 220000, 170000, "30x40"),
        ("A004", "Mural flores", "Arte", "Activo", 1, "", 350000, 280000, "50x70"),
    ]
    for r in range(2, CAT_LAST + 1):
        i = r - 2
        if i < len(sample):
            cod, nom, cat, est, mn, img, pdet, pmay, med = sample[i]
            s.text(r, 1, cod, S_INPUT); s.text(r, 2, nom, S_INPUT); s.text(r, 3, cat, S_INPUT)
            s.text(r, 4, est, S_INPUT); s.num(r, 5, mn, S_INPUT_INT)
            s.text(r, 13, img, S_INPUT)
            s.num(r, 14, pdet, S_INPUT_MONEY); s.num(r, 15, pmay, S_INPUT_MONEY)
            if med:
                s.text(r, 18, med, S_INPUT)
            else:
                s.blank(r, 18, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT); s.blank(r, 2, S_INPUT); s.blank(r, 3, S_INPUT)
            s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT_INT)
            s.blank(r, 13, S_INPUT); s.blank(r, 14, S_INPUT_MONEY); s.blank(r, 15, S_INPUT_MONEY)
            s.blank(r, 18, S_INPUT)
        A = "$A%d" % r
        # F costo promedio
        s.formula(r, 6, 'IF(%s="","",IFERROR(SUMIF(%s!$B:$B,%s,%s!$F:$F)/SUMIF(%s!$B:$B,%s,%s!$D:$D),0))'
                  % (A, q(COMP), A, q(COMP), q(COMP), A, q(COMP)), S_MONEY_A)
        # G stock casa
        s.formula(r, 7,
                  'IF(%s="","",SUMIF(%s!$B:$B,%s,%s!$D:$D)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$G:$G,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$M:$M,%s&"|%s",%s!$F:$F)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D))'
                  % (A, q(COMP), A, q(COMP),
                     q(TRA), A, CASA, q(TRA),
                     q(TRA), A, CASA, q(TRA),
                     q(VEN), A, CASA, q(VEN),
                     q(AJU), A, CASA, q(AJU)), S_INT_A)
        # H stock ML
        s.formula(r, 8,
                  'IF(%s="","",SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$G:$G,%s&"|%s",%s!$D:$D)'
                  '-SUMIF(%s!$M:$M,%s&"|%s",%s!$F:$F)'
                  '+SUMIF(%s!$H:$H,%s&"|%s",%s!$D:$D))'
                  % (A, q(TRA), A, ML, q(TRA),
                     q(TRA), A, ML, q(TRA),
                     q(VEN), A, ML, q(VEN),
                     q(AJU), A, ML, q(AJU)), S_INT_A)
        # I total
        s.formula(r, 9, 'IF(%s="","",$G%d+$H%d)' % (A, r, r), S_INT_A)
        # J valor inventario
        s.formula(r, 10, 'IF(%s="","",$I%d*$F%d)' % (A, r, r), S_MONEY_A)
        # K estado stock
        s.formula(r, 11, 'IF(%s="","",IF($I%d<=0,"Agotado",IF($I%d<=$E%d,"Stock bajo","Disponible")))'
                  % (A, r, r, r), S_TEXT_A)
        # L IVA % (según categoría, tomado de la tabla Config!D:E). En blanco si la categoría no tiene IVA.
        s.formula(r, 12, 'IF(%s="","",IFERROR(VLOOKUP($C%d,%s!$D$2:$E$30,2,FALSE),0))'
                  % (A, r, q(CFG)), S_INT_A)
        # S (_orden): área del cuadro a partir de "AOxB" (ej. 30x40 -> 1200). 0 si no hay medidas.
        #   Sirve para ordenar los artículos (arte) de menor a mayor tamaño.
        s.formula(r, 19,
                  'IF(%s="","",IFERROR('
                  'VALUE(TRIM(LEFT($R%d,FIND("x",LOWER($R%d))-1)))*'
                  'VALUE(TRIM(MID($R%d,FIND("x",LOWER($R%d))+1,20))),0))'
                  % (A, r, r, r, r), S_INT_A)
        # P (_rank): posición del producto dentro de su categoría (solo activos Y con stock > 0),
        #   ORDENADO por área (_orden) y, a igualdad, por nombre. Para el Catálogo Venta.
        #   Si el producto no tiene stock (aún no se le ha hecho ninguna Compra, o se agotó),
        #   NO entra al ranking y por lo tanto no sale en el Catálogo Venta, sin tener que
        #   marcarlo manualmente como Inactivo.
        s.formula(r, 16,
                  'IF(OR(%s="",$D%d<>"Activo",$I%d<=0),"",'
                  'SUMPRODUCT(($C$2:$C$%d=$C%d)*($D$2:$D$%d="Activo")*($I$2:$I$%d>0)*('
                  '($S$2:$S$%d<$S%d)+(($S$2:$S$%d=$S%d)*($B$2:$B$%d<$B%d))))+1)'
                  % (A, r, r, CAT_LAST, r, CAT_LAST, CAT_LAST,
                     CAT_LAST, r, CAT_LAST, r, CAT_LAST, r), S_INT_A)
        # Q (_key): "categoría|posición" para buscar con coincidencia exacta desde Catálogo Venta.
        s.formula(r, 17, 'IF($P%d="","",$C%d&"|"&$P%d)' % (r, r, r), S_TEXT_A)
        # T (_rankAll): orden GLOBAL de todos los activos CON STOCK (categoría -> área -> nombre).
        #   Permite mostrar TODAS las categorías juntas en el Catálogo Venta, sin productos agotados.
        s.formula(r, 20,
                  'IF(OR(%s="",$D%d<>"Activo",$I%d<=0),"",'
                  'SUMPRODUCT(($D$2:$D$%d="Activo")*($I$2:$I$%d>0)*('
                  '($C$2:$C$%d<$C%d)'
                  '+(($C$2:$C$%d=$C%d)*($S$2:$S$%d<$S%d))'
                  '+(($C$2:$C$%d=$C%d)*($S$2:$S$%d=$S%d)*($B$2:$B$%d<$B%d))))+1)'
                  % (A, r, r, CAT_LAST, CAT_LAST,
                     CAT_LAST, r,
                     CAT_LAST, r, CAT_LAST, r,
                     CAT_LAST, r, CAT_LAST, r, CAT_LAST, r), S_INT_A)
        # U (_keyAll): "TODAS|posición".
        s.formula(r, 21, 'IF($T%d="","","TODAS|"&$T%d)' % (r, r), S_TEXT_A)
        # V Fecha de la 1ª compra (cuándo entró por primera vez al inventario).
        s.formula(r, 22,
                  'IF(%s="","",IFERROR(IF(COUNTIF(%s!$B:$B,$A%d)=0,"",MINIFS(%s!$A:$A,%s!$B:$B,$A%d)),""))'
                  % (A, q(COMP), r, q(COMP), q(COMP), r), S_DATE)
        # W Días en bodega (desde la 1ª compra).
        s.formula(r, 23, 'IF(OR(%s="",$V%d=""),"",TODAY()-$V%d)' % (A, r, r), S_INT_A)
        # X Unidades vendidas (total histórico).
        s.formula(r, 24, 'IF(%s="","",SUMIF(%s!$D:$D,$A%d,%s!$F:$F))' % (A, q(VEN), r, q(VEN)), S_INT_A)
        # Y Sugerencia de reabastecer (rotación rápida o stock bajo/agotado).
        s.formula(r, 25,
                  'IF(%s="","",'
                  'IF(AND($X%d>=3,$W%d<>"",$W%d<=15),"🔥 Se vende rápido — REABASTECE",'
                  'IF($K%d="Agotado","🔴 Agotado — reabastece",'
                  'IF($K%d="Stock bajo","🟠 Stock bajo","✅ Ok"))))'
                  % (A, r, r, r, r, r), S_TEXT_A)
    # validaciones
    s.validate_list(2, 3, CAT_LAST, 3, "%s!$D$2:$D$30" % q(CFG))
    s.validate_list(2, 4, CAT_LAST, 4, "%s!$H$2:$H$3" % q(CFG))
    return s

# ---------- COMPRAS ----------
def build_compras():
    s = Sheet(COMP, freeze_row=1, freeze_col=3)
    # IMPORTANTE: B=Código, D=Cantidad, F=Costo Total se mantienen (el Catálogo depende de ellas).
    # Los costos fijos ($ Mary + bolsa) NO van aquí: son por VENTA (ver hoja Ventas).
    headers = ["Fecha", "Código", "Producto", "Cantidad", "Costo Unitario", "Costo Total",
               "Medio de Pago", "¿A crédito?", "Fecha Cierre", "Fecha Límite Pago",
               "Abonado", "Saldo", "Estado Pago", "Alerta Pago",
               "Fecha Llegada", "Días en Llegar"]
    widths = [12, 12, 26, 10, 14, 14, 18, 11, 13, 14, 13, 14, 14, 30, 13, 13]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    # fecha(pedido), cod, cant, costo unit, medio pago, a credito(Si/No), abonado, fecha llegada
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 1), "P001", 20, 35000, "Tarjeta Crédito 1", "Sí", 0, datetime.date(2026, 6, 5)),
        (datetime.date(2026, 6, 1), "P002", 30, 12000, "Efectivo", "No", 0, datetime.date(2026, 6, 4)),
        (datetime.date(2026, 6, 2), "P003", 40, 9000, "Tarjeta Crédito 1", "Sí", 200000, datetime.date(2026, 6, 7)),
        (datetime.date(2026, 6, 2), "P004", 15, 22000, "Transferencia", "No", 0, datetime.date(2026, 6, 6)),
        (datetime.date(2026, 6, 3), "P006", 18, 28000, "Tarjeta Crédito 2", "Sí", 0, datetime.date(2026, 6, 9)),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, cu, medio, cred, abonado, lleg = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT)
            s.num(r, 4, qy, S_INPUT_INT); s.num(r, 5, cu, S_INPUT_MONEY)
            s.text(r, 7, medio, S_INPUT); s.text(r, 8, cred, S_INPUT)
            s.num(r, 11, abonado, S_INPUT_MONEY); s.date(r, 15, lleg, S_INPUT_DATE)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT)
            s.blank(r, 4, S_INPUT_INT); s.blank(r, 5, S_INPUT_MONEY)
            s.blank(r, 7, S_INPUT); s.blank(r, 8, S_INPUT)
            s.blank(r, 11, S_INPUT_MONEY); s.blank(r, 15, S_INPUT_DATE)
        B = "$B%d" % r
        # C producto (depende de B)
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        # F Costo Total = Cantidad * Costo Unitario  (NO se toca: el Catálogo lo usa)
        s.formula(r, 6, 'IF(%s="","",$D%d*$E%d)' % (B, r, r), S_MONEY_A)
        # I Fecha de cierre: si la compra es DESPUÉS del día de cierre, cierra el mes siguiente.
        #   (usa Config!B13 = día de cierre, p. ej. 30)
        s.formula(r, 9,
                  'IF(OR(%s="",$H%d<>"Sí"),"",'
                  'IF(DAY($A%d)<=%s!$B$13,'
                  'DATE(YEAR($A%d),MONTH($A%d),%s!$B$13),'
                  'EOMONTH($A%d,0)+%s!$B$13))'
                  % (B, r, r, q(CFG), r, r, q(CFG), r, q(CFG)), S_DATE)
        # J Fecha límite de pago = mes siguiente al cierre, día de pago (Config!B14, p. ej. 16)
        s.formula(r, 10,
                  'IF($I%d="","",DATE(YEAR(EOMONTH($I%d,0)+1),MONTH(EOMONTH($I%d,0)+1),%s!$B$14))'
                  % (r, r, r, q(CFG)), S_DATE)
        # L Saldo = Costo Total - Abonado (lo que aún debes de ese pedido)
        s.formula(r, 12, 'IF(%s="","",MAX($F%d-$K%d,0))' % (B, r, r), S_MONEY_A)
        # M Estado de pago
        s.formula(r, 13,
                  'IF(%s="","",IF($H%d<>"Sí","Pagado (contado)",'
                  'IF($L%d<=0,"Pagado",IF($K%d>0,"Abono parcial","Pendiente"))))'
                  % (B, r, r, r), S_TEXT_A)
        # N Alerta de pago: avisa si está pendiente y se acerca/pasó la fecha límite.
        s.formula(r, 14,
                  'IF(OR(%s="",$H%d<>"Sí",$L%d<=0),"",'
                  'IF(TODAY()>$J%d,"🔴 VENCIDO hace "&(TODAY()-$J%d)&" día(s)",'
                  'IF($J%d-TODAY()<=%s!$B$15,"🟠 Faltan "&($J%d-TODAY())&" día(s) para pagar",'
                  '"🟢 Al día (paga el "&TEXT($J%d,"dd/mm")&")")))'
                  % (B, r, r, r, r, r, q(CFG), r, r), S_TEXT_A)
        # P Días en llegar = Fecha Llegada - Fecha pedido (cuánto se demoró el proveedor).
        s.formula(r, 16, 'IF(OR(%s="",$O%d=""),"",$O%d-$A%d)' % (B, r, r, r), S_INT_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 7, MOV_LAST, 7, "%s!$R$2:$R$30" % q(CFG))   # medio de pago (billetera o tarjeta)
    s.validate_list(2, 8, MOV_LAST, 8, '"Sí,No"')                  # a crédito
    return s

# ---------- VENTAS ----------
def build_ventas():
    s = Sheet(VEN, freeze_row=1)
    # IMPORTANTE: F=Cantidad y M=clave se mantienen (el Catálogo depende de ellas).
    headers = ["Fecha", "N° Factura", "Cliente", "Código", "Producto", "Cantidad",
               "Ubicación", "Valor Recibido", "Costo Prom. Unit.", "Costo Total",
               "Costos Fijos", "Utilidad Neta", "clave", "Margen %", "Forma de Cobro",
               "Costo Mary (editable)", "Costo Papelería (editable)", "Publicidad (editable)"]
    widths = [12, 12, 22, 12, 28, 10, 16, 15, 15, 14, 13, 14, 14, 10, 18, 16, 18, 16]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 3), "FAC-0001", "Juan Pérez", "P001", 2, CASA, 110000, "Efectivo"),
        (datetime.date(2026, 6, 3), "FAC-0001", "Juan Pérez", "P002", 1, CASA, 22000, "Nequi"),
        (datetime.date(2026, 6, 4), "FAC-0002", "María Gómez", "P003", 3, ML, 45000, "Mercado Libre"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, fac, cli, cod, qy, ub, val, cobro = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, fac, S_INPUT); s.text(r, 3, cli, S_INPUT)
            s.text(r, 4, cod, S_INPUT); s.num(r, 6, qy, S_INPUT_INT); s.text(r, 7, ub, S_INPUT)
            s.num(r, 8, val, S_INPUT_MONEY); s.text(r, 15, cobro, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 3, S_INPUT)
            s.blank(r, 4, S_INPUT); s.blank(r, 6, S_INPUT_INT); s.blank(r, 7, S_INPUT)
            s.blank(r, 8, S_INPUT_MONEY); s.blank(r, 15, S_INPUT)
        # P/Q Costo Mary y Costo Papelería: editables, precargados con el valor de Config
        # (B11/B12) para que no toque escribirlos siempre, pero se pueden cambiar en
        # cualquier fila si esa venta tuvo un costo distinto (ej. envío más caro ese día).
        # R Publicidad: editable, en blanco por defecto (solo se llena si esa venta
        # tuvo costo de publicidad/pauta; muchas ventas no la tienen).
        s.num(r, 16, 1000, S_INPUT_MONEY)
        s.num(r, 17, 1000, S_INPUT_MONEY)
        s.blank(r, 18, S_INPUT_MONEY)
        D = "$D%d" % r
        # E producto
        s.formula(r, 5, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (D, D, q(CAT)), S_TEXT_A)
        # I costo promedio unitario
        s.formula(r, 9, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$F,6,FALSE),0))' % (D, D, q(CAT)), S_MONEY_A)
        # J costo total = cantidad * costo prom
        s.formula(r, 10, 'IF(%s="","",$F%d*$I%d)' % (D, r, r), S_MONEY_A)
        # K Costos Fijos = suma de los 3 costos editables de esta venta (P Mary + Q Papelería + R Publicidad).
        # Antes salía fijo de Config; ahora cada venta puede tener su propio valor.
        s.formula(r, 11, 'IF(%s="","",SUM($P%d:$R%d))' % (D, r, r), S_MONEY_A)
        # L Utilidad NETA = Valor Recibido - Costo Total - Costos fijos
        s.formula(r, 12, 'IF(%s="","",$H%d-$J%d-$K%d)' % (D, r, r, r), S_MONEY_A)
        # M clave (la usa el Catálogo: NO mover)
        s.formula(r, 13, 'IF(%s="","",%s&"|"&$G%d)' % (D, D, r), S_TEXT_A)
        # N Margen % sobre el valor recibido (con utilidad neta)
        s.formula(r, 14, 'IF(%s="","",IF($H%d=0,0,$L%d/$H%d))' % (D, r, r, r), S_PCT_A)
    s.validate_list(2, 3, MOV_LAST, 3, "%s!$B$2:$B$%d" % (q(CLI), CLI_LAST))
    s.validate_list(2, 4, MOV_LAST, 4, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 7, MOV_LAST, 7, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(2, 15, MOV_LAST, 15, "%s!$P$2:$P$15" % q(CFG))   # forma de cobro
    s.text(1, 20, "Costo Mary y Papelería vienen precargados (los de Config) pero puedes cambiarlos "
                  "en cualquier venta. Publicidad se deja vacía y solo la llenas si esa venta tuvo pauta.",
           S_FOOTER_SM)
    return s

# ---------- TRASLADOS ----------
def build_traslados():
    s = Sheet(TRA, freeze_row=1)
    headers = ["Fecha", "Código", "Producto", "Cantidad", "Origen", "Destino", "claveOrigen", "claveDestino"]
    widths = [13, 13, 30, 11, 16, 16, 16, 16]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 3), "P001", 5, CASA, ML),
        (datetime.date(2026, 6, 3), "P003", 10, CASA, ML),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, ori, des = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT); s.num(r, 4, qy, S_INPUT_INT)
            s.text(r, 5, ori, S_INPUT); s.text(r, 6, des, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 4, S_INPUT_INT)
            s.blank(r, 5, S_INPUT); s.blank(r, 6, S_INPUT)
        B = "$B%d" % r
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        s.formula(r, 7, 'IF(%s="","",%s&"|"&$E%d)' % (B, B, r), S_TEXT_A)
        s.formula(r, 8, 'IF(%s="","",%s&"|"&$F%d)' % (B, B, r), S_TEXT_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 5, MOV_LAST, 5, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(2, 6, MOV_LAST, 6, "%s!$F$2:$F$3" % q(CFG))
    return s

# ---------- AJUSTES ----------
def build_ajustes():
    s = Sheet(AJU, freeze_row=1)
    headers = ["Fecha", "Código", "Producto", "Cantidad (+/-)", "Ubicación", "Motivo", "Nota", "clave"]
    widths = [13, 13, 30, 14, 16, 22, 30, 14]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 3), "P004", -1, CASA, "Producto dañado", "Llegó abollado"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, cod, qy, ub, mot, nota = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, cod, S_INPUT); s.num(r, 4, qy, S_SIGNED)
            s.text(r, 5, ub, S_INPUT); s.text(r, 6, mot, S_INPUT); s.text(r, 7, nota, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 4, S_SIGNED)
            s.blank(r, 5, S_INPUT); s.blank(r, 6, S_INPUT); s.blank(r, 7, S_INPUT)
        B = "$B%d" % r
        s.formula(r, 3, 'IF(%s="","",IFERROR(VLOOKUP(%s,%s!$A:$B,2,FALSE),"⚠ Código no existe"))'
                  % (B, B, q(CAT)), S_TEXT_A)
        s.formula(r, 8, 'IF(%s="","",%s&"|"&$E%d)' % (B, B, r), S_TEXT_A)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$A$2:$A$%d" % (q(CAT), CAT_LAST))
    s.validate_list(2, 5, MOV_LAST, 5, "%s!$F$2:$F$3" % q(CFG))
    s.validate_list(2, 6, MOV_LAST, 6, "%s!$J$2:$J$6" % q(CFG))
    return s

# ---------- CLIENTES ----------
def build_clientes():
    s = Sheet(CLI, freeze_row=1)
    headers = ["ID", "Nombre", "Documento / NIT", "Teléfono", "Correo", "Dirección", "Ciudad", "Notas"]
    widths = [10, 24, 18, 16, 26, 28, 14, 26]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        ("Juan Pérez", "1.111.111", "573001112233", "juan@correo.com", "Cra 1 #2-3", "Bogotá", ""),
        ("María Gómez", "2.222.222", "573004445566", "maria@correo.com", "Cll 4 #5-6", "Medellín", ""),
    ]
    for r in range(2, CLI_LAST + 1):
        i = r - 2
        s.formula(r, 1, 'IF($B%d="","","CLI-"&TEXT(ROW()-1,"000"))' % r, S_TEXT_A)
        if i < len(sample):
            nom, doc, tel, cor, dirx, ciu, notas = sample[i]
            s.text(r, 2, nom, S_INPUT); s.text(r, 3, doc, S_INPUT); s.text(r, 4, tel, S_INPUT)
            s.text(r, 5, cor, S_INPUT); s.text(r, 6, dirx, S_INPUT); s.text(r, 7, ciu, S_INPUT)
            s.blank(r, 8, S_INPUT)
        else:
            for c in range(2, 9):
                s.blank(r, c, S_INPUT)
    return s

# ---------- INVENTARIO (reporte) ----------
def build_inventario():
    s = Sheet(INV, freeze_row=1)
    headers = ["Código", "Producto", "Categoría", "Costo Promedio", "Stock Casa",
               "Stock Mercado Libre", "Stock Total", "Valor Inventario", "Estado"]
    widths = [12, 30, 18, 15, 12, 18, 12, 16, 14]
    srccols = [1, 2, 3, 6, 7, 8, 9, 10, 11]
    styles = [S_TEXT_A, S_TEXT_A, S_TEXT_A, S_MONEY_A, S_INT_A, S_INT_A, S_INT_A, S_MONEY_A, S_TEXT_A]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    for r in range(2, CAT_LAST + 1):
        for j, sc in enumerate(srccols):
            col = col_letter(sc)
            s.formula(r, j + 1, 'IF(%s!$A%d="","",%s!$%s%d)' % (q(CAT), r, q(CAT), col, r), styles[j])
    return s

# ---------- DASHBOARD ----------
def build_dashboard():
    s = Sheet(DASH, hide_gridlines=True, freeze_row=2)
    for c in range(1, 13):
        s.colw(c, 13)
    # Título
    s.text(1, 1, "DASHBOARD EJECUTIVO", S_TITLE); s.merge(1, 1, 1, 12); s.rowh(1, 40)
    s.formula(2, 1, '"Mes en curso: "&TEXT(TODAY(),"mmmm yyyy")&"   ·   Actualizado: "&TEXT(NOW(),"dd/mm/yyyy hh:mm")', S_SUBTITLE)
    s.merge(2, 1, 2, 12); s.rowh(2, 22)

    ini = 'EOMONTH(TODAY(),-1)+1'
    fin = 'EOMONTH(TODAY(),0)'
    df = 'SUMIFS(%s!$H:$H,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), ini, q(VEN), fin)
    gf = 'SUMIFS(%s!$L:$L,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), ini, q(VEN), fin)
    cf = 'SUMIFS(%s!$F:$F,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(COMP), q(COMP), ini, q(COMP), fin)

    def card(row, col, label, formula, lab_style, val_style):
        s.text(row, col, label, lab_style); s.merge(row, col, row, col + 2)
        s.formula(row + 1, col, formula, val_style); s.merge(row + 1, col, row + 1, col + 2)
        s.rowh(row, 22); s.rowh(row + 1, 38)

    card(4, 1, "VENTAS DEL MES", df, S_KPI_LAB_SEC, S_KPI_MONEY)
    card(4, 4, "GANANCIA DEL MES", gf, S_KPI_LAB_ACC, S_KPI_MONEY)
    card(4, 7, "COMPRAS DEL MES", cf, S_KPI_LAB_ADV, S_KPI_MONEY)
    card(4, 10, "VALOR DEL INVENTARIO", 'SUM(%s!$J$2:$J$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_PRI, S_KPI_MONEY)
    card(7, 1, "INVENTARIO EN CASA (UNID)", 'SUM(%s!$G$2:$G$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_DARK, S_KPI_INT)
    card(7, 4, "INVENTARIO EN ML (UNID)", 'SUM(%s!$H$2:$H$%d)' % (q(CAT), CAT_LAST), S_KPI_LAB_DARK, S_KPI_INT)
    card(7, 7, "PRODUCTOS CON STOCK BAJO", 'COUNTIF(%s!$K$2:$K$%d,"Stock bajo")' % (q(CAT), CAT_LAST), S_KPI_LAB_ADV, S_KPI_INT)
    card(7, 10, "PRODUCTOS AGOTADOS", 'COUNTIF(%s!$K$2:$K$%d,"Agotado")' % (q(CAT), CAT_LAST), S_KPI_LAB_ALE, S_KPI_INT)

    # Tabla productos más vendidos (top 8) usando columnas auxiliares ocultas (P..R)
    # P: nombre producto, Q: unidades vendidas, R: unidades ajustadas (desempate)
    auxcol = 16  # P
    for r in range(2, CAT_LAST + 1):
        s.formula(r, auxcol, 'IF(%s!$A%d="","",%s!$B%d)' % (q(CAT), r, q(CAT), r))
        s.formula(r, auxcol + 1, 'IF(%s!$A%d="","",SUMIF(%s!$D:$D,%s!$A%d,%s!$F:$F))'
                  % (q(CAT), r, q(VEN), q(CAT), r, q(VEN)))
        s.formula(r, auxcol + 2, 'IF(%s!$A%d="","",$Q%d+ROW()/100000)' % (q(CAT), r, r))
    s.colw(auxcol, 8); s.colw(auxcol + 1, 8); s.colw(auxcol + 2, 8)

    s.text(11, 1, "🏆 PRODUCTOS MÁS VENDIDOS", S_BANNER_PRIM); s.merge(11, 1, 11, 5)
    s.text(12, 1, "Producto", S_HDR); s.merge(12, 1, 12, 4)
    s.text(12, 5, "Unid.", S_HDR)
    rng_adj = "$R$2:$R$%d" % CAT_LAST
    rng_nom = "$P$2:$P$%d" % CAT_LAST
    for k in range(1, 9):
        rr = 12 + k
        large = 'LARGE(%s,%d)' % (rng_adj, k)
        s.formula(rr, 1, 'IFERROR(IF(INT(%s)=0,"",INDEX(%s,MATCH(%s,%s,0))),"")'
                  % (large, rng_nom, large, rng_adj), S_TEXT); s.merge(rr, 1, rr, 4)
        s.formula(rr, 5, 'IFERROR(IF(INT(%s)=0,"",INT(%s)),"")' % (large, large), S_INT)

    # Tabla alertas de inventario
    s.text(11, 7, "⚠️ ALERTAS DE INVENTARIO (bajo / agotado)", S_BANNER_ALERT); s.merge(11, 7, 11, 12)
    s.text(12, 7, "Producto", S_HDR); s.merge(12, 7, 12, 9)
    s.text(12, 10, "Casa", S_HDR); s.text(12, 11, "ML", S_HDR); s.text(12, 12, "Total", S_HDR)
    # listamos los primeros productos con estado de alerta (hasta 12)
    for k in range(1, 13):
        rr = 12 + k
        cat_row = k + 1  # mapea a filas 2.. del catálogo (vista simple de primeros productos)
        cond = '%s!$K%d="Agotado",%s!$K%d="Stock bajo"' % (q(CAT), cat_row, q(CAT), cat_row)
        s.formula(rr, 7, 'IF(OR(%s),%s!$B%d,"")' % (cond, q(CAT), cat_row), S_TEXT); s.merge(rr, 7, rr, 9)
        s.formula(rr, 10, 'IF(OR(%s),%s!$G%d,"")' % (cond, q(CAT), cat_row), S_INT)
        s.formula(rr, 11, 'IF(OR(%s),%s!$H%d,"")' % (cond, q(CAT), cat_row), S_INT)
        s.formula(rr, 12, 'IF(OR(%s),%s!$I%d,"")' % (cond, q(CAT), cat_row), S_INT)

    # Tabla de tendencias (6 meses) - sirve para que el usuario inserte un gráfico
    s.text(26, 1, "📈 TENDENCIAS (últimos 6 meses) — selecciona A27:D33 e inserta un gráfico", S_BANNER_SEC)
    s.merge(26, 1, 26, 12)
    s.text(27, 1, "Mes", S_HDR); s.text(27, 2, "Ventas", S_HDR); s.text(27, 3, "Utilidad", S_HDR); s.text(27, 4, "Compras", S_HDR)
    for i in range(6):
        rr = 28 + i
        off = -5 + i
        mini = 'DATE(YEAR(EOMONTH(TODAY(),%d)),MONTH(EOMONTH(TODAY(),%d)),1)' % (off, off)
        mfin = 'EOMONTH(TODAY(),%d)' % off
        s.formula(rr, 1, 'TEXT(%s,"mmm yy")' % mini, S_TEXT)
        s.formula(rr, 2, 'SUMIFS(%s!$H:$H,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), mini, q(VEN), mfin), S_MONEY)
        s.formula(rr, 3, 'SUMIFS(%s!$L:$L,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(VEN), q(VEN), mini, q(VEN), mfin), S_MONEY)
        s.formula(rr, 4, 'SUMIFS(%s!$F:$F,%s!$A:$A,">="&%s,%s!$A:$A,"<="&%s)' % (q(COMP), q(COMP), mini, q(COMP), mfin), S_MONEY)

    # ----- Bloque: PAGOS DE COMPRAS A CRÉDITO / TARJETAS -----
    # KPIs de deuda (usan la función card local: label, formula, estilo_etiqueta, estilo_valor).
    card(35, 1, "DEUDA TOTAL EN TARJETAS",
         "SUM('%s'!$D$3:$D$10)" % TARJ, S_KPI_LAB_ALE, S_KPI_MONEY)
    card(35, 4, "PEDIDOS POR PAGAR",
         'COUNTIF(%s!$M:$M,"Pendiente")+COUNTIF(%s!$M:$M,"Abono parcial")'
         % (q(COMP), q(COMP)), S_KPI_LAB_ADV, S_KPI_INT)
    card(35, 7, "PAGOS VENCIDOS",
         'SUMPRODUCT((%s!$N$2:$N$%d<>"")*(LEFT(%s!$N$2:$N$%d,4)="🔴"))'
         % (q(COMP), MOV_LAST, q(COMP), MOV_LAST), S_KPI_LAB_ALE, S_KPI_INT)

    # Tabla con los pedidos a crédito pendientes y su alerta.
    s.text(38, 1, "💳 PAGOS DE TARJETA / CRÉDITO PENDIENTES", S_BANNER_ALERT); s.merge(38, 1, 38, 12)
    s.text(39, 1, "Fecha", S_HDR); s.text(39, 2, "Producto", S_HDR); s.merge(39, 2, 39, 4)
    s.text(39, 5, "Medio de pago", S_HDR); s.merge(39, 5, 39, 6)
    s.text(39, 7, "Saldo", S_HDR)
    s.text(39, 8, "Límite pago", S_HDR)
    s.text(39, 9, "Aviso", S_HDR); s.merge(39, 9, 39, 12)
    # Listamos los primeros pedidos pendientes a crédito (vista de los primeros registros).
    for k in range(1, 16):
        rr = 39 + k
        cr = k + 1  # fila de Compras
        cond = '%s!$H%d="Sí",%s!$L%d>0' % (q(COMP), cr, q(COMP), cr)
        s.formula(rr, 1, 'IF(AND(%s),%s!$A%d,"")' % (cond, q(COMP), cr), S_DATE)
        s.formula(rr, 2, 'IF(AND(%s),%s!$C%d,"")' % (cond, q(COMP), cr), S_TEXT); s.merge(rr, 2, rr, 4)
        s.formula(rr, 5, 'IF(AND(%s),%s!$G%d,"")' % (cond, q(COMP), cr), S_TEXT); s.merge(rr, 5, rr, 6)
        s.formula(rr, 7, 'IF(AND(%s),%s!$L%d,"")' % (cond, q(COMP), cr), S_MONEY)
        s.formula(rr, 8, 'IF(AND(%s),%s!$J%d,"")' % (cond, q(COMP), cr), S_DATE)
        s.formula(rr, 9, 'IF(AND(%s),%s!$N%d,"")' % (cond, q(COMP), cr), S_TEXT); s.merge(rr, 9, rr, 12)

    # ----- Bloque: PRÉSTAMOS A PERSONAS (dinero prestado) -----
    s.text(56, 1, "🤝 PRÉSTAMOS A PERSONAS (dinero de la empresa prestado)", S_BANNER_PRIM); s.merge(56, 1, 56, 12)
    card(57, 1, "PRESTADO ACTUALMENTE", 'SUM(%s!$G:$G)' % q(PREST), S_KPI_LAB_ALE, S_KPI_MONEY)
    card(57, 4, "PERSONAS QUE ME DEBEN",
         'COUNTIF(%s!$H:$H,"Pendiente")+COUNTIF(%s!$H:$H,"Abono parcial")' % (q(PREST), q(PREST)),
         S_KPI_LAB_ADV, S_KPI_INT)
    card(57, 7, "DÍAS DEL MÁS ANTIGUO SIN PAGAR", 'IFERROR(MAX(%s!$I:$I),0)' % q(PREST),
         S_KPI_LAB_DARK, S_KPI_INT)
    s.text(60, 1, "📋 QUIÉN ME DEBE — del más antiguo primero (empieza a cobrar por arriba)", S_BANNER_ALERT)
    s.merge(60, 1, 60, 12)
    s.formula(61, 1,
              'IFERROR(QUERY(%s!$A$2:$J$2000,'
              '"select B,C,E,F,G,I where G>0 order by I desc '
              "label B 'Persona', C 'Teléfono', E 'Prestado', F 'Abonado', G 'Saldo', I 'Días'\",0),"
              '"Nadie te debe 🎉")' % q(PREST), S_TEXT)
    return s
def build_catalogo_venta():
    """Hoja para generar catálogos comerciales en PDF, con foto del producto,
    filtrados por categoría (o TODAS) y mostrando el precio que elijas (Detal o Mayor).
    Los productos salen ordenados por categoría y por tamaño (medidas)."""
    s = Sheet(CVENTA, hide_gridlines=True)
    NCAT = 200  # cuántos productos puede mostrar el catálogo

    # Anchos: A=imagen, B=nombre, C=descr/medidas, D=precio, E=disponibles Casa
    s.colw(1, 22); s.colw(2, 30); s.colw(3, 30); s.colw(4, 20); s.colw(5, 16)

    # Título (nombre de empresa) y subtítulo dinámico.
    s.formula(1, 1, "%s!$B$2" % q(CFG), S_CV_TITLE); s.merge(1, 1, 1, 5); s.rowh(1, 44)
    s.formula(2, 1, '"Catálogo de productos · "&$B$3&" · Precio "&$D$3', S_CV_SUB)
    s.merge(2, 1, 2, 5); s.rowh(2, 24)

    # Controles (lo que el usuario elige): fila 3.
    s.text(3, 1, "Categoría:", S_LABEL_RIGHT)
    s.text(3, 2, "TODAS", S_INPUT)
    s.text(3, 3, "Tipo de precio:", S_LABEL_RIGHT)
    s.text(3, 4, "Detal", S_INPUT)
    s.rowh(3, 22)
    # La lista de categorías incluye "TODAS" + todas las del negocio (Config col L).
    s.validate_list(3, 2, 3, 2, "%s!$L$2:$L$30" % q(CFG))
    s.validate_list(3, 4, 3, 4, '"Detal,Mayor"')

    # Encabezado de la tabla del catálogo.
    s.text(5, 1, "Imagen", S_HDR); s.text(5, 2, "Producto", S_HDR)
    s.text(5, 3, "Medidas / Código", S_HDR); s.text(5, 4, "Precio", S_HDR)
    s.text(5, 5, "Disp. Casa (und)", S_HDR)
    s.rowh(5, 24)

    # MÉTODO ROBUSTO (sin huecos): una sola fórmula FILTER+SORT trae TODOS los
    # productos activos de la categoría elegida (o TODAS), ordenados por
    # categoría -> medidas -> nombre. Se vuelca en columnas auxiliares ocultas
    # (H..Y) y las columnas visibles A..E solo la muestran.
    # Además de excluir "Inactivo", exige Stock Casa > 0 (Catálogo!G): así un
    # producto que aún no tiene ninguna Compra registrada, o que ya se agotó,
    # desaparece solo del catálogo comercial sin que tengas que tocar el campo Estado.
    NSHOW = 250
    filt = (
        'IFERROR(SORT(FILTER(%s!$A$2:$R$2000,'
        '(%s!$A$2:$A$2000<>"")*'
        '(%s!$D$2:$D$2000<>"Inactivo")*'
        '(%s!$G$2:$G$2000>0)*'
        '(($B$3="TODAS")+(%s!$C$2:$C$2000=$B$3))),3,TRUE,18,TRUE,2,TRUE),"")'
        % (q(CAT), q(CAT), q(CAT), q(CAT), q(CAT))
    )
    s.formula(6, 8, filt, S_CV_CODE)   # H6: la "máquina" que trae y ordena los productos
    for c in range(8, 26):             # ocultar columnas auxiliares H..Y
        s.hide_col(c, 10)

    # Columnas auxiliares (resultado de FILTER que empieza en H=col8):
    #   H=Código, I=Nombre, N=Stock Casa, T=URL Imagen, U=Precio Detal, V=Precio Mayor, Y=Medidas
    first = 6
    for i in range(NSHOW):
        r = first + i
        s.formula(r, 1, 'IF($H%d="","",IFERROR(IMAGE($T%d,1),""))' % (r, r), S_CV_IMG)
        s.formula(r, 2, 'IF($H%d="","",$I%d)' % (r, r), S_CV_NAME)
        s.formula(r, 3,
                  'IF($H%d="","",IF($Y%d="","Cód: "&$H%d,"Medidas: "&$Y%d&"  ·  Cód: "&$H%d))'
                  % (r, r, r, r, r), S_CV_CODE)
        s.formula(r, 4, 'IF($H%d="","",IF($D$3="Mayor",$V%d,$U%d))' % (r, r, r), S_CV_PRICE)
        # E: unidades disponibles en CASA (lo que tú tienes para vender; ML no se cuenta).
        s.formula(r, 5, 'IF($H%d="","",$N%d)' % (r, r), S_INT)
        s.rowh(r, 90)

    return s


# ---------- CIERRE DIARIO / CUADRE DE CAJA ----------
def build_cierre_diario():
    s = Sheet(CIERRE, hide_gridlines=True)
    s.colw(1, 38); s.colw(2, 20); s.colw(3, 46)
    F = "'%s'!$A:$A" % VEN      # fecha ventas
    H = "'%s'!$H:$H" % VEN      # valor recibido
    K = "'%s'!$K:$K" % VEN      # costos fijos
    L = "'%s'!$L:$L" % VEN      # utilidad neta
    O = "'%s'!$O:$O" % VEN      # forma de cobro
    fecha = "$B$3"

    def sumcobro(metodo):
        return 'SUMIFS(%s,%s,%s,%s,"%s")' % (H, F, fecha, O, metodo)

    # Título y fecha
    s.text(1, 1, "CIERRE DIARIO · CUADRE DE CAJA", S_TITLE); s.merge(1, 1, 1, 3); s.rowh(1, 40)
    s.text(3, 1, "Fecha del cierre:", S_LABEL_RIGHT)
    s.formula(3, 2, "TODAY()", S_INPUT_DATE)
    s.text(3, 3, "← cámbiala para ver otro día", S_FOOTER_SM)

    # INGRESOS (entran a caja / cuentas)
    s.text(5, 1, "INGRESOS DEL DÍA (entran a tu caja / cuentas)", S_BANNER_SEC); s.merge(5, 1, 5, 3)
    metodos = ["Efectivo", "Nequi", "Bancolombia", "Bancolombia MM", "Daviplata", "Davivienda", "Skydrops"]
    row = 6
    for m in metodos:
        s.text(row, 1, m, S_LABEL); s.formula(row, 2, sumcobro(m), S_MONEY)
        row += 1
    # row = 13
    s.text(13, 1, "Subtotal ventas a caja (sin MELI)", S_LABEL); s.formula(13, 2, "SUM($B$6:$B$12)", S_MONEY)
    s.text(14, 1, "+ Abonos de préstamos recibidos hoy", S_LABEL)
    s.formula(14, 2, "SUMIFS('%s'!$C:$C,'%s'!$A:$A,%s)" % (ABONOS, ABONOS, fecha), S_MONEY)
    s.text(14, 3, "Plata que te devolvieron hoy de un préstamo.", S_FOOTER_SM)
    s.text(15, 1, "= TOTAL INGRESOS A CAJA", S_TOTLAB); s.formula(15, 2, "$B$13+$B$14", S_TOTVAL)

    # MELI aparte
    s.text(17, 1, "APARTE — Mercado Libre (NO entra a tu caja)", S_BANNER_ALERT); s.merge(17, 1, 17, 3)
    s.text(18, 1, "Mercado Libre", S_LABEL); s.formula(18, 2, sumcobro("Mercado Libre"), S_MONEY)
    s.text(18, 3, "MELI te paga directo a tu cuenta; no se mezcla con la caja del día.", S_FOOTER_SM)

    # SALIDAS del día (sale de caja)
    s.text(20, 1, "SALIDAS DEL DÍA (sale de caja)", S_BANNER_SEC); s.merge(20, 1, 20, 3)
    s.text(21, 1, "Costos fijos por ventas (Mary + bolsa/etiqueta)", S_LABEL)
    s.formula(21, 2, 'SUMIFS(%s,%s,%s)' % (K, F, fecha), S_MONEY)
    s.text(22, 1, "Préstamos entregados hoy", S_LABEL)
    s.formula(22, 2, "SUMIFS('%s'!$E:$E,'%s'!$A:$A,%s)" % (PREST, PREST, fecha), S_MONEY)
    s.text(22, 3, "Plata que prestaste hoy (sale de la caja).", S_FOOTER_SM)
    s.text(23, 1, "Otro gasto 1 (escribe descripción →)", S_LABEL); s.blank(23, 2, S_INPUT_MONEY)
    s.text(24, 1, "Otro gasto 2", S_LABEL); s.blank(24, 2, S_INPUT_MONEY)
    s.text(25, 1, "Otro gasto 3", S_LABEL); s.blank(25, 2, S_INPUT_MONEY)
    s.text(26, 1, "= TOTAL SALIDAS DEL DÍA", S_TOTLAB)
    s.formula(26, 2, "$B$21+$B$22+SUM($B$23:$B$25)", S_TOTVAL)

    # RESUMEN
    s.text(28, 1, "RESUMEN DEL DÍA", S_BANNER_PRIM); s.merge(28, 1, 28, 3)
    s.text(29, 1, "Total ingresos a caja", S_LABEL); s.formula(29, 2, "$B$15", S_MONEY)
    s.text(30, 1, "Total salidas de caja", S_LABEL); s.formula(30, 2, "$B$26", S_MONEY)
    s.text(31, 1, "MOVIMIENTO NETO DE CAJA HOY", S_TOTLAB); s.formula(31, 2, "$B$15-$B$26", S_TOTVAL)
    s.text(32, 1, "Ganancia neta del día (solo ventas)", S_LABEL)
    s.formula(32, 2, 'SUMIFS(%s,%s,%s)' % (L, F, fecha), S_MONEY)
    s.text(32, 3, "Los préstamos NO son ganancia ni gasto: solo mueven la caja.", S_FOOTER_SM)
    s.text(33, 1, "Ventas por Mercado Libre (aparte)", S_LABEL); s.formula(33, 2, "$B$18", S_MONEY)

    # CUADRE DE EFECTIVO físico
    s.text(35, 1, "CUADRE DE EFECTIVO (conteo físico de caja)", S_BANNER_SEC); s.merge(35, 1, 35, 3)
    s.text(36, 1, "Saldo inicial de caja (efectivo de ayer) →", S_LABEL); s.blank(36, 2, S_INPUT_MONEY)
    s.text(37, 1, "+ Efectivo de ventas hoy", S_LABEL); s.formula(37, 2, "$B$6", S_MONEY)
    s.text(38, 1, "+ Abonos recibidos en efectivo hoy (automático)", S_LABEL)
    s.formula(38, 2, "SUMIFS('%s'!$C:$C,'%s'!$A:$A,%s,'%s'!$D:$D,\"Efectivo\")" % (ABONOS, ABONOS, fecha, ABONOS), S_MONEY)
    s.text(38, 3, "Sale solo de los abonos marcados como 'Efectivo'.", S_FOOTER_SM)
    # +/- Transferencias internas que hoy movieron plata HACIA o DESDE Efectivo
    # (hoja Movimientos: retiros de banco a caja, aportes de capital en efectivo, etc.).
    # Así el cuadre no se descuadra cuando metes o sacas plata de la caja sin que sea venta ni gasto.
    s.text(39, 1, "+ Transferencias/Aportes que entraron a Efectivo hoy (automático)", S_LABEL)
    s.formula(39, 2, "SUMIFS('%s'!$E:$E,'%s'!$A:$A,%s,'%s'!$D:$D,\"Efectivo\")" % (MOV, MOV, fecha, MOV), S_MONEY)
    s.text(39, 3, "Ej: retiro del banco a caja, o aporte de capital recibido en efectivo.", S_FOOTER_SM)
    s.text(40, 1, "− Transferencias/Gastos que salieron de Efectivo hoy (automático)", S_LABEL)
    s.formula(40, 2, "SUMIFS('%s'!$E:$E,'%s'!$A:$A,%s,'%s'!$C:$C,\"Efectivo\")" % (MOV, MOV, fecha, MOV), S_MONEY)
    s.text(40, 3, "Ej: pasaste efectivo a Nequi, o registraste un gasto con Origen = Efectivo.", S_FOOTER_SM)
    s.text(41, 1, "− Salidas en efectivo (gastos / préstamos dados) →", S_LABEL); s.blank(41, 2, S_INPUT_MONEY)
    s.text(41, 3, "Solo para gastos en efectivo que NO registraste en Movimientos.", S_FOOTER_SM)
    s.text(42, 1, "= EFECTIVO QUE DEBERÍA HABER", S_TOTLAB)
    s.formula(42, 2, "$B$36+$B$37+$B$38+$B$39-$B$40-$B$41", S_TOTVAL)
    s.text(43, 1, "Efectivo realmente contado en caja →", S_LABEL); s.blank(43, 2, S_INPUT_MONEY)
    s.text(44, 1, "DIFERENCIA (contado − esperado)", S_TOTLAB)
    s.formula(44, 2, "$B$43-$B$42", S_TOTVAL)
    s.formula(44, 3, 'IF($B$43="","Cuenta el efectivo y escríbelo arriba",'
              'IF(ROUND($B$44,0)=0,"✅ CUADRA",IF($B$44>0,"🔵 Sobra efectivo","🔴 Falta efectivo")))', S_TEXT)

    # HISTORIAL: los últimos 30 días (para revisar cualquier día pasado de un vistazo).
    s.text(46, 1, "HISTORIAL — últimos 30 días", S_BANNER_PRIM); s.merge(46, 1, 46, 3)
    s.text(47, 1, "Día", S_HDR); s.text(47, 2, "Ventas a caja", S_HDR); s.text(47, 3, "Ganancia neta", S_HDR)
    for k in range(0, 30):
        rr = 48 + k
        dia = "TODAY()-%d" % k
        s.formula(rr, 1, dia, S_DATE)
        # ventas a caja del día (sin MELI) = ventas - lo de Mercado Libre
        s.formula(rr, 2,
                  'SUMIFS(%s,%s,$A%d)-SUMIFS(%s,%s,$A%d,%s,"Mercado Libre")'
                  % (H, F, rr, H, F, rr, O), S_MONEY)
        s.formula(rr, 3, 'SUMIFS(%s,%s,$A%d)' % (L, F, rr), S_MONEY)
    return s


# ---------- PRÉSTAMOS A PERSONAS ----------
def build_prestamos():
    s = Sheet(PREST, freeze_row=1)
    headers = ["Fecha", "Persona", "Teléfono", "Motivo / Nota", "Valor Prestado",
               "Abonado", "Saldo", "Estado", "Días sin pagar", "Alerta", "Billetera origen"]
    widths = [12, 22, 14, 28, 15, 14, 14, 15, 13, 26, 16]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    # fecha, persona, tel, motivo, valor, billetera (el Abonado se calcula desde "Abonos Préstamos")
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 5, 10), "Pedro Gómez", "3001112233", "Urgencia médica", 200000, "Nequi"),
        (datetime.date(2026, 6, 1), "Ana Ruiz", "3015556677", "Préstamo personal", 100000, "Efectivo"),
        (datetime.date(2026, 5, 20), "Luis Mar", "3024445566", "Imprevisto", 150000, "Bancolombia"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, per, tel, mot, val, bill = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, per, S_INPUT); s.text(r, 3, tel, S_INPUT)
            s.text(r, 4, mot, S_INPUT); s.num(r, 5, val, S_INPUT_MONEY); s.text(r, 11, bill, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT); s.blank(r, 3, S_INPUT)
            s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT_MONEY); s.blank(r, 11, S_INPUT)
        B = "$B%d" % r
        # F Abonado = suma de todos los abonos registrados para esa persona (hoja Abonos Préstamos)
        s.formula(r, 6, 'IF(%s="","",SUMIF(\'%s\'!$B:$B,$B%d,\'%s\'!$C:$C))'
                  % (B, ABONOS, r, ABONOS), S_MONEY_A)
        # G Saldo = Valor - Abonado (lo que la persona aún te debe)
        s.formula(r, 7, 'IF(%s="","",MAX($E%d-$F%d,0))' % (B, r, r), S_MONEY_A)
        # H Estado
        s.formula(r, 8,
                  'IF(%s="","",IF($G%d<=0,"Pagado",IF($F%d>0,"Abono parcial","Pendiente")))'
                  % (B, r, r), S_TEXT_A)
        # I Días sin pagar (solo si aún debe)
        s.formula(r, 9, 'IF(OR(%s="",$G%d<=0),"",TODAY()-$A%d)' % (B, r, r), S_INT_A)
        # J Alerta por antigüedad
        s.formula(r, 10,
                  'IF(%s="","",IF($G%d<=0,"✅ Pagado",'
                  'IF($I%d>30,"🔴 Cóbrale ya ("&$I%d&" días)",'
                  'IF($I%d>15,"🟠 Lleva "&$I%d&" días",'
                  '"🟢 Reciente ("&$I%d&" días)"))))'
                  % (B, r, r, r, r, r, r), S_TEXT_A)
    s.validate_list(2, 11, MOV_LAST, 11, "%s!$P$2:$P$15" % q(CFG))   # billetera origen
    return s


# ---------- ABONOS DE PRÉSTAMOS (pagos que te hacen, con fecha) ----------
def build_abonos():
    s = Sheet(ABONOS, freeze_row=1)
    headers = ["Fecha", "Persona", "Valor Abonado", "Forma de Cobro", "Nota"]
    widths = [13, 24, 16, 18, 30]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 5), "Pedro Gómez", 50000, "Efectivo", "Primer abono"),
        (datetime.date(2026, 6, 10), "Ana Ruiz", 100000, "Nequi", "Pagó todo"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, per, val, forma, nota = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, per, S_INPUT)
            s.num(r, 3, val, S_INPUT_MONEY); s.text(r, 4, forma, S_INPUT); s.text(r, 5, nota, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT)
            s.blank(r, 3, S_INPUT_MONEY); s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT)
    # La persona se elige de la lista de la hoja Préstamos; la forma de cobro de Config.
    s.validate_list(2, 2, MOV_LAST, 2, "'%s'!$B$2:$B$%d" % (PREST, MOV_LAST))
    s.validate_list(2, 4, MOV_LAST, 4, "%s!$P$2:$P$15" % q(CFG))
    s.text(1, 7, "Cada vez que te devuelven plata de un préstamo, anótalo aquí (con su fecha y cómo te pagaron).", S_FOOTER_SM)
    return s


# ---------- TESORERÍA: Billeteras, Tarjetas, Movimientos, Pagos Tarjeta ----------
_BILLETERAS = ["Efectivo", "Nequi", "Bancolombia", "Bancolombia MM", "Daviplata",
               "Davivienda", "Mercado Libre", "Skydrops"]
_TARJETAS = ["TC Bancolombia 1 (1517)", "TC Bancolombia 2 (0554)", "TC Bancolombia 3 (4582)",
             "TC Bancolombia 4 (8366)", "TC Bancolombia 5 (0623)", "TC Bancolombia MM 6 (6423)",
             "TC Bancolombia MM 7 (2459)", "TC NU 8 (3750)"]


def build_movimientos():
    """Transferencias entre billeteras, retiros de MELI/Skydrops, gastos y aportes."""
    s = Sheet(MOV, freeze_row=1)
    headers = ["Fecha", "Tipo", "Billetera Origen", "Billetera Destino", "Valor", "Nota"]
    widths = [13, 18, 18, 18, 15, 34]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 8), "Retiro plataforma", "Mercado Libre", "Bancolombia", 45000, "Retiro de MELI a Bancolombia"),
        (datetime.date(2026, 6, 9), "Transferencia", "Bancolombia", "Nequi", 100000, "Paso plata para pagar"),
        (datetime.date(2026, 6, 9), "Gasto", "Efectivo", "", 15000, "Almuerzo / varios"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, tipo, ori, des, val, nota = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, tipo, S_INPUT); s.text(r, 3, ori, S_INPUT)
            if des:
                s.text(r, 4, des, S_INPUT)
            else:
                s.blank(r, 4, S_INPUT)
            s.num(r, 5, val, S_INPUT_MONEY); s.text(r, 6, nota, S_INPUT)
        else:
            for c in (1,): s.blank(r, c, S_INPUT_DATE)
            for c in (2, 3, 4, 6): s.blank(r, c, S_INPUT)
            s.blank(r, 5, S_INPUT_MONEY)
    s.validate_list(2, 2, MOV_LAST, 2, '"Transferencia,Retiro plataforma,Gasto,Aporte,Otro"')
    s.validate_list(2, 3, MOV_LAST, 3, "%s!$P$2:$P$15" % q(CFG))   # origen
    s.validate_list(2, 4, MOV_LAST, 4, "%s!$P$2:$P$15" % q(CFG))   # destino
    s.text(1, 8, "Origen = de dónde sale la plata · Destino = a dónde llega. Para un gasto, deja el Destino vacío.", S_FOOTER_SM)
    return s


def build_pagos_tarjeta():
    """Pagos / abonos a las tarjetas de crédito, indicando de qué billetera salió."""
    s = Sheet(PAGOST, freeze_row=1)
    headers = ["Fecha", "Tarjeta", "Valor", "Billetera (de dónde salió)", "Nota"]
    widths = [13, 26, 15, 24, 30]
    for i, h in enumerate(headers):
        s.text(1, i + 1, h, S_HDR); s.colw(i + 1, widths[i])
    sample = [] if not INCLUDE_SAMPLES else [
        (datetime.date(2026, 6, 16), "TC Bancolombia 1 (1517)", 100000, "Bancolombia", "Abono a la deuda"),
    ]
    for r in range(2, MOV_LAST + 1):
        i = r - 2
        if i < len(sample):
            d, tar, val, bill, nota = sample[i]
            s.date(r, 1, d, S_INPUT_DATE); s.text(r, 2, tar, S_INPUT)
            s.num(r, 3, val, S_INPUT_MONEY); s.text(r, 4, bill, S_INPUT); s.text(r, 5, nota, S_INPUT)
        else:
            s.blank(r, 1, S_INPUT_DATE); s.blank(r, 2, S_INPUT)
            s.blank(r, 3, S_INPUT_MONEY); s.blank(r, 4, S_INPUT); s.blank(r, 5, S_INPUT)
    s.validate_list(2, 2, MOV_LAST, 2, "%s!$N$2:$N$15" % q(CFG))   # tarjeta
    s.validate_list(2, 4, MOV_LAST, 4, "%s!$P$2:$P$15" % q(CFG))   # billetera
    return s


def build_billeteras():
    """Saldo en tiempo real de cada billetera (todo automático)."""
    s = Sheet(BILL, freeze_row=2)
    s.colw(1, 22); s.colw(2, 18); s.colw(3, 18); s.colw(4, 46)
    s.text(1, 1, "SALDOS POR BILLETERA (en tiempo real)", S_TITLE); s.merge(1, 1, 1, 4); s.rowh(1, 36)
    s.text(2, 1, "Billetera", S_HDR); s.text(2, 2, "Saldo Inicial", S_HDR)
    s.text(2, 3, "Saldo Actual", S_HDR); s.text(2, 4, "Nota", S_HDR)
    notas = {"Mercado Libre": "Por cobrar — retíralo a una billetera (hoja Movimientos)",
             "Skydrops": "Por cobrar — retíralo a una billetera (hoja Movimientos)"}
    r = 3
    for w in _BILLETERAS:
        A = "$A%d" % r
        bal = ("$B%d"
               "+SUMIF('%s'!$O:$O,%s,'%s'!$H:$H)"
               "+SUMIF('%s'!$D:$D,%s,'%s'!$C:$C)"
               "+SUMIF('%s'!$D:$D,%s,'%s'!$E:$E)"
               "-SUMIF('%s'!$C:$C,%s,'%s'!$E:$E)"
               "-SUMIF('%s'!$K:$K,%s,'%s'!$E:$E)"
               "-SUMIF('%s'!$D:$D,%s,'%s'!$C:$C)"
               "-SUMIFS('%s'!$F:$F,'%s'!$H:$H,\"No\",'%s'!$G:$G,%s)"
               ) % (r, VEN, A, VEN, ABONOS, A, ABONOS, MOV, A, MOV, MOV, A, MOV,
                    PREST, A, PREST, PAGOST, A, PAGOST, COMP, COMP, COMP, A)
        s.text(r, 1, w, S_LABEL)
        s.blank(r, 2, S_INPUT_MONEY)          # Saldo inicial (lo escribes tú al empezar)
        s.formula(r, 3, bal, S_TOTVAL)        # Saldo actual = inicial + movimientos
        s.text(r, 4, notas.get(w, ""), S_FOOTER_SM)
        r += 1
    s.text(r, 1, "TOTAL en billeteras", S_TOTLAB)
    s.formula(r, 3, "SUM($C$3:$C$%d)" % (r - 1), S_TOTVAL)
    s.text(r + 2, 1, "Escribe el Saldo Inicial UNA vez (lo que tienes hoy). De ahí en adelante el Saldo Actual se mueve solo: ventas, abonos, transferencias, préstamos, pagos de tarjeta y compras de contado.", S_FOOTER_SM)
    s.merge(r + 2, 1, r + 2, 4)
    return s


def build_tarjetas():
    """Deuda actual de cada tarjeta de crédito (todo automático)."""
    s = Sheet(TARJ, freeze_row=2)
    s.colw(1, 28); s.colw(2, 18); s.colw(3, 16); s.colw(4, 16)
    s.text(1, 1, "DEUDA POR TARJETA DE CRÉDITO", S_TITLE); s.merge(1, 1, 1, 4); s.rowh(1, 36)
    s.text(2, 1, "Tarjeta", S_HDR); s.text(2, 2, "Comprado a crédito", S_HDR)
    s.text(2, 3, "Pagado", S_HDR); s.text(2, 4, "Saldo (debes)", S_HDR)
    r = 3
    for t in _TARJETAS:
        A = "$A%d" % r
        comprado = 'SUMIFS(\'%s\'!$F:$F,\'%s\'!$H:$H,"Sí",\'%s\'!$G:$G,%s)' % (COMP, COMP, COMP, A)
        pagado = 'SUMIF(\'%s\'!$B:$B,%s,\'%s\'!$C:$C)' % (PAGOST, A, PAGOST)
        s.text(r, 1, t, S_LABEL)
        s.formula(r, 2, comprado, S_MONEY)
        s.formula(r, 3, pagado, S_MONEY)
        s.formula(r, 4, "MAX($B%d-$C%d,0)" % (r, r), S_TOTVAL)
        r += 1
    s.text(r, 1, "TOTAL QUE DEBO EN TARJETAS", S_TOTLAB)
    s.formula(r, 4, "SUM($D$3:$D$%d)" % (r - 1), S_TOTVAL)
    s.text(r + 2, 1, "Las compras a crédito (hoja Compras) suman; los abonos (hoja Pagos Tarjeta) restan.", S_FOOTER_SM)
    s.merge(r + 2, 1, r + 2, 4)
    return s


# ---------- FACTURA (se autollena desde una venta ya registrada) ----------
def build_factura():
    s = Sheet(FAC, hide_gridlines=True)
    widths = [16, 30, 12, 13, 16, 17]
    for i, w in enumerate(widths):
        s.colw(i + 1, w)
    VB = "%s!$B$2:$B$2000" % q(VEN)   # columna N° Factura en Ventas

    def fromsale(col):  # trae un dato de la venta elegida (col = letra en Ventas)
        return 'IFERROR(INDEX(%s!$%s$2:$%s$2000,MATCH($E$2,%s,0)),"")' % (q(VEN), col, col, VB)

    # Encabezado empresa
    s.formula(1, 2, "%s!$B$2" % q(CFG), S_COMPANY); s.merge(1, 2, 1, 3)
    s.formula(2, 2, '"NIT: "&%s!$B$3' % q(CFG), S_TEXT); s.merge(2, 2, 2, 3)
    s.formula(3, 2, "%s!$B$4" % q(CFG), S_TEXT); s.merge(3, 2, 3, 3)
    s.formula(4, 2, '"Tel: "&%s!$B$5&"   |   "&%s!$B$6' % (q(CFG), q(CFG)), S_TEXT); s.merge(4, 2, 4, 3)
    s.formula(5, 2, "%s!$B$7" % q(CFG), S_SECOND); s.merge(5, 2, 5, 3)
    s.text(1, 1, "LOGO", S_FOOTER_SM); s.merge(1, 1, 5, 1)
    # Metadatos: aquí ELIGES la venta (N° Factura) y todo lo demás se autollena.
    s.text(1, 4, "FACTURA DE VENTA", S_FACTITLE); s.merge(1, 4, 1, 6)
    s.text(2, 4, "Venta N°:", S_LABEL_RIGHT)
    s.text(2, 5, ("FAC-0001" if INCLUDE_SAMPLES else ""), xf(fontId=F_ALERT, fillId=FILL_INPUT, borderId=BORDER_THIN, halign="left", valign="center"))
    s.merge(2, 5, 2, 6)
    s.text(3, 4, "Fecha:", S_LABEL_RIGHT)
    s.formula(3, 5, fromsale("A"), S_DATE); s.merge(3, 5, 3, 6)
    s.text(4, 4, "Pago:", S_LABEL_RIGHT); s.formula(4, 5, fromsale("O"), S_TEXT); s.merge(4, 5, 4, 6)
    s.text(5, 4, "Origen inv.:", S_LABEL_RIGHT); s.formula(5, 5, fromsale("G"), S_TEXT); s.merge(5, 5, 5, 6)
    s.rowh(6, 8)
    # Datos cliente (auto desde la venta)
    s.text(7, 1, "DATOS DEL CLIENTE", S_BANNER_SEC); s.merge(7, 1, 7, 6)
    s.text(8, 1, "Cliente:", S_LABEL_RIGHT)
    s.formula(8, 2, fromsale("C"), S_TEXT); s.merge(8, 2, 8, 3)
    s.text(8, 4, "Documento:", S_LABEL_RIGHT)
    s.formula(8, 5, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,2,FALSE),"")' % q(CLI), S_TEXT); s.merge(8, 5, 8, 6)
    s.text(9, 1, "Teléfono:", S_LABEL_RIGHT)
    s.formula(9, 2, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,3,FALSE),"")' % q(CLI), S_TEXT); s.merge(9, 2, 9, 3)
    s.text(9, 4, "Correo:", S_LABEL_RIGHT)
    s.formula(9, 5, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,4,FALSE),"")' % q(CLI), S_TEXT); s.merge(9, 5, 9, 6)
    s.text(10, 1, "Dirección:", S_LABEL_RIGHT)
    s.formula(10, 2, 'IFERROR(VLOOKUP($B$8,%s!$B:$F,5,FALSE),"")' % q(CLI), S_TEXT); s.merge(10, 2, 10, 6)
    s.rowh(11, 8)
    # Tabla productos (se autollena con los renglones de esa venta)
    s.text(12, 1, "Código", S_HDR); s.text(12, 2, "Descripción", S_HDR); s.merge(12, 2, 12, 3)
    s.text(12, 4, "Cant.", S_HDR); s.text(12, 5, "Vr. Unitario", S_HDR); s.text(12, 6, "Total", S_HDR)
    s.hide_col(7, 8); s.hide_col(8, 8)            # G=IVA%, H=IVA$ (auxiliares)
    for c in range(10, 15):                       # J..N: auxiliar con las líneas de la venta
        s.hide_col(c, 10)
    # J13: trae las líneas de la venta elegida -> Código(J), Producto(K), Cantidad(L), Ubic(M), Valor(N)
    s.formula(13, 10, 'IFERROR(FILTER(%s!$D$2:$H$2000,%s=$E$2),"")' % (q(VEN), VB), S_TEXT)
    for r in range(13, 25):
        s.formula(r, 1, 'IF($J%d="","",$J%d)' % (r, r), S_TEXT_A)
        s.formula(r, 2, 'IF($A%d="","",IFERROR(VLOOKUP($A%d,%s!$A:$B,2,FALSE),""))'
                  % (r, r, q(CAT)), S_TEXT_A); s.merge(r, 2, r, 3)
        s.formula(r, 4, 'IF($J%d="","",$L%d)' % (r, r), S_INT)
        s.formula(r, 5, 'IF(OR($J%d="",$L%d=0),"",$N%d/$L%d)' % (r, r, r, r), S_MONEY)
        s.formula(r, 6, 'IF($A%d="","",$D%d*$E%d)' % (r, r, r), S_MONEY_A)
        s.formula(r, 7, 'IF($A%d="","",IFERROR(VLOOKUP($A%d,%s!$A:$L,12,FALSE),0))' % (r, r, q(CAT)), S_INT_A)
        s.formula(r, 8, 'IF($F%d="","",ROUND($F%d*$G%d/100,0))' % (r, r, r), S_MONEY_A)
    # Observaciones + totales
    s.text(25, 1, "Observaciones:", S_LABEL_RIGHT)
    s.blank(26, 1, S_TEXT_WRAP); s.merge(26, 1, 29, 3)
    s.text(26, 4, "Subtotal:", S_LABEL_RIGHT); s.merge(26, 4, 26, 5)
    s.formula(26, 6, "SUM($F$13:$F$24)", S_MONEY)
    s.text(27, 4, "IVA:", S_LABEL_RIGHT); s.merge(27, 4, 27, 5)
    s.formula(27, 6, "SUM($H$13:$H$24)", S_MONEY)
    s.text(28, 4, "TOTAL A PAGAR:", S_TOTLAB); s.merge(28, 4, 28, 5)
    s.formula(28, 6, "$F$26+$F$27", S_TOTVAL); s.rowh(28, 26)
    # Pie
    s.formula(31, 1, "%s!$B$10" % q(CFG), S_FOOTER); s.merge(31, 1, 31, 6)
    s.text(32, 1, "Documento generado electrónicamente · Sistema de Gestión", S_FOOTER_SM); s.merge(32, 1, 32, 6)
    # Validación: elegir el N° de una venta YA registrada en la hoja Ventas.
    s.validate_list(2, 5, 2, 5, "%s!$B$2:$B$%d" % (q(VEN), MOV_LAST))
    return s


# ----------------------------------------------------------------------
# Ensamblar el .xlsx
# ----------------------------------------------------------------------
def build_workbook(path):
    sheets = [
        build_dashboard(), build_catalogo(), build_compras(), build_ventas(),
        build_cierre_diario(), build_billeteras(), build_tarjetas(), build_movimientos(),
        build_pagos_tarjeta(), build_traslados(), build_ajustes(), build_prestamos(),
        build_abonos(), build_inventario(), build_clientes(), build_factura(),
        build_catalogo_venta(), build_config(),
    ]

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + ('<numFmts count="%d">' % len(numfmts.items)) + ''.join(numfmts.items) + '</numFmts>'
        + ('<fonts count="%d">' % len(fonts.items)) + ''.join(fonts.items) + '</fonts>'
        + ('<fills count="%d">' % len(fills.items)) + ''.join(fills.items) + '</fills>'
        + ('<borders count="%d">' % len(borders.items)) + ''.join(borders.items) + '</borders>'
        + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        + ('<cellXfs count="%d">' % len(xfs.items)) + ''.join(xfs.items) + '</cellXfs>'
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>'
    )

    # workbook.xml
    wb = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    for i, sh in enumerate(sheets):
        wb.append('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (escape(sh.name), i + 1, i + 1))
    wb.append('</sheets>')
    # Áreas de impresión: para que el PDF solo muestre la parte bonita (sin columnas auxiliares).
    names_idx = {sh.name: i for i, sh in enumerate(sheets)}
    defnames = []
    if FAC in names_idx:
        defnames.append('<definedName name="_xlnm.Print_Area" localSheetId="%d">\'%s\'!$A$1:$F$33</definedName>'
                        % (names_idx[FAC], FAC))
    if CVENTA in names_idx:
        defnames.append('<definedName name="_xlnm.Print_Area" localSheetId="%d">\'%s\'!$A$1:$E$120</definedName>'
                        % (names_idx[CVENTA], CVENTA))
    if defnames:
        wb.append('<definedNames>' + ''.join(defnames) + '</definedNames>')
    wb.append('<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    workbook_xml = ''.join(wb)

    # workbook rels
    rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for i in range(len(sheets)):
        rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (i + 1, i + 1))
    style_rid = len(sheets) + 1
    rels.append('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' % style_rid)
    rels.append('</Relationships>')
    workbook_rels = ''.join(rels)

    # content types
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for i in range(len(sheets)):
        ct.append('<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % (i + 1))
    ct.append('</Types>')
    content_types = ''.join(ct)

    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                 '</Relationships>')

    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', content_types)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', workbook_xml)
        z.writestr('xl/_rels/workbook.xml.rels', workbook_rels)
        z.writestr('xl/styles.xml', styles_xml)
        for i, sh in enumerate(sheets):
            z.writestr('xl/worksheets/sheet%d.xml' % (i + 1), sh.xml())
    return path


if __name__ == "__main__":
    INCLUDE_SAMPLES = True
    print("Generado:", build_workbook("Sistema_Gestion_Inventario.xlsx"))
    INCLUDE_SAMPLES = False
    print("Generado:", build_workbook("Sistema_Gestion_Inventario_VACIO.xlsx"))
